use std::fs;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering as AtomicOrdering};
use std::sync::{Arc, Mutex, OnceLock};

use arrow_array::types::Float32Type;
use arrow_array::{
    Array, FixedSizeListArray, Float32Array, Int64Array, RecordBatch, RecordBatchIterator,
    StringArray,
};
use arrow_schema::{DataType, Field, Schema};
use futures::TryStreamExt;
use lancedb::database::CreateTableMode;
use lancedb::index::Index as LanceIndex;
use lancedb::query::{ExecutableQuery, QueryBase, Select};
use lancedb::{connect as connect_lancedb, Table as LanceTable};
use ort::{session::Session as OrtSession, value::Tensor as OrtTensor};
use rusqlite::params;
use tauri::{AppHandle, Manager};
use tokenizers::Tokenizer;

use crate::db::{index_meta_dir, index_vector_dir, open_database};
use crate::types::{SearchHit, SemanticCandidate, SemanticIndexMeta, SemanticRuntime};
use crate::util::{file_name_from_relative, now_ms, path_display};
use crate::CommandResult;

pub(crate) const SEMANTIC_TABLE_NAME: &str = "semantic_hits_v2";
pub(crate) const SEMANTIC_META_FILE_NAME: &str = "semantic-index-meta-v2.json";
pub(crate) const SEMANTIC_MAX_DOCUMENTS: usize = 2_000_000;
pub(crate) const SEMANTIC_EMBED_BATCH: usize = 24;
pub(crate) const SEMANTIC_MAX_TOKENS: usize = 192;
pub(crate) const SEMANTIC_MIN_QUERY_CHARS: usize = 3;

static SEMANTIC_RUNTIME: OnceLock<Mutex<SemanticRuntime>> = OnceLock::new();
static SEMANTIC_REBUILD_IN_FLIGHT: AtomicBool = AtomicBool::new(false);

pub(crate) fn semantic_db_dir(app: &AppHandle) -> CommandResult<PathBuf> {
    index_vector_dir(app)
}

fn semantic_meta_path(app: &AppHandle) -> CommandResult<PathBuf> {
    Ok(index_meta_dir(app)?.join(SEMANTIC_META_FILE_NAME))
}

fn resolve_semantic_resource_path(app: &AppHandle, file_name: &str) -> CommandResult<PathBuf> {
    let mut candidates = Vec::new();
    if let Ok(resource_dir) = app.path().resource_dir() {
        candidates.push(resource_dir.join(file_name));
        candidates.push(resource_dir.join("resources").join(file_name));
    }
    let manifest_resources = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("resources");
    candidates.push(manifest_resources.join(file_name));

    for path in candidates {
        if path.exists() {
            return Ok(path);
        }
    }

    Err(format!(
        "Missing semantic resource '{file_name}'. Expected it under the app resource directory or '{}'",
        path_display(&manifest_resources)
    ))
}

fn build_semantic_runtime(app: &AppHandle) -> CommandResult<SemanticRuntime> {
    let model_path = resolve_semantic_resource_path(app, "model.onnx")?;
    let tokenizer_path = resolve_semantic_resource_path(app, "tokenizer.json")?;
    let tokenizer = Tokenizer::from_file(&tokenizer_path).map_err(|error| {
        format!(
            "Could not load tokenizer '{}': {error}",
            path_display(&tokenizer_path)
        )
    })?;
    let mut builder = OrtSession::builder().map_err(|error| {
        format!(
            "Could not create ONNX session builder for '{}': {error}",
            path_display(&model_path)
        )
    })?;
    if let Ok(parallelism) = std::thread::available_parallelism() {
        let threads = parallelism.get().clamp(1, 8);
        builder = builder
            .with_intra_threads(threads)
            .map_err(|error| format!("Could not set ONNX thread count: {error}"))?;
    }
    let session = builder.commit_from_file(&model_path).map_err(|error| {
        format!(
            "Could not load ONNX model '{}': {error}",
            path_display(&model_path)
        )
    })?;
    let output_name = session
        .outputs()
        .first()
        .map(|entry| entry.name().to_string())
        .ok_or_else(|| "ONNX model has no outputs".to_string())?;

    Ok(SemanticRuntime {
        tokenizer,
        session,
        output_name,
    })
}

fn load_semantic_runtime(app: &AppHandle) -> CommandResult<&'static Mutex<SemanticRuntime>> {
    if let Some(runtime) = SEMANTIC_RUNTIME.get() {
        return Ok(runtime);
    }

    let runtime = build_semantic_runtime(app)?;
    let _ = SEMANTIC_RUNTIME.set(Mutex::new(runtime));
    SEMANTIC_RUNTIME
        .get()
        .ok_or_else(|| "Could not initialize semantic runtime".to_string())
}

fn semantic_root_fingerprint_ms(connection: &rusqlite::Connection) -> CommandResult<i64> {
    connection
        .query_row(
            "SELECT COALESCE(MAX(last_indexed_ms), 0) FROM roots",
            [],
            |row| row.get::<_, i64>(0),
        )
        .map_err(|error| format!("Could not read semantic index fingerprint: {error}"))
}

fn read_semantic_meta(app: &AppHandle) -> CommandResult<SemanticIndexMeta> {
    let path = semantic_meta_path(app)?;
    if !path.exists() {
        return Ok(SemanticIndexMeta::default());
    }
    let raw = fs::read_to_string(&path).map_err(|error| {
        format!(
            "Could not read semantic metadata '{}': {error}",
            path_display(&path)
        )
    })?;
    serde_json::from_str::<SemanticIndexMeta>(&raw).map_err(|error| {
        format!(
            "Could not parse semantic metadata '{}': {error}",
            path_display(&path)
        )
    })
}

fn write_semantic_meta(app: &AppHandle, meta: &SemanticIndexMeta) -> CommandResult<()> {
    let path = semantic_meta_path(app)?;
    let raw = serde_json::to_vec_pretty(meta)
        .map_err(|error| format!("Could not serialize semantic metadata: {error}"))?;
    fs::write(&path, raw).map_err(|error| {
        format!(
            "Could not write semantic metadata '{}': {error}",
            path_display(&path)
        )
    })
}

fn semantic_index_is_stale(app: &AppHandle) -> CommandResult<bool> {
    let connection = open_database(app)?;
    let fingerprint = semantic_root_fingerprint_ms(&connection)?;
    if fingerprint == 0 {
        return Ok(false);
    }
    let meta = read_semantic_meta(app).unwrap_or_default();
    Ok(meta.root_fingerprint_ms < fingerprint)
}

fn semantic_embedding_text(text: &str) -> String {
    let mut value = text.trim().to_string();
    if value.chars().count() > 720 {
        value = value.chars().take(720).collect();
    }
    value
}

fn load_semantic_candidates(
    connection: &rusqlite::Connection,
    max_documents: usize,
) -> CommandResult<Vec<SemanticCandidate>> {
    if max_documents == 0 {
        return Ok(Vec::new());
    }
    let mut candidates = Vec::new();
    let mut semantic_id = 1_i64;
    let max_documents_i64 = i64::try_from(max_documents).unwrap_or(i64::MAX);

    {
        let mut statement = connection
            .prepare(
                "
                SELECT
                  root_id,
                  file_id,
                  file_name,
                  relative_path,
                  absolute_path,
                  heading_level,
                  heading_text,
                  heading_order,
                  author_text,
                  chunk_text
                FROM chunks
                ORDER BY root_id ASC, file_id ASC, chunk_order ASC
                LIMIT ?1
                ",
            )
            .map_err(|error| {
                format!("Could not prepare semantic chunk candidates query: {error}")
            })?;

        let rows = statement
            .query_map(params![max_documents_i64], |row| {
                Ok((
                    row.get::<_, i64>(0)?,
                    row.get::<_, i64>(1)?,
                    row.get::<_, String>(2)?,
                    row.get::<_, String>(3)?,
                    row.get::<_, String>(4)?,
                    row.get::<_, Option<i64>>(5)?,
                    row.get::<_, Option<String>>(6)?,
                    row.get::<_, Option<i64>>(7)?,
                    row.get::<_, Option<String>>(8)?,
                    row.get::<_, String>(9)?,
                ))
            })
            .map_err(|error| format!("Could not run semantic chunk candidates query: {error}"))?;

        for row in rows {
            if candidates.len() >= max_documents {
                break;
            }
            let (
                root_id,
                file_id,
                file_name,
                relative_path,
                absolute_path,
                heading_level,
                heading_text,
                heading_order,
                author_text,
                chunk_text,
            ) =
                row.map_err(|error| format!("Could not parse semantic chunk candidate: {error}"))?;

            let trimmed_chunk = chunk_text.trim();
            if trimmed_chunk.is_empty() {
                continue;
            }

            let semantic_text = semantic_embedding_text(&format!(
                "heading: {}\nauthor: {}\nchunk: {}\npath: {}\nfile: {}",
                heading_text.clone().unwrap_or_default(),
                author_text.clone().unwrap_or_default(),
                trimmed_chunk,
                relative_path,
                file_name
            ));
            let kind = if author_text.is_some() {
                "author".to_string()
            } else if heading_text.is_some() {
                "heading".to_string()
            } else {
                "file".to_string()
            };
            candidates.push(SemanticCandidate {
                semantic_id,
                root_id,
                kind,
                file_id,
                file_name,
                relative_path,
                absolute_path,
                heading_level,
                heading_text,
                heading_order,
                semantic_text,
            });
            semantic_id += 1;
        }
    }

    if !candidates.is_empty() {
        return Ok(candidates);
    }

    // Fallback for roots indexed before chunk rows were written.
    let mut statement = connection
        .prepare(
            "
            SELECT root_id, id, relative_path, absolute_path
            FROM files
            ORDER BY modified_ms DESC, id DESC
            LIMIT ?1
            ",
        )
        .map_err(|error| format!("Could not prepare semantic fallback file query: {error}"))?;
    let rows = statement
        .query_map(params![max_documents_i64], |row| {
            Ok((
                row.get::<_, i64>(0)?,
                row.get::<_, i64>(1)?,
                row.get::<_, String>(2)?,
                row.get::<_, String>(3)?,
            ))
        })
        .map_err(|error| format!("Could not run semantic fallback file query: {error}"))?;
    for row in rows {
        if candidates.len() >= max_documents {
            break;
        }
        let (root_id, file_id, relative_path, absolute_path) =
            row.map_err(|error| format!("Could not parse semantic fallback file row: {error}"))?;
        let file_name = file_name_from_relative(&relative_path);
        let semantic_text =
            semantic_embedding_text(&format!("file: {}\npath: {}", file_name, relative_path));
        candidates.push(SemanticCandidate {
            semantic_id,
            root_id,
            kind: "file".to_string(),
            file_id,
            file_name,
            relative_path,
            absolute_path,
            heading_level: None,
            heading_text: None,
            heading_order: None,
            semantic_text,
        });
        semantic_id += 1;
    }

    Ok(candidates)
}

fn normalize_vector_l2(values: &mut [f32]) {
    let norm = values
        .iter()
        .fold(0.0_f32, |acc, value| acc + (value * value))
        .sqrt();
    if norm <= 0.0 {
        return;
    }
    for value in values {
        *value /= norm;
    }
}

fn encode_semantic_batch(
    tokenizer: &Tokenizer,
    texts: &[String],
    max_tokens: usize,
) -> CommandResult<(Vec<i64>, Vec<i64>, Vec<i64>, usize, usize)> {
    let batch_size = texts.len();
    if batch_size == 0 {
        return Ok((Vec::new(), Vec::new(), Vec::new(), 0, 0));
    }
    let seq_len = max_tokens.max(8);
    let mut input_ids = Vec::with_capacity(batch_size.saturating_mul(seq_len));
    let mut attention_mask = Vec::with_capacity(batch_size.saturating_mul(seq_len));
    let mut token_type_ids = Vec::with_capacity(batch_size.saturating_mul(seq_len));

    for text in texts {
        let encoding = tokenizer
            .encode(text.as_str(), true)
            .map_err(|error| format!("Could not tokenize semantic input: {error}"))?;
        let mut ids = encoding
            .get_ids()
            .iter()
            .take(seq_len)
            .map(|value| i64::from(*value))
            .collect::<Vec<i64>>();
        let mut mask = encoding
            .get_attention_mask()
            .iter()
            .take(seq_len)
            .map(|value| i64::from(*value))
            .collect::<Vec<i64>>();
        let mut type_ids = encoding
            .get_type_ids()
            .iter()
            .take(seq_len)
            .map(|value| i64::from(*value))
            .collect::<Vec<i64>>();

        if ids.is_empty() {
            ids.push(101);
            mask.push(1);
            type_ids.push(0);
        }
        while mask.len() < ids.len() {
            mask.push(1);
        }
        while type_ids.len() < ids.len() {
            type_ids.push(0);
        }

        ids.resize(seq_len, 0);
        mask.resize(seq_len, 0);
        type_ids.resize(seq_len, 0);

        input_ids.extend(ids);
        attention_mask.extend(mask);
        token_type_ids.extend(type_ids);
    }

    Ok((
        input_ids,
        attention_mask,
        token_type_ids,
        batch_size,
        seq_len,
    ))
}

pub(crate) fn embed_semantic_texts(
    app: &AppHandle,
    texts: &[String],
) -> CommandResult<Vec<Vec<f32>>> {
    if texts.is_empty() {
        return Ok(Vec::new());
    }
    let runtime = load_semantic_runtime(app)?;
    let mut runtime = runtime
        .lock()
        .map_err(|_| "Could not lock semantic runtime".to_string())?;
    let output_name = runtime.output_name.clone();
    let expects_token_type_ids = runtime
        .session
        .inputs()
        .iter()
        .any(|entry| entry.name() == "token_type_ids");

    let (input_ids, attention_mask, token_type_ids, batch_size, seq_len) =
        encode_semantic_batch(&runtime.tokenizer, texts, SEMANTIC_MAX_TOKENS)?;
    if batch_size == 0 || seq_len == 0 {
        return Ok(Vec::new());
    }

    let shape = vec![
        i64::try_from(batch_size).unwrap_or(0),
        i64::try_from(seq_len).unwrap_or(0),
    ];
    let primary_input_ids = OrtTensor::from_array((shape.clone(), input_ids.clone()))
        .map_err(|error| format!("Could not create semantic input_ids tensor: {error}"))?;
    let primary_attention_mask = OrtTensor::from_array((shape.clone(), attention_mask.clone()))
        .map_err(|error| format!("Could not create semantic attention_mask tensor: {error}"))?;
    let outputs = if expects_token_type_ids {
        let primary_token_type_ids = OrtTensor::from_array((shape.clone(), token_type_ids.clone()))
            .map_err(|error| format!("Could not create semantic token_type_ids tensor: {error}"))?;
        runtime.session.run(ort::inputs! {
            "input_ids" => primary_input_ids,
            "attention_mask" => primary_attention_mask,
            "token_type_ids" => primary_token_type_ids
        })
    } else {
        runtime.session.run(ort::inputs! {
            "input_ids" => primary_input_ids,
            "attention_mask" => primary_attention_mask
        })
    }
    .map_err(|error| format!("Semantic model inference failed: {error}"))?;

    let output = if outputs.contains_key(output_name.as_str()) {
        &outputs[output_name.as_str()]
    } else {
        &outputs[0]
    };
    let output = output
        .try_extract_array::<f32>()
        .map_err(|error| format!("Could not extract semantic output tensor: {error}"))?;

    if output.ndim() != 3 {
        return Err(format!(
            "Semantic model output rank {} is unsupported (expected 3)",
            output.ndim()
        ));
    }
    let output_shape = output.shape();
    let output_batch = output_shape[0];
    let output_seq = output_shape[1];
    let embedding_dim = output_shape[2];
    if output_batch != batch_size || embedding_dim == 0 {
        return Err("Semantic model output shape does not match request".to_string());
    }

    let mut vectors = Vec::with_capacity(batch_size);
    for batch_index in 0..batch_size {
        let mut pooled = vec![0.0_f32; embedding_dim];
        let mut token_count = 0.0_f32;
        let max_steps = output_seq.min(seq_len);

        for token_index in 0..max_steps {
            if attention_mask[batch_index * seq_len + token_index] == 0 {
                continue;
            }
            token_count += 1.0;
            for dim in 0..embedding_dim {
                pooled[dim] += output[[batch_index, token_index, dim]];
            }
        }

        if token_count <= 0.0 {
            for dim in 0..embedding_dim {
                pooled[dim] = output[[batch_index, 0, dim]];
            }
        } else {
            for value in &mut pooled {
                *value /= token_count;
            }
        }
        normalize_vector_l2(&mut pooled);
        vectors.push(pooled);
    }

    Ok(vectors)
}

fn semantic_schema(embedding_dim: usize) -> Arc<Schema> {
    Arc::new(Schema::new(vec![
        Field::new("semantic_id", DataType::Int64, false),
        Field::new("root_id", DataType::Int64, false),
        Field::new("kind", DataType::Utf8, false),
        Field::new("file_id", DataType::Int64, false),
        Field::new("file_name", DataType::Utf8, false),
        Field::new("relative_path", DataType::Utf8, false),
        Field::new("absolute_path", DataType::Utf8, false),
        Field::new("heading_level", DataType::Int64, true),
        Field::new("heading_text", DataType::Utf8, true),
        Field::new("heading_order", DataType::Int64, true),
        Field::new(
            "vector",
            DataType::FixedSizeList(
                Arc::new(Field::new("item", DataType::Float32, true)),
                i32::try_from(embedding_dim).unwrap_or(0),
            ),
            false,
        ),
    ]))
}

fn semantic_record_batch(
    schema: Arc<Schema>,
    candidates: &[SemanticCandidate],
    embeddings: &[Vec<f32>],
    embedding_dim: usize,
) -> CommandResult<RecordBatch> {
    if candidates.len() != embeddings.len() {
        return Err("Semantic candidate/embedding batch size mismatch".to_string());
    }
    let semantic_ids =
        Int64Array::from_iter_values(candidates.iter().map(|candidate| candidate.semantic_id));
    let root_ids =
        Int64Array::from_iter_values(candidates.iter().map(|candidate| candidate.root_id));
    let kinds =
        StringArray::from_iter_values(candidates.iter().map(|candidate| candidate.kind.as_str()));
    let file_ids =
        Int64Array::from_iter_values(candidates.iter().map(|candidate| candidate.file_id));
    let file_names = StringArray::from_iter_values(
        candidates
            .iter()
            .map(|candidate| candidate.file_name.as_str()),
    );
    let relative_paths = StringArray::from_iter_values(
        candidates
            .iter()
            .map(|candidate| candidate.relative_path.as_str()),
    );
    let absolute_paths = StringArray::from_iter_values(
        candidates
            .iter()
            .map(|candidate| candidate.absolute_path.as_str()),
    );
    let heading_levels = Int64Array::from(
        candidates
            .iter()
            .map(|candidate| candidate.heading_level)
            .collect::<Vec<_>>(),
    );
    let heading_texts = StringArray::from(
        candidates
            .iter()
            .map(|candidate| candidate.heading_text.clone())
            .collect::<Vec<Option<String>>>(),
    );
    let heading_orders = Int64Array::from(
        candidates
            .iter()
            .map(|candidate| candidate.heading_order)
            .collect::<Vec<_>>(),
    );

    let vectors = FixedSizeListArray::from_iter_primitive::<Float32Type, _, _>(
        embeddings.iter().map(|embedding| {
            let mut row = embedding
                .iter()
                .take(embedding_dim)
                .map(|value| Some(*value))
                .collect::<Vec<Option<f32>>>();
            row.resize(embedding_dim, Some(0.0));
            Some(row)
        }),
        i32::try_from(embedding_dim).unwrap_or(0),
    );

    RecordBatch::try_new(
        schema,
        vec![
            Arc::new(semantic_ids),
            Arc::new(root_ids),
            Arc::new(kinds),
            Arc::new(file_ids),
            Arc::new(file_names),
            Arc::new(relative_paths),
            Arc::new(absolute_paths),
            Arc::new(heading_levels),
            Arc::new(heading_texts),
            Arc::new(heading_orders),
            Arc::new(vectors),
        ],
    )
    .map_err(|error| format!("Could not build semantic record batch: {error}"))
}

async fn rebuild_semantic_index(app: AppHandle, force: bool) -> CommandResult<()> {
    let connection = open_database(&app)?;
    let root_fingerprint_ms = semantic_root_fingerprint_ms(&connection)?;
    if root_fingerprint_ms <= 0 {
        return Ok(());
    }

    let previous_meta = read_semantic_meta(&app).unwrap_or_default();
    if !force && previous_meta.root_fingerprint_ms >= root_fingerprint_ms {
        return Ok(());
    }

    let candidates = load_semantic_candidates(&connection, SEMANTIC_MAX_DOCUMENTS)?;
    if candidates.is_empty() {
        let meta = SemanticIndexMeta {
            root_fingerprint_ms,
            item_count: 0,
            embedding_dim: 0,
            updated_at_ms: now_ms(),
        };
        write_semantic_meta(&app, &meta)?;
        return Ok(());
    }

    let mut schema: Option<Arc<Schema>> = None;
    let mut batches = Vec::new();
    let mut embedding_dim = 0_usize;

    for chunk in candidates.chunks(SEMANTIC_EMBED_BATCH) {
        let texts = chunk
            .iter()
            .map(|candidate| candidate.semantic_text.clone())
            .collect::<Vec<String>>();
        let app_for_embedding = app.clone();
        let embeddings = tauri::async_runtime::spawn_blocking(move || {
            embed_semantic_texts(&app_for_embedding, &texts)
        })
        .await
        .map_err(|error| format!("Semantic embedding task failed: {error}"))??;
        if embeddings.is_empty() {
            continue;
        }
        let current_dim = embeddings[0].len();
        if current_dim == 0 {
            continue;
        }
        if embedding_dim == 0 {
            embedding_dim = current_dim;
            schema = Some(semantic_schema(embedding_dim));
        }
        if current_dim != embedding_dim {
            continue;
        }
        let batch = semantic_record_batch(
            schema
                .clone()
                .ok_or_else(|| "Semantic schema was not initialized".to_string())?,
            chunk,
            &embeddings,
            embedding_dim,
        )?;
        batches.push(batch);
    }

    if batches.is_empty() || embedding_dim == 0 {
        return Ok(());
    }

    let semantic_dir = semantic_db_dir(&app)?;
    fs::create_dir_all(&semantic_dir).map_err(|error| {
        format!(
            "Could not create semantic DB directory '{}': {error}",
            path_display(&semantic_dir)
        )
    })?;
    let uri = path_display(&semantic_dir);
    let db = connect_lancedb(&uri)
        .execute()
        .await
        .map_err(|error| format!("Could not open LanceDB at '{}': {error}", uri))?;

    let schema = schema.ok_or_else(|| "Semantic schema was not created".to_string())?;
    let reader = RecordBatchIterator::new(batches.into_iter().map(Ok), schema.clone());
    let table = db
        .create_table(SEMANTIC_TABLE_NAME, Box::new(reader))
        .mode(CreateTableMode::Overwrite)
        .execute()
        .await
        .map_err(|error| format!("Could not write semantic LanceDB table: {error}"))?;

    if candidates.len() >= 4_096 {
        table
            .create_index(&["vector"], LanceIndex::Auto)
            .execute()
            .await
            .map_err(|error| format!("Could not create semantic vector index: {error}"))?;
    }

    let meta = SemanticIndexMeta {
        root_fingerprint_ms,
        item_count: candidates.len(),
        embedding_dim,
        updated_at_ms: now_ms(),
    };
    write_semantic_meta(&app, &meta)?;
    Ok(())
}

pub(crate) fn trigger_semantic_rebuild(app: AppHandle, force: bool) {
    let should_rebuild = force || semantic_index_is_stale(&app).unwrap_or(false);
    if !should_rebuild {
        return;
    }
    if SEMANTIC_REBUILD_IN_FLIGHT
        .compare_exchange(false, true, AtomicOrdering::SeqCst, AtomicOrdering::SeqCst)
        .is_err()
    {
        return;
    }
    tauri::async_runtime::spawn(async move {
        if let Err(error) = rebuild_semantic_index(app.clone(), force).await {
            eprintln!("Semantic index rebuild failed: {error}");
        }
        SEMANTIC_REBUILD_IN_FLIGHT.store(false, AtomicOrdering::SeqCst);
    });
}

pub(crate) fn semantic_hits_from_batches(
    batches: &[RecordBatch],
    limit: usize,
) -> CommandResult<Vec<SearchHit>> {
    let mut hits = Vec::new();
    let mut seen = std::collections::HashSet::new();
    for batch in batches {
        let file_id_col = batch
            .column_by_name("file_id")
            .and_then(|column| column.as_any().downcast_ref::<Int64Array>())
            .ok_or_else(|| "Semantic result batch missing file_id column".to_string())?;
        let kind_col = batch
            .column_by_name("kind")
            .and_then(|column| column.as_any().downcast_ref::<StringArray>())
            .ok_or_else(|| "Semantic result batch missing kind column".to_string())?;
        let file_name_col = batch
            .column_by_name("file_name")
            .and_then(|column| column.as_any().downcast_ref::<StringArray>())
            .ok_or_else(|| "Semantic result batch missing file_name column".to_string())?;
        let relative_path_col = batch
            .column_by_name("relative_path")
            .and_then(|column| column.as_any().downcast_ref::<StringArray>())
            .ok_or_else(|| "Semantic result batch missing relative_path column".to_string())?;
        let absolute_path_col = batch
            .column_by_name("absolute_path")
            .and_then(|column| column.as_any().downcast_ref::<StringArray>())
            .ok_or_else(|| "Semantic result batch missing absolute_path column".to_string())?;
        let heading_level_col = batch
            .column_by_name("heading_level")
            .and_then(|column| column.as_any().downcast_ref::<Int64Array>());
        let heading_text_col = batch
            .column_by_name("heading_text")
            .and_then(|column| column.as_any().downcast_ref::<StringArray>());
        let heading_order_col = batch
            .column_by_name("heading_order")
            .and_then(|column| column.as_any().downcast_ref::<Int64Array>());
        let distance_f32 = batch
            .column_by_name("_distance")
            .and_then(|column| column.as_any().downcast_ref::<Float32Array>());

        for row_index in 0..batch.num_rows() {
            if hits.len() >= limit {
                return Ok(hits);
            }
            let file_id = file_id_col.value(row_index);
            let kind = kind_col.value(row_index).to_string();
            let heading_level = heading_level_col
                .and_then(|column| (!column.is_null(row_index)).then_some(column.value(row_index)));
            let heading_text = heading_text_col.and_then(|column| {
                (!column.is_null(row_index)).then_some(column.value(row_index).to_string())
            });
            let heading_order = heading_order_col
                .and_then(|column| (!column.is_null(row_index)).then_some(column.value(row_index)));
            let dedupe_key = format!(
                "{}:{}:{}:{}",
                file_id,
                kind,
                heading_order.unwrap_or(0),
                heading_text.clone().unwrap_or_default()
            );
            if !seen.insert(dedupe_key) {
                continue;
            }

            let distance = distance_f32
                .and_then(|column| {
                    (!column.is_null(row_index)).then_some(f64::from(column.value(row_index)))
                })
                .unwrap_or(1.0);
            hits.push(SearchHit {
                source: "semantic".to_string(),
                kind,
                file_id,
                file_name: file_name_col.value(row_index).to_string(),
                relative_path: relative_path_col.value(row_index).to_string(),
                absolute_path: absolute_path_col.value(row_index).to_string(),
                heading_level,
                heading_text,
                heading_order,
                score: 7000.0 + (distance * 1000.0),
            });
        }
    }
    Ok(hits)
}

pub(crate) async fn semantic_search(
    app: &AppHandle,
    query: &str,
    requested_root_id: Option<i64>,
    limit: usize,
) -> CommandResult<Vec<SearchHit>> {
    let semantic_dir = semantic_db_dir(app)?;
    if !semantic_dir.exists() {
        return Ok(Vec::new());
    }
    let uri = path_display(&semantic_dir);
    let db = connect_lancedb(&uri)
        .execute()
        .await
        .map_err(|error| format!("Could not open semantic LanceDB: {error}"))?;
    let table: LanceTable = match db.open_table(SEMANTIC_TABLE_NAME).execute().await {
        Ok(table) => table,
        Err(_) => return Ok(Vec::new()),
    };

    let app_for_embedding = app.clone();
    let query_text = query.to_string();
    let query_embedding = tauri::async_runtime::spawn_blocking(move || {
        embed_semantic_texts(&app_for_embedding, &[query_text])
    })
    .await
    .map_err(|error| format!("Semantic query embedding task failed: {error}"))??;
    if query_embedding.is_empty() || query_embedding[0].is_empty() {
        return Ok(Vec::new());
    }

    let mut vector_query = table
        .query()
        .nearest_to(query_embedding[0].as_slice())
        .map_err(|error| format!("Could not build semantic vector query: {error}"))?
        .limit(limit.saturating_mul(2))
        .select(Select::columns(&[
            "file_id",
            "kind",
            "file_name",
            "relative_path",
            "absolute_path",
            "heading_level",
            "heading_text",
            "heading_order",
        ]))
        .nprobes(18)
        .refine_factor(2);

    if let Some(root_id) = requested_root_id {
        vector_query = vector_query.only_if(format!("root_id = {root_id}"));
    }

    let batches = vector_query
        .execute()
        .await
        .map_err(|error| format!("Semantic search execution failed: {error}"))?
        .try_collect::<Vec<RecordBatch>>()
        .await
        .map_err(|error| format!("Semantic search result stream failed: {error}"))?;

    semantic_hits_from_batches(&batches, limit)
}
