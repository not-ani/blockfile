use std::collections::HashSet;
use std::fs;
use std::sync::{Mutex, OnceLock};
use std::time::Instant;

use tantivy::collector::TopDocs;
use tantivy::query::QueryParser;
use tantivy::schema::{
    Field, IndexRecordOption, NumericOptions, Schema, TextFieldIndexing, TextOptions, Value,
    STORED, STRING, TEXT,
};
use tantivy::tokenizer::{LowerCaser, NgramTokenizer, TextAnalyzer};
use tantivy::{doc, Index, IndexReader, ReloadPolicy, TantivyDocument};
use tauri::AppHandle;

use crate::db::index_lexical_dir;
use crate::search::normalize_for_search;
use crate::types::SearchHit;
use crate::CommandResult;

const PREFIX_TOKENIZER: &str = "bf_prefix";
const NGRAM_TOKENIZER: &str = "bf_ngram";
const MIN_FETCH_MULTIPLIER: usize = 5;
const MIN_FETCH_FLOOR: usize = 80;
const MAX_FETCH_LIMIT: usize = 1_800;

#[derive(Clone)]
pub(crate) struct LexicalDocument {
    pub root_id: i64,
    pub file_id: i64,
    pub kind: String,
    pub file_name: String,
    pub relative_path: String,
    pub absolute_path: String,
    pub heading_level: Option<i64>,
    pub heading_text: Option<String>,
    pub heading_order: Option<i64>,
    pub author_text: Option<String>,
    pub chunk_text: Option<String>,
}

#[derive(Clone)]
struct LexicalFields {
    kind: Field,
    root_id: Field,
    file_id: Field,
    file_name: Field,
    relative_path: Field,
    absolute_path: Field,
    heading_level: Field,
    heading_text: Field,
    heading_order: Field,
    author_text: Field,
    chunk_text: Field,
    query_text: Field,
    prefix_text: Field,
    ngram_text: Field,
}

struct LexicalRuntime {
    index: Index,
    reader: IndexReader,
    fields: LexicalFields,
}

static LEXICAL_RUNTIME: OnceLock<Mutex<LexicalRuntime>> = OnceLock::new();

fn stored_text_options(tokenizer: &str) -> TextOptions {
    TextOptions::default().set_stored().set_indexing_options(
        TextFieldIndexing::default()
            .set_tokenizer(tokenizer)
            .set_index_option(IndexRecordOption::WithFreqsAndPositions),
    )
}

fn build_schema() -> Schema {
    let mut builder = Schema::builder();

    let numeric = NumericOptions::default()
        .set_fast()
        .set_stored()
        .set_indexed();
    builder.add_text_field("kind", STRING | STORED);
    builder.add_u64_field("root_id", numeric.clone());
    builder.add_u64_field("file_id", numeric.clone());
    builder.add_text_field("file_name", TEXT | STORED);
    builder.add_text_field("relative_path", TEXT | STORED);
    builder.add_text_field("absolute_path", STRING | STORED);
    builder.add_i64_field("heading_level", numeric.clone());
    builder.add_text_field("heading_text", TEXT | STORED);
    builder.add_i64_field("heading_order", numeric);
    builder.add_text_field("author_text", TEXT | STORED);
    builder.add_text_field("chunk_text", TEXT | STORED);
    builder.add_text_field("query_text", stored_text_options("default"));
    builder.add_text_field("prefix_text", stored_text_options(PREFIX_TOKENIZER));
    builder.add_text_field("ngram_text", stored_text_options(NGRAM_TOKENIZER));

    builder.build()
}

fn has_required_fields(schema: &Schema) -> bool {
    schema.get_field("query_text").is_ok()
        && schema.get_field("prefix_text").is_ok()
        && schema.get_field("ngram_text").is_ok()
}

fn register_tokenizers(index: &Index) -> CommandResult<()> {
    let prefix_tokenizer = NgramTokenizer::new(2, 18, true)
        .map_err(|error| format!("Could not build lexical prefix tokenizer: {error}"))?;
    let ngram_tokenizer = NgramTokenizer::new(3, 4, false)
        .map_err(|error| format!("Could not build lexical ngram tokenizer: {error}"))?;

    index.tokenizers().register(
        PREFIX_TOKENIZER,
        TextAnalyzer::builder(prefix_tokenizer)
            .filter(LowerCaser)
            .build(),
    );
    index.tokenizers().register(
        NGRAM_TOKENIZER,
        TextAnalyzer::builder(ngram_tokenizer)
            .filter(LowerCaser)
            .build(),
    );
    Ok(())
}

fn field(schema: &Schema, name: &str) -> CommandResult<Field> {
    schema
        .get_field(name)
        .map_err(|error| format!("Missing lexical schema field '{name}': {error}"))
}

fn lexical_fields(schema: &Schema) -> CommandResult<LexicalFields> {
    Ok(LexicalFields {
        kind: field(schema, "kind")?,
        root_id: field(schema, "root_id")?,
        file_id: field(schema, "file_id")?,
        file_name: field(schema, "file_name")?,
        relative_path: field(schema, "relative_path")?,
        absolute_path: field(schema, "absolute_path")?,
        heading_level: field(schema, "heading_level")?,
        heading_text: field(schema, "heading_text")?,
        heading_order: field(schema, "heading_order")?,
        author_text: field(schema, "author_text")?,
        chunk_text: field(schema, "chunk_text")?,
        query_text: field(schema, "query_text")?,
        prefix_text: field(schema, "prefix_text")?,
        ngram_text: field(schema, "ngram_text")?,
    })
}

fn init_runtime(app: &AppHandle) -> CommandResult<LexicalRuntime> {
    let schema = build_schema();
    let path = index_lexical_dir(app)?;
    fs::create_dir_all(&path).map_err(|error| {
        format!(
            "Could not create lexical index directory '{}': {error}",
            path.display()
        )
    })?;

    let recreate = match Index::open_in_dir(&path) {
        Ok(index) => !has_required_fields(&index.schema()),
        Err(_) => true,
    };

    let index = if recreate {
        let _ = fs::remove_dir_all(&path);
        fs::create_dir_all(&path).map_err(|error| {
            format!(
                "Could not reset lexical index directory '{}': {error}",
                path.display()
            )
        })?;
        Index::create_in_dir(&path, schema.clone())
            .map_err(|error| format!("Could not recreate lexical index: {error}"))?
    } else {
        Index::open_in_dir(&path)
            .map_err(|error| format!("Could not open lexical index: {error}"))?
    };

    register_tokenizers(&index)?;
    let fields = lexical_fields(&index.schema())?;
    let reader = index
        .reader_builder()
        .reload_policy(ReloadPolicy::Manual)
        .try_into()
        .map_err(|error| format!("Could not build lexical index reader: {error}"))?;

    Ok(LexicalRuntime {
        index,
        reader,
        fields,
    })
}

fn lexical_runtime(app: &AppHandle) -> CommandResult<&'static Mutex<LexicalRuntime>> {
    if let Some(runtime) = LEXICAL_RUNTIME.get() {
        return Ok(runtime);
    }
    let runtime = init_runtime(app)?;
    let _ = LEXICAL_RUNTIME.set(Mutex::new(runtime));
    LEXICAL_RUNTIME
        .get()
        .ok_or_else(|| "Could not initialize lexical runtime".to_string())
}

fn field_text(document: &TantivyDocument, field: Field) -> Option<String> {
    document
        .get_first(field)
        .and_then(|value| value.as_str())
        .map(|value| value.to_string())
}

fn field_i64(document: &TantivyDocument, field: Field) -> Option<i64> {
    document.get_first(field).and_then(|value| value.as_i64())
}

fn field_u64(document: &TantivyDocument, field: Field) -> Option<u64> {
    document.get_first(field).and_then(|value| value.as_u64())
}

fn ngrams_for_query(normalized_query: &str) -> String {
    let compact = normalized_query.replace(' ', "");
    let chars = compact.chars().collect::<Vec<char>>();
    if chars.len() <= 4 {
        return normalized_query.to_string();
    }

    let mut ngrams = Vec::new();
    for start in 0..chars.len().saturating_sub(2) {
        let end = (start + 4).min(chars.len());
        let gram = chars[start..end].iter().collect::<String>();
        if gram.len() >= 3 {
            ngrams.push(gram);
        }
    }
    ngrams.join(" ")
}

fn dedupe_key(hit: &SearchHit) -> String {
    format!(
        "{}:{}:{}:{}:{}",
        hit.kind,
        hit.file_id,
        hit.heading_order.unwrap_or(0),
        hit.heading_text.clone().unwrap_or_default(),
        hit.relative_path
    )
}

fn build_hit(
    document: &TantivyDocument,
    fields: &LexicalFields,
    score: f64,
    requested_root_id: Option<i64>,
    file_name_only: bool,
) -> Option<SearchHit> {
    let root_id = i64::try_from(field_u64(document, fields.root_id)?).ok()?;
    if let Some(requested) = requested_root_id {
        if requested != root_id {
            return None;
        }
    }

    let file_id = i64::try_from(field_u64(document, fields.file_id)?).ok()?;
    let kind = field_text(document, fields.kind).unwrap_or_else(|| "file".to_string());
    if file_name_only && kind != "file" {
        return None;
    }
    let file_name = field_text(document, fields.file_name)?;
    let relative_path = field_text(document, fields.relative_path)?;
    let absolute_path = field_text(document, fields.absolute_path).unwrap_or_default();
    let heading_level = field_i64(document, fields.heading_level);
    let heading_order = field_i64(document, fields.heading_order);
    let heading_text = field_text(document, fields.heading_text)
        .or_else(|| field_text(document, fields.author_text))
        .or_else(|| {
            field_text(document, fields.chunk_text).map(|value| {
                let trimmed = value.trim();
                if trimmed.chars().count() > 180 {
                    trimmed.chars().take(180).collect::<String>()
                } else {
                    trimmed.to_string()
                }
            })
        });

    let mapped_kind = if kind == "author" {
        "author".to_string()
    } else if kind == "file" {
        "file".to_string()
    } else {
        "heading".to_string()
    };

    Some(SearchHit {
        source: "lexical".to_string(),
        kind: mapped_kind,
        file_id,
        file_name,
        relative_path,
        absolute_path,
        heading_level,
        heading_text,
        heading_order,
        score,
    })
}

pub(crate) fn replace_all_documents(
    app: &AppHandle,
    documents: &[LexicalDocument],
) -> CommandResult<()> {
    let runtime = lexical_runtime(app)?;
    let runtime = runtime
        .lock()
        .map_err(|_| "Could not lock lexical runtime".to_string())?;

    let mut writer = runtime
        .index
        .writer(96_000_000)
        .map_err(|error| format!("Could not create lexical index writer: {error}"))?;

    writer
        .delete_all_documents()
        .map_err(|error| format!("Could not clear lexical index: {error}"))?;

    for entry in documents {
        let heading_text = entry.heading_text.clone().unwrap_or_default();
        let author_text = entry.author_text.clone().unwrap_or_default();
        let chunk_text = entry.chunk_text.clone().unwrap_or_default();
        let query_text = format!(
            "{}\n{}\n{}\n{}\n{}",
            heading_text, author_text, chunk_text, entry.file_name, entry.relative_path
        );
        let prefix_text = format!(
            "{} {} {} {}",
            heading_text, author_text, entry.file_name, entry.relative_path
        );
        let ngram_text = format!(
            "{} {} {} {} {}",
            heading_text, author_text, chunk_text, entry.file_name, entry.relative_path
        );

        let mut document = doc!(
            runtime.fields.kind => entry.kind.as_str(),
            runtime.fields.root_id => u64::try_from(entry.root_id).unwrap_or(0),
            runtime.fields.file_id => u64::try_from(entry.file_id).unwrap_or(0),
            runtime.fields.file_name => entry.file_name.as_str(),
            runtime.fields.relative_path => entry.relative_path.as_str(),
            runtime.fields.absolute_path => entry.absolute_path.as_str(),
            runtime.fields.query_text => query_text,
            runtime.fields.prefix_text => prefix_text,
            runtime.fields.ngram_text => ngram_text,
            runtime.fields.chunk_text => chunk_text,
        );

        if let Some(level) = entry.heading_level {
            document.add_i64(runtime.fields.heading_level, level);
        }
        if let Some(order) = entry.heading_order {
            document.add_i64(runtime.fields.heading_order, order);
        }
        if !heading_text.is_empty() {
            document.add_text(runtime.fields.heading_text, heading_text);
        }
        if !author_text.is_empty() {
            document.add_text(runtime.fields.author_text, author_text);
        }

        writer.add_document(document).map_err(|error| {
            format!(
                "Could not add lexical document for '{}': {error}",
                entry.relative_path
            )
        })?;
    }

    writer
        .commit()
        .map_err(|error| format!("Could not commit lexical index: {error}"))?;
    runtime
        .reader
        .reload()
        .map_err(|error| format!("Could not reload lexical reader: {error}"))?;

    Ok(())
}

pub(crate) fn search(
    app: &AppHandle,
    query: &str,
    requested_root_id: Option<i64>,
    limit: usize,
    file_name_only: bool,
) -> CommandResult<Vec<SearchHit>> {
    let started = Instant::now();
    let normalized = normalize_for_search(query);
    if normalized.is_empty() {
        return Ok(Vec::new());
    }

    let runtime = lexical_runtime(app)?;
    let runtime = runtime
        .lock()
        .map_err(|_| "Could not lock lexical runtime".to_string())?;
    let searcher = runtime.reader.searcher();

    let target_limit = limit.clamp(10, 400);
    let fetch_limit = target_limit
        .saturating_mul(MIN_FETCH_MULTIPLIER)
        .clamp(MIN_FETCH_FLOOR, MAX_FETCH_LIMIT);
    let mut results = Vec::new();
    let mut seen = HashSet::new();

    let lexical_fields = if file_name_only {
        vec![runtime.fields.file_name]
    } else {
        vec![
            runtime.fields.query_text,
            runtime.fields.heading_text,
            runtime.fields.author_text,
            runtime.fields.file_name,
            runtime.fields.relative_path,
            runtime.fields.chunk_text,
        ]
    };
    let prefix_fields = if file_name_only {
        vec![runtime.fields.file_name]
    } else {
        vec![
            runtime.fields.prefix_text,
            runtime.fields.heading_text,
            runtime.fields.file_name,
            runtime.fields.relative_path,
        ]
    };
    let ngram_fields = if file_name_only {
        Vec::new()
    } else {
        vec![runtime.fields.ngram_text]
    };

    let run_tier = |query_text: &str,
                    fields: Vec<Field>,
                    conjunction: bool|
     -> CommandResult<Vec<TantivyDocument>> {
        let mut parser = QueryParser::for_index(&runtime.index, fields);
        if conjunction {
            parser.set_conjunction_by_default();
        }
        let parsed = match parser.parse_query(query_text) {
            Ok(parsed) => parsed,
            Err(_) => return Ok(Vec::new()),
        };
        let docs = searcher
            .search(&parsed, &TopDocs::with_limit(fetch_limit))
            .map_err(|error| format!("Lexical search execution failed: {error}"))?;
        let mut output = Vec::with_capacity(docs.len());
        for (_score, address) in docs {
            let doc = searcher
                .doc::<TantivyDocument>(address)
                .map_err(|error| format!("Could not read lexical result document: {error}"))?;
            output.push(doc);
        }
        Ok(output)
    };

    let mut tiers = vec![
        (normalized.clone(), lexical_fields, true, 1_000.0_f64),
        (
            normalized
                .split_whitespace()
                .map(|token| format!("{token}*"))
                .collect::<Vec<String>>()
                .join(" "),
            prefix_fields,
            true,
            2_000.0_f64,
        ),
    ];
    if !ngram_fields.is_empty() {
        tiers.push((
            ngrams_for_query(&normalized),
            ngram_fields,
            false,
            3_000.0_f64,
        ));
    }

    for (query_text, fields, conjunction, score_base) in tiers {
        if query_text.trim().is_empty() {
            continue;
        }
        let tier_documents = run_tier(&query_text, fields, conjunction)?;
        for (rank, document) in tier_documents.into_iter().enumerate() {
            if results.len() >= target_limit {
                break;
            }
            let score = score_base + f64::from(rank as u32);
            let Some(hit) = build_hit(
                &document,
                &runtime.fields,
                score,
                requested_root_id,
                file_name_only,
            ) else {
                continue;
            };
            let key = dedupe_key(&hit);
            if !seen.insert(key) {
                continue;
            }
            results.push(hit);
        }
        if results.len() >= target_limit {
            break;
        }
    }

    if started.elapsed().as_millis() > 80 {
        eprintln!(
            "Lexical search exceeded 80ms budget: {}ms query='{}'",
            started.elapsed().as_millis(),
            normalized
        );
    }

    Ok(results)
}
