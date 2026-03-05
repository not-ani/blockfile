use std::collections::HashSet;
use std::fs;
use std::sync::{Mutex, OnceLock};
use std::time::Instant;

use rusqlite::Connection;
use tantivy::collector::TopDocs;
use tantivy::query::{BooleanQuery, Occur, Query, QueryParser, TermQuery};
use tantivy::schema::{
    Field, IndexRecordOption, NumericOptions, Schema, TextFieldIndexing, TextOptions, Value,
    STORED, STRING, TEXT,
};
use tantivy::tokenizer::{LowerCaser, NgramTokenizer, TextAnalyzer};
use tantivy::{doc, Index, IndexReader, ReloadPolicy, TantivyDocument, Term};
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
const CHUNK_PREVIEW_CHARS: usize = 480;
const LEXICAL_WRITER_HEAP_BYTES: usize = 512_000_000;

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
    chunk_preview: Field,
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

fn indexed_text_options(tokenizer: &str) -> TextOptions {
    TextOptions::default().set_indexing_options(
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
    builder.add_text_field("chunk_text", indexed_text_options("default"));
    builder.add_text_field("chunk_preview", STORED);
    builder.add_text_field("query_text", indexed_text_options("default"));
    builder.add_text_field("prefix_text", indexed_text_options(PREFIX_TOKENIZER));
    builder.add_text_field("ngram_text", indexed_text_options(NGRAM_TOKENIZER));

    builder.build()
}

fn has_required_fields(schema: &Schema) -> bool {
    schema.get_field("query_text").is_ok()
        && schema.get_field("prefix_text").is_ok()
        && schema.get_field("ngram_text").is_ok()
        && schema.get_field("chunk_preview").is_ok()
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
        chunk_preview: field(schema, "chunk_preview")?,
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
    file_name_only: bool,
) -> Option<SearchHit> {
    let _root_id = i64::try_from(field_u64(document, fields.root_id)?).ok()?;

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
        .or_else(|| field_text(document, fields.chunk_preview));

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

fn preview_text_for_chunk(chunk_text: &str) -> String {
    let trimmed = chunk_text.trim();
    if trimmed.is_empty() {
        return String::new();
    }
    if trimmed.chars().count() <= CHUNK_PREVIEW_CHARS {
        return trimmed.to_string();
    }
    trimmed
        .chars()
        .take(CHUNK_PREVIEW_CHARS)
        .collect::<String>()
}

fn add_document_to_writer(
    writer: &mut tantivy::IndexWriter,
    fields: &LexicalFields,
    entry: &LexicalDocument,
) -> CommandResult<()> {
    let heading_text = entry.heading_text.clone().unwrap_or_default();
    let author_text = entry.author_text.clone().unwrap_or_default();
    let chunk_text = entry.chunk_text.clone().unwrap_or_default();
    let chunk_preview = preview_text_for_chunk(&chunk_text);
    let query_text = format!(
        "{}\n{}\n{}\n{}",
        heading_text, author_text, entry.file_name, entry.relative_path
    );
    let prefix_text = format!(
        "{} {} {} {}",
        heading_text, author_text, entry.file_name, entry.relative_path
    );
    let ngram_text = format!(
        "{} {} {} {} {}",
        heading_text, author_text, chunk_preview, entry.file_name, entry.relative_path
    );

    let mut document = doc!(
        fields.kind => entry.kind.as_str(),
        fields.root_id => u64::try_from(entry.root_id).unwrap_or(0),
        fields.file_id => u64::try_from(entry.file_id).unwrap_or(0),
        fields.file_name => entry.file_name.as_str(),
        fields.relative_path => entry.relative_path.as_str(),
        fields.absolute_path => entry.absolute_path.as_str(),
        fields.query_text => query_text,
        fields.prefix_text => prefix_text,
        fields.ngram_text => ngram_text,
    );

    if let Some(level) = entry.heading_level {
        document.add_i64(fields.heading_level, level);
    }
    if let Some(order) = entry.heading_order {
        document.add_i64(fields.heading_order, order);
    }
    if !heading_text.is_empty() {
        document.add_text(fields.heading_text, heading_text);
    }
    if !author_text.is_empty() {
        document.add_text(fields.author_text, author_text);
    }
    if !chunk_text.is_empty() {
        document.add_text(fields.chunk_text, chunk_text);
        document.add_text(fields.chunk_preview, chunk_preview);
    }

    writer.add_document(document).map_err(|error| {
        format!(
            "Could not add lexical document for '{}': {error}",
            entry.relative_path
        )
    })?;
    Ok(())
}

pub(crate) fn replace_all_documents_from_connection(
    app: &AppHandle,
    connection: &Connection,
) -> CommandResult<()> {
    let runtime = lexical_runtime(app)?;
    let runtime = runtime
        .lock()
        .map_err(|_| "Could not lock lexical runtime".to_string())?;

    let mut writer = runtime
        .index
        .writer(LEXICAL_WRITER_HEAP_BYTES)
        .map_err(|error| format!("Could not create lexical index writer: {error}"))?;

    writer
        .delete_all_documents()
        .map_err(|error| format!("Could not clear lexical index: {error}"))?;

    {
        let mut statement = connection
            .prepare(
                "
                SELECT root_id, id, relative_path, absolute_path
                FROM files
                ORDER BY root_id ASC, relative_path ASC
                ",
            )
            .map_err(|error| format!("Could not prepare lexical file rows query: {error}"))?;

        let rows = statement
            .query_map([], |row| {
                Ok((
                    row.get::<_, i64>(0)?,
                    row.get::<_, i64>(1)?,
                    row.get::<_, String>(2)?,
                    row.get::<_, String>(3)?,
                ))
            })
            .map_err(|error| format!("Could not read lexical file rows: {error}"))?;

        for row in rows {
            let (root_id, file_id, relative_path, absolute_path) =
                row.map_err(|error| format!("Could not parse lexical file row: {error}"))?;
            let file_name = crate::util::file_name_from_relative(&relative_path);
            let entry = LexicalDocument {
                root_id,
                file_id,
                kind: "file".to_string(),
                file_name,
                relative_path,
                absolute_path,
                heading_level: None,
                heading_text: None,
                heading_order: None,
                author_text: None,
                chunk_text: None,
            };
            add_document_to_writer(&mut writer, &runtime.fields, &entry)?;
        }
    }

    {
        let mut statement = connection
            .prepare(
                "
                SELECT
                  f.root_id,
                  f.id,
                  f.relative_path,
                  f.absolute_path,
                  h.level,
                  h.text,
                  h.heading_order
                FROM headings h
                JOIN files f ON f.id = h.file_id
                ORDER BY f.root_id ASC, f.id ASC, h.heading_order ASC
                ",
            )
            .map_err(|error| format!("Could not prepare lexical heading rows query: {error}"))?;

        let rows = statement
            .query_map([], |row| {
                Ok((
                    row.get::<_, i64>(0)?,
                    row.get::<_, i64>(1)?,
                    row.get::<_, String>(2)?,
                    row.get::<_, String>(3)?,
                    row.get::<_, i64>(4)?,
                    row.get::<_, String>(5)?,
                    row.get::<_, i64>(6)?,
                ))
            })
            .map_err(|error| format!("Could not read lexical heading rows: {error}"))?;

        for row in rows {
            let (
                root_id,
                file_id,
                relative_path,
                absolute_path,
                level,
                heading_text,
                heading_order,
            ) = row.map_err(|error| format!("Could not parse lexical heading row: {error}"))?;
            let file_name = crate::util::file_name_from_relative(&relative_path);
            let entry = LexicalDocument {
                root_id,
                file_id,
                kind: "heading".to_string(),
                file_name,
                relative_path,
                absolute_path,
                heading_level: Some(level),
                heading_text: Some(heading_text),
                heading_order: Some(heading_order),
                author_text: None,
                chunk_text: None,
            };
            add_document_to_writer(&mut writer, &runtime.fields, &entry)?;
        }
    }

    {
        let mut statement = connection
            .prepare(
                "
                SELECT
                  f.root_id,
                  f.id,
                  f.relative_path,
                  f.absolute_path,
                  a.text,
                  a.author_order
                FROM authors a
                JOIN files f ON f.id = a.file_id
                ORDER BY f.root_id ASC, f.id ASC, a.author_order ASC
                ",
            )
            .map_err(|error| format!("Could not prepare lexical author rows query: {error}"))?;

        let rows = statement
            .query_map([], |row| {
                Ok((
                    row.get::<_, i64>(0)?,
                    row.get::<_, i64>(1)?,
                    row.get::<_, String>(2)?,
                    row.get::<_, String>(3)?,
                    row.get::<_, String>(4)?,
                    row.get::<_, i64>(5)?,
                ))
            })
            .map_err(|error| format!("Could not read lexical author rows: {error}"))?;

        for row in rows {
            let (root_id, file_id, relative_path, absolute_path, author_text, author_order) =
                row.map_err(|error| format!("Could not parse lexical author row: {error}"))?;
            let file_name = crate::util::file_name_from_relative(&relative_path);
            let entry = LexicalDocument {
                root_id,
                file_id,
                kind: "author".to_string(),
                file_name,
                relative_path,
                absolute_path,
                heading_level: None,
                heading_text: Some(author_text.clone()),
                heading_order: Some(author_order),
                author_text: Some(author_text),
                chunk_text: None,
            };
            add_document_to_writer(&mut writer, &runtime.fields, &entry)?;
        }
    }

    {
        let mut statement = connection
            .prepare(
                "
                SELECT
                  root_id,
                  file_id,
                  relative_path,
                  absolute_path,
                  heading_level,
                  heading_text,
                  heading_order,
                  author_text,
                  chunk_text
                FROM chunks
                ORDER BY root_id ASC, file_id ASC, chunk_order ASC
                ",
            )
            .map_err(|error| format!("Could not prepare lexical chunk rows query: {error}"))?;

        let rows = statement
            .query_map([], |row| {
                Ok((
                    row.get::<_, i64>(0)?,
                    row.get::<_, i64>(1)?,
                    row.get::<_, String>(2)?,
                    row.get::<_, String>(3)?,
                    row.get::<_, Option<i64>>(4)?,
                    row.get::<_, Option<String>>(5)?,
                    row.get::<_, Option<i64>>(6)?,
                    row.get::<_, Option<String>>(7)?,
                    row.get::<_, String>(8)?,
                ))
            })
            .map_err(|error| format!("Could not read lexical chunk rows: {error}"))?;

        for row in rows {
            let (
                root_id,
                file_id,
                relative_path,
                absolute_path,
                heading_level,
                heading_text,
                heading_order,
                author_text,
                chunk_text,
            ) = row.map_err(|error| format!("Could not parse lexical chunk row: {error}"))?;

            if chunk_text.trim().is_empty() {
                continue;
            }

            let file_name = crate::util::file_name_from_relative(&relative_path);
            let entry = LexicalDocument {
                root_id,
                file_id,
                kind: "chunk".to_string(),
                file_name,
                relative_path,
                absolute_path,
                heading_level,
                heading_text,
                heading_order,
                author_text,
                chunk_text: Some(chunk_text),
            };
            add_document_to_writer(&mut writer, &runtime.fields, &entry)?;
        }
    }

    writer
        .commit()
        .map_err(|error| format!("Could not commit lexical index: {error}"))?;

    let segment_ids = runtime
        .index
        .searchable_segment_ids()
        .map_err(|error| format!("Could not read lexical segment IDs: {error}"))?;
    if segment_ids.len() > 1 {
        writer
            .merge(&segment_ids)
            .wait()
            .map_err(|error| format!("Could not compact lexical segments: {error}"))?;
    }
    writer
        .wait_merging_threads()
        .map_err(|error| format!("Could not finalize lexical merge threads: {error}"))?;

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
    let (index, searcher, runtime_fields) = {
        let runtime = runtime
            .lock()
            .map_err(|_| "Could not lock lexical runtime".to_string())?;
        (
            runtime.index.clone(),
            runtime.reader.searcher(),
            runtime.fields.clone(),
        )
    };

    let target_limit = limit.clamp(10, 400);
    let mut results = Vec::with_capacity(target_limit);
    let mut seen = HashSet::with_capacity(target_limit.saturating_mul(2));

    let strict_fields = if file_name_only {
        vec![runtime_fields.file_name]
    } else {
        vec![
            runtime_fields.query_text,
            runtime_fields.heading_text,
            runtime_fields.author_text,
            runtime_fields.file_name,
            runtime_fields.relative_path,
        ]
    };
    let recall_fields = if file_name_only {
        vec![runtime_fields.file_name]
    } else {
        vec![
            runtime_fields.query_text,
            runtime_fields.heading_text,
            runtime_fields.author_text,
            runtime_fields.file_name,
            runtime_fields.relative_path,
            runtime_fields.chunk_text,
        ]
    };
    let prefix_fields = if file_name_only {
        vec![runtime_fields.file_name]
    } else {
        vec![
            runtime_fields.prefix_text,
            runtime_fields.heading_text,
            runtime_fields.file_name,
            runtime_fields.relative_path,
        ]
    };
    let ngram_fields = if file_name_only {
        Vec::new()
    } else {
        vec![runtime_fields.ngram_text]
    };

    let run_tier = |query_text: &str,
                    query_fields: Vec<Field>,
                    fetch_limit: usize,
                    conjunction: bool|
     -> CommandResult<Vec<TantivyDocument>> {
        let mut parser = QueryParser::for_index(&index, query_fields);
        if conjunction {
            parser.set_conjunction_by_default();
        }
        let parsed = match parser.parse_query(query_text) {
            Ok(parsed) => parsed,
            Err(_) => return Ok(Vec::new()),
        };
        let query: Box<dyn Query> = if let Some(root_id) = requested_root_id {
            let Ok(root_id_u64) = u64::try_from(root_id) else {
                return Ok(Vec::new());
            };
            let root_term = Term::from_field_u64(runtime_fields.root_id, root_id_u64);
            let root_query: Box<dyn Query> =
                Box::new(TermQuery::new(root_term, IndexRecordOption::Basic));
            Box::new(BooleanQuery::new(vec![
                (Occur::Must, parsed),
                (Occur::Must, root_query),
            ]))
        } else {
            parsed
        };

        let docs = searcher
            .search(&query, &TopDocs::with_limit(fetch_limit))
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

    let mut tiers = vec![(normalized.clone(), strict_fields, true, 1_000.0_f64)];
    if !file_name_only {
        tiers.push((normalized.clone(), recall_fields, false, 1_450.0_f64));
    }
    tiers.push((
        normalized
            .split_whitespace()
            .map(|token| format!("{token}*"))
            .collect::<Vec<String>>()
            .join(" "),
        prefix_fields,
        true,
        2_000.0_f64,
    ));
    if !ngram_fields.is_empty() {
        tiers.push((
            ngrams_for_query(&normalized),
            ngram_fields,
            false,
            3_000.0_f64,
        ));
    }

    for (query_text, tier_fields, conjunction, score_base) in tiers {
        if query_text.trim().is_empty() {
            continue;
        }
        let remaining = target_limit.saturating_sub(results.len()).max(10);
        let fetch_limit = remaining
            .saturating_mul(MIN_FETCH_MULTIPLIER)
            .clamp(MIN_FETCH_FLOOR, MAX_FETCH_LIMIT);
        let tier_documents = run_tier(&query_text, tier_fields, fetch_limit, conjunction)?;
        for (rank, document) in tier_documents.into_iter().enumerate() {
            if results.len() >= target_limit {
                break;
            }
            let score = score_base + f64::from(rank as u32);
            let Some(hit) = build_hit(&document, &runtime_fields, score, file_name_only) else {
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
