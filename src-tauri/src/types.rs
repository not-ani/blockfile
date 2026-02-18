use std::collections::HashSet;
use std::path::PathBuf;

use ort::session::Session as OrtSession;
use serde::{Deserialize, Serialize};
use tokenizers::Tokenizer;

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct RootSummary {
    pub path: String,
    pub file_count: i64,
    pub heading_count: i64,
    pub added_at_ms: i64,
    pub last_indexed_ms: i64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct IndexStats {
    pub scanned: usize,
    pub updated: usize,
    pub skipped: usize,
    pub removed: usize,
    pub headings_extracted: usize,
    pub elapsed_ms: i64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FolderEntry {
    pub path: String,
    pub name: String,
    pub parent_path: Option<String>,
    pub depth: usize,
    pub file_count: usize,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct IndexedFile {
    pub id: i64,
    pub file_name: String,
    pub relative_path: String,
    pub folder_path: String,
    pub modified_ms: i64,
    pub heading_count: i64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct IndexSnapshot {
    pub root_path: String,
    pub indexed_at_ms: i64,
    pub folders: Vec<FolderEntry>,
    pub files: Vec<IndexedFile>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FileHeading {
    pub id: i64,
    pub order: i64,
    pub level: i64,
    pub text: String,
    pub copy_text: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TaggedBlock {
    pub order: i64,
    pub style_label: String,
    pub text: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FilePreview {
    pub file_id: i64,
    pub file_name: String,
    pub relative_path: String,
    pub absolute_path: String,
    pub heading_count: i64,
    pub headings: Vec<FileHeading>,
    pub f8_cites: Vec<TaggedBlock>,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SearchHit {
    pub source: String,
    pub kind: String,
    pub file_id: i64,
    pub file_name: String,
    pub relative_path: String,
    pub absolute_path: String,
    pub heading_level: Option<i64>,
    pub heading_text: Option<String>,
    pub heading_order: Option<i64>,
    pub score: f64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct CaptureInsertResult {
    pub capture_path: String,
    pub marker: String,
    pub target_relative_path: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct CaptureTarget {
    pub relative_path: String,
    pub absolute_path: String,
    pub exists: bool,
    pub entry_count: i64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct CaptureTargetPreview {
    pub relative_path: String,
    pub absolute_path: String,
    pub exists: bool,
    pub heading_count: i64,
    pub headings: Vec<FileHeading>,
}

#[derive(Clone)]
pub(crate) struct ExistingFileMeta {
    pub id: i64,
    pub modified_ms: i64,
    pub size: i64,
    pub file_hash: String,
}

#[derive(Clone)]
pub(crate) struct ParsedHeading {
    pub order: i64,
    pub level: i64,
    pub text: String,
}

#[derive(Clone)]
pub(crate) struct ParsedParagraph {
    pub order: i64,
    pub text: String,
    pub heading_level: Option<i64>,
    pub style_label: Option<String>,
    pub is_f8_cite: bool,
}

#[derive(Clone)]
pub(crate) struct HeadingRange {
    pub order: i64,
    pub level: i64,
    pub start_index: usize,
    pub end_index: usize,
}

#[derive(Clone)]
pub(crate) struct FileRecord {
    pub id: i64,
    pub relative_path: String,
    pub modified_ms: i64,
    pub heading_count: i64,
}

#[derive(Clone)]
pub(crate) struct IndexCandidate {
    pub relative_path: String,
    pub absolute_path: PathBuf,
    pub modified_ms: i64,
    pub size: i64,
    pub file_hash: String,
}

pub(crate) struct ParsedIndexCandidate {
    pub candidate: IndexCandidate,
    pub headings: Vec<ParsedHeading>,
    pub authors: Vec<(i64, String)>,
    pub chunks: Vec<ParsedChunk>,
}

#[derive(Clone)]
pub(crate) struct ParsedChunk {
    pub chunk_order: i64,
    pub heading_order: Option<i64>,
    pub heading_level: Option<i64>,
    pub heading_text: Option<String>,
    pub author_text: Option<String>,
    pub chunk_text: String,
}

#[derive(Clone)]
pub(crate) struct SemanticCandidate {
    pub semantic_id: i64,
    pub root_id: i64,
    pub kind: String,
    pub file_id: i64,
    pub file_name: String,
    pub relative_path: String,
    pub absolute_path: String,
    pub heading_level: Option<i64>,
    pub heading_text: Option<String>,
    pub heading_order: Option<i64>,
    pub semantic_text: String,
}

#[derive(Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SemanticIndexMeta {
    pub root_fingerprint_ms: i64,
    pub item_count: usize,
    pub embedding_dim: usize,
    pub updated_at_ms: i64,
}

pub(crate) struct SemanticRuntime {
    pub tokenizer: Tokenizer,
    pub session: OrtSession,
    pub output_name: String,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct IndexProgress {
    pub root_path: String,
    pub phase: String,
    pub discovered: usize,
    pub changed: usize,
    pub processed: usize,
    pub updated: usize,
    pub skipped: usize,
    pub removed: usize,
    pub elapsed_ms: i64,
    pub current_file: Option<String>,
}

#[derive(Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct BenchmarkLatencyStats {
    pub runs: usize,
    pub min_ms: f64,
    pub p50_ms: f64,
    pub p95_ms: f64,
    pub max_ms: f64,
    pub mean_ms: f64,
}

#[derive(Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct BenchmarkTaskResult {
    pub enabled: bool,
    pub error: Option<String>,
    pub total_hits: usize,
    pub latency: BenchmarkLatencyStats,
}

#[derive(Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct BenchmarkSearchSummary {
    pub query_count: usize,
    pub iterations: usize,
    pub limit: usize,
    pub lexical_raw: BenchmarkTaskResult,
    pub lexical_cached: BenchmarkTaskResult,
    pub hybrid: BenchmarkTaskResult,
    pub semantic: BenchmarkTaskResult,
}

#[derive(Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct BenchmarkPreviewSummary {
    pub snapshot_ms: f64,
    pub file_preview: BenchmarkTaskResult,
    pub heading_preview_html: BenchmarkTaskResult,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct BenchmarkReport {
    pub root_path: String,
    pub index_full: IndexStats,
    pub index_incremental: IndexStats,
    pub queries: Vec<String>,
    pub search: BenchmarkSearchSummary,
    pub preview: BenchmarkPreviewSummary,
    pub generated_at_ms: i64,
    pub elapsed_ms: i64,
}

pub(crate) struct StyledSection {
    pub paragraph_xml: Vec<String>,
    pub style_ids: HashSet<String>,
    pub relationship_ids: HashSet<String>,
    pub used_source_xml: bool,
}

pub(crate) struct SourceStyleDefinition {
    pub xml: String,
    pub dependencies: Vec<String>,
}

#[derive(Clone, Eq, PartialEq)]
pub(crate) struct RelationshipDef {
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}
