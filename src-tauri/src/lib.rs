use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::{Component, Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use docx_rs::Docx;
use rayon::prelude::*;
use roxmltree::{Document, Node};
use rusqlite::{params, Connection, OptionalExtension};
use serde::Serialize;
use tauri::{AppHandle, Emitter, Manager};
use walkdir::{DirEntry, WalkDir};
use zip::ZipArchive;

type CommandResult<T> = Result<T, String>;

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct RootSummary {
    path: String,
    file_count: i64,
    heading_count: i64,
    added_at_ms: i64,
    last_indexed_ms: i64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct IndexStats {
    scanned: usize,
    updated: usize,
    skipped: usize,
    removed: usize,
    headings_extracted: usize,
    elapsed_ms: i64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct FolderEntry {
    path: String,
    name: String,
    parent_path: Option<String>,
    depth: usize,
    file_count: usize,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct IndexedFile {
    id: i64,
    file_name: String,
    relative_path: String,
    folder_path: String,
    modified_ms: i64,
    heading_count: i64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct IndexSnapshot {
    root_path: String,
    indexed_at_ms: i64,
    folders: Vec<FolderEntry>,
    files: Vec<IndexedFile>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct FileHeading {
    id: i64,
    order: i64,
    level: i64,
    text: String,
    copy_text: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct TaggedBlock {
    order: i64,
    style_label: String,
    text: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct FilePreview {
    file_id: i64,
    file_name: String,
    relative_path: String,
    absolute_path: String,
    heading_count: i64,
    headings: Vec<FileHeading>,
    f8_cites: Vec<TaggedBlock>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct SearchHit {
    kind: String,
    file_id: i64,
    file_name: String,
    relative_path: String,
    absolute_path: String,
    heading_level: Option<i64>,
    heading_text: Option<String>,
    heading_order: Option<i64>,
    score: f64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct CaptureInsertResult {
    capture_path: String,
    marker: String,
    target_relative_path: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct CaptureTarget {
    relative_path: String,
    absolute_path: String,
    exists: bool,
    entry_count: i64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct CaptureTargetPreview {
    relative_path: String,
    absolute_path: String,
    exists: bool,
    heading_count: i64,
    headings: Vec<FileHeading>,
}

#[derive(Clone)]
struct ExistingFileMeta {
    id: i64,
    modified_ms: i64,
    size: i64,
}

#[derive(Clone)]
struct ParsedHeading {
    order: i64,
    level: i64,
    text: String,
}

#[derive(Clone)]
struct ParsedParagraph {
    order: i64,
    text: String,
    heading_level: Option<i64>,
    style_label: Option<String>,
    is_f8_cite: bool,
}

#[derive(Clone)]
struct HeadingRange {
    order: i64,
    level: i64,
    start_index: usize,
    end_index: usize,
}

#[derive(Clone)]
struct FileRecord {
    id: i64,
    relative_path: String,
    modified_ms: i64,
    heading_count: i64,
}

#[derive(Clone)]
struct IndexCandidate {
    relative_path: String,
    absolute_path: PathBuf,
    modified_ms: i64,
    size: i64,
}

struct ParsedIndexCandidate {
    candidate: IndexCandidate,
    headings: Vec<ParsedHeading>,
    authors: Vec<(i64, String)>,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct IndexProgress {
    root_path: String,
    phase: String,
    discovered: usize,
    changed: usize,
    processed: usize,
    updated: usize,
    skipped: usize,
    removed: usize,
    elapsed_ms: i64,
    current_file: Option<String>,
}

struct StyledSection {
    paragraph_xml: Vec<String>,
    style_ids: HashSet<String>,
    relationship_ids: HashSet<String>,
    used_source_xml: bool,
}

struct SourceStyleDefinition {
    xml: String,
    dependencies: Vec<String>,
}

#[derive(Clone, Eq, PartialEq)]
struct RelationshipDef {
    rel_type: String,
    target: String,
    target_mode: Option<String>,
}

const DEFAULT_CAPTURE_TARGET: &str = "BlockFile-Captures.docx";
const INDEX_PROGRESS_EVENT: &str = "index-progress";
const INDEX_PROGRESS_EMIT_INTERVAL_MS: i64 = 120;

fn now_ms() -> i64 {
    epoch_ms(SystemTime::now())
}

fn epoch_ms(time: SystemTime) -> i64 {
    time.duration_since(UNIX_EPOCH)
        .ok()
        .and_then(|duration| i64::try_from(duration.as_millis()).ok())
        .unwrap_or(0)
}

fn path_display(path: &Path) -> String {
    path.to_string_lossy().into_owned()
}

fn suggested_parse_chunk_size() -> usize {
    std::thread::available_parallelism()
        .map(|parallelism| parallelism.get().saturating_mul(2))
        .unwrap_or(8)
        .clamp(8, 64)
}

fn emit_index_progress(
    app: &AppHandle,
    started_at: i64,
    progress: &IndexProgress,
    last_emitted_ms: &mut i64,
    force: bool,
) {
    let now = now_ms();
    if !force && now - *last_emitted_ms < INDEX_PROGRESS_EMIT_INTERVAL_MS {
        return;
    }

    let mut payload = progress.clone();
    payload.elapsed_ms = now - started_at;
    let _ = app.emit(INDEX_PROGRESS_EVENT, payload);
    *last_emitted_ms = now;
}

fn canonicalize_folder(path: &str) -> CommandResult<PathBuf> {
    let canonical = fs::canonicalize(path)
        .map_err(|error| format!("Could not access folder '{path}': {error}"))?;
    if !canonical.is_dir() {
        return Err(format!("Path is not a folder: {path}"));
    }
    Ok(canonical)
}

fn root_index_marker_path(root: &Path) -> PathBuf {
    root.join(".blockfile-index.json")
}

fn normalize_capture_target_path(target_path: Option<&str>) -> CommandResult<String> {
    let raw = target_path
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .unwrap_or(DEFAULT_CAPTURE_TARGET);

    let candidate = Path::new(raw);
    let mut normalized = if candidate.is_absolute() {
        PathBuf::from(candidate)
    } else {
        let mut value = PathBuf::new();
        for component in candidate.components() {
            match component {
                Component::Normal(part) => value.push(part),
                Component::CurDir => {}
                Component::ParentDir | Component::RootDir | Component::Prefix(_) => return Err(
                    "Capture target path cannot use '..' or root-prefix components when relative."
                        .to_string(),
                ),
            }
        }
        value
    };

    if normalized.as_os_str().is_empty() {
        return Err("Capture target path cannot be empty.".to_string());
    }

    if normalized
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.eq_ignore_ascii_case("docx"))
        != Some(true)
    {
        normalized.set_extension("docx");
    }

    Ok(path_display(&normalized))
}

fn capture_docx_path(root: &Path, target_relative_path: &str) -> PathBuf {
    root.join(target_relative_path)
}

fn capture_marker(entry_id: i64) -> String {
    format!("BF-{entry_id:06}")
}

fn write_root_index_marker(root: &Path, last_indexed_ms: i64) -> CommandResult<()> {
    let marker_path = root_index_marker_path(root);
    let marker = serde_json::json!({
        "version": 1,
        "rootPath": path_display(root),
        "lastIndexedMs": last_indexed_ms,
    });
    let content = serde_json::to_string_pretty(&marker)
        .map_err(|error| format!("Could not serialize index marker JSON: {error}"))?;
    fs::write(&marker_path, content).map_err(|error| {
        format!(
            "Could not write index marker '{}': {error}",
            path_display(&marker_path)
        )
    })
}

fn xml_escape_text(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn xml_escape_attr(value: &str) -> String {
    xml_escape_text(value)
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn paragraph_xml_plain(text: &str) -> String {
    if text.is_empty() {
        return "<w:p/>".to_string();
    }
    format!(
        "<w:p><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
        xml_escape_text(text)
    )
}

fn paragraph_xml_bold(text: &str) -> String {
    format!(
        "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
        xml_escape_text(text)
    )
}

fn fallback_styled_section(content: &str) -> StyledSection {
    let mut paragraph_xml = content
        .split('\n')
        .map(|line| line.strip_suffix('\r').unwrap_or(line))
        .map(paragraph_xml_plain)
        .collect::<Vec<String>>();

    if paragraph_xml.is_empty() {
        paragraph_xml.push("<w:p/>".to_string());
    }

    StyledSection {
        paragraph_xml,
        style_ids: HashSet::new(),
        relationship_ids: HashSet::new(),
        used_source_xml: false,
    }
}

fn extract_styled_section(
    source_file_path: &Path,
    heading_order: Option<i64>,
    fallback_content: &str,
) -> StyledSection {
    let Some(heading_order) = heading_order else {
        return fallback_styled_section(fallback_content);
    };

    let Ok(paragraphs) = parse_docx_paragraphs(source_file_path) else {
        return fallback_styled_section(fallback_content);
    };

    let Some((start_index, start_paragraph)) = paragraphs
        .iter()
        .enumerate()
        .find(|(_, paragraph)| paragraph.order == heading_order)
    else {
        return fallback_styled_section(fallback_content);
    };

    let Some(start_level) = start_paragraph.heading_level else {
        return fallback_styled_section(fallback_content);
    };

    let mut end_index = paragraphs.len();
    for candidate_index in (start_index + 1)..paragraphs.len() {
        let candidate = &paragraphs[candidate_index];
        let Some(candidate_level) = candidate.heading_level else {
            continue;
        };

        if is_probable_author_line(&candidate.text) {
            continue;
        }

        if candidate_level <= start_level {
            end_index = candidate_index;
            break;
        }
    }

    if start_index >= end_index {
        return fallback_styled_section(fallback_content);
    }

    let file = match File::open(source_file_path) {
        Ok(file) => file,
        Err(_) => return fallback_styled_section(fallback_content),
    };
    let mut archive = match ZipArchive::new(file) {
        Ok(archive) => archive,
        Err(_) => return fallback_styled_section(fallback_content),
    };

    let Some(document_xml) = read_zip_file(&mut archive, "word/document.xml") else {
        return fallback_styled_section(fallback_content);
    };
    let Ok(document) = Document::parse(&document_xml) else {
        return fallback_styled_section(fallback_content);
    };

    let paragraph_nodes = document
        .descendants()
        .filter(|node| has_tag(*node, "p"))
        .collect::<Vec<Node<'_, '_>>>();

    let mut paragraph_xml = Vec::new();
    for node in paragraph_nodes
        .iter()
        .skip(start_index)
        .take(end_index - start_index)
    {
        let range = node.range();
        if range.end > document_xml.len() || range.start >= range.end {
            continue;
        }
        let snippet = document_xml[range].to_string();
        if !snippet.trim().is_empty() {
            paragraph_xml.push(snippet);
        }
    }

    if paragraph_xml.is_empty() {
        return fallback_styled_section(fallback_content);
    }

    let wrapped = format!(
        "<w:root xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">{}</w:root>",
        paragraph_xml.join("")
    );

    let mut style_ids = HashSet::new();
    let mut relationship_ids = HashSet::new();
    if let Ok(wrapper_document) = Document::parse(&wrapped) {
        for node in wrapper_document
            .descendants()
            .filter(|node| node.is_element())
        {
            if has_tag(node, "pStyle") || has_tag(node, "rStyle") {
                if let Some(style_id) = attribute_value(node, "val") {
                    if !style_id.is_empty() {
                        style_ids.insert(style_id.to_string());
                    }
                }
            }

            if has_tag(node, "hyperlink") {
                if let Some(rel_id) = attribute_value(node, "id") {
                    if !rel_id.is_empty() {
                        relationship_ids.insert(rel_id.to_string());
                    }
                }
            }

            if has_tag(node, "blip") {
                if let Some(rel_id) = attribute_value(node, "embed") {
                    if !rel_id.is_empty() {
                        relationship_ids.insert(rel_id.to_string());
                    }
                }
                if let Some(rel_id) = attribute_value(node, "link") {
                    if !rel_id.is_empty() {
                        relationship_ids.insert(rel_id.to_string());
                    }
                }
            }
        }
    }

    StyledSection {
        paragraph_xml,
        style_ids,
        relationship_ids,
        used_source_xml: true,
    }
}

fn create_blank_docx(capture_path: &Path) -> CommandResult<()> {
    let mut output = File::create(capture_path).map_err(|error| {
        format!(
            "Could not create capture docx '{}': {error}",
            path_display(capture_path)
        )
    })?;
    Docx::new().build().pack(&mut output).map_err(|error| {
        format!(
            "Could not initialize capture docx '{}': {error}",
            path_display(capture_path)
        )
    })
}

fn ensure_valid_capture_docx(capture_path: &Path) -> CommandResult<()> {
    if !capture_path.is_file() {
        return create_blank_docx(capture_path);
    }

    let file = File::open(capture_path).map_err(|error| {
        format!(
            "Could not open capture docx '{}': {error}",
            path_display(capture_path)
        )
    })?;

    let mut archive = ZipArchive::new(file).map_err(|error| {
        format!(
            "Could not read capture docx '{}': {error}",
            path_display(capture_path)
        )
    })?;

    if read_zip_file(&mut archive, "word/document.xml").is_some() {
        return Ok(());
    }

    let backup_path = capture_path.with_extension("docx.bak");
    let _ = fs::copy(capture_path, &backup_path);
    create_blank_docx(capture_path)
}

fn read_docx_part(path: &Path, part_name: &str) -> CommandResult<Option<String>> {
    let file = File::open(path)
        .map_err(|error| format!("Could not open '{}': {error}", path_display(path)))?;
    let mut archive = ZipArchive::new(file)
        .map_err(|error| format!("Could not read '{}': {error}", path_display(path)))?;
    Ok(read_zip_file(&mut archive, part_name))
}

fn document_has_body_content(document_xml: &str) -> bool {
    let Ok(document) = Document::parse(document_xml) else {
        return document_xml.contains("<w:p") || document_xml.contains("<w:tbl");
    };

    let Some(body) = document.descendants().find(|node| has_tag(*node, "body")) else {
        return false;
    };

    body.children()
        .any(|node| node.is_element() && !has_tag(node, "sectPr"))
}

fn body_bounds(document_xml: &str) -> CommandResult<(usize, usize)> {
    let body_open = document_xml
        .find("<w:body")
        .ok_or_else(|| "Could not find <w:body> in destination document.xml".to_string())?;
    let body_open_end = document_xml[body_open..]
        .find('>')
        .map(|offset| body_open + offset + 1)
        .ok_or_else(|| "Could not parse <w:body> opening tag".to_string())?;
    let body_close = document_xml
        .rfind("</w:body>")
        .ok_or_else(|| "Could not find </w:body> in destination document.xml".to_string())?;

    Ok((body_open_end, body_close))
}

fn fallback_body_insertion_index(document_xml: &str) -> CommandResult<usize> {
    let (body_open_end, body_close) = body_bounds(document_xml)?;

    let body_slice = &document_xml[body_open_end..body_close];
    Ok(
        if let Some(section_props_index) = body_slice.rfind("<w:sectPr") {
            body_open_end + section_props_index
        } else {
            body_close
        },
    )
}

fn insertion_index_after_paragraph_count(
    document_xml: &str,
    paragraph_count: usize,
) -> Option<usize> {
    if paragraph_count == 0 {
        return body_bounds(document_xml).ok().map(|(open, _)| open);
    }

    let document = Document::parse(document_xml).ok()?;
    let paragraphs = document
        .descendants()
        .filter(|node| has_tag(*node, "p"))
        .collect::<Vec<Node<'_, '_>>>();

    let paragraph_index = paragraph_count.saturating_sub(1);
    let paragraph = paragraphs.get(paragraph_index)?;
    let range = paragraph.range();
    (range.end <= document_xml.len()).then_some(range.end)
}

fn insert_fragment_into_document_xml(
    document_xml: &str,
    fragment: &str,
    after_paragraph_count: Option<usize>,
) -> CommandResult<String> {
    let fallback_index = fallback_body_insertion_index(document_xml)?;
    let insertion_index = after_paragraph_count
        .and_then(|count| insertion_index_after_paragraph_count(document_xml, count))
        .unwrap_or(fallback_index);

    let mut updated = String::with_capacity(document_xml.len() + fragment.len() + 32);
    updated.push_str(&document_xml[..insertion_index]);
    updated.push_str(fragment);
    updated.push_str(&document_xml[insertion_index..]);
    Ok(updated)
}

fn parse_source_style_definitions(styles_xml: &str) -> HashMap<String, SourceStyleDefinition> {
    let mut definitions = HashMap::new();
    let Ok(document) = Document::parse(styles_xml) else {
        return definitions;
    };

    for style in document
        .descendants()
        .filter(|node| has_tag(*node, "style"))
    {
        let Some(style_id) = attribute_value(style, "styleId") else {
            continue;
        };

        let range = style.range();
        if range.end > styles_xml.len() || range.start >= range.end {
            continue;
        }

        let mut dependencies = Vec::new();
        for dependency_node in style.children().filter(|node| node.is_element()) {
            if !(has_tag(dependency_node, "basedOn")
                || has_tag(dependency_node, "next")
                || has_tag(dependency_node, "link"))
            {
                continue;
            }

            if let Some(value) = attribute_value(dependency_node, "val") {
                if !value.is_empty() {
                    dependencies.push(value.to_string());
                }
            }
        }

        definitions.insert(
            style_id.to_string(),
            SourceStyleDefinition {
                xml: styles_xml[range].to_string(),
                dependencies,
            },
        );
    }

    definitions
}

fn parse_style_ids(styles_xml: &str) -> HashSet<String> {
    let mut ids = HashSet::new();
    let Ok(document) = Document::parse(styles_xml) else {
        return ids;
    };

    for style in document
        .descendants()
        .filter(|node| has_tag(*node, "style"))
    {
        if let Some(style_id) = attribute_value(style, "styleId") {
            if !style_id.is_empty() {
                ids.insert(style_id.to_string());
            }
        }
    }

    ids
}

fn collect_required_style_ids(
    requested_ids: &HashSet<String>,
    definitions: &HashMap<String, SourceStyleDefinition>,
) -> Vec<String> {
    fn visit(
        style_id: &str,
        definitions: &HashMap<String, SourceStyleDefinition>,
        seen: &mut HashSet<String>,
        ordered: &mut Vec<String>,
    ) {
        if !seen.insert(style_id.to_string()) {
            return;
        }

        if let Some(definition) = definitions.get(style_id) {
            for dependency in &definition.dependencies {
                visit(dependency, definitions, seen, ordered);
            }
            ordered.push(style_id.to_string());
        }
    }

    let mut seen = HashSet::new();
    let mut ordered = Vec::new();
    for style_id in requested_ids {
        visit(style_id, definitions, &mut seen, &mut ordered);
    }
    ordered
}

fn merge_missing_styles(
    target_styles_xml: &str,
    source_styles_xml: &str,
    requested_style_ids: &HashSet<String>,
) -> String {
    if requested_style_ids.is_empty() {
        return target_styles_xml.to_string();
    }

    let definitions = parse_source_style_definitions(source_styles_xml);
    if definitions.is_empty() {
        return target_styles_xml.to_string();
    }

    let required_ids = collect_required_style_ids(requested_style_ids, &definitions);
    if required_ids.is_empty() {
        return target_styles_xml.to_string();
    }

    let mut existing_ids = parse_style_ids(target_styles_xml);
    let mut to_append = Vec::new();
    for style_id in required_ids {
        if existing_ids.contains(&style_id) {
            continue;
        }
        if let Some(definition) = definitions.get(&style_id) {
            to_append.push(definition.xml.clone());
            existing_ids.insert(style_id);
        }
    }

    if to_append.is_empty() {
        return target_styles_xml.to_string();
    }

    if let Some(styles_close) = target_styles_xml.rfind("</w:styles>") {
        let mut updated = String::with_capacity(target_styles_xml.len() + to_append.join("").len());
        updated.push_str(&target_styles_xml[..styles_close]);
        for snippet in &to_append {
            updated.push_str(snippet);
        }
        updated.push_str(&target_styles_xml[styles_close..]);
        return updated;
    }

    let mut fallback = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
    );
    for snippet in &to_append {
        fallback.push_str(snippet);
    }
    fallback.push_str("</w:styles>");
    fallback
}

fn parse_relationships(relationships_xml: &str) -> HashMap<String, RelationshipDef> {
    let mut relationships = HashMap::new();
    let Ok(document) = Document::parse(relationships_xml) else {
        return relationships;
    };

    for relationship in document
        .descendants()
        .filter(|node| has_tag(*node, "Relationship"))
    {
        let Some(id) = attribute_value(relationship, "Id") else {
            continue;
        };
        let Some(rel_type) = attribute_value(relationship, "Type") else {
            continue;
        };
        let Some(target) = attribute_value(relationship, "Target") else {
            continue;
        };
        let target_mode = attribute_value(relationship, "TargetMode").map(str::to_string);

        relationships.insert(
            id.to_string(),
            RelationshipDef {
                rel_type: rel_type.to_string(),
                target: target.to_string(),
                target_mode,
            },
        );
    }

    relationships
}

fn next_relationship_id(existing_ids: &HashSet<String>) -> String {
    let mut max_numeric = 0_i64;
    for id in existing_ids {
        if let Some(raw) = id.strip_prefix("rId") {
            if let Ok(value) = raw.parse::<i64>() {
                max_numeric = max_numeric.max(value);
            }
        }
    }

    let mut next = max_numeric + 1;
    loop {
        let candidate = format!("rId{next}");
        if !existing_ids.contains(&candidate) {
            return candidate;
        }
        next += 1;
    }
}

fn relationship_xml(id: &str, definition: &RelationshipDef) -> String {
    let mut xml = format!(
        "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"",
        xml_escape_attr(id),
        xml_escape_attr(&definition.rel_type),
        xml_escape_attr(&definition.target)
    );
    if let Some(target_mode) = &definition.target_mode {
        xml.push_str(&format!(" TargetMode=\"{}\"", xml_escape_attr(target_mode)));
    }
    xml.push_str("/>");
    xml
}

fn merge_relationships(
    target_relationships_xml: &str,
    source_relationships_xml: &str,
    requested_relationship_ids: &HashSet<String>,
) -> (String, HashMap<String, String>) {
    if requested_relationship_ids.is_empty() {
        return (target_relationships_xml.to_string(), HashMap::new());
    }

    let source_relationships = parse_relationships(source_relationships_xml);
    if source_relationships.is_empty() {
        return (target_relationships_xml.to_string(), HashMap::new());
    }

    let mut target_relationships = parse_relationships(target_relationships_xml);
    let mut existing_ids = target_relationships
        .keys()
        .cloned()
        .collect::<HashSet<String>>();
    let mut id_remap = HashMap::new();
    let mut appended_xml = Vec::new();

    for requested_id in requested_relationship_ids {
        let Some(source_definition) = source_relationships.get(requested_id) else {
            continue;
        };

        if let Some(existing_definition) = target_relationships.get(requested_id) {
            if existing_definition == source_definition {
                continue;
            }
        } else {
            target_relationships.insert(requested_id.to_string(), source_definition.clone());
            existing_ids.insert(requested_id.to_string());
            appended_xml.push(relationship_xml(requested_id, source_definition));
            continue;
        }

        if let Some((existing_id, _)) = target_relationships
            .iter()
            .find(|(_, definition)| *definition == source_definition)
        {
            id_remap.insert(requested_id.to_string(), existing_id.to_string());
            continue;
        }

        let new_id = next_relationship_id(&existing_ids);
        existing_ids.insert(new_id.clone());
        target_relationships.insert(new_id.clone(), source_definition.clone());
        id_remap.insert(requested_id.to_string(), new_id.clone());
        appended_xml.push(relationship_xml(&new_id, source_definition));
    }

    if appended_xml.is_empty() {
        return (target_relationships_xml.to_string(), id_remap);
    }

    if let Some(close_index) = target_relationships_xml.rfind("</Relationships>") {
        let mut updated = String::with_capacity(
            target_relationships_xml.len() + appended_xml.join("").len() + 32,
        );
        updated.push_str(&target_relationships_xml[..close_index]);
        for snippet in &appended_xml {
            updated.push_str(snippet);
        }
        updated.push_str(&target_relationships_xml[close_index..]);
        return (updated, id_remap);
    }

    let mut fallback = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for snippet in &appended_xml {
        fallback.push_str(snippet);
    }
    fallback.push_str("</Relationships>");
    (fallback, id_remap)
}

fn remap_relationship_ids(paragraph_xml: &mut [String], id_remap: &HashMap<String, String>) {
    if id_remap.is_empty() {
        return;
    }

    for paragraph in paragraph_xml.iter_mut() {
        let mut updated = paragraph.clone();
        for (from, to) in id_remap {
            for attribute in ["r:id", "r:embed", "r:link"] {
                updated = updated.replace(
                    &format!("{}=\"{}\"", attribute, from),
                    &format!("{}=\"{}\"", attribute, to),
                );
                updated = updated.replace(
                    &format!("{}='{}'", attribute, from),
                    &format!("{}='{}'", attribute, to),
                );
            }
        }
        *paragraph = updated;
    }
}

fn rewrite_docx_with_parts(
    capture_path: &Path,
    replacements: &HashMap<String, Vec<u8>>,
) -> CommandResult<()> {
    let source_file = File::open(capture_path).map_err(|error| {
        format!(
            "Could not open capture docx '{}' for update: {error}",
            path_display(capture_path)
        )
    })?;
    let mut archive = ZipArchive::new(source_file).map_err(|error| {
        format!(
            "Could not read capture docx '{}' for update: {error}",
            path_display(capture_path)
        )
    })?;

    let temp_path = capture_path.with_extension("docx.tmp");
    let temp_file = File::create(&temp_path).map_err(|error| {
        format!(
            "Could not create temporary capture file '{}': {error}",
            path_display(&temp_path)
        )
    })?;
    let mut writer = zip::ZipWriter::new(temp_file);
    let mut copied_names = HashSet::new();

    for index in 0..archive.len() {
        let mut entry = archive
            .by_index(index)
            .map_err(|error| format!("Could not read capture docx entry: {error}"))?;
        let name = entry.name().to_string();
        if entry.is_dir() {
            continue;
        }

        let options =
            zip::write::SimpleFileOptions::default().compression_method(entry.compression());
        writer
            .start_file(name.clone(), options)
            .map_err(|error| format!("Could not write capture zip entry '{name}': {error}"))?;

        if let Some(updated_bytes) = replacements.get(&name) {
            writer
                .write_all(updated_bytes)
                .map_err(|error| format!("Could not write capture zip entry '{name}': {error}"))?;
        } else {
            let mut original = Vec::new();
            entry
                .read_to_end(&mut original)
                .map_err(|error| format!("Could not read capture zip entry '{name}': {error}"))?;
            writer
                .write_all(&original)
                .map_err(|error| format!("Could not write capture zip entry '{name}': {error}"))?;
        }

        copied_names.insert(name);
    }

    for (name, updated_bytes) in replacements {
        if copied_names.contains(name) {
            continue;
        }

        writer
            .start_file(name, zip::write::SimpleFileOptions::default())
            .map_err(|error| format!("Could not add capture zip entry '{name}': {error}"))?;
        writer
            .write_all(updated_bytes)
            .map_err(|error| format!("Could not add capture zip entry '{name}': {error}"))?;
    }

    writer
        .finish()
        .map_err(|error| format!("Could not finish capture zip rewrite: {error}"))?;

    match fs::rename(&temp_path, capture_path) {
        Ok(()) => Ok(()),
        Err(_) => {
            fs::remove_file(capture_path).map_err(|error| {
                format!(
                    "Could not replace capture docx '{}': {error}",
                    path_display(capture_path)
                )
            })?;
            fs::rename(&temp_path, capture_path).map_err(|error| {
                format!(
                    "Could not move updated capture docx into place '{}': {error}",
                    path_display(capture_path)
                )
            })
        }
    }
}

fn append_capture_to_docx(
    capture_path: &Path,
    source_file_path: &Path,
    heading_level: Option<i64>,
    selected_target_heading_order: Option<i64>,
    styled_section: &StyledSection,
) -> CommandResult<()> {
    if let Some(parent) = capture_path.parent() {
        fs::create_dir_all(parent).map_err(|error| {
            format!(
                "Could not create capture target folder '{}': {error}",
                path_display(parent)
            )
        })?;
    }

    ensure_valid_capture_docx(capture_path)?;

    let target_document_xml =
        read_docx_part(capture_path, "word/document.xml")?.ok_or_else(|| {
            format!(
                "Missing word/document.xml in '{}' after initialization",
                path_display(capture_path)
            )
        })?;
    let mut target_styles_xml = read_docx_part(capture_path, "word/styles.xml")?.unwrap_or_else(|| {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"></w:styles>".to_string()
    });
    let mut target_relationships_xml = read_docx_part(capture_path, "word/_rels/document.xml.rels")?
        .unwrap_or_else(|| {
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>".to_string()
        });

    let mut section_paragraph_xml = styled_section.paragraph_xml.clone();
    let destination_paragraphs = parse_docx_paragraphs(capture_path).unwrap_or_default();

    if styled_section.used_source_xml {
        if !styled_section.style_ids.is_empty() {
            if let Ok(Some(source_styles_xml)) = read_docx_part(source_file_path, "word/styles.xml")
            {
                target_styles_xml = merge_missing_styles(
                    &target_styles_xml,
                    &source_styles_xml,
                    &styled_section.style_ids,
                );
            }
        }

        if !styled_section.relationship_ids.is_empty() {
            if let Ok(Some(source_relationships_xml)) =
                read_docx_part(source_file_path, "word/_rels/document.xml.rels")
            {
                let (merged_relationships, id_remap) = merge_relationships(
                    &target_relationships_xml,
                    &source_relationships_xml,
                    &styled_section.relationship_ids,
                );
                target_relationships_xml = merged_relationships;
                remap_relationship_ids(&mut section_paragraph_xml, &id_remap);
            }
        }
    }

    let mut fragment = String::new();
    if !document_has_body_content(&target_document_xml) {
        fragment.push_str(&paragraph_xml_bold("Block File Captures"));
    }

    for paragraph in &section_paragraph_xml {
        fragment.push_str(paragraph);
    }
    fragment.push_str("<w:p/>");

    let insert_after_order = resolve_insert_after_order(
        &destination_paragraphs,
        selected_target_heading_order,
        heading_level,
    );
    let insert_after_paragraph_count =
        insert_after_order.and_then(|value| usize::try_from(value).ok());

    let updated_document_xml = insert_fragment_into_document_xml(
        &target_document_xml,
        &fragment,
        insert_after_paragraph_count,
    )?;

    let mut replacements = HashMap::new();
    replacements.insert(
        "word/document.xml".to_string(),
        updated_document_xml.into_bytes(),
    );
    replacements.insert(
        "word/styles.xml".to_string(),
        target_styles_xml.into_bytes(),
    );
    replacements.insert(
        "word/_rels/document.xml.rels".to_string(),
        target_relationships_xml.into_bytes(),
    );

    rewrite_docx_with_parts(capture_path, &replacements)
}

fn database_path(app: &AppHandle) -> CommandResult<PathBuf> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Could not resolve app data dir: {error}"))?;
    fs::create_dir_all(&app_data).map_err(|error| {
        format!(
            "Could not create app data dir '{}': {error}",
            path_display(&app_data)
        )
    })?;
    Ok(app_data.join("blockfile-index-v1.sqlite3"))
}

fn table_has_column(connection: &Connection, table: &str, column: &str) -> CommandResult<bool> {
    let mut statement = connection
        .prepare(&format!("PRAGMA table_info({table})"))
        .map_err(|error| format!("Could not inspect table schema for '{table}': {error}"))?;

    let rows = statement
        .query_map([], |row| row.get::<_, String>(1))
        .map_err(|error| format!("Could not iterate schema for '{table}': {error}"))?;

    for row in rows {
        if row.map_err(|error| format!("Could not parse schema row for '{table}': {error}"))?
            == column
        {
            return Ok(true);
        }
    }

    Ok(false)
}

fn ensure_capture_schema(connection: &Connection) -> CommandResult<()> {
    if !table_has_column(connection, "captures", "target_relative_path")? {
        connection
            .execute(
                "ALTER TABLE captures ADD COLUMN target_relative_path TEXT NOT NULL DEFAULT 'BlockFile-Captures.docx'",
                [],
            )
            .map_err(|error| format!("Could not add captures.target_relative_path: {error}"))?;
    }

    if !table_has_column(connection, "captures", "heading_level")? {
        connection
            .execute("ALTER TABLE captures ADD COLUMN heading_level INTEGER", [])
            .map_err(|error| format!("Could not add captures.heading_level: {error}"))?;
    }

    connection
        .execute(
            "UPDATE captures SET target_relative_path = 'BlockFile-Captures.docx' WHERE target_relative_path IS NULL OR target_relative_path = ''",
            [],
        )
        .map_err(|error| format!("Could not backfill capture target paths: {error}"))?;

    connection
        .execute_batch(
            "CREATE INDEX IF NOT EXISTS idx_captures_root_target ON captures(root_id, target_relative_path, id);",
        )
        .map_err(|error| format!("Could not create captures target index: {error}"))?;

    Ok(())
}

fn open_database(app: &AppHandle) -> CommandResult<Connection> {
    let db_path = database_path(app)?;
    let connection = Connection::open(&db_path).map_err(|error| {
        format!(
            "Could not open database '{}': {error}",
            path_display(&db_path)
        )
    })?;

    connection
        .query_row("PRAGMA journal_mode = WAL", [], |row| {
            row.get::<_, String>(0)
        })
        .map_err(|error| format!("Could not set journal mode: {error}"))?;

    connection
        .execute_batch(
            "
            PRAGMA foreign_keys = ON;
            PRAGMA synchronous = NORMAL;
            PRAGMA temp_store = MEMORY;

            CREATE TABLE IF NOT EXISTS roots (
              id INTEGER PRIMARY KEY,
              path TEXT NOT NULL UNIQUE,
              added_at_ms INTEGER NOT NULL,
              last_indexed_ms INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS files (
              id INTEGER PRIMARY KEY,
              root_id INTEGER NOT NULL,
              relative_path TEXT NOT NULL,
              absolute_path TEXT NOT NULL,
              modified_ms INTEGER NOT NULL,
              size INTEGER NOT NULL,
              heading_count INTEGER NOT NULL DEFAULT 0,
              UNIQUE(root_id, relative_path),
              FOREIGN KEY(root_id) REFERENCES roots(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS headings (
              id INTEGER PRIMARY KEY,
              file_id INTEGER NOT NULL,
              heading_order INTEGER NOT NULL,
              level INTEGER NOT NULL,
              text TEXT NOT NULL,
              normalized TEXT NOT NULL,
              file_name TEXT NOT NULL,
              relative_path TEXT NOT NULL,
              FOREIGN KEY(file_id) REFERENCES files(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS authors (
              id INTEGER PRIMARY KEY,
              file_id INTEGER NOT NULL,
              author_order INTEGER NOT NULL,
              text TEXT NOT NULL,
              normalized TEXT NOT NULL,
              file_name TEXT NOT NULL,
              relative_path TEXT NOT NULL,
              FOREIGN KEY(file_id) REFERENCES files(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS captures (
              id INTEGER PRIMARY KEY,
              root_id INTEGER NOT NULL,
              source_path TEXT NOT NULL,
              section_title TEXT NOT NULL,
              target_relative_path TEXT NOT NULL DEFAULT 'BlockFile-Captures.docx',
              heading_level INTEGER,
              content TEXT NOT NULL,
              created_at_ms INTEGER NOT NULL,
              FOREIGN KEY(root_id) REFERENCES roots(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_files_root_relative ON files(root_id, relative_path);
            CREATE INDEX IF NOT EXISTS idx_headings_file ON headings(file_id);
            CREATE INDEX IF NOT EXISTS idx_authors_file ON authors(file_id);
            CREATE INDEX IF NOT EXISTS idx_captures_root ON captures(root_id, id);

            CREATE VIRTUAL TABLE IF NOT EXISTS search_fts USING fts5(
              heading_text,
              normalized,
              file_name,
              relative_path,
              tokenize = 'unicode61 remove_diacritics 2'
            );

            CREATE VIRTUAL TABLE IF NOT EXISTS author_fts USING fts5(
              author_text,
              normalized,
              file_name,
              relative_path,
              tokenize = 'unicode61 remove_diacritics 2'
            );

            CREATE TRIGGER IF NOT EXISTS headings_insert_fts AFTER INSERT ON headings BEGIN
              INSERT INTO search_fts(rowid, heading_text, normalized, file_name, relative_path)
              VALUES (new.id, new.text, new.normalized, new.file_name, new.relative_path);
            END;

            CREATE TRIGGER IF NOT EXISTS headings_delete_fts AFTER DELETE ON headings BEGIN
              DELETE FROM search_fts WHERE rowid = old.id;
            END;

            CREATE TRIGGER IF NOT EXISTS headings_update_fts AFTER UPDATE ON headings BEGIN
              UPDATE search_fts
              SET heading_text = new.text,
                  normalized = new.normalized,
                  file_name = new.file_name,
                  relative_path = new.relative_path
              WHERE rowid = old.id;
            END;

            CREATE TRIGGER IF NOT EXISTS authors_insert_fts AFTER INSERT ON authors BEGIN
              INSERT INTO author_fts(rowid, author_text, normalized, file_name, relative_path)
              VALUES (new.id, new.text, new.normalized, new.file_name, new.relative_path);
            END;

            CREATE TRIGGER IF NOT EXISTS authors_delete_fts AFTER DELETE ON authors BEGIN
              DELETE FROM author_fts WHERE rowid = old.id;
            END;

            CREATE TRIGGER IF NOT EXISTS authors_update_fts AFTER UPDATE ON authors BEGIN
              UPDATE author_fts
              SET author_text = new.text,
                  normalized = new.normalized,
                  file_name = new.file_name,
                  relative_path = new.relative_path
              WHERE rowid = old.id;
            END;
            ",
        )
        .map_err(|error| format!("Could not initialize index database: {error}"))?;

    ensure_capture_schema(&connection)?;

    Ok(connection)
}

fn file_name_from_relative(relative_path: &str) -> String {
    Path::new(relative_path)
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| relative_path.to_string())
}

fn folder_from_relative(relative_path: &str) -> String {
    relative_path
        .rsplit_once('/')
        .map(|(folder, _)| folder.to_string())
        .unwrap_or_default()
}

fn is_visible_entry(entry: &DirEntry) -> bool {
    let name = entry.file_name().to_string_lossy();
    !name.starts_with('.')
}

fn relative_path(root: &Path, file_path: &Path) -> CommandResult<String> {
    let relative = file_path
        .strip_prefix(root)
        .map_err(|error| format!("Failed to strip root prefix: {error}"))?;
    Ok(relative.to_string_lossy().replace('\\', "/"))
}

fn normalize_for_search(text: &str) -> String {
    let mut normalized = String::with_capacity(text.len());
    let mut previous_space = false;
    for character in text.chars() {
        if character.is_alphanumeric() {
            previous_space = false;
            for lower in character.to_lowercase() {
                normalized.push(lower);
            }
        } else if !previous_space {
            normalized.push(' ');
            previous_space = true;
        }
    }
    normalized.trim().to_string()
}

fn contains_year_token(text: &str) -> bool {
    for token in text
        .split(|character: char| !character.is_ascii_digit())
        .filter(|token| token.len() == 4)
    {
        if let Ok(year) = token.parse::<i32>() {
            if (1900..=2099).contains(&year) {
                return true;
            }
        }
    }
    false
}

fn is_probable_author_line(text: &str) -> bool {
    let normalized = normalize_for_search(text);
    if normalized.is_empty() {
        return false;
    }

    let word_count = normalized.split_whitespace().count();
    if !(3..=90).contains(&word_count) {
        return false;
    }

    let has_year = contains_year_token(&normalized);
    if !has_year {
        return false;
    }

    let comma_count = text.matches(',').count();
    let has_source_marker = normalized.contains("journal")
        || normalized.contains("university")
        || normalized.contains("postdoctoral")
        || normalized.contains("vol ")
        || normalized.contains("edition")
        || normalized.contains("press")
        || normalized.contains("retrieved")
        || normalized.contains("archive");
    let looks_like_url_line = normalized.contains("http") || normalized.contains("doi");

    (comma_count >= 2 || has_source_marker || looks_like_url_line) && word_count >= 5
}

fn extract_author_candidates(paragraphs: &[ParsedParagraph]) -> Vec<(i64, String)> {
    let mut seen = HashSet::new();
    let mut authors = Vec::new();

    for paragraph in paragraphs {
        if !is_probable_author_line(&paragraph.text) {
            continue;
        }

        let normalized = normalize_for_search(&paragraph.text);
        if normalized.is_empty() || !seen.insert(normalized) {
            continue;
        }

        authors.push((paragraph.order, paragraph.text.clone()));
        if authors.len() >= 120 {
            break;
        }
    }

    authors
}

fn tokenize_for_fts(query: &str) -> String {
    normalize_for_search(query)
        .split_whitespace()
        .take(12)
        .map(|token| format!("{token}*"))
        .collect::<Vec<String>>()
        .join(" AND ")
}

fn normalized_levenshtein_similarity(left: &str, right: &str) -> f64 {
    if left.is_empty() || right.is_empty() {
        return 0.0;
    }
    if left == right {
        return 1.0;
    }

    let left_chars = left.chars().collect::<Vec<char>>();
    let right_chars = right.chars().collect::<Vec<char>>();
    let left_len = left_chars.len();
    let right_len = right_chars.len();
    if left_len == 0 || right_len == 0 {
        return 0.0;
    }

    let mut previous_row = (0..=right_len).collect::<Vec<usize>>();
    let mut current_row = vec![0_usize; right_len + 1];

    for (left_index, left_char) in left_chars.iter().enumerate() {
        current_row[0] = left_index + 1;

        for (right_index, right_char) in right_chars.iter().enumerate() {
            let substitution_cost = usize::from(left_char != right_char);
            let deletion = previous_row[right_index + 1] + 1;
            let insertion = current_row[right_index] + 1;
            let substitution = previous_row[right_index] + substitution_cost;
            current_row[right_index + 1] = deletion.min(insertion).min(substitution);
        }

        std::mem::swap(&mut previous_row, &mut current_row);
    }

    let edit_distance = previous_row[right_len];
    let max_len = left_len.max(right_len);
    1.0 - (edit_distance as f64 / max_len as f64)
}

fn fuzzy_similarity(query: &str, candidate: &str) -> f64 {
    if query.is_empty() || candidate.is_empty() {
        return 0.0;
    }

    if candidate.contains(query) {
        return 0.96;
    }
    if query.contains(candidate) {
        return 0.88;
    }

    let edit_similarity = normalized_levenshtein_similarity(query, candidate);

    let query_tokens = query.split_whitespace().collect::<Vec<&str>>();
    let candidate_tokens = candidate.split_whitespace().collect::<Vec<&str>>();

    let mut best_token_similarity = 0.0_f64;
    for query_token in &query_tokens {
        for candidate_token in &candidate_tokens {
            let similarity = normalized_levenshtein_similarity(query_token, candidate_token);
            if similarity > best_token_similarity {
                best_token_similarity = similarity;
            }
        }
    }

    (edit_similarity * 0.72) + (best_token_similarity * 0.28)
}

fn fuzzy_threshold(query: &str) -> f64 {
    let query_len = query.chars().count();
    if query_len <= 4 {
        0.58
    } else if query_len <= 7 {
        0.64
    } else if query_len <= 12 {
        0.70
    } else {
        0.74
    }
}

fn has_tag(node: Node<'_, '_>, expected: &str) -> bool {
    node.is_element() && node.tag_name().name() == expected
}

fn attribute_value<'a>(node: Node<'a, 'a>, key: &str) -> Option<&'a str> {
    if let Some(value) = node.attribute(key) {
        return Some(value);
    }
    node.attributes()
        .find_map(|attribute| (attribute.name().ends_with(key)).then_some(attribute.value()))
}

fn parse_trailing_level(value: &str) -> Option<i64> {
    let lowered = value.to_ascii_lowercase();

    if let Some(without_h) = lowered.strip_prefix('h') {
        if let Ok(level) = without_h.parse::<i64>() {
            if (1..=9).contains(&level) {
                return Some(level);
            }
        }
    }

    if let Some(index) = lowered.find("heading") {
        let tail = &lowered[index + "heading".len()..];
        let digits: String = tail
            .chars()
            .filter(|character| character.is_ascii_digit())
            .collect();
        if let Ok(level) = digits.parse::<i64>() {
            if (1..=9).contains(&level) {
                return Some(level);
            }
        }
    }

    None
}

fn read_zip_file(archive: &mut ZipArchive<File>, entry_name: &str) -> Option<String> {
    let mut entry = archive.by_name(entry_name).ok()?;
    let mut value = String::new();
    entry.read_to_string(&mut value).ok()?;
    Some(value)
}

fn read_style_map(styles_xml: Option<String>) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let Some(styles_xml) = styles_xml else {
        return map;
    };

    let Ok(document) = Document::parse(&styles_xml) else {
        return map;
    };

    for style in document
        .descendants()
        .filter(|node| has_tag(*node, "style"))
    {
        let Some(style_id) = attribute_value(style, "styleId") else {
            continue;
        };

        let mut display_name = style_id.to_string();
        if let Some(name_node) = style.children().find(|node| has_tag(*node, "name")) {
            if let Some(value) = attribute_value(name_node, "val") {
                display_name = value.to_string();
            }
        }

        map.insert(style_id.to_string(), display_name);
    }

    map
}

fn extract_paragraph_text(paragraph: Node<'_, '_>) -> String {
    let mut value = String::new();

    for node in paragraph.descendants().filter(|node| node.is_element()) {
        if has_tag(node, "t") {
            if let Some(text) = node.text() {
                value.push_str(text);
            }
        } else if has_tag(node, "tab") {
            value.push('\t');
        } else if has_tag(node, "br") || has_tag(node, "cr") {
            value.push('\n');
        }
    }

    value
}

fn html_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&#39;")
}

fn push_escaped_text_with_breaks(target: &mut String, text: &str) {
    for (index, segment) in text.split('\n').enumerate() {
        if index > 0 {
            target.push_str("<br/>");
        }
        target.push_str(&html_escape(segment));
    }
}

fn run_properties_node<'a>(run: Node<'a, 'a>) -> Option<Node<'a, 'a>> {
    run.children().find(|node| has_tag(*node, "rPr"))
}

fn run_has_property(run: Node<'_, '_>, property_tag: &str) -> bool {
    run_properties_node(run)
        .and_then(|props| props.children().find(|node| has_tag(*node, property_tag)))
        .is_some()
}

fn run_has_active_underline(run: Node<'_, '_>) -> bool {
    let Some(props) = run_properties_node(run) else {
        return false;
    };

    let Some(underline) = props.children().find(|node| has_tag(*node, "u")) else {
        return false;
    };

    let Some(value) = attribute_value(underline, "val") else {
        return true;
    };

    !(value.eq_ignore_ascii_case("none")
        || value.eq_ignore_ascii_case("false")
        || value.eq_ignore_ascii_case("0"))
}

fn run_highlight_class(run: Node<'_, '_>) -> Option<&'static str> {
    let props = run_properties_node(run)?;
    let highlight = props.children().find(|node| has_tag(*node, "highlight"))?;
    let value = attribute_value(highlight, "val")?
        .trim()
        .to_ascii_lowercase();

    match value.as_str() {
        "yellow" | "darkyellow" => Some("yellow"),
        "green" | "darkgreen" => Some("green"),
        "cyan" | "darkcyan" | "turquoise" => Some("cyan"),
        "magenta" | "darkmagenta" | "pink" => Some("magenta"),
        "blue" | "darkblue" => Some("blue"),
        "gray" | "grey" | "lightgray" | "darkgray" | "gray25" | "gray50" => Some("gray"),
        _ => None,
    }
}

fn render_preview_run(run: Node<'_, '_>) -> String {
    let mut body = String::new();
    for node in run.descendants().filter(|node| node.is_element()) {
        if has_tag(node, "t") {
            if let Some(text) = node.text() {
                push_escaped_text_with_breaks(&mut body, text);
            }
        } else if has_tag(node, "tab") {
            body.push('\t');
        } else if has_tag(node, "br") || has_tag(node, "cr") {
            body.push_str("<br/>");
        }
    }

    if body.is_empty() {
        return String::new();
    }

    let mut classes = vec!["bf-run".to_string()];
    if run_has_property(run, "b") {
        classes.push("bf-run-bold".to_string());
    }
    if run_has_property(run, "i") {
        classes.push("bf-run-italic".to_string());
    }
    if run_has_active_underline(run) {
        classes.push("bf-run-underline".to_string());
    }
    if run_has_property(run, "smallCaps") || run_has_property(run, "caps") {
        classes.push("bf-run-smallcaps".to_string());
    }
    if let Some(highlight_class) = run_highlight_class(run) {
        classes.push("bf-run-highlight".to_string());
        classes.push(format!("bf-hl-{highlight_class}"));
    }

    format!("<span class=\"{}\">{body}</span>", classes.join(" "))
}

fn render_preview_inline_nodes(node: Node<'_, '_>, output: &mut String) {
    if !node.is_element() {
        return;
    }

    if has_tag(node, "hyperlink") {
        let mut link_body = String::new();
        for child in node.children() {
            render_preview_inline_nodes(child, &mut link_body);
        }
        if !link_body.is_empty() {
            output.push_str("<a class=\"bf-preview-link\">");
            output.push_str(&link_body);
            output.push_str("</a>");
        }
        return;
    }

    if has_tag(node, "r") {
        output.push_str(&render_preview_run(node));
        return;
    }

    if has_tag(node, "t") {
        if let Some(text) = node.text() {
            push_escaped_text_with_breaks(output, text);
        }
        return;
    }

    if has_tag(node, "tab") {
        output.push('\t');
        return;
    }

    if has_tag(node, "br") || has_tag(node, "cr") {
        output.push_str("<br/>");
        return;
    }

    for child in node.children() {
        render_preview_inline_nodes(child, output);
    }
}

fn preview_paragraph_class(heading_level: Option<i64>) -> &'static str {
    match heading_level {
        Some(1) => "bf-preview-h1",
        Some(2) => "bf-preview-h2",
        Some(3) => "bf-preview-h3",
        Some(4) => "bf-preview-h4",
        _ => "bf-preview-p",
    }
}

fn render_preview_paragraph(
    paragraph_node: Node<'_, '_>,
    heading_level: Option<i64>,
    fallback_text: &str,
) -> String {
    let mut body = String::new();
    for child in paragraph_node.children() {
        render_preview_inline_nodes(child, &mut body);
    }

    if body.trim().is_empty() && !fallback_text.trim().is_empty() {
        push_escaped_text_with_breaks(&mut body, fallback_text);
    }
    if body.trim().is_empty() {
        body.push_str("&nbsp;");
    }

    format!(
        "<p class=\"{}\">{body}</p>",
        preview_paragraph_class(heading_level)
    )
}

fn extract_heading_preview_html(file_path: &Path, heading_order: i64) -> CommandResult<String> {
    let paragraphs = parse_docx_paragraphs(file_path)?;
    let heading_ranges = build_heading_ranges(&paragraphs);
    let Some(target_range) = heading_ranges
        .iter()
        .find(|range| range.order == heading_order)
    else {
        return Ok(String::new());
    };

    let file = File::open(file_path)
        .map_err(|error| format!("Could not open '{}': {error}", path_display(file_path)))?;
    let mut archive = ZipArchive::new(file)
        .map_err(|error| format!("Could not read '{}': {error}", path_display(file_path)))?;
    let document_xml = read_zip_file(&mut archive, "word/document.xml").ok_or_else(|| {
        format!(
            "Missing word/document.xml in '{}'. Is this a valid docx file?",
            path_display(file_path)
        )
    })?;
    let document = Document::parse(&document_xml).map_err(|error| {
        format!(
            "Could not parse preview XML '{}': {error}",
            path_display(file_path)
        )
    })?;

    let paragraph_nodes = document
        .descendants()
        .filter(|node| has_tag(*node, "p"))
        .collect::<Vec<Node<'_, '_>>>();

    let start = target_range.start_index;
    let end = target_range
        .end_index
        .min(paragraph_nodes.len())
        .min(paragraphs.len());
    if start >= end {
        return Ok(String::new());
    }

    let mut html = String::new();
    for index in start..end {
        let paragraph_node = paragraph_nodes[index];
        let paragraph_meta = &paragraphs[index];
        html.push_str(&render_preview_paragraph(
            paragraph_node,
            paragraph_meta.heading_level,
            &paragraph_meta.text,
        ));
    }

    Ok(html)
}

fn detect_heading_level(
    paragraph: Node<'_, '_>,
    style_map: &HashMap<String, String>,
) -> Option<i64> {
    let paragraph_props = paragraph.children().find(|node| has_tag(*node, "pPr"))?;

    if let Some(outline_level_node) = paragraph_props
        .children()
        .find(|node| has_tag(*node, "outlineLvl"))
    {
        if let Some(raw_level) = attribute_value(outline_level_node, "val") {
            if let Ok(level_zero_based) = raw_level.parse::<i64>() {
                let level = level_zero_based + 1;
                if (1..=9).contains(&level) {
                    return Some(level);
                }
            }
        }
    }

    let style_node = paragraph_props
        .children()
        .find(|node| has_tag(*node, "pStyle"))?;
    let style_id = attribute_value(style_node, "val")?;

    if let Some(level) = parse_trailing_level(style_id) {
        return Some(level);
    }

    if let Some(style_name) = style_map.get(style_id) {
        return parse_trailing_level(style_name);
    }

    None
}

fn paragraph_style_label(
    paragraph: Node<'_, '_>,
    style_map: &HashMap<String, String>,
) -> Option<String> {
    let paragraph_props = paragraph.children().find(|node| has_tag(*node, "pPr"))?;
    let style_node = paragraph_props
        .children()
        .find(|node| has_tag(*node, "pStyle"))?;
    let style_id = attribute_value(style_node, "val")?;
    let style_name = style_map
        .get(style_id)
        .cloned()
        .unwrap_or_else(|| style_id.to_string());
    Some(format!("{style_name} ({style_id})"))
}

fn is_f8_cite_style(style_label: &str) -> bool {
    let normalized = normalize_for_search(style_label);
    normalized.contains("f8 cite") || normalized.contains("f8cite")
}

fn parse_docx_paragraphs(file_path: &Path) -> CommandResult<Vec<ParsedParagraph>> {
    let file = File::open(file_path)
        .map_err(|error| format!("Could not open '{}': {error}", path_display(file_path)))?;
    let mut archive = ZipArchive::new(file)
        .map_err(|error| format!("Could not read '{}': {error}", path_display(file_path)))?;

    let document_xml = read_zip_file(&mut archive, "word/document.xml").ok_or_else(|| {
        format!(
            "Missing word/document.xml in '{}'. Is this a valid docx file?",
            path_display(file_path)
        )
    })?;

    let style_map = read_style_map(read_zip_file(&mut archive, "word/styles.xml"));

    let document = Document::parse(&document_xml).map_err(|error| {
        format!(
            "Could not parse XML in '{}': {error}",
            path_display(file_path)
        )
    })?;

    let mut order = 0_i64;
    let mut paragraphs = Vec::new();

    for paragraph in document.descendants().filter(|node| has_tag(*node, "p")) {
        let text = extract_paragraph_text(paragraph);

        order += 1;
        let style_label = paragraph_style_label(paragraph, &style_map);
        let is_f8_cite = style_label
            .as_ref()
            .map(|label| is_f8_cite_style(label))
            .unwrap_or(false);
        let mut heading_level = detect_heading_level(paragraph, &style_map);
        if heading_level.is_some() && (is_probable_author_line(&text) || is_f8_cite) {
            heading_level = None;
        }

        paragraphs.push(ParsedParagraph {
            order,
            text,
            heading_level,
            style_label,
            is_f8_cite,
        });
    }

    Ok(paragraphs)
}

fn build_heading_ranges(paragraphs: &[ParsedParagraph]) -> Vec<HeadingRange> {
    let mut heading_indices = Vec::new();
    for (index, paragraph) in paragraphs.iter().enumerate() {
        if paragraph.heading_level.is_some() {
            heading_indices.push(index);
        }
    }

    let mut ranges = Vec::new();
    for (heading_position, start_index) in heading_indices.iter().enumerate() {
        let paragraph = &paragraphs[*start_index];
        let Some(level) = paragraph.heading_level else {
            continue;
        };

        let mut end_index = paragraphs.len();
        for candidate_index in heading_indices.iter().skip(heading_position + 1) {
            if let Some(candidate_level) = paragraphs[*candidate_index].heading_level {
                if is_probable_author_line(&paragraphs[*candidate_index].text) {
                    continue;
                }
                if candidate_level <= level {
                    end_index = *candidate_index;
                    break;
                }
            }
        }

        ranges.push(HeadingRange {
            order: paragraph.order,
            level,
            start_index: *start_index,
            end_index,
        });
    }

    ranges
}

fn resolve_insert_after_order(
    paragraphs: &[ParsedParagraph],
    selected_target_heading_order: Option<i64>,
    incoming_heading_level: Option<i64>,
) -> Option<i64> {
    let heading_ranges = build_heading_ranges(paragraphs);
    if heading_ranges.is_empty() {
        return None;
    }

    let end_order = |range: &HeadingRange| {
        paragraphs
            .get(range.end_index.saturating_sub(1))
            .map(|paragraph| paragraph.order)
    };

    if let Some(selected_order) = selected_target_heading_order {
        if let Some(selected_range) = heading_ranges
            .iter()
            .find(|range| range.order == selected_order)
        {
            if let Some(incoming_level) = incoming_heading_level {
                if incoming_level < selected_range.level {
                    let mut ancestor_match: Option<&HeadingRange> = None;
                    for candidate in &heading_ranges {
                        if candidate.start_index >= selected_range.start_index {
                            break;
                        }
                        if candidate.level < incoming_level
                            && candidate.end_index > selected_range.start_index
                        {
                            ancestor_match = Some(candidate);
                        }
                    }

                    if let Some(ancestor) = ancestor_match {
                        return end_order(ancestor);
                    }

                    if let Some(last_at_or_above) = heading_ranges
                        .iter()
                        .rev()
                        .find(|range| range.level <= incoming_level)
                    {
                        return end_order(last_at_or_above);
                    }
                }
            }

            return end_order(selected_range);
        }
    }

    if let Some(incoming_level) = incoming_heading_level {
        if let Some(last_same_level) = heading_ranges
            .iter()
            .rev()
            .find(|range| range.level == incoming_level)
        {
            return end_order(last_same_level);
        }

        if let Some(last_parent_level) = heading_ranges
            .iter()
            .rev()
            .find(|range| range.level < incoming_level)
        {
            return end_order(last_parent_level);
        }
    }

    heading_ranges.last().and_then(end_order)
}

fn extract_preview_content(
    file_path: &Path,
) -> CommandResult<(Vec<FileHeading>, Vec<TaggedBlock>)> {
    let paragraphs = parse_docx_paragraphs(file_path)?;

    let mut heading_indices = Vec::new();
    for (index, paragraph) in paragraphs.iter().enumerate() {
        if paragraph.heading_level.is_some() {
            heading_indices.push(index);
        }
    }

    let mut headings = Vec::new();
    for (heading_position, start_index) in heading_indices.iter().enumerate() {
        let paragraph = &paragraphs[*start_index];
        let Some(level) = paragraph.heading_level else {
            continue;
        };

        let mut end_index = paragraphs.len();
        for candidate_index in heading_indices.iter().skip(heading_position + 1) {
            if let Some(candidate_level) = paragraphs[*candidate_index].heading_level {
                if is_probable_author_line(&paragraphs[*candidate_index].text) {
                    continue;
                }
                if candidate_level <= level {
                    end_index = *candidate_index;
                    break;
                }
            }
        }

        let section_lines = paragraphs[*start_index..end_index]
            .iter()
            .map(|entry| entry.text.as_str())
            .collect::<Vec<&str>>();
        let copy_text = section_lines.join("\n");

        headings.push(FileHeading {
            id: paragraph.order,
            order: paragraph.order,
            level,
            text: paragraph.text.clone(),
            copy_text,
        });
    }

    let mut f8_cites = Vec::new();
    let mut cursor = 0_usize;
    while cursor < paragraphs.len() {
        let paragraph = &paragraphs[cursor];
        if !paragraph.is_f8_cite {
            cursor += 1;
            continue;
        }

        let start_order = paragraph.order;
        let style_label = paragraph
            .style_label
            .clone()
            .unwrap_or_else(|| "F8 Cite".to_string());
        let mut lines = vec![paragraph.text.clone()];

        cursor += 1;
        while cursor < paragraphs.len() && paragraphs[cursor].is_f8_cite {
            lines.push(paragraphs[cursor].text.clone());
            cursor += 1;
        }

        let text = lines.join("\n");
        if text.trim().is_empty() {
            continue;
        }

        f8_cites.push(TaggedBlock {
            order: start_order,
            style_label,
            text,
        });
    }

    Ok((headings, f8_cites))
}

fn extract_docx_headings_and_authors(
    file_path: &Path,
) -> CommandResult<(Vec<ParsedHeading>, Vec<(i64, String)>)> {
    let paragraphs = parse_docx_paragraphs(file_path)?;
    let mut headings = Vec::new();

    for paragraph in &paragraphs {
        let Some(level) = paragraph.heading_level else {
            continue;
        };
        headings.push(ParsedHeading {
            order: paragraph.order,
            level,
            text: paragraph.text.clone(),
        });
    }

    let authors = extract_author_candidates(&paragraphs);
    Ok((headings, authors))
}

fn root_id(connection: &Connection, root_path: &str) -> CommandResult<Option<i64>> {
    connection
        .query_row(
            "SELECT id FROM roots WHERE path = ?1",
            params![root_path],
            |row| row.get(0),
        )
        .optional()
        .map_err(|error| format!("Could not query root path '{root_path}': {error}"))
}

fn add_or_get_root_id(connection: &Connection, root_path: &str) -> CommandResult<i64> {
    connection
        .execute(
            "INSERT INTO roots(path, added_at_ms, last_indexed_ms) VALUES(?1, ?2, 0)
             ON CONFLICT(path) DO NOTHING",
            params![root_path, now_ms()],
        )
        .map_err(|error| format!("Could not store root path '{root_path}': {error}"))?;

    root_id(connection, root_path)?
        .ok_or_else(|| format!("Could not find root row for '{root_path}'"))
}

fn load_existing_files(
    connection: &Connection,
    root_id: i64,
) -> CommandResult<HashMap<String, ExistingFileMeta>> {
    let mut statement = connection
        .prepare("SELECT id, relative_path, modified_ms, size FROM files WHERE root_id = ?1")
        .map_err(|error| format!("Could not prepare file metadata query: {error}"))?;

    let rows = statement
        .query_map(params![root_id], |row| {
            Ok((
                row.get::<_, i64>(0)?,
                row.get::<_, String>(1)?,
                row.get::<_, i64>(2)?,
                row.get::<_, i64>(3)?,
            ))
        })
        .map_err(|error| format!("Could not iterate existing files: {error}"))?;

    let mut metadata = HashMap::new();
    for row in rows {
        let (id, relative_path, modified_ms, size) =
            row.map_err(|error| format!("Could not parse existing file metadata row: {error}"))?;
        metadata.insert(
            relative_path,
            ExistingFileMeta {
                id,
                modified_ms,
                size,
            },
        );
    }

    Ok(metadata)
}

#[tauri::command]
fn add_root(app: AppHandle, path: String) -> CommandResult<String> {
    let canonical = canonicalize_folder(&path)?;
    let canonical_string = path_display(&canonical);

    let connection = open_database(&app)?;
    add_or_get_root_id(&connection, &canonical_string)?;
    write_root_index_marker(&canonical, 0)?;
    Ok(canonical_string)
}

#[tauri::command]
fn remove_root(app: AppHandle, path: String) -> CommandResult<()> {
    let canonical_path = canonicalize_folder(&path).ok();
    let canonical_string = canonical_path
        .as_ref()
        .map(|path| path_display(path))
        .unwrap_or(path);
    let connection = open_database(&app)?;
    connection
        .execute(
            "DELETE FROM roots WHERE path = ?1",
            params![canonical_string],
        )
        .map_err(|error| format!("Could not remove root: {error}"))?;

    if let Some(root_path) = canonical_path {
        let marker_path = root_index_marker_path(&root_path);
        let _ = fs::remove_file(marker_path);
    }
    Ok(())
}

#[tauri::command]
fn insert_capture(
    app: AppHandle,
    root_path: String,
    source_path: String,
    section_title: String,
    content: String,
    target_path: Option<String>,
    heading_level: Option<i64>,
    heading_order: Option<i64>,
    selected_target_heading_order: Option<i64>,
) -> CommandResult<CaptureInsertResult> {
    let content_value = content;
    if content_value.trim().is_empty() {
        return Err("Cannot insert empty content into capture file.".to_string());
    }

    let canonical_root = canonicalize_folder(&root_path)?;
    let target_relative_path = normalize_capture_target_path(target_path.as_deref())?;
    let normalized_heading_level = heading_level.filter(|level| (1..=9).contains(level));
    let normalized_target_heading_order = selected_target_heading_order.filter(|value| *value > 0);
    let root_path_string = path_display(&canonical_root);
    let connection = open_database(&app)?;
    let root_id = add_or_get_root_id(&connection, &root_path_string)?;

    let created_at_ms = now_ms();
    connection
        .execute(
            "
            INSERT INTO captures(
              root_id,
              source_path,
              section_title,
              target_relative_path,
              heading_level,
              content,
              created_at_ms
            )
            VALUES(?1, ?2, ?3, ?4, ?5, ?6, ?7)
            ",
            params![
                root_id,
                &source_path,
                &section_title,
                &target_relative_path,
                normalized_heading_level,
                &content_value,
                created_at_ms
            ],
        )
        .map_err(|error| format!("Could not insert capture entry: {error}"))?;

    let capture_id = connection.last_insert_rowid();
    let capture_path = capture_docx_path(&canonical_root, &target_relative_path);
    let source_file_path = Path::new(&source_path);
    let styled_section = extract_styled_section(source_file_path, heading_order, &content_value);
    append_capture_to_docx(
        &capture_path,
        source_file_path,
        normalized_heading_level,
        normalized_target_heading_order,
        &styled_section,
    )?;

    Ok(CaptureInsertResult {
        capture_path: path_display(&capture_path),
        marker: capture_marker(capture_id),
        target_relative_path,
    })
}

#[tauri::command]
fn list_capture_targets(app: AppHandle, root_path: String) -> CommandResult<Vec<CaptureTarget>> {
    let canonical_root = canonicalize_folder(&root_path)?;
    let root_path_string = path_display(&canonical_root);
    let connection = open_database(&app)?;
    let root_id = add_or_get_root_id(&connection, &root_path_string)?;

    let mut by_target = HashMap::<String, i64>::new();
    by_target.insert(DEFAULT_CAPTURE_TARGET.to_string(), 0);

    let mut statement = connection
        .prepare(
            "
            SELECT target_relative_path, COUNT(*)
            FROM captures
            WHERE root_id = ?1
            GROUP BY target_relative_path
            ORDER BY target_relative_path ASC
            ",
        )
        .map_err(|error| format!("Could not prepare capture targets query: {error}"))?;

    let rows = statement
        .query_map(params![root_id], |row| {
            Ok((row.get::<_, String>(0)?, row.get::<_, i64>(1)?))
        })
        .map_err(|error| format!("Could not iterate capture targets query: {error}"))?;

    for row in rows {
        let (target, count) =
            row.map_err(|error| format!("Could not parse capture target row: {error}"))?;
        by_target.insert(target, count);
    }

    let mut targets = by_target
        .into_iter()
        .map(|(relative_path, entry_count)| {
            let absolute_path = capture_docx_path(&canonical_root, &relative_path);
            CaptureTarget {
                relative_path,
                absolute_path: path_display(&absolute_path),
                exists: absolute_path.is_file(),
                entry_count,
            }
        })
        .collect::<Vec<CaptureTarget>>();

    targets.sort_by(|left, right| {
        (left.relative_path != DEFAULT_CAPTURE_TARGET)
            .cmp(&(right.relative_path != DEFAULT_CAPTURE_TARGET))
            .then(left.relative_path.cmp(&right.relative_path))
    });

    Ok(targets)
}

fn capture_target_preview_for_path(
    canonical_root: &Path,
    normalized_target: &str,
) -> CaptureTargetPreview {
    let absolute_path = capture_docx_path(canonical_root, normalized_target);

    if !absolute_path.is_file() {
        return CaptureTargetPreview {
            relative_path: normalized_target.to_string(),
            absolute_path: path_display(&absolute_path),
            exists: false,
            heading_count: 0,
            headings: Vec::new(),
        };
    }

    let (mut headings, _) = extract_preview_content(&absolute_path).unwrap_or_default();
    headings.sort_by(|left, right| left.order.cmp(&right.order));

    CaptureTargetPreview {
        relative_path: normalized_target.to_string(),
        absolute_path: path_display(&absolute_path),
        exists: true,
        heading_count: i64::try_from(headings.len()).unwrap_or(0),
        headings,
    }
}

#[tauri::command]
fn get_capture_target_preview(
    _app: AppHandle,
    root_path: String,
    target_path: String,
) -> CommandResult<CaptureTargetPreview> {
    let canonical_root = canonicalize_folder(&root_path)?;
    let normalized_target = normalize_capture_target_path(Some(&target_path))?;
    Ok(capture_target_preview_for_path(
        &canonical_root,
        &normalized_target,
    ))
}

#[tauri::command]
fn delete_capture_heading(
    _app: AppHandle,
    root_path: String,
    target_path: String,
    heading_order: i64,
) -> CommandResult<CaptureTargetPreview> {
    let canonical_root = canonicalize_folder(&root_path)?;
    let normalized_target = normalize_capture_target_path(Some(&target_path))?;
    let absolute_path = capture_docx_path(&canonical_root, &normalized_target);

    if !absolute_path.is_file() {
        return Err(format!(
            "Target capture file does not exist: {}",
            path_display(&absolute_path)
        ));
    }

    ensure_valid_capture_docx(&absolute_path)?;
    let paragraphs = parse_docx_paragraphs(&absolute_path)?;
    let heading_ranges = build_heading_ranges(&paragraphs);
    let target_range = heading_ranges
        .iter()
        .find(|range| range.order == heading_order)
        .cloned()
        .ok_or_else(|| format!("Heading order {heading_order} not found in target document."))?;

    let document_xml = read_docx_part(&absolute_path, "word/document.xml")?.ok_or_else(|| {
        format!(
            "Missing word/document.xml in '{}'",
            path_display(&absolute_path)
        )
    })?;
    let document = Document::parse(&document_xml).map_err(|error| {
        format!(
            "Could not parse destination document XML '{}': {error}",
            path_display(&absolute_path)
        )
    })?;
    let paragraph_nodes = document
        .descendants()
        .filter(|node| has_tag(*node, "p"))
        .collect::<Vec<Node<'_, '_>>>();

    if target_range.start_index >= paragraph_nodes.len()
        || target_range.end_index == 0
        || target_range.end_index > paragraph_nodes.len()
    {
        return Err("Heading range is out of bounds in destination document.".to_string());
    }

    let start = paragraph_nodes[target_range.start_index].range().start;
    let end = paragraph_nodes[target_range.end_index - 1].range().end;
    if start >= end || end > document_xml.len() {
        return Err("Could not resolve heading XML range in destination document.".to_string());
    }

    let mut updated_document_xml =
        String::with_capacity(document_xml.len().saturating_sub(end.saturating_sub(start)));
    updated_document_xml.push_str(&document_xml[..start]);
    updated_document_xml.push_str(&document_xml[end..]);

    let mut replacements = HashMap::new();
    replacements.insert(
        "word/document.xml".to_string(),
        updated_document_xml.into_bytes(),
    );
    rewrite_docx_with_parts(&absolute_path, &replacements)?;

    Ok(capture_target_preview_for_path(
        &canonical_root,
        &normalized_target,
    ))
}

#[tauri::command]
fn move_capture_heading(
    _app: AppHandle,
    root_path: String,
    target_path: String,
    source_heading_order: i64,
    target_heading_order: i64,
) -> CommandResult<CaptureTargetPreview> {
    let canonical_root = canonicalize_folder(&root_path)?;
    let normalized_target = normalize_capture_target_path(Some(&target_path))?;
    let absolute_path = capture_docx_path(&canonical_root, &normalized_target);

    if source_heading_order == target_heading_order {
        return Ok(capture_target_preview_for_path(
            &canonical_root,
            &normalized_target,
        ));
    }

    if !absolute_path.is_file() {
        return Err(format!(
            "Target capture file does not exist: {}",
            path_display(&absolute_path)
        ));
    }

    ensure_valid_capture_docx(&absolute_path)?;
    let paragraphs = parse_docx_paragraphs(&absolute_path)?;
    let heading_ranges = build_heading_ranges(&paragraphs);

    let source_range = heading_ranges
        .iter()
        .find(|range| range.order == source_heading_order)
        .cloned()
        .ok_or_else(|| {
            format!("Source heading order {source_heading_order} not found in target document.")
        })?;
    let target_range = heading_ranges
        .iter()
        .find(|range| range.order == target_heading_order)
        .cloned()
        .ok_or_else(|| {
            format!("Target heading order {target_heading_order} not found in target document.")
        })?;

    if target_range.start_index >= source_range.start_index
        && target_range.start_index < source_range.end_index
    {
        return Err("Cannot move a heading into its own subtree.".to_string());
    }

    let document_xml = read_docx_part(&absolute_path, "word/document.xml")?.ok_or_else(|| {
        format!(
            "Missing word/document.xml in '{}'",
            path_display(&absolute_path)
        )
    })?;
    let document = Document::parse(&document_xml).map_err(|error| {
        format!(
            "Could not parse destination document XML '{}': {error}",
            path_display(&absolute_path)
        )
    })?;
    let paragraph_nodes = document
        .descendants()
        .filter(|node| has_tag(*node, "p"))
        .collect::<Vec<Node<'_, '_>>>();

    if source_range.start_index >= paragraph_nodes.len()
        || source_range.end_index == 0
        || source_range.end_index > paragraph_nodes.len()
        || target_range.start_index >= paragraph_nodes.len()
        || target_range.end_index == 0
        || target_range.end_index > paragraph_nodes.len()
    {
        return Err("Heading range is out of bounds in destination document.".to_string());
    }

    let source_start = paragraph_nodes[source_range.start_index].range().start;
    let source_end = paragraph_nodes[source_range.end_index - 1].range().end;
    if source_start >= source_end || source_end > document_xml.len() {
        return Err("Could not resolve source heading XML range.".to_string());
    }

    let moved_fragment = document_xml[source_start..source_end].to_string();
    let mut without_source =
        String::with_capacity(document_xml.len() - (source_end - source_start));
    without_source.push_str(&document_xml[..source_start]);
    without_source.push_str(&document_xml[source_end..]);

    let source_len = source_range
        .end_index
        .saturating_sub(source_range.start_index);
    let mut insertion_paragraph_count = target_range.end_index;
    if source_range.start_index < target_range.end_index {
        insertion_paragraph_count = insertion_paragraph_count.saturating_sub(source_len);
    }

    let insertion_index =
        insertion_index_after_paragraph_count(&without_source, insertion_paragraph_count)
            .unwrap_or(fallback_body_insertion_index(&without_source)?);

    let mut updated_document_xml =
        String::with_capacity(without_source.len().saturating_add(moved_fragment.len()));
    updated_document_xml.push_str(&without_source[..insertion_index]);
    updated_document_xml.push_str(&moved_fragment);
    updated_document_xml.push_str(&without_source[insertion_index..]);

    let mut replacements = HashMap::new();
    replacements.insert(
        "word/document.xml".to_string(),
        updated_document_xml.into_bytes(),
    );
    rewrite_docx_with_parts(&absolute_path, &replacements)?;

    Ok(capture_target_preview_for_path(
        &canonical_root,
        &normalized_target,
    ))
}

#[tauri::command]
fn list_roots(app: AppHandle) -> CommandResult<Vec<RootSummary>> {
    let connection = open_database(&app)?;
    let mut statement = connection
        .prepare(
            "
            SELECT
              r.path,
              r.added_at_ms,
              r.last_indexed_ms,
              (SELECT COUNT(*) FROM files f WHERE f.root_id = r.id) AS file_count,
              (
                SELECT COUNT(*)
                FROM headings h
                JOIN files f ON f.id = h.file_id
                WHERE f.root_id = r.id
              ) AS heading_count
            FROM roots r
            ORDER BY r.path
            ",
        )
        .map_err(|error| format!("Could not prepare roots query: {error}"))?;

    let rows = statement
        .query_map([], |row| {
            Ok(RootSummary {
                path: row.get(0)?,
                added_at_ms: row.get(1)?,
                last_indexed_ms: row.get(2)?,
                file_count: row.get(3)?,
                heading_count: row.get(4)?,
            })
        })
        .map_err(|error| format!("Could not iterate roots query: {error}"))?;

    let mut roots = Vec::new();
    for row in rows {
        roots.push(row.map_err(|error| format!("Could not parse roots row: {error}"))?);
    }

    Ok(roots)
}

#[tauri::command]
fn index_root(app: AppHandle, path: String) -> CommandResult<IndexStats> {
    let started_at = now_ms();
    let canonical_root = canonicalize_folder(&path)?;
    let root_path = path_display(&canonical_root);

    let mut connection = open_database(&app)?;
    let root_id = add_or_get_root_id(&connection, &root_path)?;
    let existing_files = load_existing_files(&connection, root_id)?;

    let mut scanned = 0_usize;
    let mut updated = 0_usize;
    let mut skipped = 0_usize;
    let mut removed = 0_usize;
    let mut headings_extracted = 0_usize;
    let mut seen_relative_paths = HashSet::new();
    let mut indexing_candidates = Vec::new();

    let mut progress = IndexProgress {
        root_path: root_path.clone(),
        phase: "discovering".to_string(),
        discovered: 0,
        changed: 0,
        processed: 0,
        updated: 0,
        skipped: 0,
        removed: 0,
        elapsed_ms: 0,
        current_file: None,
    };
    let mut last_progress_emit_ms = 0_i64;
    emit_index_progress(
        &app,
        started_at,
        &progress,
        &mut last_progress_emit_ms,
        true,
    );

    for entry in WalkDir::new(&canonical_root)
        .follow_links(false)
        .into_iter()
        .filter_entry(is_visible_entry)
    {
        let Ok(entry) = entry else {
            continue;
        };

        if !entry.file_type().is_file() {
            continue;
        }

        let is_docx = entry
            .path()
            .extension()
            .and_then(|extension| extension.to_str())
            .map(|extension| extension.eq_ignore_ascii_case("docx"))
            .unwrap_or(false);
        if !is_docx {
            continue;
        }

        scanned += 1;
        let absolute_path = entry.path().to_path_buf();
        let relative_path = relative_path(&canonical_root, &absolute_path)?;
        seen_relative_paths.insert(relative_path.clone());

        let metadata = fs::metadata(&absolute_path).map_err(|error| {
            format!(
                "Could not read metadata for '{}': {error}",
                path_display(&absolute_path)
            )
        })?;
        let modified_ms = metadata.modified().map(epoch_ms).unwrap_or(0);
        let size = i64::try_from(metadata.len()).unwrap_or(0);

        if let Some(existing) = existing_files.get(&relative_path) {
            if existing.modified_ms == modified_ms && existing.size == size {
                skipped += 1;
            } else {
                indexing_candidates.push(IndexCandidate {
                    relative_path: relative_path.clone(),
                    absolute_path,
                    modified_ms,
                    size,
                });
            }
        } else {
            indexing_candidates.push(IndexCandidate {
                relative_path: relative_path.clone(),
                absolute_path,
                modified_ms,
                size,
            });
        }

        progress.discovered = scanned;
        progress.changed = indexing_candidates.len();
        progress.skipped = skipped;
        progress.current_file = Some(relative_path);
        emit_index_progress(
            &app,
            started_at,
            &progress,
            &mut last_progress_emit_ms,
            false,
        );
    }

    let stale_entries = existing_files
        .iter()
        .filter_map(|(relative_path, existing)| {
            (!seen_relative_paths.contains(relative_path))
                .then_some((relative_path.clone(), existing.id))
        })
        .collect::<Vec<(String, i64)>>();

    progress.phase = "indexing".to_string();
    progress.current_file = None;
    progress.discovered = scanned;
    progress.changed = indexing_candidates.len();
    progress.skipped = skipped;
    emit_index_progress(
        &app,
        started_at,
        &progress,
        &mut last_progress_emit_ms,
        true,
    );

    let parse_chunk_size = suggested_parse_chunk_size();
    let transaction = connection
        .transaction()
        .map_err(|error| format!("Could not start index transaction: {error}"))?;

    for chunk in indexing_candidates.chunks(parse_chunk_size) {
        let parsed_chunk = chunk
            .par_iter()
            .map(|candidate| {
                let (headings, authors) =
                    extract_docx_headings_and_authors(&candidate.absolute_path).unwrap_or_default();
                ParsedIndexCandidate {
                    candidate: candidate.clone(),
                    headings,
                    authors,
                }
            })
            .collect::<Vec<ParsedIndexCandidate>>();

        for parsed in parsed_chunk {
            let relative_path = parsed.candidate.relative_path;
            let absolute_path_string = path_display(&parsed.candidate.absolute_path);
            let modified_ms = parsed.candidate.modified_ms;
            let size = parsed.candidate.size;
            let heading_count = i64::try_from(parsed.headings.len()).unwrap_or(0);
            headings_extracted += parsed.headings.len();

            let file_name = file_name_from_relative(&relative_path);

            let file_id = if let Some(existing) = existing_files.get(&relative_path) {
                transaction
                    .execute(
                        "UPDATE files
                         SET absolute_path = ?1, modified_ms = ?2, size = ?3, heading_count = ?4
                         WHERE id = ?5",
                        params![
                            absolute_path_string,
                            modified_ms,
                            size,
                            heading_count,
                            existing.id
                        ],
                    )
                    .map_err(|error| {
                        format!("Could not update indexed file '{}': {error}", relative_path)
                    })?;
                existing.id
            } else {
                transaction
                    .execute(
                        "INSERT INTO files(root_id, relative_path, absolute_path, modified_ms, size, heading_count)
                         VALUES(?1, ?2, ?3, ?4, ?5, ?6)",
                        params![
                            root_id,
                            relative_path.as_str(),
                            absolute_path_string,
                            modified_ms,
                            size,
                            heading_count
                        ],
                    )
                    .map_err(|error| {
                        format!("Could not insert indexed file '{}': {error}", relative_path)
                    })?;
                transaction.last_insert_rowid()
            };

            transaction
                .execute("DELETE FROM headings WHERE file_id = ?1", params![file_id])
                .map_err(|error| {
                    format!(
                        "Could not clear old headings for '{}': {error}",
                        relative_path
                    )
                })?;

            transaction
                .execute("DELETE FROM authors WHERE file_id = ?1", params![file_id])
                .map_err(|error| {
                    format!(
                        "Could not clear old author rows for '{}': {error}",
                        relative_path
                    )
                })?;

            for heading in parsed.headings {
                let normalized = normalize_for_search(&heading.text);
                transaction
                    .execute(
                        "INSERT INTO headings(file_id, heading_order, level, text, normalized, file_name, relative_path)
                         VALUES(?1, ?2, ?3, ?4, ?5, ?6, ?7)",
                        params![
                            file_id,
                            heading.order,
                            heading.level,
                            heading.text,
                            normalized,
                            file_name.as_str(),
                            relative_path.as_str()
                        ],
                    )
                    .map_err(|error| {
                        format!("Could not insert heading for '{}': {error}", relative_path)
                    })?;
            }

            for (author_order, author_text) in parsed.authors {
                let normalized_author = normalize_for_search(&author_text);
                transaction
                    .execute(
                        "INSERT INTO authors(file_id, author_order, text, normalized, file_name, relative_path)
                         VALUES(?1, ?2, ?3, ?4, ?5, ?6)",
                        params![
                            file_id,
                            author_order,
                            author_text,
                            normalized_author,
                            file_name.as_str(),
                            relative_path.as_str()
                        ],
                    )
                    .map_err(|error| {
                        format!(
                            "Could not insert author metadata for '{}': {error}",
                            relative_path
                        )
                    })?;
            }

            updated += 1;
            progress.processed = updated;
            progress.updated = updated;
            progress.current_file = Some(relative_path);
            emit_index_progress(
                &app,
                started_at,
                &progress,
                &mut last_progress_emit_ms,
                false,
            );
        }
    }

    progress.phase = "cleaning".to_string();
    progress.current_file = None;
    emit_index_progress(
        &app,
        started_at,
        &progress,
        &mut last_progress_emit_ms,
        true,
    );

    for (relative_path, file_id) in stale_entries {
        transaction
            .execute("DELETE FROM files WHERE id = ?1", params![file_id])
            .map_err(|error| {
                format!(
                    "Could not remove stale index row '{}': {error}",
                    relative_path
                )
            })?;
        removed += 1;

        progress.removed = removed;
        progress.current_file = Some(relative_path);
        emit_index_progress(
            &app,
            started_at,
            &progress,
            &mut last_progress_emit_ms,
            false,
        );
    }

    let finished_at_ms = now_ms();

    transaction
        .execute(
            "UPDATE roots SET last_indexed_ms = ?1 WHERE id = ?2",
            params![finished_at_ms, root_id],
        )
        .map_err(|error| format!("Could not update root index timestamp: {error}"))?;

    transaction
        .commit()
        .map_err(|error| format!("Could not commit index transaction: {error}"))?;

    write_root_index_marker(&canonical_root, finished_at_ms)?;

    progress.phase = "complete".to_string();
    progress.current_file = None;
    progress.discovered = scanned;
    progress.changed = indexing_candidates.len();
    progress.processed = updated;
    progress.updated = updated;
    progress.skipped = skipped;
    progress.removed = removed;
    emit_index_progress(
        &app,
        started_at,
        &progress,
        &mut last_progress_emit_ms,
        true,
    );

    Ok(IndexStats {
        scanned,
        updated,
        skipped,
        removed,
        headings_extracted,
        elapsed_ms: finished_at_ms - started_at,
    })
}

fn ensure_folder_with_ancestors(folders: &mut HashMap<String, FolderEntry>, folder_path: &str) {
    let mut current = folder_path.to_string();

    loop {
        if !folders.contains_key(&current) {
            let parent_path = current
                .rsplit_once('/')
                .map(|(parent, _)| parent.to_string());
            let name = if current.is_empty() {
                "Root".to_string()
            } else {
                current
                    .rsplit_once('/')
                    .map(|(_, name)| name.to_string())
                    .unwrap_or_else(|| current.clone())
            };
            let depth = if current.is_empty() {
                0
            } else {
                current.split('/').count()
            };

            folders.insert(
                current.clone(),
                FolderEntry {
                    path: current.clone(),
                    name,
                    parent_path,
                    depth,
                    file_count: 0,
                },
            );
        }

        if current.is_empty() {
            break;
        }

        current = current
            .rsplit_once('/')
            .map(|(parent, _)| parent.to_string())
            .unwrap_or_default();
    }
}

#[tauri::command]
fn get_index_snapshot(app: AppHandle, path: String) -> CommandResult<IndexSnapshot> {
    let canonical_path = canonicalize_folder(&path)
        .map(|canonical| path_display(&canonical))
        .unwrap_or(path);

    let connection = open_database(&app)?;
    let root_id = root_id(&connection, &canonical_path)?.ok_or_else(|| {
        format!(
            "No index found for '{}'. Add the folder first.",
            canonical_path
        )
    })?;

    let indexed_at_ms = connection
        .query_row(
            "SELECT last_indexed_ms FROM roots WHERE id = ?1",
            params![root_id],
            |row| row.get::<_, i64>(0),
        )
        .map_err(|error| format!("Could not read root timestamp: {error}"))?;

    let mut statement = connection
        .prepare(
            "
            SELECT id, relative_path, modified_ms, heading_count
            FROM files
            WHERE root_id = ?1
            ORDER BY relative_path
            ",
        )
        .map_err(|error| format!("Could not prepare file snapshot query: {error}"))?;

    let rows = statement
        .query_map(params![root_id], |row| {
            Ok(FileRecord {
                id: row.get(0)?,
                relative_path: row.get(1)?,
                modified_ms: row.get(2)?,
                heading_count: row.get(3)?,
            })
        })
        .map_err(|error| format!("Could not iterate indexed files: {error}"))?;

    let mut files = Vec::new();
    let mut folders = HashMap::new();
    ensure_folder_with_ancestors(&mut folders, "");

    for row in rows {
        let record = row.map_err(|error| format!("Could not parse indexed file row: {error}"))?;
        let folder_path = folder_from_relative(&record.relative_path);
        ensure_folder_with_ancestors(&mut folders, &folder_path);

        let mut current_folder = folder_path.clone();
        loop {
            if let Some(folder_entry) = folders.get_mut(&current_folder) {
                folder_entry.file_count += 1;
            }

            if current_folder.is_empty() {
                break;
            }

            current_folder = current_folder
                .rsplit_once('/')
                .map(|(parent, _)| parent.to_string())
                .unwrap_or_default();
        }

        files.push(IndexedFile {
            id: record.id,
            file_name: file_name_from_relative(&record.relative_path),
            relative_path: record.relative_path,
            folder_path,
            modified_ms: record.modified_ms,
            heading_count: record.heading_count,
        });
    }

    let mut folder_values = folders.into_values().collect::<Vec<FolderEntry>>();
    folder_values.sort_by(|left, right| {
        left.depth
            .cmp(&right.depth)
            .then(left.path.cmp(&right.path))
    });

    Ok(IndexSnapshot {
        root_path: canonical_path,
        indexed_at_ms,
        folders: folder_values,
        files,
    })
}

#[tauri::command]
fn get_file_preview(app: AppHandle, file_id: i64) -> CommandResult<FilePreview> {
    let connection = open_database(&app)?;

    let (relative_path, absolute_path, heading_count) = connection
        .query_row(
            "SELECT relative_path, absolute_path, heading_count FROM files WHERE id = ?1",
            params![file_id],
            |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, i64>(2)?,
                ))
            },
        )
        .map_err(|error| format!("Could not load file preview metadata: {error}"))?;
    let (mut headings, mut f8_cites) =
        extract_preview_content(Path::new(&absolute_path)).unwrap_or_default();

    headings.sort_by(|left, right| left.order.cmp(&right.order));
    f8_cites.sort_by(|left, right| left.order.cmp(&right.order));

    Ok(FilePreview {
        file_id,
        file_name: file_name_from_relative(&relative_path),
        relative_path,
        absolute_path,
        heading_count: i64::try_from(headings.len()).unwrap_or(heading_count),
        headings,
        f8_cites,
    })
}

#[tauri::command]
fn get_heading_preview_html(
    app: AppHandle,
    file_id: i64,
    heading_order: i64,
) -> CommandResult<String> {
    if heading_order <= 0 {
        return Ok(String::new());
    }

    let connection = open_database(&app)?;
    let absolute_path = connection
        .query_row(
            "SELECT absolute_path FROM files WHERE id = ?1",
            params![file_id],
            |row| row.get::<_, String>(0),
        )
        .map_err(|error| format!("Could not load heading preview source file: {error}"))?;

    extract_heading_preview_html(Path::new(&absolute_path), heading_order)
}

#[tauri::command]
fn search_index(
    app: AppHandle,
    query: String,
    root_path: Option<String>,
    limit: Option<usize>,
) -> CommandResult<Vec<SearchHit>> {
    let cleaned_query = query.trim();
    if cleaned_query.len() < 2 {
        return Ok(Vec::new());
    }
    let normalized_query = normalize_for_search(cleaned_query);
    if normalized_query.is_empty() {
        return Ok(Vec::new());
    }

    let fts_query = tokenize_for_fts(cleaned_query);
    if fts_query.is_empty() {
        return Ok(Vec::new());
    }

    let connection = open_database(&app)?;
    let requested_root_id = if let Some(root) = root_path {
        let canonical = canonicalize_folder(&root)
            .map(|path| path_display(&path))
            .unwrap_or(root);
        root_id(&connection, &canonical)?
    } else {
        None
    };

    let max_results = i64::try_from(limit.unwrap_or(120))
        .unwrap_or(120)
        .clamp(10, 400);
    let max_results_usize = usize::try_from(max_results).unwrap_or(120);
    let mut results = Vec::new();
    let mut seen_file_ids = HashSet::new();
    let mut seen_heading_keys = HashSet::new();
    let mut seen_author_keys = HashSet::new();

    let heading_key = |hit: &SearchHit| -> Option<String> {
        let heading_text = hit.heading_text.as_ref()?;
        let heading_level = hit.heading_level.unwrap_or(0);
        let heading_order = hit.heading_order.unwrap_or(0);
        Some(format!(
            "{}:{heading_level}:{heading_order}:{}",
            hit.file_id, heading_text
        ))
    };

    let author_key = |hit: &SearchHit| -> Option<String> {
        let author_text = hit.heading_text.as_ref()?;
        let author_order = hit.heading_order.unwrap_or(0);
        Some(format!("{}:{author_order}:{author_text}", hit.file_id))
    };

    {
        let mut statement = connection
            .prepare(
                "
                SELECT
                  f.id,
                  f.relative_path,
                  f.absolute_path,
                  h.level,
                  h.text,
                  h.heading_order,
                  bm25(search_fts, 12.0, 6.0, 1.5, 1.0) AS score
                FROM search_fts
                JOIN headings h ON h.id = search_fts.rowid
                JOIN files f ON f.id = h.file_id
                WHERE search_fts MATCH ?1
                  AND (?2 IS NULL OR f.root_id = ?2)
                ORDER BY score
                LIMIT ?3
                ",
            )
            .map_err(|error| format!("Could not prepare heading search query: {error}"))?;

        let rows = statement
            .query_map(params![fts_query, requested_root_id, max_results], |row| {
                let file_id: i64 = row.get(0)?;
                let relative_path: String = row.get(1)?;
                Ok(SearchHit {
                    kind: "heading".to_string(),
                    file_id,
                    file_name: file_name_from_relative(&relative_path),
                    relative_path,
                    absolute_path: row.get(2)?,
                    heading_level: row.get(3)?,
                    heading_text: row.get(4)?,
                    heading_order: row.get(5)?,
                    score: row.get(6)?,
                })
            })
            .map_err(|error| format!("Could not run heading search query: {error}"))?;

        for row in rows {
            let result =
                row.map_err(|error| format!("Could not parse heading search row: {error}"))?;
            seen_file_ids.insert(result.file_id);
            if let Some(key) = heading_key(&result) {
                seen_heading_keys.insert(key);
            }
            results.push(result);
        }
    }

    {
        let mut statement = connection
            .prepare(
                "
                SELECT
                  f.id,
                  f.relative_path,
                  f.absolute_path,
                  a.text,
                  a.author_order,
                  bm25(author_fts, 16.0, 7.0, 1.5, 1.0) AS score
                FROM author_fts
                JOIN authors a ON a.id = author_fts.rowid
                JOIN files f ON f.id = a.file_id
                WHERE author_fts MATCH ?1
                  AND (?2 IS NULL OR f.root_id = ?2)
                ORDER BY score
                LIMIT ?3
                ",
            )
            .map_err(|error| format!("Could not prepare author search query: {error}"))?;

        let rows = statement
            .query_map(params![fts_query, requested_root_id, max_results], |row| {
                let file_id: i64 = row.get(0)?;
                let relative_path: String = row.get(1)?;
                Ok(SearchHit {
                    kind: "author".to_string(),
                    file_id,
                    file_name: file_name_from_relative(&relative_path),
                    relative_path,
                    absolute_path: row.get(2)?,
                    heading_level: None,
                    heading_text: row.get(3)?,
                    heading_order: row.get(4)?,
                    score: row.get::<_, f64>(5)? + 400.0,
                })
            })
            .map_err(|error| format!("Could not run author search query: {error}"))?;

        for row in rows {
            let result =
                row.map_err(|error| format!("Could not parse author search row: {error}"))?;
            if let Some(key) = author_key(&result) {
                if !seen_author_keys.insert(key) {
                    continue;
                }
            }
            seen_file_ids.insert(result.file_id);
            results.push(result);
        }
    }

    let remaining = max_results.saturating_sub(i64::try_from(results.len()).unwrap_or(0));
    if remaining > 0 {
        let like_pattern = format!("%{}%", cleaned_query.to_ascii_lowercase());
        let mut statement = connection
            .prepare(
                "
                SELECT id, relative_path, absolute_path
                FROM files
                WHERE (?1 IS NULL OR root_id = ?1)
                  AND lower(relative_path) LIKE ?2
                ORDER BY relative_path
                LIMIT ?3
                ",
            )
            .map_err(|error| format!("Could not prepare file search query: {error}"))?;

        let rows = statement
            .query_map(params![requested_root_id, like_pattern, remaining], |row| {
                let file_id: i64 = row.get(0)?;
                let relative_path: String = row.get(1)?;
                Ok(SearchHit {
                    kind: "file".to_string(),
                    file_id,
                    file_name: file_name_from_relative(&relative_path),
                    relative_path,
                    absolute_path: row.get(2)?,
                    heading_level: None,
                    heading_text: None,
                    heading_order: None,
                    score: 9999.0,
                })
            })
            .map_err(|error| format!("Could not run file search query: {error}"))?;

        for row in rows {
            let result =
                row.map_err(|error| format!("Could not parse file search row: {error}"))?;
            if seen_file_ids.insert(result.file_id) {
                results.push(result);
            }
        }
    }

    if results.len() < max_results_usize {
        let threshold = fuzzy_threshold(&normalized_query);
        let query_len_chars = i64::try_from(normalized_query.chars().count()).unwrap_or(1);
        let min_heading_len = (query_len_chars - 6).max(1);
        let max_heading_len = query_len_chars + 36;
        let min_path_len = (query_len_chars - 6).max(1);
        let max_path_len = query_len_chars + 160;

        let heading_candidate_limit =
            i64::try_from((max_results_usize.saturating_mul(14)).clamp(120, 1800)).unwrap_or(600);
        let file_candidate_limit =
            i64::try_from((max_results_usize.saturating_mul(8)).clamp(80, 1200)).unwrap_or(400);

        let mut fuzzy_candidates = Vec::new();

        {
            let mut statement = connection
                .prepare(
                    "
                    SELECT
                      f.id,
                      f.relative_path,
                      f.absolute_path,
                      h.level,
                      h.text,
                      h.heading_order
                    FROM headings h
                    JOIN files f ON f.id = h.file_id
                    WHERE (?1 IS NULL OR f.root_id = ?1)
                      AND length(h.normalized) BETWEEN ?2 AND ?3
                    ORDER BY f.modified_ms DESC, h.id DESC
                    LIMIT ?4
                    ",
                )
                .map_err(|error| format!("Could not prepare fuzzy heading query: {error}"))?;

            let rows = statement
                .query_map(
                    params![
                        requested_root_id,
                        min_heading_len,
                        max_heading_len,
                        heading_candidate_limit
                    ],
                    |row| {
                        Ok((
                            row.get::<_, i64>(0)?,
                            row.get::<_, String>(1)?,
                            row.get::<_, String>(2)?,
                            row.get::<_, i64>(3)?,
                            row.get::<_, String>(4)?,
                            row.get::<_, i64>(5)?,
                        ))
                    },
                )
                .map_err(|error| format!("Could not run fuzzy heading query: {error}"))?;

            for row in rows {
                let (
                    file_id,
                    relative_path,
                    absolute_path,
                    heading_level,
                    heading_text,
                    heading_order,
                ) = row.map_err(|error| format!("Could not parse fuzzy heading row: {error}"))?;

                let heading_normalized = normalize_for_search(&heading_text);
                if heading_normalized.is_empty() {
                    continue;
                }

                let heading_similarity = fuzzy_similarity(&normalized_query, &heading_normalized);
                let path_similarity =
                    fuzzy_similarity(&normalized_query, &normalize_for_search(&relative_path))
                        * 0.84;
                let similarity = heading_similarity.max(path_similarity);
                if similarity < threshold {
                    continue;
                }

                fuzzy_candidates.push(SearchHit {
                    kind: "heading".to_string(),
                    file_id,
                    file_name: file_name_from_relative(&relative_path),
                    relative_path,
                    absolute_path,
                    heading_level: Some(heading_level),
                    heading_text: Some(heading_text),
                    heading_order: Some(heading_order),
                    score: 2000.0 + ((1.0 - similarity) * 1000.0),
                });
            }
        }

        {
            let mut statement = connection
                .prepare(
                    "
                    SELECT id, relative_path, absolute_path
                    FROM files
                    WHERE (?1 IS NULL OR root_id = ?1)
                      AND length(relative_path) BETWEEN ?2 AND ?3
                    ORDER BY modified_ms DESC, id DESC
                    LIMIT ?4
                    ",
                )
                .map_err(|error| format!("Could not prepare fuzzy file query: {error}"))?;

            let rows = statement
                .query_map(
                    params![
                        requested_root_id,
                        min_path_len,
                        max_path_len,
                        file_candidate_limit
                    ],
                    |row| {
                        Ok((
                            row.get::<_, i64>(0)?,
                            row.get::<_, String>(1)?,
                            row.get::<_, String>(2)?,
                        ))
                    },
                )
                .map_err(|error| format!("Could not run fuzzy file query: {error}"))?;

            for row in rows {
                let (file_id, relative_path, absolute_path) =
                    row.map_err(|error| format!("Could not parse fuzzy file row: {error}"))?;

                let file_name = file_name_from_relative(&relative_path);
                let path_similarity =
                    fuzzy_similarity(&normalized_query, &normalize_for_search(&relative_path));
                let name_similarity =
                    fuzzy_similarity(&normalized_query, &normalize_for_search(&file_name)) * 0.94;
                let similarity = path_similarity.max(name_similarity);
                if similarity < threshold {
                    continue;
                }

                fuzzy_candidates.push(SearchHit {
                    kind: "file".to_string(),
                    file_id,
                    file_name,
                    relative_path,
                    absolute_path,
                    heading_level: None,
                    heading_text: None,
                    heading_order: None,
                    score: 4000.0 + ((1.0 - similarity) * 1000.0),
                });
            }
        }

        {
            let author_candidate_limit =
                i64::try_from((max_results_usize.saturating_mul(10)).clamp(100, 1500))
                    .unwrap_or(500);
            let mut statement = connection
                .prepare(
                    "
                    SELECT
                      f.id,
                      f.relative_path,
                      f.absolute_path,
                      a.text,
                      a.author_order
                    FROM authors a
                    JOIN files f ON f.id = a.file_id
                    WHERE (?1 IS NULL OR f.root_id = ?1)
                      AND length(a.normalized) BETWEEN ?2 AND ?3
                    ORDER BY f.modified_ms DESC, a.id DESC
                    LIMIT ?4
                    ",
                )
                .map_err(|error| format!("Could not prepare fuzzy author query: {error}"))?;

            let rows = statement
                .query_map(
                    params![
                        requested_root_id,
                        min_heading_len,
                        max_heading_len + 100,
                        author_candidate_limit
                    ],
                    |row| {
                        Ok((
                            row.get::<_, i64>(0)?,
                            row.get::<_, String>(1)?,
                            row.get::<_, String>(2)?,
                            row.get::<_, String>(3)?,
                            row.get::<_, i64>(4)?,
                        ))
                    },
                )
                .map_err(|error| format!("Could not run fuzzy author query: {error}"))?;

            for row in rows {
                let (file_id, relative_path, absolute_path, author_text, author_order) =
                    row.map_err(|error| format!("Could not parse fuzzy author row: {error}"))?;

                let similarity =
                    fuzzy_similarity(&normalized_query, &normalize_for_search(&author_text));
                if similarity < threshold {
                    continue;
                }

                fuzzy_candidates.push(SearchHit {
                    kind: "author".to_string(),
                    file_id,
                    file_name: file_name_from_relative(&relative_path),
                    relative_path,
                    absolute_path,
                    heading_level: None,
                    heading_text: Some(author_text),
                    heading_order: Some(author_order),
                    score: 3000.0 + ((1.0 - similarity) * 1000.0),
                });
            }
        }

        fuzzy_candidates.sort_by(|left, right| {
            left.score
                .partial_cmp(&right.score)
                .unwrap_or(Ordering::Equal)
                .then(left.relative_path.cmp(&right.relative_path))
        });

        for candidate in fuzzy_candidates {
            if results.len() >= max_results_usize {
                break;
            }

            if candidate.kind == "file" {
                if !seen_file_ids.insert(candidate.file_id) {
                    continue;
                }
                results.push(candidate);
                continue;
            }

            if candidate.kind == "author" {
                if let Some(key) = author_key(&candidate) {
                    if !seen_author_keys.insert(key) {
                        continue;
                    }
                }
                seen_file_ids.insert(candidate.file_id);
                results.push(candidate);
                continue;
            }

            if let Some(key) = heading_key(&candidate) {
                if !seen_heading_keys.insert(key) {
                    continue;
                }
            }

            seen_file_ids.insert(candidate.file_id);
            results.push(candidate);
        }
    }

    Ok(results)
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            add_root,
            remove_root,
            insert_capture,
            list_capture_targets,
            get_capture_target_preview,
            delete_capture_heading,
            move_capture_heading,
            list_roots,
            index_root,
            get_index_snapshot,
            get_file_preview,
            get_heading_preview_html,
            search_index
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
