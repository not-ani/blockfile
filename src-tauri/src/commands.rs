use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::Path;
use std::time::Instant;

use rayon::prelude::*;
use rusqlite::{params, Connection};
use tauri::AppHandle;
use walkdir::WalkDir;

use crate::chunking::build_chunks;
use crate::db::{add_or_get_root_id, load_existing_files, open_database, root_id};
use crate::docx_capture::{
    append_capture_to_docx, ensure_valid_capture_docx, extract_styled_section,
    paragraph_xml_heading, rewrite_docx_with_parts,
};
use crate::docx_parse::{build_heading_ranges, has_tag, parse_docx_paragraphs, read_docx_part};
use crate::indexer::rebuild_lexical_index;
use crate::lexical;
use crate::preview::{extract_heading_preview_html, extract_preview_content};
use crate::query_engine;
use crate::search::normalize_for_search;
use crate::types::*;
use crate::util::*;
use crate::CommandResult;
use crate::DEFAULT_CAPTURE_TARGET;

use crate::docx_capture::{fallback_body_insertion_index, insertion_index_after_paragraph_count};

use roxmltree::{Document, Node};

#[tauri::command]
pub(crate) fn add_root(app: AppHandle, path: String) -> CommandResult<String> {
    let canonical = canonicalize_folder(&path)?;
    let canonical_string = path_display(&canonical);

    let connection = open_database(&app)?;
    add_or_get_root_id(&connection, &canonical_string)?;
    write_root_index_marker(&canonical, 0)?;
    Ok(canonical_string)
}

#[tauri::command]
pub(crate) fn remove_root(app: AppHandle, path: String) -> CommandResult<()> {
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
    rebuild_lexical_index(&app)?;
    Ok(())
}

#[tauri::command]
pub(crate) fn insert_capture(
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
pub(crate) fn list_capture_targets(
    app: AppHandle,
    root_path: String,
) -> CommandResult<Vec<CaptureTarget>> {
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
pub(crate) fn get_capture_target_preview(
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
pub(crate) fn delete_capture_heading(
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
pub(crate) fn move_capture_heading(
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
pub(crate) fn add_capture_heading(
    _app: AppHandle,
    root_path: String,
    target_path: String,
    heading_level: i64,
    heading_text: String,
    selected_target_heading_order: Option<i64>,
) -> CommandResult<CaptureTargetPreview> {
    if !(1..=4).contains(&heading_level) {
        return Err("Heading level must be H1, H2, H3, or H4.".to_string());
    }

    let trimmed_text = heading_text.trim();
    if trimmed_text.is_empty() {
        return Err("Heading name cannot be empty.".to_string());
    }

    let canonical_root = canonicalize_folder(&root_path)?;
    let normalized_target = normalize_capture_target_path(Some(&target_path))?;
    let absolute_path = capture_docx_path(&canonical_root, &normalized_target);

    let styled_section = StyledSection {
        paragraph_xml: vec![paragraph_xml_heading(heading_level, trimmed_text)],
        style_ids: HashSet::new(),
        relationship_ids: HashSet::new(),
        used_source_xml: false,
    };

    append_capture_to_docx(
        &absolute_path,
        &absolute_path,
        Some(heading_level),
        selected_target_heading_order.filter(|value| *value > 0),
        &styled_section,
    )?;

    Ok(capture_target_preview_for_path(
        &canonical_root,
        &normalized_target,
    ))
}

#[tauri::command]
pub(crate) fn list_roots(app: AppHandle) -> CommandResult<Vec<RootSummary>> {
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
pub(crate) fn index_root(app: AppHandle, path: String) -> CommandResult<IndexStats> {
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
        let relative_path_value = relative_path(&canonical_root, &absolute_path)?;
        seen_relative_paths.insert(relative_path_value.clone());

        let metadata = fs::metadata(&absolute_path).map_err(|error| {
            format!(
                "Could not read metadata for '{}': {error}",
                path_display(&absolute_path)
            )
        })?;
        let modified_ms = metadata.modified().map(epoch_ms).unwrap_or(0);
        let size = i64::try_from(metadata.len()).unwrap_or(0);

        if let Some(existing) = existing_files.get(&relative_path_value) {
            if existing.modified_ms == modified_ms
                && existing.size == size
                && !existing.file_hash.is_empty()
            {
                skipped += 1;
            } else {
                let file_hash = fast_file_hash(&absolute_path)?;
                if existing.file_hash == file_hash {
                    skipped += 1;
                } else {
                    indexing_candidates.push(IndexCandidate {
                        relative_path: relative_path_value.clone(),
                        absolute_path,
                        modified_ms,
                        size,
                        file_hash,
                    });
                }
            }
        } else {
            let file_hash = fast_file_hash(&absolute_path)?;
            indexing_candidates.push(IndexCandidate {
                relative_path: relative_path_value.clone(),
                absolute_path,
                modified_ms,
                size,
                file_hash,
            });
        }

        progress.discovered = scanned;
        progress.changed = indexing_candidates.len();
        progress.skipped = skipped;
        progress.current_file = Some(relative_path_value);
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
                let paragraphs =
                    parse_docx_paragraphs(&candidate.absolute_path).unwrap_or_default();
                let headings = paragraphs
                    .iter()
                    .filter_map(|paragraph| {
                        paragraph.heading_level.map(|level| ParsedHeading {
                            order: paragraph.order,
                            level,
                            text: paragraph.text.clone(),
                        })
                    })
                    .collect::<Vec<ParsedHeading>>();
                let authors = extract_author_candidates(&paragraphs);
                let chunks = build_chunks(&paragraphs);
                ParsedIndexCandidate {
                    candidate: candidate.clone(),
                    headings,
                    authors,
                    chunks,
                }
            })
            .collect::<Vec<ParsedIndexCandidate>>();

        for parsed in parsed_chunk {
            let relative_path_value = parsed.candidate.relative_path;
            let absolute_path_string = path_display(&parsed.candidate.absolute_path);
            let modified_ms = parsed.candidate.modified_ms;
            let size = parsed.candidate.size;
            let heading_count = i64::try_from(parsed.headings.len()).unwrap_or(0);
            headings_extracted += parsed.headings.len();

            let file_name = file_name_from_relative(&relative_path_value);

            let file_id = if let Some(existing) = existing_files.get(&relative_path_value) {
                transaction
                    .execute(
                        "UPDATE files
                         SET absolute_path = ?1, modified_ms = ?2, size = ?3, file_hash = ?4, heading_count = ?5
                         WHERE id = ?6",
                        params![
                            absolute_path_string,
                            modified_ms,
                            size,
                            parsed.candidate.file_hash.as_str(),
                            heading_count,
                            existing.id
                        ],
                    )
                    .map_err(|error| {
                        format!(
                            "Could not update indexed file '{}': {error}",
                            relative_path_value
                        )
                    })?;
                existing.id
            } else {
                transaction
                    .execute(
                        "INSERT INTO files(root_id, relative_path, absolute_path, modified_ms, size, file_hash, heading_count)
                         VALUES(?1, ?2, ?3, ?4, ?5, ?6, ?7)",
                        params![
                            root_id,
                            relative_path_value.as_str(),
                            absolute_path_string,
                            modified_ms,
                            size,
                            parsed.candidate.file_hash.as_str(),
                            heading_count
                        ],
                    )
                    .map_err(|error| {
                        format!(
                            "Could not insert indexed file '{}': {error}",
                            relative_path_value
                        )
                    })?;
                transaction.last_insert_rowid()
            };

            transaction
                .execute("DELETE FROM headings WHERE file_id = ?1", params![file_id])
                .map_err(|error| {
                    format!(
                        "Could not clear old headings for '{}': {error}",
                        relative_path_value
                    )
                })?;

            transaction
                .execute("DELETE FROM authors WHERE file_id = ?1", params![file_id])
                .map_err(|error| {
                    format!(
                        "Could not clear old author rows for '{}': {error}",
                        relative_path_value
                    )
                })?;

            transaction
                .execute("DELETE FROM chunks WHERE file_id = ?1", params![file_id])
                .map_err(|error| {
                    format!(
                        "Could not clear old chunks for '{}': {error}",
                        relative_path_value
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
                            relative_path_value.as_str()
                        ],
                    )
                    .map_err(|error| {
                        format!(
                            "Could not insert heading for '{}': {error}",
                            relative_path_value
                        )
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
                            relative_path_value.as_str()
                        ],
                    )
                    .map_err(|error| {
                        format!(
                            "Could not insert author metadata for '{}': {error}",
                            relative_path_value
                        )
                    })?;
            }

            for chunk in parsed.chunks {
                let chunk_id = format!("{}:{}:{}", root_id, file_id, chunk.chunk_order);
                transaction
                    .execute(
                        "
                        INSERT INTO chunks(
                          chunk_id,
                          root_id,
                          file_id,
                          chunk_order,
                          heading_order,
                          heading_level,
                          heading_text,
                          author_text,
                          chunk_text,
                          file_name,
                          relative_path,
                          absolute_path
                        )
                        VALUES(?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12)
                        ",
                        params![
                            chunk_id,
                            root_id,
                            file_id,
                            chunk.chunk_order,
                            chunk.heading_order,
                            chunk.heading_level,
                            chunk.heading_text,
                            chunk.author_text,
                            chunk.chunk_text,
                            file_name.as_str(),
                            relative_path_value.as_str(),
                            absolute_path_string.as_str()
                        ],
                    )
                    .map_err(|error| {
                        format!(
                            "Could not insert chunk row for '{}': {error}",
                            relative_path_value
                        )
                    })?;
            }

            updated += 1;
            progress.processed = updated;
            progress.updated = updated;
            progress.current_file = Some(relative_path_value);
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

    for (relative_path_value, file_id) in stale_entries {
        transaction
            .execute("DELETE FROM files WHERE id = ?1", params![file_id])
            .map_err(|error| {
                format!(
                    "Could not remove stale index row '{}': {error}",
                    relative_path_value
                )
            })?;
        removed += 1;

        progress.removed = removed;
        progress.current_file = Some(relative_path_value);
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

    rebuild_lexical_index(&app)?;

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

    // Rebuild vector index asynchronously after lexical/index metadata updates complete.
    crate::vector::trigger_rebuild(app.clone(), true);

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
pub(crate) fn get_index_snapshot(app: AppHandle, path: String) -> CommandResult<IndexSnapshot> {
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
pub(crate) fn get_file_preview(app: AppHandle, file_id: i64) -> CommandResult<FilePreview> {
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
pub(crate) fn get_heading_preview_html(
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
pub(crate) async fn search_index(
    app: AppHandle,
    query: String,
    root_path: Option<String>,
    limit: Option<usize>,
) -> CommandResult<Vec<SearchHit>> {
    tauri::async_runtime::spawn_blocking(move || {
        query_engine::search_lexical(&app, &query, root_path, limit)
    })
    .await
    .map_err(|error| format!("Lexical search command failed: {error}"))?
}

#[tauri::command]
pub(crate) async fn search_index_semantic(
    app: AppHandle,
    query: String,
    root_path: Option<String>,
    limit: Option<usize>,
) -> CommandResult<Vec<SearchHit>> {
    query_engine::search_semantic(&app, &query, root_path, limit).await
}

#[tauri::command]
pub(crate) async fn search_index_hybrid(
    app: AppHandle,
    query: String,
    root_path: Option<String>,
    limit: Option<usize>,
    file_name_only: Option<bool>,
    semantic_enabled: Option<bool>,
) -> CommandResult<Vec<SearchHit>> {
    query_engine::search_hybrid(
        &app,
        &query,
        root_path,
        limit,
        file_name_only.unwrap_or(false),
        semantic_enabled.unwrap_or(true),
    )
    .await
}

fn elapsed_ms(started: Instant) -> f64 {
    started.elapsed().as_secs_f64() * 1000.0
}

fn percentile(sorted_samples: &[f64], percentile: f64) -> f64 {
    if sorted_samples.is_empty() {
        return 0.0;
    }
    let clamped = percentile.clamp(0.0, 1.0);
    let last = sorted_samples.len().saturating_sub(1);
    let index = ((last as f64) * clamped).round() as usize;
    sorted_samples[index.min(last)]
}

fn latency_stats(samples: &[f64]) -> BenchmarkLatencyStats {
    if samples.is_empty() {
        return BenchmarkLatencyStats::default();
    }

    let mut sorted = samples.to_vec();
    sorted.sort_by(|left, right| {
        left.partial_cmp(right)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    let sum = sorted.iter().copied().sum::<f64>();
    BenchmarkLatencyStats {
        runs: sorted.len(),
        min_ms: *sorted.first().unwrap_or(&0.0),
        p50_ms: percentile(&sorted, 0.50),
        p95_ms: percentile(&sorted, 0.95),
        max_ms: *sorted.last().unwrap_or(&0.0),
        mean_ms: sum / (sorted.len() as f64),
    }
}

fn build_task_result(
    enabled: bool,
    samples: &[f64],
    total_hits: usize,
    error: Option<String>,
) -> BenchmarkTaskResult {
    BenchmarkTaskResult {
        enabled,
        error,
        total_hits,
        latency: latency_stats(samples),
    }
}

fn query_candidates_from_text(text: &str) -> Vec<String> {
    let normalized = normalize_for_search(text);
    if normalized.is_empty() {
        return Vec::new();
    }

    let tokens = normalized
        .split_whitespace()
        .filter(|token| token.len() >= 3)
        .map(|token| token.to_string())
        .collect::<Vec<String>>();
    if tokens.is_empty() {
        return Vec::new();
    }

    let mut candidates = Vec::new();
    let head_three = tokens.iter().take(3).cloned().collect::<Vec<String>>();
    if !head_three.is_empty() {
        candidates.push(head_three.join(" "));
    }
    if tokens.len() >= 2 {
        candidates.push(tokens.iter().take(2).cloned().collect::<Vec<String>>().join(" "));
    }
    candidates.push(tokens[0].clone());
    if tokens.len() >= 4 {
        let tail_two = tokens[tokens.len().saturating_sub(2)..]
            .iter()
            .cloned()
            .collect::<Vec<String>>()
            .join(" ");
        candidates.push(tail_two);
    }

    candidates
}

fn push_query_candidate(
    target: &mut Vec<String>,
    seen: &mut HashSet<String>,
    candidate: String,
    max_queries: usize,
) {
    let normalized = normalize_for_search(&candidate);
    if normalized.chars().count() < 2 {
        return;
    }
    if seen.insert(normalized.clone()) {
        target.push(normalized);
    }
    if target.len() > max_queries {
        target.truncate(max_queries);
    }
}

fn collect_benchmark_queries(
    connection: &Connection,
    root_id_value: i64,
    provided_queries: &[String],
    max_queries: usize,
) -> CommandResult<Vec<String>> {
    let mut queries = Vec::new();
    let mut seen = HashSet::new();

    for provided in provided_queries {
        for candidate in query_candidates_from_text(provided) {
            if queries.len() >= max_queries {
                return Ok(queries);
            }
            push_query_candidate(&mut queries, &mut seen, candidate, max_queries);
        }
    }

    let mut heading_statement = connection
        .prepare(
            "
            SELECT h.text
            FROM headings h
            JOIN files f ON f.id = h.file_id
            WHERE f.root_id = ?1
            ORDER BY length(h.text) DESC, h.id ASC
            LIMIT 240
            ",
        )
        .map_err(|error| format!("Could not prepare benchmark heading query source: {error}"))?;
    let heading_rows = heading_statement
        .query_map(params![root_id_value], |row| row.get::<_, String>(0))
        .map_err(|error| format!("Could not load benchmark heading query source: {error}"))?;
    for row in heading_rows {
        if queries.len() >= max_queries {
            break;
        }
        let text = row.map_err(|error| format!("Could not parse benchmark heading text: {error}"))?;
        for candidate in query_candidates_from_text(&text) {
            if queries.len() >= max_queries {
                break;
            }
            push_query_candidate(&mut queries, &mut seen, candidate, max_queries);
        }
    }

    let mut author_statement = connection
        .prepare(
            "
            SELECT a.text
            FROM authors a
            JOIN files f ON f.id = a.file_id
            WHERE f.root_id = ?1
            ORDER BY a.id DESC
            LIMIT 120
            ",
        )
        .map_err(|error| format!("Could not prepare benchmark author query source: {error}"))?;
    let author_rows = author_statement
        .query_map(params![root_id_value], |row| row.get::<_, String>(0))
        .map_err(|error| format!("Could not load benchmark author query source: {error}"))?;
    for row in author_rows {
        if queries.len() >= max_queries {
            break;
        }
        let text = row.map_err(|error| format!("Could not parse benchmark author text: {error}"))?;
        for candidate in query_candidates_from_text(&text) {
            if queries.len() >= max_queries {
                break;
            }
            push_query_candidate(&mut queries, &mut seen, candidate, max_queries);
        }
    }

    let mut file_statement = connection
        .prepare(
            "
            SELECT relative_path
            FROM files
            WHERE root_id = ?1
            ORDER BY heading_count DESC, modified_ms DESC
            LIMIT 180
            ",
        )
        .map_err(|error| format!("Could not prepare benchmark file query source: {error}"))?;
    let file_rows = file_statement
        .query_map(params![root_id_value], |row| row.get::<_, String>(0))
        .map_err(|error| format!("Could not load benchmark file query source: {error}"))?;
    for row in file_rows {
        if queries.len() >= max_queries {
            break;
        }
        let relative_path_value =
            row.map_err(|error| format!("Could not parse benchmark file relative path: {error}"))?;
        let file_name = file_name_from_relative(&relative_path_value);
        for candidate in query_candidates_from_text(&file_name) {
            if queries.len() >= max_queries {
                break;
            }
            push_query_candidate(&mut queries, &mut seen, candidate, max_queries);
        }
    }

    if queries.is_empty() {
        for fallback in [
            "introduction",
            "method",
            "results",
            "discussion",
            "conclusion",
            "references",
        ] {
            push_query_candidate(
                &mut queries,
                &mut seen,
                fallback.to_string(),
                max_queries,
            );
        }
    }

    Ok(queries)
}

fn sample_file_ids(
    connection: &Connection,
    root_id_value: i64,
    limit: usize,
) -> CommandResult<Vec<i64>> {
    if limit == 0 {
        return Ok(Vec::new());
    }

    let limit_i64 = i64::try_from(limit).unwrap_or(i64::MAX);
    let mut statement = connection
        .prepare(
            "
            SELECT id
            FROM files
            WHERE root_id = ?1
            ORDER BY heading_count DESC, modified_ms DESC, id DESC
            LIMIT ?2
            ",
        )
        .map_err(|error| format!("Could not prepare benchmark file preview sample query: {error}"))?;
    let rows = statement
        .query_map(params![root_id_value, limit_i64], |row| row.get::<_, i64>(0))
        .map_err(|error| format!("Could not run benchmark file preview sample query: {error}"))?;

    let mut output = Vec::new();
    for row in rows {
        output.push(row.map_err(|error| format!("Could not parse sampled file id: {error}"))?);
    }
    Ok(output)
}

fn sample_heading_refs(
    connection: &Connection,
    root_id_value: i64,
    limit: usize,
) -> CommandResult<Vec<(i64, i64)>> {
    if limit == 0 {
        return Ok(Vec::new());
    }

    let limit_i64 = i64::try_from(limit).unwrap_or(i64::MAX);
    let mut statement = connection
        .prepare(
            "
            SELECT h.file_id, h.heading_order
            FROM headings h
            JOIN files f ON f.id = h.file_id
            WHERE f.root_id = ?1
            ORDER BY f.heading_count DESC, h.heading_order ASC
            LIMIT ?2
            ",
        )
        .map_err(|error| {
            format!("Could not prepare benchmark heading preview sample query: {error}")
        })?;
    let rows = statement
        .query_map(params![root_id_value, limit_i64], |row| {
            Ok((row.get::<_, i64>(0)?, row.get::<_, i64>(1)?))
        })
        .map_err(|error| format!("Could not run benchmark heading preview sample query: {error}"))?;

    let mut output = Vec::new();
    for row in rows {
        output.push(
            row.map_err(|error| format!("Could not parse sampled heading reference: {error}"))?,
        );
    }
    Ok(output)
}

#[tauri::command]
pub(crate) async fn benchmark_root_performance(
    app: AppHandle,
    path: String,
    queries: Option<Vec<String>>,
    iterations: Option<usize>,
    limit: Option<usize>,
    include_semantic: Option<bool>,
    preview_samples: Option<usize>,
) -> CommandResult<BenchmarkReport> {
    let benchmark_started = Instant::now();
    let canonical_root = canonicalize_folder(&path)?;
    let root_path = path_display(&canonical_root);

    add_root(app.clone(), root_path.clone())?;
    let index_full = index_root(app.clone(), root_path.clone())?;
    let index_incremental = index_root(app.clone(), root_path.clone())?;

    let connection = open_database(&app)?;
    let root_id_value = root_id(&connection, &root_path)?.ok_or_else(|| {
        format!(
            "Benchmark root id missing for '{}'. Try indexing again.",
            root_path
        )
    })?;

    let benchmark_iterations = iterations.unwrap_or(3).clamp(1, 12);
    let benchmark_limit = limit.unwrap_or(80).clamp(10, 400);
    let benchmark_include_semantic = include_semantic.unwrap_or(true);
    let benchmark_preview_samples = preview_samples.unwrap_or(16).clamp(0, 240);
    let provided_queries = queries.unwrap_or_default();
    let benchmark_queries =
        collect_benchmark_queries(&connection, root_id_value, &provided_queries, 32)?;

    let mut search = BenchmarkSearchSummary {
        query_count: benchmark_queries.len(),
        iterations: benchmark_iterations,
        limit: benchmark_limit,
        ..BenchmarkSearchSummary::default()
    };

    let mut lexical_raw_samples = Vec::new();
    let mut lexical_raw_hits = 0_usize;
    let mut lexical_raw_error: Option<String> = None;
    'lexical_raw: for _ in 0..benchmark_iterations {
        for query in &benchmark_queries {
            let started = Instant::now();
            match lexical::search(&app, query, Some(root_id_value), benchmark_limit, false) {
                Ok(hits) => {
                    lexical_raw_samples.push(elapsed_ms(started));
                    lexical_raw_hits = lexical_raw_hits.saturating_add(hits.len());
                }
                Err(error) => {
                    lexical_raw_error = Some(error);
                    break 'lexical_raw;
                }
            }
        }
    }
    search.lexical_raw = build_task_result(
        true,
        &lexical_raw_samples,
        lexical_raw_hits,
        lexical_raw_error,
    );

    query_engine::clear_query_cache();
    for query in &benchmark_queries {
        let _ = query_engine::search_lexical(
            &app,
            query,
            Some(root_path.clone()),
            Some(benchmark_limit),
        );
    }
    let mut lexical_cached_samples = Vec::new();
    let mut lexical_cached_hits = 0_usize;
    let mut lexical_cached_error: Option<String> = None;
    'lexical_cached: for _ in 0..benchmark_iterations {
        for query in &benchmark_queries {
            let started = Instant::now();
            match query_engine::search_lexical(
                &app,
                query,
                Some(root_path.clone()),
                Some(benchmark_limit),
            ) {
                Ok(hits) => {
                    lexical_cached_samples.push(elapsed_ms(started));
                    lexical_cached_hits = lexical_cached_hits.saturating_add(hits.len());
                }
                Err(error) => {
                    lexical_cached_error = Some(error);
                    break 'lexical_cached;
                }
            }
        }
    }
    search.lexical_cached = build_task_result(
        true,
        &lexical_cached_samples,
        lexical_cached_hits,
        lexical_cached_error,
    );

    if benchmark_include_semantic {
        query_engine::clear_query_cache();
        for query in &benchmark_queries {
            let _ = query_engine::search_hybrid(
                &app,
                query,
                Some(root_path.clone()),
                Some(benchmark_limit),
                false,
                true,
            )
            .await;
        }

        let mut hybrid_samples = Vec::new();
        let mut hybrid_hits = 0_usize;
        let mut hybrid_error: Option<String> = None;
        'hybrid: for _ in 0..benchmark_iterations {
            for query in &benchmark_queries {
                let started = Instant::now();
                match query_engine::search_hybrid(
                    &app,
                    query,
                    Some(root_path.clone()),
                    Some(benchmark_limit),
                    false,
                    true,
                )
                .await
                {
                    Ok(hits) => {
                        hybrid_samples.push(elapsed_ms(started));
                        hybrid_hits = hybrid_hits.saturating_add(hits.len());
                    }
                    Err(error) => {
                        hybrid_error = Some(error);
                        break 'hybrid;
                    }
                }
            }
        }
        search.hybrid = build_task_result(true, &hybrid_samples, hybrid_hits, hybrid_error);

        let mut semantic_samples = Vec::new();
        let mut semantic_hits = 0_usize;
        let mut semantic_error: Option<String> = None;
        if let Some(warm_query) = benchmark_queries.first() {
            let _ = query_engine::search_semantic(
                &app,
                warm_query,
                Some(root_path.clone()),
                Some(benchmark_limit),
            )
            .await;
        }
        'semantic: for _ in 0..benchmark_iterations {
            for query in &benchmark_queries {
                let started = Instant::now();
                match query_engine::search_semantic(
                    &app,
                    query,
                    Some(root_path.clone()),
                    Some(benchmark_limit),
                )
                .await
                {
                    Ok(hits) => {
                        semantic_samples.push(elapsed_ms(started));
                        semantic_hits = semantic_hits.saturating_add(hits.len());
                    }
                    Err(error) => {
                        semantic_error = Some(error);
                        break 'semantic;
                    }
                }
            }
        }
        search.semantic = build_task_result(true, &semantic_samples, semantic_hits, semantic_error);
    } else {
        search.hybrid = build_task_result(false, &[], 0, None);
        search.semantic = build_task_result(false, &[], 0, None);
    }

    let snapshot_started = Instant::now();
    let _ = get_index_snapshot(app.clone(), root_path.clone())?;
    let mut preview = BenchmarkPreviewSummary {
        snapshot_ms: elapsed_ms(snapshot_started),
        ..BenchmarkPreviewSummary::default()
    };

    let sampled_file_ids = sample_file_ids(&connection, root_id_value, benchmark_preview_samples)?;
    let mut file_preview_samples = Vec::new();
    let mut file_preview_hits = 0_usize;
    let mut file_preview_error: Option<String> = None;
    for file_id in sampled_file_ids {
        let started = Instant::now();
        match get_file_preview(app.clone(), file_id) {
            Ok(file_preview) => {
                file_preview_samples.push(elapsed_ms(started));
                file_preview_hits =
                    file_preview_hits.saturating_add(usize::try_from(file_preview.heading_count).unwrap_or(0));
            }
            Err(error) => {
                file_preview_error = Some(error);
                break;
            }
        }
    }
    preview.file_preview = build_task_result(
        benchmark_preview_samples > 0,
        &file_preview_samples,
        file_preview_hits,
        file_preview_error,
    );

    let sampled_heading_refs =
        sample_heading_refs(&connection, root_id_value, benchmark_preview_samples)?;
    let mut heading_preview_samples = Vec::new();
    let mut heading_preview_hits = 0_usize;
    let mut heading_preview_error: Option<String> = None;
    for (file_id, heading_order) in sampled_heading_refs {
        let started = Instant::now();
        match get_heading_preview_html(app.clone(), file_id, heading_order) {
            Ok(html) => {
                heading_preview_samples.push(elapsed_ms(started));
                if !html.trim().is_empty() {
                    heading_preview_hits = heading_preview_hits.saturating_add(1);
                }
            }
            Err(error) => {
                heading_preview_error = Some(error);
                break;
            }
        }
    }
    preview.heading_preview_html = build_task_result(
        benchmark_preview_samples > 0,
        &heading_preview_samples,
        heading_preview_hits,
        heading_preview_error,
    );

    Ok(BenchmarkReport {
        root_path,
        index_full,
        index_incremental,
        queries: benchmark_queries,
        search,
        preview,
        generated_at_ms: now_ms(),
        elapsed_ms: elapsed_ms(benchmark_started).round() as i64,
    })
}
