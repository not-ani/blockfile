use std::collections::HashSet;
use std::fs;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Component, Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use tauri::{AppHandle, Emitter};
use walkdir::DirEntry;

use crate::search::normalize_for_search;
use crate::types::{IndexProgress, ParsedParagraph};
use crate::CommandResult;
use crate::DEFAULT_CAPTURE_TARGET;

pub(crate) const INDEX_PROGRESS_EVENT: &str = "index-progress";
pub(crate) const INDEX_PROGRESS_EMIT_INTERVAL_MS: i64 = 120;

pub(crate) fn now_ms() -> i64 {
    epoch_ms(SystemTime::now())
}

pub(crate) fn epoch_ms(time: SystemTime) -> i64 {
    time.duration_since(UNIX_EPOCH)
        .ok()
        .and_then(|duration| i64::try_from(duration.as_millis()).ok())
        .unwrap_or(0)
}

pub(crate) fn path_display(path: &Path) -> String {
    path.to_string_lossy().into_owned()
}

pub(crate) fn suggested_parse_chunk_size() -> usize {
    std::thread::available_parallelism()
        .map(|parallelism| parallelism.get().saturating_mul(2))
        .unwrap_or(8)
        .clamp(8, 64)
}

pub(crate) fn emit_index_progress(
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

pub(crate) fn canonicalize_folder(path: &str) -> CommandResult<PathBuf> {
    let canonical = fs::canonicalize(path)
        .map_err(|error| format!("Could not access folder '{path}': {error}"))?;
    if !canonical.is_dir() {
        return Err(format!("Path is not a folder: {path}"));
    }
    Ok(canonical)
}

pub(crate) fn root_index_marker_path(root: &Path) -> PathBuf {
    root.join(".blockfile-index.json")
}

pub(crate) fn normalize_capture_target_path(target_path: Option<&str>) -> CommandResult<String> {
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

pub(crate) fn capture_docx_path(root: &Path, target_relative_path: &str) -> PathBuf {
    root.join(target_relative_path)
}

pub(crate) fn capture_marker(entry_id: i64) -> String {
    format!("BF-{entry_id:06}")
}

pub(crate) fn write_root_index_marker(root: &Path, last_indexed_ms: i64) -> CommandResult<()> {
    let marker_path = root_index_marker_path(root);
    let marker = serde_json::json!({
        "version": 2,
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

pub(crate) fn fast_file_hash(path: &Path) -> CommandResult<String> {
    const WINDOW_BYTES: usize = 64 * 1024;
    let mut file = fs::File::open(path)
        .map_err(|error| format!("Could not open '{}': {error}", path_display(path)))?;
    let metadata = file.metadata().map_err(|error| {
        format!(
            "Could not read metadata for '{}': {error}",
            path_display(path)
        )
    })?;
    let file_len = metadata.len();

    let mut hasher = blake3::Hasher::new();
    hasher.update(&file_len.to_le_bytes());

    let mut buffer = vec![0_u8; WINDOW_BYTES];
    let front_bytes = file.read(&mut buffer).map_err(|error| {
        format!(
            "Could not read hash prefix for '{}': {error}",
            path_display(path)
        )
    })?;
    hasher.update(&buffer[..front_bytes]);

    if file_len > u64::try_from(WINDOW_BYTES).unwrap_or(0) {
        let start = file_len.saturating_sub(u64::try_from(WINDOW_BYTES).unwrap_or(0));
        file.seek(SeekFrom::Start(start)).map_err(|error| {
            format!(
                "Could not seek hash suffix for '{}': {error}",
                path_display(path)
            )
        })?;
        let tail_bytes = file.read(&mut buffer).map_err(|error| {
            format!(
                "Could not read hash suffix for '{}': {error}",
                path_display(path)
            )
        })?;
        hasher.update(&buffer[..tail_bytes]);
    }

    Ok(hasher.finalize().to_hex().to_string())
}

pub(crate) fn file_name_from_relative(relative_path: &str) -> String {
    Path::new(relative_path)
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| relative_path.to_string())
}

pub(crate) fn folder_from_relative(relative_path: &str) -> String {
    relative_path
        .rsplit_once('/')
        .map(|(folder, _)| folder.to_string())
        .unwrap_or_default()
}

pub(crate) fn is_visible_entry(entry: &DirEntry) -> bool {
    let name = entry.file_name().to_string_lossy();
    !name.starts_with('.')
}

pub(crate) fn relative_path(root: &Path, file_path: &Path) -> CommandResult<String> {
    let relative = file_path
        .strip_prefix(root)
        .map_err(|error| format!("Failed to strip root prefix: {error}"))?;
    Ok(relative.to_string_lossy().replace('\\', "/"))
}

pub(crate) fn contains_year_token(text: &str) -> bool {
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

pub(crate) fn is_probable_author_line(text: &str) -> bool {
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

pub(crate) fn extract_author_candidates(paragraphs: &[ParsedParagraph]) -> Vec<(i64, String)> {
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
