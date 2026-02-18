use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

use rusqlite::{params, Connection, OptionalExtension};
use tauri::{AppHandle, Manager};

use crate::types::ExistingFileMeta;
use crate::util::{now_ms, path_display};
use crate::CommandResult;

pub(crate) const INDEX_LAYOUT_VERSION: i64 = 2;
const INDEX_LAYOUT_DIR_NAME: &str = "index-v2";
const INDEX_META_DIR_NAME: &str = "meta";
const INDEX_LEXICAL_DIR_NAME: &str = "lexical";
const INDEX_VECTOR_DIR_NAME: &str = "vector";
const INDEX_LAYOUT_FILE_NAME: &str = "layout.json";
const DATABASE_FILE_NAME: &str = "blockfile-meta-v2.sqlite3";
const LEGACY_DATABASE_FILE_NAME: &str = "blockfile-index-v1.sqlite3";
const LEGACY_SEMANTIC_DIR_NAME: &str = "semantic-lancedb";
const LEGACY_SEMANTIC_META_FILE_NAME: &str = "semantic-index-meta-v1.json";

pub(crate) fn app_data_dir(app: &AppHandle) -> CommandResult<PathBuf> {
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
    Ok(app_data)
}

pub(crate) fn index_layout_dir(app: &AppHandle) -> CommandResult<PathBuf> {
    Ok(app_data_dir(app)?.join(INDEX_LAYOUT_DIR_NAME))
}

pub(crate) fn index_meta_dir(app: &AppHandle) -> CommandResult<PathBuf> {
    Ok(index_layout_dir(app)?.join(INDEX_META_DIR_NAME))
}

pub(crate) fn index_lexical_dir(app: &AppHandle) -> CommandResult<PathBuf> {
    Ok(index_layout_dir(app)?.join(INDEX_LEXICAL_DIR_NAME))
}

pub(crate) fn index_vector_dir(app: &AppHandle) -> CommandResult<PathBuf> {
    Ok(index_layout_dir(app)?.join(INDEX_VECTOR_DIR_NAME))
}

fn remove_path_if_exists(path: &PathBuf) -> CommandResult<()> {
    if !path.exists() {
        return Ok(());
    }
    if path.is_dir() {
        fs::remove_dir_all(path).map_err(|error| {
            format!(
                "Could not remove directory '{}': {error}",
                path_display(path)
            )
        })?;
        return Ok(());
    }
    fs::remove_file(path)
        .map_err(|error| format!("Could not remove file '{}': {error}", path_display(path)))
}

fn ensure_index_layout(app: &AppHandle) -> CommandResult<()> {
    let app_data = app_data_dir(app)?;
    let layout_dir = app_data.join(INDEX_LAYOUT_DIR_NAME);
    let layout_file = layout_dir.join(INDEX_LAYOUT_FILE_NAME);
    let current_version = fs::read_to_string(&layout_file).ok().and_then(|raw| {
        serde_json::from_str::<serde_json::Value>(&raw)
            .ok()
            .and_then(|value| value.get("version").and_then(|version| version.as_i64()))
    });

    if current_version == Some(INDEX_LAYOUT_VERSION) {
        fs::create_dir_all(layout_dir.join(INDEX_META_DIR_NAME)).map_err(|error| {
            format!(
                "Could not create index meta dir '{}': {error}",
                path_display(&layout_dir.join(INDEX_META_DIR_NAME))
            )
        })?;
        fs::create_dir_all(layout_dir.join(INDEX_LEXICAL_DIR_NAME)).map_err(|error| {
            format!(
                "Could not create lexical index dir '{}': {error}",
                path_display(&layout_dir.join(INDEX_LEXICAL_DIR_NAME))
            )
        })?;
        fs::create_dir_all(layout_dir.join(INDEX_VECTOR_DIR_NAME)).map_err(|error| {
            format!(
                "Could not create vector index dir '{}': {error}",
                path_display(&layout_dir.join(INDEX_VECTOR_DIR_NAME))
            )
        })?;
        return Ok(());
    }

    // Hard reset path: v1 compatibility is intentionally removed.
    remove_path_if_exists(&app_data.join(LEGACY_DATABASE_FILE_NAME))?;
    remove_path_if_exists(&app_data.join(LEGACY_SEMANTIC_DIR_NAME))?;
    remove_path_if_exists(&app_data.join(LEGACY_SEMANTIC_META_FILE_NAME))?;
    remove_path_if_exists(&layout_dir)?;

    fs::create_dir_all(layout_dir.join(INDEX_META_DIR_NAME)).map_err(|error| {
        format!(
            "Could not create index meta dir '{}': {error}",
            path_display(&layout_dir.join(INDEX_META_DIR_NAME))
        )
    })?;
    fs::create_dir_all(layout_dir.join(INDEX_LEXICAL_DIR_NAME)).map_err(|error| {
        format!(
            "Could not create lexical index dir '{}': {error}",
            path_display(&layout_dir.join(INDEX_LEXICAL_DIR_NAME))
        )
    })?;
    fs::create_dir_all(layout_dir.join(INDEX_VECTOR_DIR_NAME)).map_err(|error| {
        format!(
            "Could not create vector index dir '{}': {error}",
            path_display(&layout_dir.join(INDEX_VECTOR_DIR_NAME))
        )
    })?;

    let manifest = serde_json::json!({
        "version": INDEX_LAYOUT_VERSION,
        "updatedAtMs": now_ms(),
    });
    let manifest_raw = serde_json::to_string_pretty(&manifest)
        .map_err(|error| format!("Could not serialize index layout manifest: {error}"))?;
    fs::write(&layout_file, manifest_raw).map_err(|error| {
        format!(
            "Could not write index layout manifest '{}': {error}",
            path_display(&layout_file)
        )
    })?;

    Ok(())
}

pub(crate) fn database_path(app: &AppHandle) -> CommandResult<PathBuf> {
    Ok(index_meta_dir(app)?.join(DATABASE_FILE_NAME))
}

pub(crate) fn table_has_column(
    connection: &Connection,
    table: &str,
    column: &str,
) -> CommandResult<bool> {
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

pub(crate) fn ensure_capture_schema(connection: &Connection) -> CommandResult<()> {
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

pub(crate) fn open_database(app: &AppHandle) -> CommandResult<Connection> {
    ensure_index_layout(app)?;
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
              file_hash TEXT NOT NULL DEFAULT '',
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

            CREATE TABLE IF NOT EXISTS chunks (
              id INTEGER PRIMARY KEY,
              chunk_id TEXT NOT NULL UNIQUE,
              root_id INTEGER NOT NULL,
              file_id INTEGER NOT NULL,
              chunk_order INTEGER NOT NULL,
              heading_order INTEGER,
              heading_level INTEGER,
              heading_text TEXT,
              author_text TEXT,
              chunk_text TEXT NOT NULL,
              file_name TEXT NOT NULL,
              relative_path TEXT NOT NULL,
              absolute_path TEXT NOT NULL,
              FOREIGN KEY(root_id) REFERENCES roots(id) ON DELETE CASCADE,
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
            CREATE INDEX IF NOT EXISTS idx_files_root_modified ON files(root_id, modified_ms DESC, id DESC);
            CREATE INDEX IF NOT EXISTS idx_headings_file ON headings(file_id);
            CREATE INDEX IF NOT EXISTS idx_headings_file_order ON headings(file_id, heading_order);
            CREATE INDEX IF NOT EXISTS idx_headings_normalized_length ON headings(length(normalized));
            CREATE INDEX IF NOT EXISTS idx_authors_file ON authors(file_id);
            CREATE INDEX IF NOT EXISTS idx_authors_file_order ON authors(file_id, author_order);
            CREATE INDEX IF NOT EXISTS idx_authors_normalized_length ON authors(length(normalized));
            CREATE INDEX IF NOT EXISTS idx_chunks_file_order ON chunks(file_id, chunk_order);
            CREATE INDEX IF NOT EXISTS idx_chunks_root_file ON chunks(root_id, file_id);
            CREATE INDEX IF NOT EXISTS idx_chunks_root_file_order ON chunks(root_id, file_id, chunk_order);
            CREATE INDEX IF NOT EXISTS idx_files_relative_length ON files(length(relative_path));
            CREATE INDEX IF NOT EXISTS idx_captures_root ON captures(root_id, id);
            ",
        )
        .map_err(|error| format!("Could not initialize index database: {error}"))?;

    let _ = connection.query_row("PRAGMA cache_size = -65536", [], |row| row.get::<_, i64>(0));
    let _ = connection.query_row("PRAGMA mmap_size = 268435456", [], |row| {
        row.get::<_, i64>(0)
    });
    let _ = connection.query_row("PRAGMA wal_autocheckpoint = 1000", [], |row| {
        row.get::<_, i64>(0)
    });

    ensure_capture_schema(&connection)?;

    Ok(connection)
}

pub(crate) fn root_id(connection: &Connection, root_path: &str) -> CommandResult<Option<i64>> {
    connection
        .query_row(
            "SELECT id FROM roots WHERE path = ?1",
            params![root_path],
            |row| row.get(0),
        )
        .optional()
        .map_err(|error| format!("Could not query root path '{root_path}': {error}"))
}

pub(crate) fn add_or_get_root_id(connection: &Connection, root_path: &str) -> CommandResult<i64> {
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

pub(crate) fn load_existing_files(
    connection: &Connection,
    root_id: i64,
) -> CommandResult<HashMap<String, ExistingFileMeta>> {
    let mut statement = connection
        .prepare(
            "SELECT id, relative_path, modified_ms, size, file_hash FROM files WHERE root_id = ?1",
        )
        .map_err(|error| format!("Could not prepare file metadata query: {error}"))?;

    let rows = statement
        .query_map(params![root_id], |row| {
            Ok((
                row.get::<_, i64>(0)?,
                row.get::<_, String>(1)?,
                row.get::<_, i64>(2)?,
                row.get::<_, i64>(3)?,
                row.get::<_, String>(4)?,
            ))
        })
        .map_err(|error| format!("Could not iterate existing files: {error}"))?;

    let mut metadata = HashMap::new();
    for row in rows {
        let (id, relative_path, modified_ms, size, file_hash) =
            row.map_err(|error| format!("Could not parse existing file metadata row: {error}"))?;
        metadata.insert(
            relative_path,
            ExistingFileMeta {
                id,
                modified_ms,
                size,
                file_hash,
            },
        );
    }

    Ok(metadata)
}
