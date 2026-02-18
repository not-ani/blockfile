use tauri::AppHandle;

use crate::db::open_database;
use crate::lexical::{self, LexicalDocument};
use crate::CommandResult;

pub(crate) fn rebuild_lexical_index(app: &AppHandle) -> CommandResult<()> {
    let connection = open_database(app)?;
    let mut documents = Vec::<LexicalDocument>::new();

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
            documents.push(LexicalDocument {
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
            });
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
            documents.push(LexicalDocument {
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
            });
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
            documents.push(LexicalDocument {
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
            });
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
            documents.push(LexicalDocument {
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
            });
        }
    }

    lexical::replace_all_documents(app, &documents)?;
    Ok(())
}
