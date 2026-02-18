use crate::types::{ParsedChunk, ParsedParagraph};
use crate::util::is_probable_author_line;

const CHUNK_MIN_CHARS: usize = 420;
const CHUNK_MAX_CHARS: usize = 860;
const CHUNK_OVERLAP_CHARS: usize = 160;

fn split_text_into_chunks(text: &str) -> Vec<String> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return Vec::new();
    }

    let chars = trimmed.chars().collect::<Vec<char>>();
    if chars.len() <= CHUNK_MAX_CHARS {
        return vec![trimmed.to_string()];
    }

    let mut chunks = Vec::new();
    let mut start = 0_usize;

    while start < chars.len() {
        let max_end = (start + CHUNK_MAX_CHARS).min(chars.len());
        let min_end = (start + CHUNK_MIN_CHARS).min(max_end);
        let mut cut = max_end;

        for index in (min_end..max_end).rev() {
            if chars[index].is_whitespace() {
                cut = index;
                break;
            }
        }

        if cut <= start {
            cut = max_end;
        }

        let chunk_text = chars[start..cut]
            .iter()
            .collect::<String>()
            .trim()
            .to_string();
        if !chunk_text.is_empty() {
            chunks.push(chunk_text);
        }

        if cut >= chars.len() {
            break;
        }

        let next_start = cut.saturating_sub(CHUNK_OVERLAP_CHARS);
        if next_start <= start {
            start = cut;
        } else {
            start = next_start;
        }
    }

    chunks
}

pub(crate) fn build_chunks(paragraphs: &[ParsedParagraph]) -> Vec<ParsedChunk> {
    let mut chunks = Vec::new();
    let mut chunk_order = 1_i64;

    let mut current_heading_order: Option<i64> = None;
    let mut current_heading_level: Option<i64> = None;
    let mut current_heading_text: Option<String> = None;
    let mut section_author: Option<String> = None;
    let mut section_lines = Vec::<String>::new();

    let flush_section = |chunks: &mut Vec<ParsedChunk>,
                         chunk_order: &mut i64,
                         lines: &mut Vec<String>,
                         heading_order: Option<i64>,
                         heading_level: Option<i64>,
                         heading_text: Option<String>,
                         author_text: Option<String>| {
        if lines.is_empty() {
            return;
        }

        let section_text = lines.join("\n");
        lines.clear();

        for chunk_text in split_text_into_chunks(&section_text) {
            chunks.push(ParsedChunk {
                chunk_order: *chunk_order,
                heading_order,
                heading_level,
                heading_text: heading_text.clone(),
                author_text: author_text.clone(),
                chunk_text,
            });
            *chunk_order += 1;
        }
    };

    for paragraph in paragraphs {
        let text = paragraph.text.trim();
        if text.is_empty() {
            continue;
        }

        if let Some(level) = paragraph.heading_level {
            flush_section(
                &mut chunks,
                &mut chunk_order,
                &mut section_lines,
                current_heading_order,
                current_heading_level,
                current_heading_text.clone(),
                section_author.clone(),
            );

            current_heading_order = Some(paragraph.order);
            current_heading_level = Some(level);
            current_heading_text = Some(text.to_string());
            section_author = None;

            // Keep structure searchable even when body text is short.
            chunks.push(ParsedChunk {
                chunk_order,
                heading_order: current_heading_order,
                heading_level: current_heading_level,
                heading_text: current_heading_text.clone(),
                author_text: None,
                chunk_text: text.to_string(),
            });
            chunk_order += 1;
            continue;
        }

        if section_author.is_none() && is_probable_author_line(text) {
            section_author = Some(text.to_string());
        }
        section_lines.push(text.to_string());
    }

    flush_section(
        &mut chunks,
        &mut chunk_order,
        &mut section_lines,
        current_heading_order,
        current_heading_level,
        current_heading_text,
        section_author,
    );

    chunks
}
