use std::fs::File;
use std::path::Path;

use roxmltree::{Document, Node};
use zip::ZipArchive;

use crate::docx_parse::{
    build_heading_ranges, has_tag, html_escape, parse_docx_paragraphs, read_zip_file,
    run_has_active_underline, run_has_property, run_highlight_class,
};
use crate::types::{FileHeading, TaggedBlock};
use crate::util::{is_probable_author_line, path_display};
use crate::CommandResult;

fn push_escaped_text_with_breaks(target: &mut String, text: &str) {
    for (index, segment) in text.split('\n').enumerate() {
        if index > 0 {
            target.push_str("<br/>");
        }
        target.push_str(&html_escape(segment));
    }
}

pub(crate) fn render_preview_run(run: Node<'_, '_>) -> String {
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

pub(crate) fn render_preview_inline_nodes(node: Node<'_, '_>, output: &mut String) {
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

pub(crate) fn preview_paragraph_class(heading_level: Option<i64>) -> &'static str {
    match heading_level {
        Some(1) => "bf-preview-h1",
        Some(2) => "bf-preview-h2",
        Some(3) => "bf-preview-h3",
        Some(4) => "bf-preview-h4",
        _ => "bf-preview-p",
    }
}

pub(crate) fn render_preview_paragraph(
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

pub(crate) fn extract_heading_preview_html(
    file_path: &Path,
    heading_order: i64,
) -> CommandResult<String> {
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

pub(crate) fn extract_preview_content(
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
