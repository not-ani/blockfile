use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;

use roxmltree::{Document, Node};
use zip::ZipArchive;

use crate::search::normalize_for_search;
use crate::types::{HeadingRange, ParsedHeading, ParsedParagraph};
use crate::util::{is_probable_author_line, path_display};
use crate::CommandResult;

pub(crate) fn has_tag(node: Node<'_, '_>, expected: &str) -> bool {
    node.is_element() && node.tag_name().name() == expected
}

pub(crate) fn attribute_value<'a>(node: Node<'a, 'a>, key: &str) -> Option<&'a str> {
    if let Some(value) = node.attribute(key) {
        return Some(value);
    }
    node.attributes()
        .find_map(|attribute| (attribute.name().ends_with(key)).then_some(attribute.value()))
}

pub(crate) fn parse_trailing_level(value: &str) -> Option<i64> {
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

pub(crate) fn read_zip_file(archive: &mut ZipArchive<File>, entry_name: &str) -> Option<String> {
    let mut entry = archive.by_name(entry_name).ok()?;
    let mut value = String::new();
    entry.read_to_string(&mut value).ok()?;
    Some(value)
}

pub(crate) fn read_docx_part(path: &Path, part_name: &str) -> CommandResult<Option<String>> {
    let file = File::open(path)
        .map_err(|error| format!("Could not open '{}': {error}", path_display(path)))?;
    let mut archive = ZipArchive::new(file)
        .map_err(|error| format!("Could not read '{}': {error}", path_display(path)))?;
    Ok(read_zip_file(&mut archive, part_name))
}

pub(crate) fn read_style_map(styles_xml: Option<String>) -> HashMap<String, String> {
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

pub(crate) fn extract_paragraph_text(paragraph: Node<'_, '_>) -> String {
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

pub(crate) fn html_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&#39;")
}

pub(crate) fn run_properties_node<'a>(run: Node<'a, 'a>) -> Option<Node<'a, 'a>> {
    run.children().find(|node| has_tag(*node, "rPr"))
}

pub(crate) fn run_has_property(run: Node<'_, '_>, property_tag: &str) -> bool {
    run_properties_node(run)
        .and_then(|props| props.children().find(|node| has_tag(*node, property_tag)))
        .is_some()
}

pub(crate) fn run_has_active_underline(run: Node<'_, '_>) -> bool {
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

pub(crate) fn run_highlight_class(run: Node<'_, '_>) -> Option<&'static str> {
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

pub(crate) fn detect_heading_level(
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

pub(crate) fn paragraph_style_label(
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

pub(crate) fn is_f8_cite_style(style_label: &str) -> bool {
    let normalized = normalize_for_search(style_label);
    normalized.contains("f8 cite") || normalized.contains("f8cite")
}

pub(crate) fn parse_docx_paragraphs(file_path: &Path) -> CommandResult<Vec<ParsedParagraph>> {
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

pub(crate) fn build_heading_ranges(paragraphs: &[ParsedParagraph]) -> Vec<HeadingRange> {
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

pub(crate) fn resolve_insert_after_order(
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

#[allow(dead_code)]
pub(crate) fn extract_docx_headings_and_authors(
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

    let authors = crate::util::extract_author_candidates(&paragraphs);
    Ok((headings, authors))
}
