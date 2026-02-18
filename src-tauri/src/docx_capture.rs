use std::collections::{HashMap, HashSet};
use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::Path;

use docx_rs::Docx;
use roxmltree::{Document, Node};
use zip::ZipArchive;

use crate::docx_parse::{
    attribute_value, has_tag, parse_docx_paragraphs, read_docx_part, read_zip_file,
    resolve_insert_after_order,
};
use crate::types::{RelationshipDef, SourceStyleDefinition, StyledSection};
use crate::util::{is_probable_author_line, path_display};
use crate::CommandResult;

const CITATION_STYLE_PLACEHOLDER: &str = "__BF_CITATION_STYLE__";

pub(crate) fn xml_escape_text(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

pub(crate) fn xml_escape_attr(value: &str) -> String {
    xml_escape_text(value)
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

pub(crate) fn paragraph_xml_plain(text: &str) -> String {
    if text.is_empty() {
        return "<w:p/>".to_string();
    }
    format!(
        "<w:p><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
        xml_escape_text(text)
    )
}

pub(crate) fn paragraph_xml_bold(text: &str) -> String {
    format!(
        "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
        xml_escape_text(text)
    )
}

pub(crate) fn paragraph_xml_heading(level: i64, text: &str) -> String {
    let style_id = format!("Heading{}", level);
    format!(
        "<w:p><w:pPr><w:pStyle w:val=\"{}\"/></w:pPr><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
        xml_escape_attr(&style_id),
        xml_escape_text(text)
    )
}

pub(crate) fn fallback_styled_section(content: &str) -> StyledSection {
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

pub(crate) fn extract_styled_section(
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

pub(crate) fn create_blank_docx(capture_path: &Path) -> CommandResult<()> {
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

pub(crate) fn ensure_valid_capture_docx(capture_path: &Path) -> CommandResult<()> {
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

pub(crate) fn document_has_body_content(document_xml: &str) -> bool {
    let Ok(document) = Document::parse(document_xml) else {
        return document_xml.contains("<w:p") || document_xml.contains("<w:tbl");
    };

    let Some(body) = document.descendants().find(|node| has_tag(*node, "body")) else {
        return false;
    };

    body.children()
        .any(|node| node.is_element() && !has_tag(node, "sectPr"))
}

pub(crate) fn body_bounds(document_xml: &str) -> CommandResult<(usize, usize)> {
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

pub(crate) fn fallback_body_insertion_index(document_xml: &str) -> CommandResult<usize> {
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

pub(crate) fn insertion_index_after_paragraph_count(
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

pub(crate) fn insert_fragment_into_document_xml(
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

pub(crate) fn merge_missing_styles(
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

pub(crate) fn parse_relationships(relationships_xml: &str) -> HashMap<String, RelationshipDef> {
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

pub(crate) fn merge_relationships(
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

pub(crate) fn remap_relationship_ids(
    paragraph_xml: &mut [String],
    id_remap: &HashMap<String, String>,
) {
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

fn citation_style_score(style_id: &str, style_name: &str) -> i32 {
    let combined = format!("{} {}", style_id, style_name).to_lowercase();
    let has_f8 = combined.contains("f8");
    let has_citation = combined.contains("citation");
    let has_cite = combined.contains("cite");
    let has_quote = combined.contains("quote");

    if has_f8 && (has_cite || has_citation) {
        return 600;
    }
    if has_citation {
        return 520;
    }
    if has_cite {
        return 430;
    }
    if has_quote {
        return 280;
    }
    if combined == "normal" {
        return -100;
    }
    0
}

fn resolve_citation_paragraph_style_id(styles_xml: &str) -> Option<String> {
    let Ok(document) = Document::parse(styles_xml) else {
        return None;
    };

    let mut best_match: Option<(i32, String)> = None;
    let mut quote_style_id: Option<String> = None;

    for style_node in document
        .descendants()
        .filter(|node| has_tag(*node, "style"))
    {
        let style_type = attribute_value(style_node, "type").unwrap_or("");
        if !style_type.eq_ignore_ascii_case("paragraph") {
            continue;
        }

        let Some(style_id_raw) = attribute_value(style_node, "styleId") else {
            continue;
        };
        let style_id = style_id_raw.trim();
        if style_id.is_empty() {
            continue;
        }

        let style_name = style_node
            .children()
            .find(|child| has_tag(*child, "name"))
            .and_then(|name_node| attribute_value(name_node, "val"))
            .unwrap_or("")
            .trim();

        let score = citation_style_score(style_id, style_name);
        if score > 0 {
            let replace_current = best_match
                .as_ref()
                .map(|(best_score, _)| score > *best_score)
                .unwrap_or(true);
            if replace_current {
                best_match = Some((score, style_id.to_string()));
            }
        }

        let style_id_lower = style_id.to_lowercase();
        let style_name_lower = style_name.to_lowercase();
        if quote_style_id.is_none()
            && (style_id_lower == "quote"
                || style_name_lower == "quote"
                || style_name_lower == "intense quote")
        {
            quote_style_id = Some(style_id.to_string());
        }
    }

    if let Some((_, style_id)) = best_match {
        return Some(style_id);
    }
    quote_style_id
}

fn apply_citation_style_placeholders(
    paragraph_xml: &mut [String],
    citation_style_id: Option<&str>,
) {
    let citation_style_id = citation_style_id.unwrap_or("Quote");
    let escaped_style_id = xml_escape_attr(citation_style_id);
    for paragraph in paragraph_xml.iter_mut() {
        if paragraph.contains(CITATION_STYLE_PLACEHOLDER) {
            *paragraph = paragraph.replace(CITATION_STYLE_PLACEHOLDER, &escaped_style_id);
        }
    }
}

pub(crate) fn rewrite_docx_with_parts(
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

pub(crate) fn append_capture_to_docx(
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

    let citation_paragraph_style_id = resolve_citation_paragraph_style_id(&target_styles_xml);
    apply_citation_style_placeholders(
        &mut section_paragraph_xml,
        citation_paragraph_style_id.as_deref(),
    );

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
