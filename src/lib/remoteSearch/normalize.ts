import { buildDebatifyRichText } from "./richText";
import type { DebatifyApiTagHit, DebatifyTagHit } from "./types";

const REMOTE_SOURCE_URL = "https://api.debatify.app/search";
const CITATION_STYLE_PLACEHOLDER = "__BF_CITATION_STYLE__";

const asRecord = (value: unknown): Record<string, unknown> | null => {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  return value as Record<string, unknown>;
};

const asTrimmedString = (value: unknown) => (typeof value === "string" ? value.trim() : "");

const buildResultId = (tag: string, citation: string, index: number) => {
  const input = `${tag}|${citation}|${index}`;
  let hash = 0;
  for (let i = 0; i < input.length; i += 1) {
    hash = (hash * 31 + input.charCodeAt(i)) >>> 0;
  }
  return `debatify-${hash.toString(16)}`;
};

const buildSourcePath = (query: string, index: number) => {
  const encoded = encodeURIComponent(query);
  return `${REMOTE_SOURCE_URL}?q=${encoded}#result-${index + 1}`;
};

const formatCopyText = (tag: string, citation: string, plainText: string) => {
  const sections = [tag, citation, plainText].filter((value) => value.length > 0);
  return sections.join("\n\n");
};

const escapeXml = (value: string) =>
  value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");

const applyParagraphStyle = (paragraphXml: string, styleId: string) => {
  const trimmed = paragraphXml.trim();
  if (!trimmed.startsWith("<w:p")) return trimmed;
  if (trimmed === "<w:p/>") {
    return `<w:p><w:pPr><w:pStyle w:val="${escapeXml(styleId)}"/></w:pPr></w:p>`;
  }

  const paragraphMatch = trimmed.match(/^<w:p>([\s\S]*)<\/w:p>$/);
  if (!paragraphMatch) return trimmed;
  return `<w:p><w:pPr><w:pStyle w:val="${escapeXml(styleId)}"/></w:pPr>${paragraphMatch[1]}</w:p>`;
};

const cleanParagraphXmlList = (paragraphs: string[]) =>
  paragraphs.map((paragraph) => paragraph.trim()).filter((paragraph) => paragraph.length > 0);

const unwrapSingleParagraphHtml = (value: string) => {
  const trimmed = value.trim();
  const match = trimmed.match(/^<p>([\s\S]*)<\/p>$/i);
  return match ? match[1] : trimmed;
};

const buildPreviewRichHtml = (tagHtml: string, citationHtml: string, bodyHtml: string) => {
  const fragments: string[] = [];

  if (tagHtml.trim()) {
    fragments.push(`<p class="bf-preview-h4">${unwrapSingleParagraphHtml(tagHtml)}</p>`);
  }

  if (citationHtml.trim()) {
    fragments.push(`<p class="bf-preview-p"><em>${unwrapSingleParagraphHtml(citationHtml)}</em></p>`);
  }

  if (bodyHtml.trim()) {
    fragments.push(bodyHtml.trim());
  }

  return fragments.join("");
};

const buildInsertParagraphXml = (tagParagraphs: string[], citationParagraphs: string[], bodyParagraphs: string[]) => {
  const result: string[] = [];

  const cleanedTag = cleanParagraphXmlList(tagParagraphs).map((paragraph) => applyParagraphStyle(paragraph, "Heading4"));
  const cleanedCitation = cleanParagraphXmlList(citationParagraphs).map((paragraph) =>
    applyParagraphStyle(paragraph, CITATION_STYLE_PLACEHOLDER),
  );
  const cleanedBody = cleanParagraphXmlList(bodyParagraphs);

  result.push(...cleanedTag);
  result.push(...cleanedCitation);
  result.push(...cleanedBody);
  return result;
};

const normalizeApiHit = (value: unknown): DebatifyApiTagHit | null => {
  const record = asRecord(value);
  if (!record) return null;

  const tag = asTrimmedString(record.tag);
  const citation = asTrimmedString(record.citation);
  const markdown = asTrimmedString(record.markdown);

  if (!tag && !citation && !markdown) return null;
  return { tag, citation, markdown };
};

export const normalizeDebatifySearchResponse = (payload: unknown, query: string): DebatifyTagHit[] => {
  if (!Array.isArray(payload)) return [];

  const trimmedQuery = query.trim();
  if (!trimmedQuery) return [];

  const normalized: DebatifyTagHit[] = [];
  for (let index = 0; index < payload.length; index += 1) {
    const apiHit = normalizeApiHit(payload[index]);
    if (!apiHit) continue;

    const tag = apiHit.tag || `Tag ${index + 1}`;
    const tagRichText = buildDebatifyRichText(tag);
    const citationRichText = buildDebatifyRichText(apiHit.citation);
    const bodyRichText = buildDebatifyRichText(apiHit.markdown);
    const copyText = formatCopyText(tagRichText.plainText, citationRichText.plainText, bodyRichText.plainText);
    if (!copyText) continue;

    normalized.push({
      id: buildResultId(tag, apiHit.citation, index),
      tag,
      citation: apiHit.citation,
      richHtml: buildPreviewRichHtml(tagRichText.richHtml, citationRichText.richHtml, bodyRichText.richHtml),
      plainText: bodyRichText.plainText,
      copyText,
      paragraphXml: buildInsertParagraphXml(tagRichText.paragraphXml, citationRichText.paragraphXml, bodyRichText.paragraphXml),
      sourcePath: buildSourcePath(trimmedQuery, index),
    });
  }

  return normalized;
};
