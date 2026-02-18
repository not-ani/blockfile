type RunStyle = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  highlight: boolean;
  fontSizeHalfPoints: number | null;
};

type DebatifyRichText = {
  richHtml: string;
  plainText: string;
  paragraphXml: string[];
};

const BASE_STYLE: RunStyle = {
  bold: false,
  italic: false,
  underline: false,
  highlight: false,
  fontSizeHalfPoints: null,
};

const BLOCK_TAGS = new Set(["p", "div"]);

const escapeHtml = (value: string) =>
  value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");

const escapeXml = (value: string) =>
  value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");

const normalizeWhitespace = (value: string) =>
  value
    .replace(/\r\n/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

const parseFontSizeHalfPoints = (styleValue: string | null) => {
  if (!styleValue) return null;
  const match = styleValue.match(/font-size\s*:\s*([0-9]*\.?[0-9]+)\s*pt/i);
  if (!match) return null;
  const parsed = Number.parseFloat(match[1]);
  if (!Number.isFinite(parsed) || parsed <= 0) return null;
  const clampedPt = Math.max(6, Math.min(48, parsed));
  return Math.round(clampedPt * 2);
};

const readCssProp = (styleValue: string | null, propName: string) => {
  if (!styleValue) return "";
  const entries = styleValue
    .split(";")
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);

  const lowerPropName = propName.toLowerCase();
  for (const entry of entries) {
    const split = entry.indexOf(":");
    if (split <= 0) continue;
    const key = entry.slice(0, split).trim().toLowerCase();
    if (key !== lowerPropName) continue;
    return entry.slice(split + 1).trim();
  }
  return "";
};

const styleHasUnderline = (styleValue: string | null) => {
  const textDecoration = readCssProp(styleValue, "text-decoration").toLowerCase();
  const textDecorationLine = readCssProp(styleValue, "text-decoration-line").toLowerCase();
  return textDecoration.includes("underline") || textDecorationLine.includes("underline");
};

const styleHasBold = (styleValue: string | null) => {
  const fontWeight = readCssProp(styleValue, "font-weight").toLowerCase();
  if (!fontWeight) return false;
  if (fontWeight === "bold" || fontWeight === "bolder") return true;
  const numeric = Number.parseInt(fontWeight, 10);
  return Number.isFinite(numeric) && numeric >= 600;
};

const styleHasItalic = (styleValue: string | null) => {
  const fontStyle = readCssProp(styleValue, "font-style").toLowerCase();
  return fontStyle === "italic" || fontStyle === "oblique";
};

const styleHasHighlight = (styleValue: string | null) => {
  const background = readCssProp(styleValue, "background-color").toLowerCase();
  if (!background) return false;
  return background !== "transparent" && background !== "none";
};

const fontSizePtString = (halfPoints: number | null) => {
  if (!halfPoints) return "";
  const points = halfPoints / 2;
  return Number.isInteger(points) ? String(points) : points.toFixed(1).replace(/\.0$/, "");
};

const styleForTag = (tag: string, element: Element, current: RunStyle): RunStyle => {
  const next: RunStyle = { ...current };
  if (tag === "strong" || tag === "b") {
    next.bold = true;
  } else if (tag === "em" || tag === "i") {
    next.italic = true;
  } else if (tag === "u") {
    next.underline = true;
  } else if (tag === "mark") {
    next.highlight = true;
  } else if (tag === "span") {
    const inlineStyle = element.getAttribute("style");
    const fontSizeHalfPoints = parseFontSizeHalfPoints(inlineStyle);
    if (fontSizeHalfPoints) {
      next.fontSizeHalfPoints = fontSizeHalfPoints;
    }
    if (styleHasUnderline(inlineStyle)) {
      next.underline = true;
    }
    if (styleHasBold(inlineStyle)) {
      next.bold = true;
    }
    if (styleHasItalic(inlineStyle)) {
      next.italic = true;
    }
    if (styleHasHighlight(inlineStyle)) {
      next.highlight = true;
    }
  }
  return next;
};

const sanitizeSpanStyle = (styleValue: string | null) => {
  if (!styleValue) return "";
  const entries: string[] = [];

  const fontSizeHalfPoints = parseFontSizeHalfPoints(styleValue);
  if (fontSizeHalfPoints) {
    const points = fontSizePtString(fontSizeHalfPoints);
    entries.push(`font-size:${points}pt`);
  }

  if (styleHasUnderline(styleValue)) {
    entries.push("text-decoration:underline");
  }

  if (styleHasBold(styleValue)) {
    entries.push("font-weight:700");
  }

  if (styleHasItalic(styleValue)) {
    entries.push("font-style:italic");
  }

  const background = readCssProp(styleValue, "background-color");
  if (background && styleHasHighlight(styleValue)) {
    entries.push(`background-color:${background}`);
  }

  return entries.join(";");
};

const sanitizeNode = (node: ChildNode): string => {
  if (node.nodeType === Node.TEXT_NODE) {
    return escapeHtml(node.textContent ?? "");
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return "";

  const element = node as Element;
  const tag = element.tagName.toLowerCase();
  const children = Array.from(element.childNodes).map(sanitizeNode).join("");

  if (tag === "br") return "<br>";

  if (BLOCK_TAGS.has(tag)) {
    if (!children.trim()) return "";
    return `<p>${children}</p>`;
  }

  if (tag === "strong" || tag === "b") {
    return `<strong>${children}</strong>`;
  }

  if (tag === "em" || tag === "i") {
    return `<em>${children}</em>`;
  }

  if (tag === "u") {
    return `<u>${children}</u>`;
  }

  if (tag === "mark") {
    return `<mark>${children}</mark>`;
  }

  if (tag === "span") {
    const safeStyle = sanitizeSpanStyle(element.getAttribute("style"));
    if (!safeStyle) {
      return `<span>${children}</span>`;
    }
    return `<span style="${escapeHtml(safeStyle)}">${children}</span>`;
  }

  return children;
};

const sanitizeDebatifyHtml = (rawHtml: string) => {
  if (!rawHtml.trim()) return "";

  const parser = new DOMParser();
  const document = parser.parseFromString(`<body>${rawHtml}</body>`, "text/html");
  return Array.from(document.body.childNodes).map(sanitizeNode).join("").trim();
};

const renderRunProps = (style: RunStyle) => {
  const entries: string[] = [];
  if (style.bold) entries.push("<w:b/>");
  if (style.italic) entries.push("<w:i/>");
  if (style.underline) entries.push('<w:u w:val="single"/>');
  if (style.highlight) entries.push('<w:highlight w:val="yellow"/>');
  if (style.fontSizeHalfPoints) {
    const size = String(style.fontSizeHalfPoints);
    entries.push(`<w:sz w:val="${size}"/>`);
    entries.push(`<w:szCs w:val="${size}"/>`);
  }

  if (entries.length === 0) return "";
  return `<w:rPr>${entries.join("")}</w:rPr>`;
};

const pushTextRuns = (runs: string[], text: string, style: RunStyle) => {
  const normalized = text.replace(/\r/g, "");
  if (!normalized) return;

  const pieces = normalized.split("\n");
  const runProps = renderRunProps(style);
  pieces.forEach((piece, index) => {
    if (piece.length > 0) {
      runs.push(`<w:r>${runProps}<w:t xml:space="preserve">${escapeXml(piece)}</w:t></w:r>`);
    }
    if (index < pieces.length - 1) {
      runs.push(`<w:r>${runProps}<w:br/></w:r>`);
    }
  });
};

const collectInlineRuns = (node: ChildNode, style: RunStyle, runs: string[]) => {
  if (node.nodeType === Node.TEXT_NODE) {
    pushTextRuns(runs, node.textContent ?? "", style);
    return;
  }

  if (node.nodeType !== Node.ELEMENT_NODE) return;

  const element = node as Element;
  const tag = element.tagName.toLowerCase();
  if (tag === "br") {
    const runProps = renderRunProps(style);
    runs.push(`<w:r>${runProps}<w:br/></w:r>`);
    return;
  }

  const nextStyle = styleForTag(tag, element, style);
  Array.from(element.childNodes).forEach((child) => collectInlineRuns(child, nextStyle, runs));
};

const renderParagraph = (runs: string[]) => {
  if (runs.length === 0) return "<w:p/>";
  return `<w:p>${runs.join("")}</w:p>`;
};

const buildParagraphXml = (safeHtml: string) => {
  if (!safeHtml) return ["<w:p/>"];

  const parser = new DOMParser();
  const document = parser.parseFromString(`<body>${safeHtml}</body>`, "text/html");
  const paragraphs: string[] = [];
  let currentRuns: string[] = [];

  const flushCurrentRuns = () => {
    if (currentRuns.length === 0) return;
    paragraphs.push(renderParagraph(currentRuns));
    currentRuns = [];
  };

  for (const node of Array.from(document.body.childNodes)) {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const element = node as Element;
      const tag = element.tagName.toLowerCase();
      if (BLOCK_TAGS.has(tag)) {
        flushCurrentRuns();
        const blockRuns: string[] = [];
        Array.from(element.childNodes).forEach((child) => collectInlineRuns(child, BASE_STYLE, blockRuns));
        paragraphs.push(renderParagraph(blockRuns));
        continue;
      }
    }

    collectInlineRuns(node, BASE_STYLE, currentRuns);
  }

  flushCurrentRuns();
  if (paragraphs.length === 0) {
    paragraphs.push("<w:p/>");
  }
  return paragraphs;
};

const extractPlainText = (safeHtml: string) => {
  if (!safeHtml) return "";
  const parser = new DOMParser();
  const document = parser.parseFromString(`<body>${safeHtml}</body>`, "text/html");
  return normalizeWhitespace(document.body.textContent ?? "");
};

export const buildDebatifyRichText = (rawHtml: string): DebatifyRichText => {
  const richHtml = sanitizeDebatifyHtml(rawHtml);
  const plainText = extractPlainText(richHtml);
  const paragraphXml = buildParagraphXml(richHtml);
  return {
    richHtml,
    plainText,
    paragraphXml,
  };
};
