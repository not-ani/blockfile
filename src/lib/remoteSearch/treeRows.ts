import type { TreeRow } from "../types";
import type { DebatifyTagHit } from "./types";

export const DEBATIFY_REMOTE_FOLDER_PATH = "remote:debatify";
const REMOTE_FOLDER_KEY = "search:folder:remote:debatify";

const headingSubLabel = (citation: string) => {
  if (citation) return `H4 - ${citation}`;
  return "H4 - Debatify API tag";
};

export const buildDebatifyTagTreeRows = (hits: DebatifyTagHit[], folderExpanded: boolean): TreeRow[] => {
  if (hits.length === 0) return [];

  const rows: TreeRow[] = [
    {
      key: REMOTE_FOLDER_KEY,
      kind: "folder",
      depth: 0,
      label: "Debatify API tags",
      subLabel: `${hits.length} match${hits.length === 1 ? "" : "es"}`,
      folderPath: DEBATIFY_REMOTE_FOLDER_PATH,
    },
  ];

  if (!folderExpanded) {
    return rows;
  }

  hits.forEach((hit, index) => {
    rows.push({
      key: `search:remote:tag:${hit.id}:${index}`,
      kind: "heading",
      depth: 1,
      label: hit.tag,
      subLabel: headingSubLabel(hit.citation),
      headingLevel: 4,
      copyText: hit.copyText,
      richHtml: hit.richHtml,
      paragraphXml: hit.paragraphXml,
      sourcePath: hit.sourcePath,
    });
  });

  return rows;
};
