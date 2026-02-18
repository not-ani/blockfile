export type RootSummary = {
  path: string;
  fileCount: number;
  headingCount: number;
  addedAtMs: number;
  lastIndexedMs: number;
};

export type FolderEntry = {
  path: string;
  name: string;
  parentPath: string | null;
  depth: number;
  fileCount: number;
};

export type IndexedFile = {
  id: number;
  fileName: string;
  relativePath: string;
  folderPath: string;
  modifiedMs: number;
  headingCount: number;
};

export type IndexSnapshot = {
  rootPath: string;
  indexedAtMs: number;
  folders: FolderEntry[];
  files: IndexedFile[];
};

export type FileHeading = {
  id: number;
  order: number;
  level: number;
  text: string;
  copyText: string;
};

export type TaggedBlock = {
  order: number;
  styleLabel: string;
  text: string;
};

export type FilePreview = {
  fileId: number;
  fileName: string;
  relativePath: string;
  absolutePath: string;
  headingCount: number;
  headings: FileHeading[];
  f8Cites: TaggedBlock[];
};

export type SearchHit = {
  source: "lexical" | "semantic" | "hybrid";
  kind: "heading" | "file" | "author";
  fileId: number;
  fileName: string;
  relativePath: string;
  absolutePath: string;
  headingLevel: number | null;
  headingText: string | null;
  headingOrder: number | null;
  score: number;
};

export type IndexStats = {
  scanned: number;
  updated: number;
  skipped: number;
  removed: number;
  headingsExtracted: number;
  elapsedMs: number;
};

export type IndexProgress = {
  rootPath: string;
  phase: "discovering" | "indexing" | "cleaning" | "complete";
  discovered: number;
  changed: number;
  processed: number;
  updated: number;
  skipped: number;
  removed: number;
  elapsedMs: number;
  currentFile: string | null;
};

export type TreeRow = {
  key: string;
  kind: "folder" | "file" | "heading" | "f8" | "author" | "loading";
  depth: number;
  label: string;
  subLabel?: string;
  headingLevel?: number;
  headingOrder?: number;
  folderPath?: string;
  fileId?: number;
  copyText?: string;
  sourcePath?: string;
  richHtml?: string;
  paragraphXml?: string[];
  searchResult?: SearchHit;
  hasChildren?: boolean;
};

export type CaptureInsertResult = {
  capturePath: string;
  marker: string;
  targetRelativePath: string;
};

export type CaptureTarget = {
  relativePath: string;
  absolutePath: string;
  exists: boolean;
  entryCount: number;
};

export type CaptureTargetPreview = {
  relativePath: string;
  absolutePath: string;
  exists: boolean;
  headingCount: number;
  headings: FileHeading[];
};

export type SidePreview = {
  title: string;
  subTitle?: string;
  text: string;
  richHtml?: string;
  headingLevel?: number | null;
  kind?: "heading" | "f8" | "author";
};
