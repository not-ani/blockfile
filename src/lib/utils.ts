import { invoke } from "@tauri-apps/api/core";
import type {
  FileHeading,
  FolderEntry,
  IndexSnapshot,
  IndexedFile,
} from "./types";
import { HEADING_LEVEL_LABELS } from "./constants";

export const invokeTyped = <T,>(command: string, args?: Record<string, unknown>) =>
  invoke<T>(command, args);

export const basename = (path: string) => {
  const segments = path.split(/[/\\]/).filter(Boolean);
  return segments.length > 0 ? segments[segments.length - 1] : path;
};

export const normalizeSlashes = (path: string) => path.replace(/\\/g, "/");

export const pathLooksAbsolute = (path: string) => /^[A-Za-z]:[/\\]|^\//.test(path);

export const folderFromRelativePath = (relativePath: string) =>
  relativePath.lastIndexOf("/") > -1
    ? relativePath.slice(0, relativePath.lastIndexOf("/"))
    : "";

export const headingRowKey = (fileId: number, heading: FileHeading) =>
  `heading:${fileId}:${heading.id}:${heading.level}:${heading.order}`;

export const f8RowKey = (fileId: number, cite: { order: number; styleLabel: string }, index: number) =>
  `f8:${fileId}:${index}:${cite.order}:${cite.styleLabel}`;

export const formatTime = (epochMs: number) => {
  if (epochMs <= 0) return "Never";
  return new Date(epochMs).toLocaleString();
};

export const headingLevelLabel = (level: number | null | undefined) => {
  if (!level) return "Heading";
  return HEADING_LEVEL_LABELS[level] ?? `H${level}`;
};

export const fallbackCopy = async (value: string) => {
  const area = document.createElement("textarea");
  area.value = value;
  area.style.position = "fixed";
  area.style.left = "-9999px";
  document.body.append(area);
  area.select();
  document.execCommand("copy");
  area.remove();
};

export const folderAncestors = (folderPath: string) => {
  const parts = folderPath ? folderPath.split("/") : [];
  const ancestors: string[] = [""];
  let current = "";
  for (const part of parts) {
    current = current ? `${current}/${part}` : part;
    ancestors.push(current);
  }
  return ancestors;
};

export const headingChainForTarget = (headings: FileHeading[], target: FileHeading) => {
  const ordered = [...headings].sort((left, right) => left.order - right.order);
  const stack: FileHeading[] = [];
  for (const heading of ordered) {
    while (stack.length > 0 && stack[stack.length - 1].level >= heading.level) {
      stack.pop();
    }
    stack.push(heading);
    if (
      heading.order === target.order &&
      heading.level === target.level &&
      heading.text === target.text
    ) {
      return [...stack];
    }
  }
  return [target];
};

export const mergeSnapshots = (snapshots: IndexSnapshot[]): IndexSnapshot => {
  if (snapshots.length === 0) {
    return {
      rootPath: "All indexed folders",
      indexedAtMs: 0,
      folders: [],
      files: [],
    };
  }

  const folders: FolderEntry[] = [];
  const files: IndexedFile[] = [];
  let latestIndexedMs = 0;

  snapshots.forEach((snapshotItem, index) => {
    latestIndexedMs = Math.max(latestIndexedMs, snapshotItem.indexedAtMs);
    const rootKey = `root-${index}`;
    const rootFolder = snapshotItem.folders.find((folder) => folder.path === "");

    folders.push({
      path: rootKey,
      name: basename(snapshotItem.rootPath),
      parentPath: "",
      depth: 1,
      fileCount: rootFolder?.fileCount ?? snapshotItem.files.length,
    });

    snapshotItem.folders
      .filter((folder) => folder.path !== "")
      .forEach((folder) => {
        const parentPath =
          folder.parentPath === null || folder.parentPath === ""
            ? rootKey
            : `${rootKey}/${folder.parentPath}`;
        folders.push({
          path: `${rootKey}/${folder.path}`,
          name: folder.name,
          parentPath,
          depth: folder.depth + 1,
          fileCount: folder.fileCount,
        });
      });

    snapshotItem.files.forEach((file) => {
      files.push({
        ...file,
        folderPath: file.folderPath ? `${rootKey}/${file.folderPath}` : rootKey,
      });
    });
  });

  return {
    rootPath: "All indexed folders",
    indexedAtMs: latestIndexedMs,
    folders,
    files,
  };
};
