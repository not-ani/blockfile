import type {
  FilePreview,
  FolderEntry,
  IndexSnapshot,
  IndexedFile,
  SearchHit,
  TreeRow,
} from "./types";
import {
  basename,
  f8RowKey,
  folderAncestors,
  headingChainForTarget,
  headingLevelLabel,
  headingRowKey,
} from "./utils";

type SnapshotIndex = {
  snap: IndexSnapshot;
  folderByPath: Map<string, FolderEntry>;
  foldersByParent: Map<string, FolderEntry[]>;
  filesByFolder: Map<string, IndexedFile[]>;
  fileById: Map<number, IndexedFile>;
};

type BuildTreeRowsArgs = {
  snapshotIndex: SnapshotIndex | null;
  searchMode: boolean;
  searchResults: SearchHit[];
  previewCache: Record<number, FilePreview>;
  expandedFolders: Set<string>;
  expandedFiles: Set<number>;
  collapsedHeadings: Set<string>;
};

export const buildSnapshotIndex = (snap: IndexSnapshot | null): SnapshotIndex | null => {
  if (!snap) return null;

  const folderByPath = new Map(snap.folders.map((folder) => [folder.path, folder]));

  const foldersByParent = new Map<string, FolderEntry[]>();
  for (const folder of snap.folders) {
    if (!folder.path) continue;
    const parent = folder.parentPath ?? "";
    const current = foldersByParent.get(parent) ?? [];
    current.push(folder);
    foldersByParent.set(parent, current);
  }
  for (const folders of foldersByParent.values()) {
    folders.sort((left, right) => left.name.localeCompare(right.name));
  }

  const filesByFolder = new Map<string, IndexedFile[]>();
  const fileById = new Map<number, IndexedFile>();
  for (const file of snap.files) {
    const current = filesByFolder.get(file.folderPath) ?? [];
    current.push(file);
    filesByFolder.set(file.folderPath, current);
    fileById.set(file.id, file);
  }
  for (const files of filesByFolder.values()) {
    files.sort((left, right) => left.fileName.localeCompare(right.fileName));
  }

  return {
    snap,
    folderByPath,
    foldersByParent,
    filesByFolder,
    fileById,
  };
};

export const buildTreeRows = (args: BuildTreeRowsArgs): TreeRow[] => {
  const {
    snapshotIndex,
    searchMode,
    searchResults,
    previewCache,
    expandedFolders,
    expandedFiles,
    collapsedHeadings,
  } = args;

  if (!snapshotIndex) return [];

  const { snap, folderByPath, foldersByParent, filesByFolder, fileById } = snapshotIndex;

  if (searchMode) {
    const rows: TreeRow[] = [];
    const seenKeys = new Set<string>();

    const pushRow = (row: TreeRow) => {
      if (seenKeys.has(row.key)) return;
      seenKeys.add(row.key);
      rows.push(row);
    };

    for (const result of searchResults) {
      const file = fileById.get(result.fileId);
      if (!file) continue;

      const ancestorPaths = folderAncestors(file.folderPath);
      ancestorPaths.forEach((ancestorPath, index) => {
        const folder = folderByPath.get(ancestorPath);
        const label = index === 0 ? basename(snap.rootPath) : folder?.name ?? basename(ancestorPath);
        const subLabel = index === 0 ? snap.rootPath : `${folder?.fileCount ?? 0} files`;
        pushRow({
          key: `search:folder:${ancestorPath || "__root__"}`,
          kind: "folder",
          depth: index,
          label,
          subLabel,
          folderPath: ancestorPath,
        });
      });

      const fileDepth = ancestorPaths.length;
      pushRow({
        key: `search:file:${file.id}`,
        kind: "file",
        depth: fileDepth,
        label: file.fileName,
        subLabel: file.relativePath,
        fileId: file.id,
        sourcePath: result.absolutePath,
        searchResult: result,
      });

      if (result.kind === "heading" && result.headingText) {
        const preview = previewCache[file.id];
        if (!preview) {
          pushRow({
            key: `search:loading:${file.id}`,
            kind: "loading",
            depth: fileDepth + 1,
            label: "Loading heading context...",
            fileId: file.id,
          });
          continue;
        }

        const targetHeading =
          preview.headings.find((entry) => entry.order === (result.headingOrder ?? -1)) ??
          preview.headings.find((entry) => entry.level === result.headingLevel && entry.text === result.headingText);

        if (!targetHeading) {
          pushRow({
            key: `search:heading:${file.id}:${result.headingOrder ?? result.headingText}`,
            kind: "heading",
            depth: fileDepth + 1,
            label: result.headingText,
            subLabel: headingLevelLabel(result.headingLevel),
            headingLevel: result.headingLevel ?? undefined,
            headingOrder: result.headingOrder ?? undefined,
            fileId: file.id,
            copyText: result.headingText,
            sourcePath: result.absolutePath,
            searchResult: result,
          });
          continue;
        }

        const chain = headingChainForTarget(preview.headings, targetHeading);
        chain.forEach((heading, index) => {
          pushRow({
            key: `search:heading:${file.id}:${heading.id}:${heading.order}:${heading.level}`,
            kind: "heading",
            depth: fileDepth + 1 + index,
            label: heading.text,
            subLabel: headingLevelLabel(heading.level),
            headingLevel: heading.level,
            headingOrder: heading.order,
            fileId: file.id,
            copyText: heading.copyText || heading.text,
            sourcePath: preview.absolutePath,
            searchResult: result,
            hasChildren: index < chain.length - 1,
          });
        });
        continue;
      }

      if (result.kind === "author" && result.headingText) {
        pushRow({
          key: `search:author:${file.id}:${result.headingOrder ?? 0}:${result.headingText}`,
          kind: "author",
          depth: fileDepth + 1,
          label: result.headingText,
          subLabel: "Author / Source",
          fileId: file.id,
          copyText: result.headingText,
          sourcePath: result.absolutePath,
          searchResult: result,
        });
      }
    }

    return rows;
  }

  const rows: TreeRow[] = [];

  const walkFolder = (folderPath: string, depth: number, isRoot = false) => {
    const folder = folderByPath.get(folderPath);
    const folderExpanded = expandedFolders.has(folderPath);

    rows.push({
      key: folderPath ? `folder:${folderPath}` : "folder:__root__",
      kind: "folder",
      depth,
      label: isRoot ? basename(snap.rootPath) : folder?.name ?? folderPath,
      subLabel: isRoot ? snap.rootPath : `${folder?.fileCount ?? 0} files`,
      folderPath,
    });

    if (!folderExpanded) return;

    for (const childFolder of foldersByParent.get(folderPath) ?? []) {
      walkFolder(childFolder.path, depth + 1, false);
    }

    for (const file of filesByFolder.get(folderPath) ?? []) {
      const fileExpanded = expandedFiles.has(file.id);
      rows.push({
        key: `file:${file.id}`,
        kind: "file",
        depth: depth + 1,
        label: file.fileName,
        subLabel: `${file.headingCount} headings`,
        fileId: file.id,
      });

      if (!fileExpanded) continue;

      const preview = previewCache[file.id];
      if (!preview) {
        rows.push({
          key: `loading:${file.id}`,
          kind: "loading",
          depth: depth + 2,
          label: "Loading headings...",
        });
        continue;
      }

      const orderedHeadings = [...preview.headings].sort((left, right) => left.order - right.order);
      const depthByHeadingIndex = new Array<number>(orderedHeadings.length).fill(0);
      const hasChildrenByHeadingIndex = new Array<boolean>(orderedHeadings.length).fill(false);
      const indexStack: number[] = [];

      for (let index = 0; index < orderedHeadings.length; index += 1) {
        const currentHeading = orderedHeadings[index];
        while (
          indexStack.length > 0 &&
          orderedHeadings[indexStack[indexStack.length - 1]].level >= currentHeading.level
        ) {
          indexStack.pop();
        }

        depthByHeadingIndex[index] = indexStack.length;
        if (indexStack.length > 0) {
          hasChildrenByHeadingIndex[indexStack[indexStack.length - 1]] = true;
        }
        indexStack.push(index);
      }

      const visibilityStack: { level: number; collapsedChain: boolean }[] = [];
      for (let index = 0; index < orderedHeadings.length; index += 1) {
        const currentHeading = orderedHeadings[index];
        while (
          visibilityStack.length > 0 &&
          visibilityStack[visibilityStack.length - 1].level >= currentHeading.level
        ) {
          visibilityStack.pop();
        }

        const key = headingRowKey(file.id, currentHeading);
        const hasChildren = hasChildrenByHeadingIndex[index];
        const parentCollapsed =
          visibilityStack.length > 0 && visibilityStack[visibilityStack.length - 1].collapsedChain;
        const selfCollapsed = collapsedHeadings.has(key);

        if (!parentCollapsed) {
          rows.push({
            key,
            kind: "heading",
            depth: depth + 2 + depthByHeadingIndex[index],
            label: currentHeading.text,
            subLabel: headingLevelLabel(currentHeading.level),
            headingLevel: currentHeading.level,
            headingOrder: currentHeading.order,
            fileId: file.id,
            copyText: currentHeading.copyText || currentHeading.text,
            sourcePath: preview.absolutePath,
            hasChildren,
          });
        }

        visibilityStack.push({
          level: currentHeading.level,
          collapsedChain: parentCollapsed || selfCollapsed,
        });
      }

      if (preview.headings.length === 0) {
        rows.push({
          key: `heading-empty:${file.id}`,
          kind: "heading",
          depth: depth + 2,
          label: "No headings detected",
          subLabel: "Heading",
          fileId: file.id,
          copyText: "",
          sourcePath: preview.absolutePath,
        });
      }

      preview.f8Cites.forEach((cite, index) => {
        rows.push({
          key: f8RowKey(file.id, cite, index),
          kind: "f8",
          depth: depth + 2,
          label: cite.text,
          subLabel: cite.styleLabel,
          fileId: file.id,
          copyText: cite.text,
          sourcePath: preview.absolutePath,
        });
      });
    }
  };

  walkFolder("", 0, true);
  return rows;
};
