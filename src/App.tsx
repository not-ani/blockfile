import { open } from "@tauri-apps/plugin-dialog";
import { openPath } from "@tauri-apps/plugin-opener";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";
import { batch, createEffect, createMemo, createSignal, onCleanup, onMount, untrack } from "solid-js";
import CaptureTargetPanel from "./components/CaptureTargetPanel.tsx";
import SidePreviewPane from "./components/SidePreviewPane.tsx";
import TopControls from "./components/TopControls.tsx";
import TreeView from "./components/TreeView.tsx";
import { ALL_ROOTS_KEY, TREE_OVERSCAN_ROWS, TREE_ROW_STRIDE_PX } from "./lib/constants";
import type {
  CaptureInsertResult,
  CaptureTarget,
  CaptureTargetPreview,
  FilePreview,
  IndexSnapshot,
  IndexProgress,
  IndexStats,
  RootSummary,
  SearchHit,
  SidePreview,
  TreeRow,
} from "./lib/types";
import {
  basename,
  fallbackCopy,
  folderFromRelativePath,
  headingLevelLabel,
  headingRowKey,
  invokeTyped,
  mergeSnapshots,
  normalizeSlashes,
  pathLooksAbsolute,
} from "./lib/utils";
import { buildSnapshotIndex, buildTreeRows } from "./lib/treeRows";

function App() {
  const LEFT_RAIL_DEFAULT_PX = 560;
  const LEFT_RAIL_MIN_PX = 420;
  const CAPTURE_TARGET_PREFS_KEY = "blockfile.captureTargetsByRoot.v1";

  const [roots, setRoots] = createSignal<RootSummary[]>([]);
  const [selectedRootPath, setSelectedRootPath] = createSignal("");
  const [snapshot, setSnapshot] = createSignal<IndexSnapshot | null>(null);
  const [allRootSnapshots, setAllRootSnapshots] = createSignal<IndexSnapshot[]>([]);
  const [previewCache, setPreviewCache] = createSignal<Record<number, FilePreview>>({});
  const [expandedFolders, setExpandedFolders] = createSignal<Set<string>>(new Set([""]));
  const [expandedFiles, setExpandedFiles] = createSignal<Set<number>>(new Set());
  const [collapsedHeadings, setCollapsedHeadings] = createSignal<Set<string>>(new Set());

  const [searchQuery, setSearchQuery] = createSignal("");
  const [searchResults, setSearchResults] = createSignal<SearchHit[]>([]);

  const [focusedNodeKey, setFocusedNodeKey] = createSignal("");
  const [sidePreview, setSidePreview] = createSignal<SidePreview | null>(null);
  const [captureByRowKey, setCaptureByRowKey] = createSignal<Record<string, CaptureInsertResult>>({});
  const [captureTargets, setCaptureTargets] = createSignal<CaptureTarget[]>([]);
  const [selectedCaptureTarget, setSelectedCaptureTarget] = createSignal("");
  const [selectedCaptureHeadingOrder, setSelectedCaptureHeadingOrder] = createSignal<number | null>(null);
  const [captureTargetPreview, setCaptureTargetPreview] = createSignal<CaptureTargetPreview | null>(null);
  const [headingPreviewHtmlCache, setHeadingPreviewHtmlCache] = createSignal<Record<string, string>>({});
  const [isLoadingCaptureTargets, setIsLoadingCaptureTargets] = createSignal(false);
  const [isLoadingCapturePreview, setIsLoadingCapturePreview] = createSignal(false);

  const [isIndexing, setIsIndexing] = createSignal(false);
  const [isSearching, setIsSearching] = createSignal(false);
  const [isLoadingSnapshot, setIsLoadingSnapshot] = createSignal(false);
  const [status, setStatus] = createSignal("Ready");
  const [indexProgress, setIndexProgress] = createSignal<IndexProgress | null>(null);
  const [copyToast, setCopyToast] = createSignal("");
  const [treeScrollTop, setTreeScrollTop] = createSignal(0);
  const [treeViewportHeight, setTreeViewportHeight] = createSignal(0);
  const [leftRailWidthPx, setLeftRailWidthPx] = createSignal(LEFT_RAIL_DEFAULT_PX);

  let treeRef: HTMLDivElement | undefined;
  let searchRequestSeq = 0;
  let previewWarmupSeq = 0;
  let headingPreviewSeq = 0;
  let focusScrollFrame = 0;
  let captureTargetsSeq = 0;
  let capturePreviewSeq = 0;
  let stopLeftRailResize: (() => void) | null = null;
  let stopIndexProgressListener: UnlistenFn | null = null;
  let searchInputRef: HTMLInputElement | undefined;
  let headingPreviewTimer = 0;
  let pendingFocusDelta = 0;
  let focusMoveFrame = 0;

  const selectedRoot = createMemo(() => roots().find((root) => root.path === selectedRootPath()) ?? null);
  const isAllRootsSelected = createMemo(() => selectedRootPath() === ALL_ROOTS_KEY);
  const activeSnapshot = createMemo(() =>
    isAllRootsSelected() ? mergeSnapshots(allRootSnapshots()) : snapshot(),
  );
  const captureRootPath = createMemo(() => (isAllRootsSelected() ? "" : selectedRootPath()));
  const searchMode = createMemo(() => searchQuery().trim().length >= 2);
  const normalizeCaptureTargetPath = (value: string) => normalizeSlashes(value).trim();

  const loadCaptureTargetPrefs = () => {
    try {
      const raw = localStorage.getItem(CAPTURE_TARGET_PREFS_KEY);
      if (!raw) return {} as Record<string, string>;
      const parsed: unknown = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
        return {} as Record<string, string>;
      }
      return parsed as Record<string, string>;
    } catch {
      return {} as Record<string, string>;
    }
  };

  const getPersistedCaptureTargetForRoot = (rootPath: string) => {
    const normalizedRoot = normalizeCaptureTargetPath(rootPath);
    const value = loadCaptureTargetPrefs()[normalizedRoot];
    return typeof value === "string" ? normalizeCaptureTargetPath(value) : "";
  };

  const persistCaptureTargetForRoot = (rootPath: string, targetPath: string) => {
    const normalizedRoot = normalizeCaptureTargetPath(rootPath);
    const normalizedTarget = normalizeCaptureTargetPath(targetPath);
    if (!normalizedRoot || !normalizedTarget) return;

    const prefs = loadCaptureTargetPrefs();
    prefs[normalizedRoot] = normalizedTarget;
    try {
      localStorage.setItem(CAPTURE_TARGET_PREFS_KEY, JSON.stringify(prefs));
    } catch {
      // Ignore storage write failures (e.g. restricted storage mode)
    }
  };

  const setCaptureTargetSelection = (targetPath: string, persist = true) => {
    const normalizedTarget = normalizeCaptureTargetPath(targetPath);
    setSelectedCaptureTarget(normalizedTarget);

    if (!persist || !normalizedTarget) return;
    const rootPath = captureRootPath();
    if (!rootPath) return;
    persistCaptureTargetForRoot(rootPath, normalizedTarget);
  };

  const selectedCaptureTargetMeta = createMemo(() => {
    const selectedPath = normalizeCaptureTargetPath(selectedCaptureTarget());
    if (!selectedPath) return null;

    return (
      captureTargets().find((target) => normalizeCaptureTargetPath(target.relativePath) === selectedPath) ?? null
    );
  });
  const captureTargetH1ToH4 = createMemo(() =>
    (captureTargetPreview()?.headings ?? []).filter((heading) => heading.level >= 1 && heading.level <= 4),
  );
  const activeRootLabel = createMemo(() =>
    isAllRootsSelected() ? "All indexed folders" : selectedRootPath() || "No folder selected",
  );
  const activeLastIndexedMs = createMemo(() =>
    isAllRootsSelected()
      ? roots().reduce((latest, root) => Math.max(latest, root.lastIndexedMs), 0)
      : selectedRoot()?.lastIndexedMs ?? 0,
  );

  const withFolderSet = (mutator: (set: Set<string>) => void) => {
    setExpandedFolders((current) => {
      const next = new Set(current);
      mutator(next);
      return next;
    });
  };

  const withFileSet = (mutator: (set: Set<number>) => void) => {
    setExpandedFiles((current) => {
      const next = new Set(current);
      mutator(next);
      return next;
    });
  };

  const withCollapsedHeadingSet = (mutator: (set: Set<string>) => void) => {
    setCollapsedHeadings((current) => {
      const next = new Set(current);
      mutator(next);
      return next;
    });
  };

  const copyText = async (text: string) => {
    try {
      if (typeof ClipboardItem !== "undefined" && navigator.clipboard?.write) {
        const escaped = text
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
        const html = `<pre style="white-space:pre-wrap">${escaped}</pre>`;
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/plain": new Blob([text], { type: "text/plain" }),
            "text/html": new Blob([html], { type: "text/html" }),
          }),
        ]);
      } else {
        await navigator.clipboard.writeText(text);
      }
    } catch {
      await fallbackCopy(text);
    }
    setCopyToast("Copied");
    setTimeout(() => setCopyToast(""), 1000);
  };

  const ensurePreviewLoaded = async (fileId: number) => {
    const cached = previewCache()[fileId];
    if (cached) return cached;

    const preview = await invokeTyped<FilePreview>("get_file_preview", { fileId });
    setPreviewCache((current) => ({ ...current, [fileId]: preview }));
    return preview;
  };

  const expandFolderAncestors = (folderPath: string) => {
    withFolderSet((next) => {
      next.add("");
      let current = folderPath;
      while (true) {
        next.add(current);
        if (!current) break;
        const split = current.lastIndexOf("/");
        current = split < 0 ? "" : current.slice(0, split);
      }
    });
  };

  const toggleFolder = (folderPath: string) => {
    withFolderSet((next) => {
      if (next.has(folderPath)) {
        next.delete(folderPath);
      } else {
        next.add(folderPath);
      }
    });
  };

  const toggleFile = async (fileId: number) => {
    const alreadyExpanded = expandedFiles().has(fileId);
    withFileSet((next) => {
      if (next.has(fileId)) {
        next.delete(fileId);
      } else {
        next.add(fileId);
      }
    });

    if (!alreadyExpanded) {
      await ensurePreviewLoaded(fileId);
    }
  };

  const toggleHeadingCollapse = (headingKey: string) => {
    withCollapsedHeadingSet((next) => {
      if (next.has(headingKey)) {
        next.delete(headingKey);
      } else {
        next.add(headingKey);
      }
    });
  };

  const syncTreeViewportState = () => {
    if (!treeRef) return;
    setTreeViewportHeight(treeRef.clientHeight);
    setTreeScrollTop(treeRef.scrollTop);
  };

  const keepFocusedRowInView = () => {
    if (!treeRef) return;
    const key = focusedNodeKey();
    if (!key) return;

    const rowIndex = treeRowIndexByKey().get(key) ?? -1;
    if (rowIndex < 0) return;

    const rowTop = rowIndex * TREE_ROW_STRIDE_PX;
    const rowBottom = rowTop + TREE_ROW_STRIDE_PX;
    const viewportTop = treeRef.scrollTop;
    const viewportBottom = viewportTop + treeRef.clientHeight;

    let nextScrollTop = viewportTop;
    if (rowTop < viewportTop) {
      nextScrollTop = rowTop;
    } else if (rowBottom > viewportBottom) {
      nextScrollTop = rowBottom - treeRef.clientHeight;
    }

    if (nextScrollTop !== viewportTop) {
      treeRef.scrollTop = nextScrollTop;
      setTreeScrollTop(nextScrollTop);
    }
  };

  const setSearchInputElement = (element: HTMLInputElement) => {
    searchInputRef = element;
  };

  const focusSearchField = () => {
    searchInputRef?.focus();
    searchInputRef?.select();
  };

  const targetIsTextEditable = (target: EventTarget | null) => {
    if (!(target instanceof HTMLElement)) return false;
    if (target.isContentEditable) return true;
    const editable = target.closest("input, textarea, select, [contenteditable=''], [contenteditable='true']");
    return editable !== null;
  };

  const scheduleFocusRelativeRow = (delta: number) => {
    pendingFocusDelta += delta;
    if (focusMoveFrame) return;

    focusMoveFrame = requestAnimationFrame(() => {
      focusMoveFrame = 0;
      const nextDelta = pendingFocusDelta;
      pendingFocusDelta = 0;
      if (nextDelta !== 0) {
        focusRelativeRow(nextDelta);
      }
    });
  };

  const loadHeadingPreviewHtml = async (row: TreeRow) => {
    if (row.kind !== "heading" || row.fileId === undefined || row.headingOrder === undefined) {
      return;
    }

    const cacheKey = `${row.fileId}:${row.headingOrder}`;
    const cached = headingPreviewHtmlCache()[cacheKey];
    if (cached) {
      setSidePreview((current) => {
        if (!current || focusedNodeKey() !== row.key) return current;
        if (current.richHtml === cached) return current;
        return { ...current, richHtml: cached };
      });
      return;
    }

    const requestId = ++headingPreviewSeq;
    try {
      const html = await invokeTyped<string>("get_heading_preview_html", {
        fileId: row.fileId,
        headingOrder: row.headingOrder,
      });
      if (requestId !== headingPreviewSeq) return;
      setHeadingPreviewHtmlCache((current) => ({ ...current, [cacheKey]: html }));
      setSidePreview((current) => {
        if (!current || focusedNodeKey() !== row.key) return current;
        return { ...current, richHtml: html };
      });
    } catch {
      // Keep plain text preview when rich preview extraction fails.
    }
  };

  const queueHeadingPreviewHtml = (rowKey: string) => {
    if (headingPreviewTimer) {
      window.clearTimeout(headingPreviewTimer);
      headingPreviewTimer = 0;
    }

    headingPreviewTimer = window.setTimeout(() => {
      headingPreviewTimer = 0;
      const activeRow = interactiveRowByKey().get(rowKey);
      if (!activeRow || activeRow.kind !== "heading") return;
      if (focusedNodeKey() !== rowKey) return;
      void loadHeadingPreviewHtml(activeRow);
    }, 90);
  };

  const startLeftRailResize = (event: MouseEvent) => {
    if (window.innerWidth < 1024) return;
    event.preventDefault();

    const startX = event.clientX;
    const initialWidth = leftRailWidthPx();
    const maxWidth = Math.max(LEFT_RAIL_MIN_PX, window.innerWidth - 520);

    const onMouseMove = (moveEvent: MouseEvent) => {
      const nextWidth = Math.min(
        maxWidth,
        Math.max(LEFT_RAIL_MIN_PX, initialWidth + (moveEvent.clientX - startX)),
      );
      setLeftRailWidthPx(nextWidth);
    };

    const onMouseUp = () => {
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
      window.removeEventListener("blur", onMouseUp);
      document.body.classList.remove("is-resizing-panels");
      stopLeftRailResize = null;
    };

    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
    window.addEventListener("blur", onMouseUp);
    document.body.classList.add("is-resizing-panels");
    stopLeftRailResize = onMouseUp;
  };

  const loadRoots = async () => {
    const loaded = await invokeTyped<RootSummary[]>("list_roots");
    const selected = selectedRootPath();
    const canKeepAllSelection = selected === ALL_ROOTS_KEY && loaded.length > 1;

    batch(() => {
      setRoots(loaded);

      if (
        loaded.length > 0 &&
        !canKeepAllSelection &&
        (!selected || !loaded.some((root) => root.path === selected))
      ) {
        setSelectedRootPath(loaded[0].path);
      }

      if (selected === ALL_ROOTS_KEY && loaded.length <= 1) {
        setSelectedRootPath(loaded[0]?.path ?? "");
      }

      if (loaded.length === 0) {
        setSelectedRootPath("");
        setSnapshot(null);
        setAllRootSnapshots([]);
      }
    });
  };

  const loadSnapshot = async (rootPath: string) => {
    if (!rootPath) return;
    setIsLoadingSnapshot(true);
    try {
      const nextSnapshot = await invokeTyped<IndexSnapshot>("get_index_snapshot", { path: rootPath });

      batch(() => {
        setSnapshot(nextSnapshot);
        setAllRootSnapshots([]);
        setExpandedFolders(new Set([""]));
        setExpandedFiles(new Set<number>());
        setCollapsedHeadings(new Set<string>());
        setPreviewCache({});
        setCaptureByRowKey({});
        setFocusedNodeKey("folder:__root__");
        setTreeScrollTop(0);
      });

      if (treeRef) {
        treeRef.scrollTop = 0;
      }

      setTimeout(() => {
        treeRef?.focus();
        syncTreeViewportState();
      }, 0);
    } catch (error) {
      setStatus(`Could not load index snapshot: ${String(error)}`);
      setSnapshot(null);
    } finally {
      setIsLoadingSnapshot(false);
    }
  };

  const loadAllSnapshots = async () => {
    const rootPaths = roots().map((root) => root.path);
    if (rootPaths.length === 0) {
      setAllRootSnapshots([]);
      setSnapshot(null);
      return;
    }

    setIsLoadingSnapshot(true);
    try {
      const results = await Promise.allSettled(
        rootPaths.map((path) => invokeTyped<IndexSnapshot>("get_index_snapshot", { path })),
      );

      const snapshots = results
        .filter((result): result is PromiseFulfilledResult<IndexSnapshot> => result.status === "fulfilled")
        .map((result) => result.value);

      const failed = results.length - snapshots.length;
      batch(() => {
        setAllRootSnapshots(snapshots);
        setSnapshot(null);
        setExpandedFolders(new Set([""]));
        setExpandedFiles(new Set<number>());
        setCollapsedHeadings(new Set<string>());
        setPreviewCache({});
        setCaptureByRowKey({});
        setFocusedNodeKey("folder:__root__");
        setTreeScrollTop(0);
      });

      if (treeRef) {
        treeRef.scrollTop = 0;
      }

      if (failed > 0) {
        setStatus(`Loaded ${snapshots.length} roots. ${failed} roots failed to load.`);
      }

      setTimeout(() => {
        treeRef?.focus();
        syncTreeViewportState();
      }, 0);
    } catch (error) {
      setStatus(`Could not load snapshots for all roots: ${String(error)}`);
      setAllRootSnapshots([]);
    } finally {
      setIsLoadingSnapshot(false);
    }
  };

  const captureEntryKey = (rowKey: string, targetRelativePath: string, headingOrder: number | null) =>
    `${targetRelativePath}::${headingOrder ?? "root"}::${rowKey}`;

  const loadCaptureTargets = async (rootPath: string) => {
    if (!rootPath) {
      return;
    }

    const requestId = ++captureTargetsSeq;
    setIsLoadingCaptureTargets(true);
    try {
      const loaded = await invokeTyped<CaptureTarget[]>("list_capture_targets", { rootPath });
      if (requestId !== captureTargetsSeq) return;

      const current = normalizeCaptureTargetPath(selectedCaptureTarget());
      const persisted = getPersistedCaptureTargetForRoot(rootPath);
      const preferredTarget = current || persisted;

      const merged = [...loaded];
      if (
        preferredTarget &&
        !merged.some((entry) => normalizeCaptureTargetPath(entry.relativePath) === preferredTarget)
      ) {
        const root = normalizeSlashes(rootPath).replace(/[\/]+$/, "");
        const absolutePath = pathLooksAbsolute(preferredTarget)
          ? preferredTarget
          : `${root}/${preferredTarget.replace(/^[\/\\]+/, "")}`;
        merged.unshift({
          relativePath: preferredTarget,
          absolutePath,
          exists: false,
          entryCount: 0,
        });
      }

      const defaultTarget =
        !preferredTarget && merged.length > 0 ? normalizeCaptureTargetPath(merged[0].relativePath) : "";
      const nextTarget = preferredTarget || defaultTarget;
      const shouldPersistDefault = !current && !persisted && !!defaultTarget;

      batch(() => {
        setCaptureTargets(merged);

        if (!nextTarget) {
          setCaptureTargetSelection("", false);
          return;
        }

        setCaptureTargetSelection(nextTarget, shouldPersistDefault);
      });
    } catch (error) {
      if (requestId === captureTargetsSeq) {
        setStatus(`Could not load capture target files: ${String(error)}`);
      }
    } finally {
      if (requestId === captureTargetsSeq) {
        setIsLoadingCaptureTargets(false);
      }
    }
  };

  const loadCaptureTargetPreview = async (rootPath: string, targetPath: string) => {
    if (!rootPath || !targetPath) {
      setCaptureTargetPreview(null);
      return;
    }

    const requestId = ++capturePreviewSeq;
    setIsLoadingCapturePreview(true);
    try {
      const preview = await invokeTyped<CaptureTargetPreview>("get_capture_target_preview", {
        rootPath,
        targetPath,
      });
      if (requestId !== capturePreviewSeq) return;
      setCaptureTargetPreview(preview);
    } catch (error) {
      if (requestId === capturePreviewSeq) {
        setStatus(`Could not load target preview: ${String(error)}`);
        setCaptureTargetPreview(null);
      }
    } finally {
      if (requestId === capturePreviewSeq) {
        setIsLoadingCapturePreview(false);
      }
    }
  };

  const deleteCaptureHeading = async (headingOrder: number) => {
    const rootPath = captureRootPath();
    const targetPath = selectedCaptureTarget();
    if (!rootPath || !targetPath) {
      setStatus("Select a destination docx file first.");
      return;
    }

    const requestId = ++capturePreviewSeq;
    setIsLoadingCapturePreview(true);
    try {
      const preview = await invokeTyped<CaptureTargetPreview>("delete_capture_heading", {
        rootPath,
        targetPath,
        headingOrder,
      });
      if (requestId !== capturePreviewSeq) return;
      batch(() => {
        setCaptureTargetPreview(preview);
        if (selectedCaptureHeadingOrder() === headingOrder) {
          setSelectedCaptureHeadingOrder(null);
        }
      });
      setStatus(`Deleted heading from ${basename(preview.absolutePath)}`);
    } catch (error) {
      if (requestId === capturePreviewSeq) {
        setStatus(`Could not delete heading: ${String(error)}`);
      }
    } finally {
      if (requestId === capturePreviewSeq) {
        setIsLoadingCapturePreview(false);
      }
    }
  };

  const moveCaptureHeading = async (sourceHeadingOrder: number, targetHeadingOrder: number) => {
    const rootPath = captureRootPath();
    const targetPath = selectedCaptureTarget();
    if (!rootPath || !targetPath) {
      setStatus("Select a destination docx file first.");
      return;
    }

    const requestId = ++capturePreviewSeq;
    setIsLoadingCapturePreview(true);
    try {
      const preview = await invokeTyped<CaptureTargetPreview>("move_capture_heading", {
        rootPath,
        targetPath,
        sourceHeadingOrder,
        targetHeadingOrder,
      });
      if (requestId !== capturePreviewSeq) return;
      batch(() => {
        setCaptureTargetPreview(preview);
        setSelectedCaptureHeadingOrder(sourceHeadingOrder);
      });
      setStatus(`Moved heading under target in ${basename(preview.absolutePath)}`);
    } catch (error) {
      if (requestId === capturePreviewSeq) {
        setStatus(`Could not move heading: ${String(error)}`);
      }
    } finally {
      if (requestId === capturePreviewSeq) {
        setIsLoadingCapturePreview(false);
      }
    }
  };

  const createCaptureTarget = async () => {
    const rootPath = captureRootPath();
    if (!rootPath) {
      setStatus("Select an individual root folder first.");
      return;
    }

    const nextName = window.prompt(
      "New capture file path (relative to root OR absolute):",
      "BlockFile-Captures-2.docx",
    )?.trim();
    if (!nextName) {
      return;
    }

    const normalized = nextName.toLowerCase().endsWith(".docx") ? nextName : `${nextName}.docx`;
    setCaptureTargetSelection(normalized);
    void loadCaptureTargets(rootPath);
    void loadCaptureTargetPreview(rootPath, normalized);
    setStatus(`Target file selected: ${normalized}`);
  };

  const selectCaptureTargetFromFilesystem = async () => {
    const rootPath = captureRootPath();
    if (!rootPath) {
      setStatus("Select an individual root folder first.");
      return;
    }

    setStatus("Opening target file picker...");
    let selectedPath: string | null = null;

    try {
      const selected: unknown = await open({
        directory: false,
        multiple: false,
        defaultPath: rootPath,
        title: "Choose destination .docx anywhere",
        filters: [{ name: "Word Document", extensions: ["docx"] }],
      });

      if (typeof selected === "string") {
        selectedPath = selected;
      } else if (Array.isArray(selected)) {
        selectedPath = selected.find((entry: unknown): entry is string => typeof entry === "string") ?? null;
      }
    } catch (error) {
      setStatus(`File picker failed: ${String(error)}`);
      return;
    }

    if (!selectedPath) {
      setStatus("Target file selection cancelled.");
      return;
    }

    const normalizedBase = normalizeSlashes(selectedPath);
    const normalized = normalizedBase.toLowerCase().endsWith(".docx") ? normalizedBase : `${normalizedBase}.docx`;
    setCaptureTargetSelection(normalized);
    void loadCaptureTargets(rootPath);
    void loadCaptureTargetPreview(rootPath, normalized);
    setStatus(`Target file selected: ${normalized}`);
  };

  const runIndex = async (rootPath: string) => {
    if (!rootPath) return;
    setIndexProgress(null);
    setIsIndexing(true);
    setStatus(`Indexing ${rootPath} (all subfolders)...`);
    try {
      const stats = await invokeTyped<IndexStats>("index_root", { path: rootPath });
      setStatus(
        `Indexed ${stats.scanned} docx. Updated ${stats.updated}, skipped ${stats.skipped}, removed ${stats.removed}.`,
      );
      await loadRoots();
      await loadSnapshot(rootPath);
    } catch (error) {
      setStatus(`Indexing failed: ${String(error)}`);
    } finally {
      setIndexProgress(null);
      setIsIndexing(false);
    }
  };

  const runIndexForSelection = async () => {
    if (isAllRootsSelected()) {
      const targetRoots = roots().map((root) => root.path);
      if (targetRoots.length === 0) return;

      setIndexProgress(null);
      setIsIndexing(true);
      try {
        let scanned = 0;
        let updated = 0;
        let skipped = 0;
        let removed = 0;
        for (const path of targetRoots) {
          const stats = await invokeTyped<IndexStats>("index_root", { path });
          scanned += stats.scanned;
          updated += stats.updated;
          skipped += stats.skipped;
          removed += stats.removed;
        }
        setStatus(
          `Reindexed ${targetRoots.length} folders. Scanned ${scanned} docx, updated ${updated}, skipped ${skipped}, removed ${removed}.`,
        );
        await loadRoots();
        await loadAllSnapshots();
      } catch (error) {
        setStatus(`Bulk reindex failed: ${String(error)}`);
      } finally {
        setIndexProgress(null);
        setIsIndexing(false);
      }
      return;
    }

    await runIndex(selectedRootPath());
  };

  const addFolder = async () => {
    setStatus("Opening folder picker...");

    let selectedPath: string | null = null;
    try {
      const selected: unknown = await open({
        directory: true,
        multiple: false,
        title: "Choose a folder to index",
      });
      if (typeof selected === "string") {
        selectedPath = selected;
      } else if (Array.isArray(selected)) {
        selectedPath = selected.find((entry: unknown): entry is string => typeof entry === "string") ?? null;
      }
    } catch (error) {
      setStatus(`Folder picker failed: ${String(error)}`);
    }

    if (!selectedPath) {
      const manual = window.prompt("Paste absolute folder path:", "")?.trim();
      if (!manual) {
        setStatus("Folder selection cancelled.");
        return;
      }
      selectedPath = manual;
    }

    try {
      const canonicalPath = await invokeTyped<string>("add_root", { path: selectedPath });
      setSelectedRootPath(canonicalPath);
      await loadRoots();
      await runIndex(canonicalPath);
    } catch (error) {
      setStatus(`Could not add folder index: ${String(error)}`);
    }
  };

  const snapshotIndex = createMemo(() => buildSnapshotIndex(activeSnapshot()));

  const treeRows = createMemo<TreeRow[]>(() =>
    buildTreeRows({
      snapshotIndex: snapshotIndex(),
      searchMode: searchMode(),
      searchResults: searchResults(),
      previewCache: previewCache(),
      expandedFolders: expandedFolders(),
      expandedFiles: expandedFiles(),
      collapsedHeadings: collapsedHeadings(),
    }),
  );

  const virtualWindow = createMemo(() => {
    const rows = treeRows();
    const viewportHeight = Math.max(treeViewportHeight(), TREE_ROW_STRIDE_PX);
    const rawStart = Math.floor(treeScrollTop() / TREE_ROW_STRIDE_PX) - TREE_OVERSCAN_ROWS;
    const start = Math.min(rows.length, Math.max(0, rawStart));
    const end = Math.min(
      rows.length,
      Math.max(start, Math.ceil((treeScrollTop() + viewportHeight) / TREE_ROW_STRIDE_PX) + TREE_OVERSCAN_ROWS),
    );

    return {
      start,
      end,
      topSpacerPx: start * TREE_ROW_STRIDE_PX,
      bottomSpacerPx: Math.max(0, (rows.length - end) * TREE_ROW_STRIDE_PX),
    };
  });

  const visibleTreeRows = createMemo<TreeRow[]>(() => {
    const rows = treeRows();
    const { start, end } = virtualWindow();
    return rows.slice(start, end);
  });

  const treeRowIndexByKey = createMemo(() => {
    const indexByKey = new Map<string, number>();
    treeRows().forEach((row, index) => indexByKey.set(row.key, index));
    return indexByKey;
  });

  const interactiveRows = createMemo(() => treeRows().filter((row) => row.kind !== "loading"));

  const interactiveRowByKey = createMemo(() => {
    const rowsByKey = new Map<string, TreeRow>();
    interactiveRows().forEach((row) => rowsByKey.set(row.key, row));
    return rowsByKey;
  });

  const interactiveRowIndexByKey = createMemo(() => {
    const indexByKey = new Map<string, number>();
    interactiveRows().forEach((row, index) => indexByKey.set(row.key, index));
    return indexByKey;
  });

  const focusedContentRow = () => {
    const row = interactiveRowByKey().get(focusedNodeKey()) ?? null;
    if (!row) return null;
    if (!row.copyText || !row.sourcePath) return null;
    if (row.kind !== "heading" && row.kind !== "f8" && row.kind !== "author") return null;
    return row;
  };

  const applyPreviewFromRow = (row: TreeRow) => {
    if ((row.kind === "heading" || row.kind === "f8" || row.kind === "author") && row.copyText) {
      const headingCacheKey =
        row.kind === "heading" && row.fileId !== undefined && row.headingOrder !== undefined
          ? `${row.fileId}:${row.headingOrder}`
          : "";
      setSidePreview({
        title: row.kind === "heading" ? row.label : row.kind === "f8" ? "F8 Cite" : "Author / Source",
        subTitle: row.subLabel,
        text: row.copyText,
        headingLevel: row.kind === "heading" ? row.headingLevel ?? null : null,
        kind: row.kind,
        richHtml: headingCacheKey ? headingPreviewHtmlCache()[headingCacheKey] : undefined,
      });

      if (row.kind === "heading") {
        queueHeadingPreviewHtml(row.key);
      }
    }
  };

  const insertRowIntoCapture = async (row: TreeRow) => {
    if (!row.copyText || !row.sourcePath) return null;
    const rootPath = captureRootPath();
    if (!rootPath) {
      setStatus("Select an individual root folder first to insert.");
      return null;
    }

    const targetPath = selectedCaptureTarget();
    if (!targetPath) {
      setStatus("Select a destination docx file first.");
      return null;
    }

    const insertionHeadingOrder = selectedCaptureHeadingOrder();
    const captureKey = captureEntryKey(row.key, targetPath, insertionHeadingOrder);
    const existing = captureByRowKey()[captureKey];
    if (existing) {
      return existing;
    }

    try {
      const inserted = await invokeTyped<CaptureInsertResult>("insert_capture", {
        rootPath,
        sourcePath: row.sourcePath,
        sectionTitle: row.label,
        content: row.copyText,
        targetPath,
        headingLevel: row.kind === "heading" ? row.headingLevel ?? null : null,
        headingOrder: row.kind === "heading" ? row.headingOrder ?? null : null,
        selectedTargetHeadingOrder: insertionHeadingOrder,
      });
      setCaptureByRowKey((current) => ({ ...current, [captureKey]: inserted }));
      void loadCaptureTargetPreview(rootPath, inserted.targetRelativePath || targetPath);
      setStatus(`Inserted into ${basename(inserted.capturePath)} [${inserted.marker}]`);
      return inserted;
    } catch (error) {
      setStatus(`Insert failed: ${String(error)}`);
      return null;
    }
  };

  const insertFocusedIntoCapture = async () => {
    const row = focusedContentRow();
    if (!row) {
      setStatus("Focus a heading/F8/author line first.");
      return;
    }

    const inserted = await insertRowIntoCapture(row);
    if (inserted) {
      setCopyToast(`Inserted ${inserted.marker}`);
      setTimeout(() => setCopyToast(""), 1200);
    }
  };

  const openCaptureForFocused = async () => {
    const row = focusedContentRow();
    if (!row) {
      setStatus("Focus a heading/F8/author line first.");
      return;
    }

    const targetPath = selectedCaptureTarget();
    if (!targetPath) {
      setStatus("Select a destination docx file first.");
      return;
    }

    const captureKey = captureEntryKey(row.key, targetPath, selectedCaptureHeadingOrder());
    let inserted: CaptureInsertResult | null = captureByRowKey()[captureKey] ?? null;
    if (!inserted) {
      inserted = await insertRowIntoCapture(row);
    }
    if (!inserted) return;

    await openPath(inserted.capturePath);
    setStatus(`Opened capture file: ${basename(inserted.capturePath)}`);
  };

  const activateRow = async (row: TreeRow, fromKeyboard = false) => {
    setFocusedNodeKey(row.key);

    if (row.kind === "folder" && row.folderPath !== undefined) {
      toggleFolder(row.folderPath);
      return;
    }

    if (row.kind === "file" && row.fileId !== undefined) {
      await toggleFile(row.fileId);
      return;
    }

    if (row.kind === "heading" && row.hasChildren && fromKeyboard) {
      toggleHeadingCollapse(row.key);
      applyPreviewFromRow(row);
      return;
    }

    if ((row.kind === "heading" || row.kind === "f8" || row.kind === "author") && row.copyText) {
      applyPreviewFromRow(row);
      if (!fromKeyboard) {
        await copyText(row.copyText);
      }
    }
  };

  const focusRelativeRow = (delta: number) => {
    const rows = interactiveRows();
    if (rows.length === 0) return;

    const currentKey = focusedNodeKey();
    const currentIndex = interactiveRowIndexByKey().get(currentKey) ?? -1;
    const nextIndex = Math.min(rows.length - 1, Math.max(0, (currentIndex < 0 ? 0 : currentIndex) + delta));
    const row = rows[nextIndex];
    setFocusedNodeKey(row.key);
    applyPreviewFromRow(row);
  };

  const copyFocusedRow = async () => {
    const row = focusedContentRow();
    if (!row) return;
    applyPreviewFromRow(row);
    await copyText(row.copyText ?? "");
  };

  const onTreeKeyDown = (event: KeyboardEvent) => {
    if (event.defaultPrevented) return;

    if (event.key === "ArrowDown") {
      event.preventDefault();
      scheduleFocusRelativeRow(1);
      return;
    }
    if (event.key === "ArrowUp") {
      event.preventDefault();
      scheduleFocusRelativeRow(-1);
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      const row = interactiveRowByKey().get(focusedNodeKey());
      if (row) {
        if (searchMode() && row.searchResult) {
          void openSearchResult(row.searchResult);
        } else {
          void activateRow(row, true);
        }
      }
      return;
    }
    if (event.key === " " || event.key === "Spacebar") {
      event.preventDefault();
      void insertFocusedIntoCapture();
      return;
    }
    if (event.key === "c" || event.key === "C") {
      event.preventDefault();
      void copyFocusedRow();
      return;
    }
    if (event.key === "o" || event.key === "O") {
      event.preventDefault();
      void openCaptureForFocused();
    }
  };

  const openSearchResult = async (result: SearchHit) => {
    const folderPath = folderFromRelativePath(result.relativePath);
    expandFolderAncestors(folderPath);

    withFileSet((set) => set.add(result.fileId));
    const preview = await ensurePreviewLoaded(result.fileId);
    setFocusedNodeKey(searchMode() ? `search:file:${result.fileId}` : `file:${result.fileId}`);

    if (result.kind === "heading") {
      const heading = preview.headings.find(
        (entry) =>
          entry.order === (result.headingOrder ?? -1) ||
          (entry.level === result.headingLevel && entry.text === result.headingText),
      );
      if (heading) {
        const key = searchMode()
          ? `search:heading:${result.fileId}:${heading.id}:${heading.order}:${heading.level}`
          : headingRowKey(result.fileId, heading);
        setFocusedNodeKey(key);
        setSidePreview({
          title: heading.text,
          subTitle: headingLevelLabel(heading.level),
          text: heading.copyText,
        });
      }
      return;
    }

    if (result.kind === "author" && result.headingText) {
      if (searchMode()) {
        setFocusedNodeKey(`search:author:${result.fileId}:${result.headingOrder ?? 0}:${result.headingText}`);
      }
      setSidePreview({
        title: "Author / Source",
        subTitle: result.fileName,
        text: result.headingText,
      });
    }
  };

  createEffect(() => {
    const rootPath = selectedRootPath();
    if (!rootPath) return;
    if (rootPath === ALL_ROOTS_KEY) {
      void loadAllSnapshots();
      return;
    }
    void loadSnapshot(rootPath);
  });

  createEffect(() => {
    const rootPath = captureRootPath();
    if (!rootPath) {
      setCaptureTargets([]);
      return;
    }
    void loadCaptureTargets(rootPath);
  });

  createEffect(() => {
    const rootPath = captureRootPath();
    const targetPath = selectedCaptureTarget();
    if (!rootPath || !targetPath) {
      setCaptureTargetPreview(null);
      setSelectedCaptureHeadingOrder(null);
      return;
    }
    void loadCaptureTargetPreview(rootPath, targetPath);
  });

  createEffect(() => {
    const selectedHeadingOrder = selectedCaptureHeadingOrder();
    if (selectedHeadingOrder === null) return;

    const exists = captureTargetH1ToH4().some((heading) => heading.order === selectedHeadingOrder);
    if (!exists) {
      setSelectedCaptureHeadingOrder(null);
    }
  });

  createEffect(() => {
    treeRows().length;
    queueMicrotask(syncTreeViewportState);
  });

  createEffect(() => {
    const rootPath = selectedRootPath();
    const searchRootPath = rootPath === ALL_ROOTS_KEY ? undefined : rootPath;
    const query = searchQuery().trim();
    if ((!rootPath && roots().length === 0) || query.length < 2) {
      setSearchResults([]);
      setIsSearching(false);
      return;
    }

    setIsSearching(true);
    const requestId = ++searchRequestSeq;
    const timer = setTimeout(() => {
      void invokeTyped<SearchHit[]>("search_index", {
        query,
        rootPath: searchRootPath,
        limit: 120,
      })
        .then((results) => {
          if (requestId === searchRequestSeq) {
            setSearchResults(results);
          }
        })
        .catch((error) => setStatus(`Search failed: ${String(error)}`))
        .finally(() => {
          if (requestId === searchRequestSeq) {
            setIsSearching(false);
          }
        });
    }, 150);

    onCleanup(() => clearTimeout(timer));
  });

  createEffect(() => {
    if (!searchMode()) return;
    const ids = Array.from(new Set(searchResults().map((result) => result.fileId))).slice(0, 48);
    if (ids.length === 0) return;

    const requestId = ++previewWarmupSeq;
    const cache = untrack(() => previewCache());
    const missing = ids.filter((id) => !cache[id]);
    if (missing.length === 0) return;

    const loadBatch = async (batchIds: number[]) => {
      await Promise.allSettled(batchIds.map((id) => ensurePreviewLoaded(id)));
    };

    void (async () => {
      const immediate = missing.slice(0, 12);
      const deferred = missing.slice(12);

      for (let index = 0; index < immediate.length; index += 4) {
        if (requestId !== previewWarmupSeq) return;
        await loadBatch(immediate.slice(index, index + 4));
      }

      for (let index = 0; index < deferred.length; index += 4) {
        if (requestId !== previewWarmupSeq) return;
        await new Promise<void>((resolve) => setTimeout(resolve, 48));
        if (requestId !== previewWarmupSeq) return;
        await loadBatch(deferred.slice(index, index + 4));
      }
    })();
  });

  createEffect(() => {
    focusedNodeKey();
    treeRows();

    if (focusScrollFrame) {
      cancelAnimationFrame(focusScrollFrame);
    }

    focusScrollFrame = requestAnimationFrame(() => {
      focusScrollFrame = 0;
      keepFocusedRowInView();
    });
  });

  onMount(() => {
    const onResize = () => syncTreeViewportState();
    const onGlobalKeyDown = (event: KeyboardEvent) => {
      const key = event.key;

      if ((event.metaKey || event.ctrlKey) && key.toLowerCase() === "k") {
        event.preventDefault();
        focusSearchField();
        return;
      }

      if (!event.metaKey && !event.ctrlKey && !event.altKey && key === "/") {
        event.preventDefault();
        focusSearchField();
        return;
      }

      if (!event.metaKey && !event.ctrlKey && !event.altKey && (key === "ArrowDown" || key === "ArrowUp")) {
        if (event.defaultPrevented) return;
        if (targetIsTextEditable(event.target)) return;
        if (treeRef && event.target instanceof Node && treeRef.contains(event.target)) {
          return;
        }

        event.preventDefault();
        treeRef?.focus();
        scheduleFocusRelativeRow(key === "ArrowDown" ? 1 : -1);
      }
    };

    let listenerDisposed = false;
    void listen<IndexProgress>("index-progress", (event) => {
      if (!isIndexing()) return;
      setIndexProgress(event.payload);
    })
      .then((unlisten) => {
        if (listenerDisposed) {
          unlisten();
          return;
        }
        stopIndexProgressListener = unlisten;
      })
      .catch((error) => {
        setStatus(`Could not subscribe to indexing progress: ${String(error)}`);
      });

    window.addEventListener("resize", onResize);
    window.addEventListener("keydown", onGlobalKeyDown, true);

    onCleanup(() => {
      listenerDisposed = true;
      stopIndexProgressListener?.();
      stopIndexProgressListener = null;
      stopLeftRailResize?.();
      window.removeEventListener("resize", onResize);
      window.removeEventListener("keydown", onGlobalKeyDown, true);
      if (headingPreviewTimer) {
        window.clearTimeout(headingPreviewTimer);
        headingPreviewTimer = 0;
      }
      if (focusMoveFrame) {
        cancelAnimationFrame(focusMoveFrame);
        focusMoveFrame = 0;
      }
      pendingFocusDelta = 0;
      if (focusScrollFrame) {
        cancelAnimationFrame(focusScrollFrame);
      }
    });

    queueMicrotask(syncTreeViewportState);

    void loadRoots().catch((error) => {
      setStatus(`Failed to load roots: ${String(error)}`);
    });
  });

  const setTreeElement = (element: HTMLDivElement) => {
    treeRef = element;
    syncTreeViewportState();
  };

  return (
    <div class="h-screen bg-[#0a0a0a]">
      <div class="flex h-full w-full flex-col px-6 py-5">
        <header class="mb-6">
          <div class="flex items-center gap-3">
            <div class="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-to-br from-blue-500 to-emerald-500 shadow-lg shadow-blue-500/20">
              <svg class="h-5 w-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" />
              </svg>
            </div>
            <div>
              <h1 class="text-xl font-semibold tracking-tight text-white">BlockFile</h1>
              <p class="text-sm text-neutral-500">Document indexer and capture tool</p>
            </div>
          </div>
        </header>

        <main class="vercel-card flex min-h-0 flex-1 flex-col overflow-hidden">
          <TopControls
            activeLastIndexedMs={activeLastIndexedMs}
            activeRootLabel={activeRootLabel}
            addFolder={addFolder}
            copyToast={copyToast}
            isIndexing={isIndexing}
            isSearching={isSearching}
            roots={roots}
            runIndexForSelection={runIndexForSelection}
            searchQuery={searchQuery}
            selectedRootPath={selectedRootPath}
            indexProgress={indexProgress}
            setSearchInputRef={setSearchInputElement}
            setSearchQuery={setSearchQuery}
            setSelectedRootPath={setSelectedRootPath}
            status={status}
          />

          <div class="vercel-divider my-5" />

          <div class="grid min-h-0 flex-1 gap-5 xl:grid-cols-[minmax(0,1fr)_420px]">
            <section class="vercel-card min-h-0 border-0 bg-neutral-950/50 p-4">
              <div class="workspace-split h-full min-h-0" style={{ "--left-rail-width": `${leftRailWidthPx()}px` }}>
                <aside class="vercel-card min-h-0 border-neutral-800/70 bg-neutral-950/60 p-4">
                  <CaptureTargetPanel
                    captureRootPath={captureRootPath}
                    captureTargetH1ToH4={captureTargetH1ToH4}
                    captureTargetPreview={captureTargetPreview}
                    captureTargets={captureTargets}
                    createCaptureTarget={createCaptureTarget}
                    deleteCaptureHeading={deleteCaptureHeading}
                    isAllRootsSelected={isAllRootsSelected}
                    isLoadingCapturePreview={isLoadingCapturePreview}
                    isLoadingCaptureTargets={isLoadingCaptureTargets}
                    moveCaptureHeading={moveCaptureHeading}
                    selectCaptureTargetFromFilesystem={selectCaptureTargetFromFilesystem}
                    selectedCaptureHeadingOrder={selectedCaptureHeadingOrder}
                    selectedCaptureTarget={selectedCaptureTarget}
                    selectedCaptureTargetMeta={selectedCaptureTargetMeta}
                    setSelectedCaptureHeadingOrder={setSelectedCaptureHeadingOrder}
                    setSelectedCaptureTarget={setCaptureTargetSelection}
                  />
                </aside>

                <button
                  aria-label="Resize insert preview panel"
                  class="panel-resize-handle hidden lg:flex"
                  onMouseDown={startLeftRailResize}
                  title="Drag to resize"
                  type="button"
                />

                <div class="vercel-card min-h-0 border-neutral-800/70 bg-neutral-950/30">
                  <TreeView
                    activateRow={activateRow}
                    applyPreviewFromRow={applyPreviewFromRow}
                    collapsedHeadings={collapsedHeadings}
                    expandedFiles={expandedFiles}
                    expandedFolders={expandedFolders}
                    focusedNodeKey={focusedNodeKey}
                    isLoadingSnapshot={isLoadingSnapshot}
                    isSearching={isSearching}
                    onTreeKeyDown={onTreeKeyDown}
                    onTreeScroll={setTreeScrollTop}
                    openSearchResult={openSearchResult}
                    searchMode={searchMode}
                    selectedRootPath={selectedRootPath}
                    setFocusedNodeKey={setFocusedNodeKey}
                    setTreeRef={setTreeElement}
                    treeRowsLength={() => treeRows().length}
                    virtualWindow={virtualWindow}
                    visibleTreeRows={visibleTreeRows}
                  />
                </div>
              </div>
            </section>

            <SidePreviewPane sidePreview={sidePreview} />
          </div>
        </main>
      </div>
    </div>
  );
}

export default App;
