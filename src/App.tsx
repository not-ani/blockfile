import { open } from "@tauri-apps/plugin-dialog";
import { openPath } from "@tauri-apps/plugin-opener";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";
import { batch, createEffect, createMemo, createSignal, onCleanup, onMount, startTransition, untrack } from "solid-js";
import CaptureTargetPanel from "./components/CaptureTargetPanel.tsx";
import SidePreviewPane from "./components/SidePreviewPane.tsx";
import TopControls from "./components/TopControls.tsx";
import TreeView from "./components/TreeView.tsx";
import { ALL_ROOTS_KEY, TREE_OVERSCAN_ROWS, TREE_ROW_STRIDE_PX } from "./lib/constants";
import { isAbortError, searchDebatifyTags } from "./lib/remoteSearch/client";
import { DEBATIFY_REMOTE_FOLDER_PATH } from "./lib/remoteSearch/treeRows";
import type { DebatifyTagHit } from "./lib/remoteSearch/types";
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
  const TREE_PANEL_MIN_PX = 360;
  const PANEL_HANDLE_WIDTH_PX = 12;
  const CAPTURE_TARGET_PREFS_KEY = "blockfile.captureTargetsByRoot.v1";
  const SEARCH_FILENAME_ONLY_PREFS_KEY = "blockfile.searchFileNamesOnly.v1";
  const SEARCH_DEBATIFY_ENABLED_PREFS_KEY = "blockfile.searchDebatifyEnabled.v1";
  const defaultExpandedFolders = () => new Set(["", DEBATIFY_REMOTE_FOLDER_PATH]);

  const loadStoredBoolean = (key: string, fallback: boolean) => {
    try {
      const raw = localStorage.getItem(key);
      if (raw === "true") return true;
      if (raw === "false") return false;
    } catch {
      // Ignore storage read failures (e.g. restricted storage mode)
    }
    return fallback;
  };

  const persistBooleanSetting = (key: string, value: boolean) => {
    try {
      localStorage.setItem(key, value ? "true" : "false");
    } catch {
      // Ignore storage write failures (e.g. restricted storage mode)
    }
  };

  const [roots, setRoots] = createSignal<RootSummary[]>([]);
  const [selectedRootPath, setSelectedRootPath] = createSignal("");
  const [snapshot, setSnapshot] = createSignal<IndexSnapshot | null>(null);
  const [allRootSnapshots, setAllRootSnapshots] = createSignal<IndexSnapshot[]>([]);
  const [previewCacheRenderTick, setPreviewCacheRenderTick] = createSignal(0);
  const [expandedFolders, setExpandedFolders] = createSignal<Set<string>>(defaultExpandedFolders());
  const [expandedFiles, setExpandedFiles] = createSignal<Set<number>>(new Set());
  const [collapsedHeadings, setCollapsedHeadings] = createSignal<Set<string>>(new Set());

  const [searchQuery, setSearchQuery] = createSignal("");
  const [searchResults, setSearchResults] = createSignal<SearchHit[]>([]);
  const [remoteTagResults, setRemoteTagResults] = createSignal<DebatifyTagHit[]>([]);
  const [searchFileNamesOnly, setSearchFileNamesOnly] = createSignal(
    loadStoredBoolean(SEARCH_FILENAME_ONLY_PREFS_KEY, false),
  );
  const [searchSemanticEnabled, setSearchSemanticEnabled] = createSignal(true);
  const [searchDebatifyEnabled, setSearchDebatifyEnabled] = createSignal(
    loadStoredBoolean(SEARCH_DEBATIFY_ENABLED_PREFS_KEY, true),
  );

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
  const [isSearchingLocal, setIsSearchingLocal] = createSignal(false);
  const [isSearchingRemote, setIsSearchingRemote] = createSignal(false);
  const [isLoadingSnapshot, setIsLoadingSnapshot] = createSignal(false);
  const [status, setStatus] = createSignal("Ready");
  const [indexProgress, setIndexProgress] = createSignal<IndexProgress | null>(null);
  const [copyToast, setCopyToast] = createSignal("");
  const [treeScrollTop, setTreeScrollTop] = createSignal(0);
  const [treeViewportHeight, setTreeViewportHeight] = createSignal(0);
  const [leftRailWidthPx, setLeftRailWidthPx] = createSignal(LEFT_RAIL_DEFAULT_PX);
  const [showCapturePanel, setShowCapturePanel] = createSignal(true);
  const [showPreviewPanel, setShowPreviewPanel] = createSignal(true);
  const [previewPanelWidthPx, setPreviewPanelWidthPx] = createSignal(420);
  const PREVIEW_PANEL_MIN_PX = 280;
  const PREVIEW_PANEL_MAX_PX = 640;
  const PREVIEW_WARMUP_LIMIT = 24;
  const PREVIEW_WARMUP_BATCH_SIZE = 2;
  const PREVIEW_WARMUP_DELAY_MS = 40;
  const PREVIEW_CACHE_FLUSH_MS = 16;

  let treeRef: HTMLDivElement | undefined;
  let searchRequestSeq = 0;
  let remoteSearchRequestSeq = 0;
  let previewWarmupSeq = 0;
  let headingPreviewSeq = 0;
  let focusScrollFrame = 0;
  let treeScrollFrame = 0;
  let pendingTreeScrollTopValue = 0;
  let captureTargetsSeq = 0;
  let capturePreviewSeq = 0;
  let stopLeftRailResize: (() => void) | null = null;
  let stopIndexProgressListener: UnlistenFn | null = null;
  let searchInputRef: HTMLInputElement | undefined;
  let headingPreviewTimer = 0;
  let previewCacheFlushTimer = 0;
  let pendingFocusDelta = 0;
  let focusMoveFrame = 0;
  const previewCacheByFileId = new Map<number, FilePreview>();
  const previewLoadsInFlight = new Map<number, Promise<FilePreview>>();
  const pendingPreviewCacheUpdates = new Map<number, FilePreview>();

  const selectedRoot = createMemo(() => roots().find((root) => root.path === selectedRootPath()) ?? null);
  const isAllRootsSelected = createMemo(() => selectedRootPath() === ALL_ROOTS_KEY);
  const activeSnapshot = createMemo(() =>
    isAllRootsSelected() ? mergeSnapshots(allRootSnapshots()) : snapshot(),
  );
  const captureRootPath = createMemo(() => (isAllRootsSelected() ? "" : selectedRootPath()));
  const searchMode = createMemo(() => searchQuery().trim().length >= 2);
  const isSearching = createMemo(() => isSearchingLocal() || isSearchingRemote());
  const activeTreePreviewFileIds = createMemo(() => {
    const ids = new Set<number>();

    if (searchMode()) {
      for (const result of searchResults()) {
        if (searchFileNamesOnly() || result.kind === "heading") {
          ids.add(result.fileId);
        }
      }
      return ids;
    }

    for (const fileId of expandedFiles()) {
      ids.add(fileId);
    }
    return ids;
  });
  const previewCacheForTree = createMemo<Record<number, FilePreview>>(() => {
    previewCacheRenderTick();
    const ids = activeTreePreviewFileIds();
    if (ids.size === 0) return {};

    const subset: Record<number, FilePreview> = {};
    for (const fileId of ids) {
      const preview = previewCacheByFileId.get(fileId);
      if (preview) {
        subset[fileId] = preview;
      }
    }
    return subset;
  });
  const normalizeCaptureTargetPath = (value: string) => normalizeSlashes(value).trim();

  const resetPreviewCacheState = () => {
    previewCacheByFileId.clear();
    previewLoadsInFlight.clear();
    pendingPreviewCacheUpdates.clear();

    if (previewCacheFlushTimer) {
      window.clearTimeout(previewCacheFlushTimer);
      previewCacheFlushTimer = 0;
    }

    setPreviewCacheRenderTick((tick) => tick + 1);
  };

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

  const flushPendingPreviewCache = () => {
    previewCacheFlushTimer = 0;
    if (pendingPreviewCacheUpdates.size === 0) return;

    const updates = Array.from(pendingPreviewCacheUpdates.entries());
    pendingPreviewCacheUpdates.clear();

    const activePreviewIds = untrack(activeTreePreviewFileIds);
    let shouldRender = false;
    for (const [fileId, preview] of updates) {
      if (previewCacheByFileId.get(fileId) !== preview) {
        previewCacheByFileId.set(fileId, preview);
        if (activePreviewIds.has(fileId)) {
          shouldRender = true;
        }
      }
    }

    if (!shouldRender) return;

    startTransition(() => {
      setPreviewCacheRenderTick((tick) => tick + 1);
    });
  };

  const queuePreviewCacheUpdate = (preview: FilePreview) => {
    pendingPreviewCacheUpdates.set(preview.fileId, preview);
    if (previewCacheFlushTimer) return;
    previewCacheFlushTimer = window.setTimeout(flushPendingPreviewCache, PREVIEW_CACHE_FLUSH_MS);
  };

  const ensurePreviewLoaded = async (fileId: number) => {
    const cached = previewCacheByFileId.get(fileId);
    if (cached) return cached;

    const inFlight = previewLoadsInFlight.get(fileId);
    if (inFlight) return inFlight;

    const loading = invokeTyped<FilePreview>("get_file_preview", { fileId })
      .then((preview) => {
        queuePreviewCacheUpdate(preview);
        return preview;
      })
      .finally(() => {
        previewLoadsInFlight.delete(fileId);
      });

    previewLoadsInFlight.set(fileId, loading);
    return loading;
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

  const scheduleTreeScrollTop = (nextScrollTop: number) => {
    pendingTreeScrollTopValue = nextScrollTop;
    if (treeScrollFrame) return;

    treeScrollFrame = requestAnimationFrame(() => {
      treeScrollFrame = 0;
      setTreeScrollTop((current) =>
        current === pendingTreeScrollTopValue ? current : pendingTreeScrollTopValue,
      );
    });
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

  const toggleFileNameSearchMode = () => {
    const next = !searchFileNamesOnly();
    setSearchFileNamesOnly(next);
    setStatus(next ? "Filename-only search enabled (press F again to exit)." : "Hybrid search restored.");
    focusSearchField();
  };

  const toggleSemanticSearchMode = () => {
    const next = !searchSemanticEnabled();
    setSearchSemanticEnabled(next);
    setStatus(next ? "AI semantic search enabled." : "AI semantic search disabled (lexical only).");
    focusSearchField();
  };

  const toggleDebatifySearchMode = () => {
    const next = !searchDebatifyEnabled();
    setSearchDebatifyEnabled(next);
    setStatus(next ? "Debatify API search enabled (press D to toggle)." : "Debatify API search disabled.");
    focusSearchField();
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

  let stopPreviewPanelResize: (() => void) | null = null;

  const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value));

  const computePreviewWidthMax = () => {
    if (!showPreviewPanel()) return PREVIEW_PANEL_MAX_PX;
    const handleCount = (showCapturePanel() ? 1 : 0) + 1;
    const reservedLeft = showCapturePanel() ? LEFT_RAIL_MIN_PX : 0;
    const maxByLayout = window.innerWidth - TREE_PANEL_MIN_PX - reservedLeft - handleCount * PANEL_HANDLE_WIDTH_PX;
    return Math.max(PREVIEW_PANEL_MIN_PX, Math.min(PREVIEW_PANEL_MAX_PX, maxByLayout));
  };

  const computeLeftRailWidthMax = (previewWidth: number) => {
    if (!showCapturePanel()) return LEFT_RAIL_MIN_PX;
    const handleCount = (showCapturePanel() ? 1 : 0) + (showPreviewPanel() ? 1 : 0);
    const reservedPreview = showPreviewPanel() ? Math.max(PREVIEW_PANEL_MIN_PX, previewWidth) : 0;
    const maxByLayout = window.innerWidth - TREE_PANEL_MIN_PX - reservedPreview - handleCount * PANEL_HANDLE_WIDTH_PX;
    return Math.max(LEFT_RAIL_MIN_PX, maxByLayout);
  };

  const constrainPanelWidths = () => {
    const nextPreviewWidth = showPreviewPanel()
      ? clamp(previewPanelWidthPx(), PREVIEW_PANEL_MIN_PX, computePreviewWidthMax())
      : previewPanelWidthPx();

    if (showPreviewPanel() && nextPreviewWidth !== previewPanelWidthPx()) {
      setPreviewPanelWidthPx(nextPreviewWidth);
    }

    if (!showCapturePanel()) return;

    const nextLeftWidth = clamp(leftRailWidthPx(), LEFT_RAIL_MIN_PX, computeLeftRailWidthMax(nextPreviewWidth));
    if (nextLeftWidth !== leftRailWidthPx()) {
      setLeftRailWidthPx(nextLeftWidth);
    }
  };

  const startLeftRailResize = (event: MouseEvent) => {
    if (window.innerWidth < 1024) return;
    event.preventDefault();

    const startX = event.clientX;
    const initialWidth = leftRailWidthPx();

    const onMouseMove = (moveEvent: MouseEvent) => {
      const maxWidth = computeLeftRailWidthMax(previewPanelWidthPx());
      const nextWidth = clamp(initialWidth + (moveEvent.clientX - startX), LEFT_RAIL_MIN_PX, maxWidth);
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

  const startPreviewPanelResize = (event: MouseEvent) => {
    event.preventDefault();

    const startX = event.clientX;
    const initialWidth = previewPanelWidthPx();

    const onMouseMove = (moveEvent: MouseEvent) => {
      const delta = startX - moveEvent.clientX;
      const maxPreviewWidth = computePreviewWidthMax();
      const nextWidth = clamp(initialWidth + delta, PREVIEW_PANEL_MIN_PX, maxPreviewWidth);

      if (showCapturePanel()) {
        const allowedLeftWidth = computeLeftRailWidthMax(nextWidth);
        setLeftRailWidthPx((current) => Math.min(current, allowedLeftWidth));
      }

      setPreviewPanelWidthPx(nextWidth);
    };

    const onMouseUp = () => {
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
      window.removeEventListener("blur", onMouseUp);
      document.body.classList.remove("is-resizing-panels");
      stopPreviewPanelResize = null;
    };

    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
    window.addEventListener("blur", onMouseUp);
    document.body.classList.add("is-resizing-panels");
    stopPreviewPanelResize = onMouseUp;
  };

  const toggleCapturePanel = () => {
    setShowCapturePanel((prev) => !prev);
    requestAnimationFrame(constrainPanelWidths);
  };

  const togglePreviewPanel = () => {
    setShowPreviewPanel((prev) => !prev);
    requestAnimationFrame(constrainPanelWidths);
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
        resetPreviewCacheState();
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
        setExpandedFolders(defaultExpandedFolders());
        setExpandedFiles(new Set<number>());
        setCollapsedHeadings(new Set<string>());
        resetPreviewCacheState();
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
        setExpandedFolders(defaultExpandedFolders());
        setExpandedFiles(new Set<number>());
        setCollapsedHeadings(new Set<string>());
        resetPreviewCacheState();
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

  const addCaptureHeading = async (headingLevel: 1 | 2 | 3 | 4, headingName: string): Promise<boolean> => {
    const rootPath = captureRootPath();
    const targetPath = selectedCaptureTarget();
    if (!rootPath || !targetPath) {
      setStatus("Select a destination docx file first.");
      return false;
    }

    const headingText = headingName.trim();
    if (!headingText) {
      setStatus("Heading name cannot be empty.");
      return false;
    }

    const requestId = ++capturePreviewSeq;
    setIsLoadingCapturePreview(true);
    try {
      const preview = await invokeTyped<CaptureTargetPreview>("add_capture_heading", {
        rootPath,
        targetPath,
        headingLevel,
        headingText,
        selectedTargetHeadingOrder: selectedCaptureHeadingOrder(),
      });
      if (requestId !== capturePreviewSeq) return false;

      const insertedHeading = [...preview.headings]
        .reverse()
        .find((heading) => heading.level === headingLevel && heading.text === headingText);

      batch(() => {
        setCaptureTargetPreview(preview);
        if (insertedHeading) {
          setSelectedCaptureHeadingOrder(insertedHeading.order);
        }
      });

      setStatus(`Added H${headingLevel} \"${headingText}\" in ${basename(preview.absolutePath)}`);
      return true;
    } catch (error) {
      if (requestId === capturePreviewSeq) {
        setStatus(`Could not add heading: ${String(error)}`);
      }
      return false;
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

  const sameSearchResult = (left: SearchHit | undefined, right: SearchHit | undefined) => {
    if (left === right) return true;
    if (!left || !right) return false;

    return (
      left.source === right.source &&
      left.kind === right.kind &&
      left.fileId === right.fileId &&
      left.fileName === right.fileName &&
      left.relativePath === right.relativePath &&
      left.absolutePath === right.absolutePath &&
      left.headingLevel === right.headingLevel &&
      left.headingText === right.headingText &&
      left.headingOrder === right.headingOrder &&
      left.score === right.score
    );
  };

  const sameParagraphXml = (left: string[] | undefined, right: string[] | undefined) => {
    if (left === right) return true;
    if (!left || !right || left.length !== right.length) return false;
    for (let index = 0; index < left.length; index += 1) {
      if (left[index] !== right[index]) return false;
    }
    return true;
  };

  const sameTreeRow = (left: TreeRow, right: TreeRow) => {
    if (left === right) return true;
    return (
      left.key === right.key &&
      left.kind === right.kind &&
      left.depth === right.depth &&
      left.label === right.label &&
      left.subLabel === right.subLabel &&
      left.headingLevel === right.headingLevel &&
      left.headingOrder === right.headingOrder &&
      left.folderPath === right.folderPath &&
      left.fileId === right.fileId &&
      left.copyText === right.copyText &&
      left.sourcePath === right.sourcePath &&
      left.richHtml === right.richHtml &&
      left.hasChildren === right.hasChildren &&
      sameParagraphXml(left.paragraphXml, right.paragraphXml) &&
      sameSearchResult(left.searchResult, right.searchResult)
    );
  };

  const reconcileTreeRows = (previousRows: TreeRow[], nextRows: TreeRow[]) => {
    if (previousRows.length === 0 || nextRows.length === 0) return nextRows;

    const previousByKey = new Map<string, TreeRow>();
    for (const previousRow of previousRows) {
      previousByKey.set(previousRow.key, previousRow);
    }

    let reusedCount = 0;
    const reconciled = nextRows.map((nextRow) => {
      const previousRow = previousByKey.get(nextRow.key);
      if (previousRow && sameTreeRow(previousRow, nextRow)) {
        reusedCount += 1;
        return previousRow;
      }
      return nextRow;
    });

    if (reusedCount === nextRows.length && previousRows.length === nextRows.length) {
      return previousRows;
    }

    return reconciled;
  };

  const treeRows = createMemo<TreeRow[]>((previousRows = []) => {
    const nextRows = buildTreeRows({
      snapshotIndex: snapshotIndex(),
      searchMode: searchMode(),
      searchFileNamesOnly: searchFileNamesOnly(),
      searchResults: searchResults(),
      remoteTagResults: remoteTagResults(),
      previewCache: previewCacheForTree(),
      expandedFolders: expandedFolders(),
      expandedFiles: expandedFiles(),
      collapsedHeadings: collapsedHeadings(),
    });

    return reconcileTreeRows(previousRows, nextRows);
  }, []);

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
      const rowRichHtml = row.richHtml;
      setSidePreview({
        title: row.kind === "heading" ? row.label : row.kind === "f8" ? "F8 Cite" : "Author / Source",
        subTitle: row.subLabel,
        text: row.copyText,
        headingLevel: row.kind === "heading" ? row.headingLevel ?? null : null,
        kind: row.kind,
        richHtml: headingCacheKey ? headingPreviewHtmlCache()[headingCacheKey] ?? rowRichHtml : rowRichHtml,
      });

      if (row.kind === "heading" && headingCacheKey) {
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
        paragraphXml: row.paragraphXml ?? null,
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
    persistBooleanSetting(SEARCH_FILENAME_ONLY_PREFS_KEY, searchFileNamesOnly());
  });

  createEffect(() => {
    persistBooleanSetting(SEARCH_DEBATIFY_ENABLED_PREFS_KEY, searchDebatifyEnabled());
  });

  createEffect(() => {
    treeRows().length;
    const frame = requestAnimationFrame(() => syncTreeViewportState());
    onCleanup(() => cancelAnimationFrame(frame));
  });

  createEffect(() => {
    const rootPath = selectedRootPath();
    const searchRootPath = rootPath === ALL_ROOTS_KEY ? undefined : rootPath;
    const query = searchQuery().trim();
    const fileNameOnly = searchFileNamesOnly();
    const semanticEnabled = searchSemanticEnabled();
    if ((!rootPath && roots().length === 0) || query.length < 2) {
      setSearchResults([]);
      setIsSearchingLocal(false);
      return;
    }

    const debounceMs = query.length > 120 ? 240 : query.length > 42 ? 170 : 120;
    setIsSearchingLocal(true);
    const requestId = ++searchRequestSeq;
    const timer = setTimeout(() => {
      const invocation = invokeTyped<SearchHit[]>("search_index_hybrid", {
        query,
        rootPath: searchRootPath,
        limit: 120,
        fileNameOnly,
        semanticEnabled,
      })
        .catch((error) => {
          if (requestId === searchRequestSeq) {
            setStatus(`Search failed: ${String(error)}`);
          }
          return [] as SearchHit[];
        })
        .then((results) => {
          if (requestId === searchRequestSeq) {
            startTransition(() => {
              setSearchResults(results);
            });
          }
        })
        .finally(() => {
          if (requestId === searchRequestSeq) {
            setIsSearchingLocal(false);
          }
        });

      const timeout = setTimeout(() => {
        if (requestId === searchRequestSeq) {
          setIsSearchingLocal(false);
        }
      }, 1500);
      void invocation.finally(() => {
        clearTimeout(timeout);
      });
    }, debounceMs);

    onCleanup(() => clearTimeout(timer));
  });

  createEffect(() => {
    const query = searchQuery().trim();
    const debatifyEnabled = searchDebatifyEnabled();
    if (!debatifyEnabled || query.length < 2) {
      setRemoteTagResults([]);
      setIsSearchingRemote(false);
      return;
    }

    const debounceMs = query.length > 120 ? 280 : query.length > 42 ? 200 : 140;
    setIsSearchingRemote(true);
    const requestId = ++remoteSearchRequestSeq;
    const controller = new AbortController();

    const timer = setTimeout(() => {
      const invocation = searchDebatifyTags(query, { signal: controller.signal })
        .catch((error) => {
          if (requestId === remoteSearchRequestSeq && !isAbortError(error)) {
            setStatus(`Debatify tag search failed: ${String(error)}`);
          }
          return [] as DebatifyTagHit[];
        })
        .then((results) => {
          if (requestId === remoteSearchRequestSeq) {
            startTransition(() => {
              setRemoteTagResults(results);
            });
          }
        })
        .finally(() => {
          if (requestId === remoteSearchRequestSeq) {
            setIsSearchingRemote(false);
          }
        });

      const timeout = setTimeout(() => {
        if (requestId === remoteSearchRequestSeq) {
          setIsSearchingRemote(false);
        }
      }, 2500);
      void invocation.finally(() => {
        clearTimeout(timeout);
      });
    }, debounceMs);

    onCleanup(() => {
      clearTimeout(timer);
      controller.abort();
    });
  });

  createEffect(() => {
    if (!searchMode()) return;
    const ids = Array.from(new Set(searchResults().map((result) => result.fileId))).slice(0, PREVIEW_WARMUP_LIMIT);
    if (ids.length === 0) return;

    const requestId = ++previewWarmupSeq;
    const missing = ids.filter((id) => !previewCacheByFileId.has(id));
    if (missing.length === 0) return;

    const loadBatch = async (batchIds: number[]) => {
      await Promise.allSettled(batchIds.map((id) => ensurePreviewLoaded(id)));
    };
    const pause = () => new Promise<void>((resolve) => setTimeout(resolve, PREVIEW_WARMUP_DELAY_MS));

    void (async () => {
      for (let index = 0; index < missing.length; index += PREVIEW_WARMUP_BATCH_SIZE) {
        if (requestId !== previewWarmupSeq) return;
        await loadBatch(missing.slice(index, index + PREVIEW_WARMUP_BATCH_SIZE));
        if (requestId !== previewWarmupSeq) return;
        await pause();
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

  createEffect(() => {
    showCapturePanel();
    showPreviewPanel();
    requestAnimationFrame(constrainPanelWidths);
  });

  onMount(() => {
    const onResize = () => {
      constrainPanelWidths();
      syncTreeViewportState();
    };
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

      if (!event.metaKey && !event.ctrlKey && !event.altKey && key.toLowerCase() === "f") {
        if (targetIsTextEditable(event.target)) return;
        event.preventDefault();
        toggleFileNameSearchMode();
        return;
      }

      if (!event.metaKey && !event.ctrlKey && !event.altKey && key.toLowerCase() === "d") {
        if (targetIsTextEditable(event.target)) return;
        event.preventDefault();
        toggleDebatifySearchMode();
        return;
      }

      if (!event.metaKey && !event.ctrlKey && !event.altKey && key.toLowerCase() === "i") {
        if (targetIsTextEditable(event.target)) return;
        event.preventDefault();
        toggleCapturePanel();
        return;
      }

      if (!event.metaKey && !event.ctrlKey && !event.altKey && key.toLowerCase() === "p") {
        if (targetIsTextEditable(event.target)) return;
        event.preventDefault();
        togglePreviewPanel();
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
      stopPreviewPanelResize?.();
      window.removeEventListener("resize", onResize);
      window.removeEventListener("keydown", onGlobalKeyDown, true);
      if (headingPreviewTimer) {
        window.clearTimeout(headingPreviewTimer);
        headingPreviewTimer = 0;
      }
      if (previewCacheFlushTimer) {
        window.clearTimeout(previewCacheFlushTimer);
        previewCacheFlushTimer = 0;
      }
      if (focusMoveFrame) {
        cancelAnimationFrame(focusMoveFrame);
        focusMoveFrame = 0;
      }
      pendingFocusDelta = 0;
      if (focusScrollFrame) {
        cancelAnimationFrame(focusScrollFrame);
      }
      if (treeScrollFrame) {
        cancelAnimationFrame(treeScrollFrame);
        treeScrollFrame = 0;
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
    <div class="h-screen bg-[#0a0a0a] overflow-hidden">
      <div class="flex h-full w-full flex-col">
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
          searchFileNamesOnly={searchFileNamesOnly}
          searchDebatifyEnabled={searchDebatifyEnabled}
          searchSemanticEnabled={searchSemanticEnabled}
          selectedRootPath={selectedRootPath}
          indexProgress={indexProgress}
          setSearchInputRef={setSearchInputElement}
          setSearchQuery={setSearchQuery}
          setSelectedRootPath={setSelectedRootPath}
          toggleFileNameSearchMode={toggleFileNameSearchMode}
          toggleDebatifySearchMode={toggleDebatifySearchMode}
          toggleSemanticSearchMode={toggleSemanticSearchMode}
          status={status}
          showCapturePanel={showCapturePanel}
          showPreviewPanel={showPreviewPanel}
          toggleCapturePanel={toggleCapturePanel}
          togglePreviewPanel={togglePreviewPanel}
        />

        <div class="flex min-h-0 flex-1">
          <div class="workspace-split h-full min-h-0 min-w-0 flex-1" style={{ "--left-rail-width": showCapturePanel() ? `${leftRailWidthPx()}px` : '0px' }}>
            {showCapturePanel() && (
              <aside class="h-full min-h-0 border-r border-neutral-800/50 bg-neutral-950/30 flex flex-col">
                <CaptureTargetPanel
                  addCaptureHeading={addCaptureHeading}
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
            )}

            {showCapturePanel() && (
              <button
                aria-label="Resize insert preview panel"
                class="panel-resize-handle hidden lg:flex"
                onMouseDown={startLeftRailResize}
                title="Drag to resize"
                type="button"
              />
            )}

            <div class="h-full min-h-0 min-w-0 flex-1 bg-neutral-950/20">
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
                onTreeScroll={scheduleTreeScrollTop}
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

          {showPreviewPanel() && (
            <>
              <button
                aria-label="Resize preview panel"
                class="panel-resize-handle flex"
                onMouseDown={startPreviewPanelResize}
                title="Drag to resize preview"
                type="button"
              />
              <SidePreviewPane
                sidePreview={sidePreview}
                width={previewPanelWidthPx}
              />
            </>
          )}
        </div>
      </div>
    </div>
  );
}

export default App;
