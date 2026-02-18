import { For, Show, type Accessor } from "solid-js";
import { ALL_ROOTS_KEY } from "../lib/constants";
import type { IndexProgress, RootSummary } from "../lib/types";
import { formatTime } from "../lib/utils";

type TopControlsProps = {
  searchQuery: Accessor<string>;
  setSearchQuery: (value: string) => void;
  setSearchInputRef: (element: HTMLInputElement) => void;
  isIndexing: Accessor<boolean>;
  addFolder: () => void;
  selectedRootPath: Accessor<string>;
  runIndexForSelection: () => Promise<void>;
  activeRootLabel: Accessor<string>;
  activeLastIndexedMs: Accessor<number>;
  isSearching: Accessor<boolean>;
  status: Accessor<string>;
  copyToast: Accessor<string>;
  roots: Accessor<RootSummary[]>;
  setSelectedRootPath: (value: string) => void;
  indexProgress: Accessor<IndexProgress | null>;
};

export default function TopControls(props: TopControlsProps) {
  const activeIndexProgress = () => {
    if (!props.isIndexing()) return null;
    return props.indexProgress();
  };

  const indexProgressPercent = () => {
    const progress = activeIndexProgress();
    if (!progress || progress.changed <= 0) return 0;
    return Math.min(100, Math.round((progress.processed / progress.changed) * 100));
  };

  const indexProgressTitle = () => {
    const progress = activeIndexProgress();
    if (!progress) return "";

    if (progress.phase === "discovering") {
      return `Scanning ${progress.discovered.toLocaleString()} .docx files`;
    }

    if (progress.phase === "indexing") {
      return `Indexing ${progress.processed.toLocaleString()} / ${progress.changed.toLocaleString()} files`;
    }

    if (progress.phase === "cleaning") {
      return "Removing stale index entries";
    }

    return "Finalizing index";
  };

  const indexProgressDetail = () => {
    const progress = activeIndexProgress();
    if (!progress) return "";
    const elapsedSeconds = Math.max(0, Math.round(progress.elapsedMs / 1000));

    const counts = [
      `${progress.updated.toLocaleString()} updated`,
      `${progress.skipped.toLocaleString()} skipped`,
      `${progress.removed.toLocaleString()} removed`,
      `${elapsedSeconds}s`,
    ];
    const currentFile = progress.currentFile ? ` - ${progress.currentFile}` : "";
    return `${counts.join(" - ")}${currentFile}`;
  };

  return (
    <div class="space-y-5 p-5">
      <div class="flex flex-col gap-4 md:flex-row">
        <div class="relative flex-1">
          <div class="pointer-events-none absolute inset-y-0 left-0 flex items-center pl-3">
            <svg class="h-4 w-4 text-neutral-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
            </svg>
          </div>
          <input
            id="app-search-input"
            class="vercel-input w-full pl-10"
            onInput={(event) => props.setSearchQuery(event.currentTarget.value)}
            placeholder="Search files, pockets, hats, blocks, and tags..."
            ref={props.setSearchInputRef}
            value={props.searchQuery()}
          />
          <Show when={props.isSearching()}>
            <div class="pointer-events-none absolute inset-y-0 right-0 flex items-center pr-3">
              <div class="h-4 w-4 animate-spin rounded-full border-2 border-neutral-600 border-t-blue-500" />
            </div>
          </Show>
        </div>
        
        <div class="flex gap-3">
          <button
            class="vercel-btn"
            disabled={props.isIndexing()}
            onClick={props.addFolder}
            type="button"
          >
            <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4" />
            </svg>
            Add Folder
          </button>
          <button
            class="vercel-btn vercel-btn-primary"
            disabled={!props.selectedRootPath() || props.isIndexing()}
            onClick={() => void props.runIndexForSelection()}
            type="button"
          >
            <Show when={props.isIndexing()} fallback={
              <>
                <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                </svg>
                Reindex
              </>
            }>
              <div class="h-4 w-4 animate-spin rounded-full border-2 border-white/30 border-t-white" />
              Indexing...
            </Show>
          </button>
        </div>
      </div>

      <div class="flex flex-wrap items-center gap-4 text-sm">
        <div class="flex items-center gap-2 text-neutral-400">
          <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z" />
          </svg>
          <span class="truncate max-w-[300px]" title={props.activeRootLabel()}>
            {props.activeRootLabel()}
          </span>
        </div>
        
        <div class="h-4 w-px bg-neutral-700" />
        
        <div class="flex items-center gap-2 text-neutral-500">
          <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
          {formatTime(props.activeLastIndexedMs())}
        </div>
        
        <div class="flex-1" />
        
        <div class="flex items-center gap-2">
          <span class="text-neutral-500 leading-relaxed">{props.status()}</span>
          <Show when={props.copyToast()}>
            <span class="vercel-badge border-emerald-500/30 bg-emerald-500/10 text-emerald-400">
              {props.copyToast()}
            </span>
          </Show>
        </div>
      </div>

      <Show when={activeIndexProgress()}>
        <div class="rounded-xl border border-blue-500/20 bg-blue-500/5 px-4 py-3">
          <div class="flex items-center justify-between gap-4 text-xs font-medium text-blue-100">
            <span>{indexProgressTitle()}</span>
            <Show when={activeIndexProgress()?.phase !== "discovering" && activeIndexProgress()?.changed !== 0}>
              <span>{indexProgressPercent()}%</span>
            </Show>
          </div>

          <div class="mt-2 h-2 overflow-hidden rounded-full bg-neutral-800">
            <Show
              when={activeIndexProgress()?.phase !== "discovering" && activeIndexProgress()?.changed !== 0}
              fallback={<div class="h-full w-1/3 animate-pulse rounded-full bg-gradient-to-r from-blue-500/50 via-emerald-400/60 to-blue-500/50" />}
            >
              <div
                class="h-full rounded-full bg-gradient-to-r from-blue-500 to-emerald-400 transition-[width] duration-200"
                style={{ width: `${indexProgressPercent()}%` }}
              />
            </Show>
          </div>

          <p class="mt-2 truncate text-xs text-neutral-500" title={indexProgressDetail()}>
            {indexProgressDetail()}
          </p>
        </div>
      </Show>

      <Show when={props.roots().length > 1}>
        <div class="flex items-center gap-4">
          <label class="text-sm text-neutral-400">Root:</label>
          <select
            class="vercel-select min-w-[240px]"
            onChange={(event) => props.setSelectedRootPath(event.currentTarget.value)}
            value={props.selectedRootPath()}
          >
            <option value={ALL_ROOTS_KEY}>All indexed folders</option>
            <For each={props.roots()}>{(root) => <option value={root.path}>{root.path}</option>}</For>
          </select>
        </div>
      </Show>
    </div>
  );
}
