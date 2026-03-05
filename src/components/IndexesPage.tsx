import { For, Show, type Accessor } from "solid-js";
import type { IndexProgress, RootSummary } from "../lib/types";
import { basename, formatTime } from "../lib/utils";

type IndexesPageProps = {
  roots: Accessor<RootSummary[]>;
  selectedRootPath: Accessor<string>;
  status: Accessor<string>;
  isIndexing: Accessor<boolean>;
  indexProgress: Accessor<IndexProgress | null>;
  addFolders: () => void;
  reindexAll: () => Promise<void>;
  reindexRoot: (rootPath: string) => Promise<void>;
  openWorkspacePage: () => void;
  openWorkspaceRoot: (rootPath: string) => void;
};

export default function IndexesPage(props: IndexesPageProps) {
  const activeIndexProgress = () => {
    if (!props.isIndexing()) return null;
    return props.indexProgress();
  };

  const indexProgressPercent = () => {
    const progress = activeIndexProgress();
    if (!progress || progress.changed <= 0) return 0;
    return Math.min(100, Math.round((progress.processed / progress.changed) * 100));
  };

  return (
    <div class="flex h-full min-h-0 flex-col">
      <header class="border-b border-neutral-800/50 bg-neutral-950/50 px-4 py-3 backdrop-blur-sm">
        <div class="flex flex-wrap items-center justify-between gap-3">
          <div class="flex items-center gap-2">
            <button
              class="inline-flex h-8 items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-900 px-3 text-xs font-medium text-neutral-200 transition hover:border-neutral-600 hover:text-white"
              onClick={props.openWorkspacePage}
              type="button"
            >
              <svg class="h-3.5 w-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 19l-7-7 7-7" />
              </svg>
              Workspace
            </button>
            <div class="h-4 w-px bg-neutral-700" />
            <h1 class="text-sm font-semibold text-white">Indexes</h1>
            <span class="vercel-badge text-[11px]">
              {props.roots().length.toLocaleString()} folders
            </span>
          </div>

          <div class="flex flex-wrap items-center gap-2">
            <button
              class="inline-flex h-8 items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-800 px-3 text-xs font-medium text-neutral-200 transition hover:border-neutral-600 hover:bg-neutral-700 hover:text-white disabled:opacity-50"
              disabled={props.isIndexing()}
              onClick={props.addFolders}
              type="button"
            >
              <svg class="h-3.5 w-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4" />
              </svg>
              Add folders
            </button>
            <button
              class="inline-flex h-8 items-center gap-1.5 rounded-md border border-blue-600 bg-blue-600 px-3 text-xs font-medium text-white transition hover:border-blue-500 hover:bg-blue-500 disabled:opacity-50"
              disabled={props.isIndexing() || props.roots().length === 0}
              onClick={() => void props.reindexAll()}
              type="button"
            >
              <Show
                when={props.isIndexing()}
                fallback={
                  <>
                    <svg class="h-3.5 w-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path
                        stroke-linecap="round"
                        stroke-linejoin="round"
                        stroke-width="2"
                        d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"
                      />
                    </svg>
                    Reindex all
                  </>
                }
              >
                <div class="h-3.5 w-3.5 animate-spin rounded-full border-2 border-white/30 border-t-white" />
              </Show>
            </button>
          </div>
        </div>

        <p class="mt-2 text-xs text-neutral-500">
          Select one or more folders in the picker to add and index them in a single pass.
        </p>
      </header>

      <Show when={activeIndexProgress()}>
        <div class="border-b border-blue-500/20 bg-blue-500/5 px-4 py-2">
          <div class="flex items-center justify-between gap-4 text-xs">
            <span class="text-blue-100">
              {activeIndexProgress()?.phase === "discovering"
                ? `Scanning ${activeIndexProgress()?.discovered.toLocaleString()} .docx files`
                : `Indexing ${activeIndexProgress()?.processed.toLocaleString()} / ${activeIndexProgress()?.changed.toLocaleString()} files`}
            </span>
            <Show when={activeIndexProgress()?.phase !== "discovering" && activeIndexProgress()?.changed !== 0}>
              <span class="text-blue-200">{indexProgressPercent()}%</span>
            </Show>
          </div>
          <div class="mt-1.5 h-1.5 overflow-hidden rounded-full bg-neutral-800">
            <Show
              when={activeIndexProgress()?.phase !== "discovering" && activeIndexProgress()?.changed !== 0}
              fallback={
                <div class="h-full w-1/3 animate-pulse rounded-full bg-gradient-to-r from-blue-500/50 via-emerald-400/60 to-blue-500/50" />
              }
            >
              <div
                class="h-full rounded-full bg-gradient-to-r from-blue-500 to-emerald-400 transition-[width] duration-200"
                style={{ width: `${indexProgressPercent()}%` }}
              />
            </Show>
          </div>
        </div>
      </Show>

      <main class="min-h-0 flex-1 overflow-auto p-4">
        <Show
          when={props.roots().length > 0}
          fallback={
            <div class="vercel-card mx-auto mt-10 max-w-xl p-6 text-center">
              <h2 class="text-base font-semibold text-white">No folders indexed yet</h2>
              <p class="mt-2 text-sm text-neutral-400">
                Add one or multiple folders, then BlockFile can index all `.docx` files for search and preview.
              </p>
              <button
                class="vercel-btn vercel-btn-primary mt-4 h-9 px-4 text-xs"
                disabled={props.isIndexing()}
                onClick={props.addFolders}
                type="button"
              >
                Add folders
              </button>
            </div>
          }
        >
          <div class="grid gap-3">
            <For each={props.roots()}>
              {(root) => {
                const isSelected = () => props.selectedRootPath() === root.path;
                return (
                  <article
                    class={`vercel-card p-4 transition ${
                      isSelected() ? "border-blue-500/50 bg-blue-500/5" : "hover:border-neutral-700"
                    }`}
                  >
                    <div class="flex flex-wrap items-start justify-between gap-3">
                      <div class="min-w-0">
                        <h2 class="truncate text-sm font-semibold text-white" title={root.path}>
                          {basename(root.path)}
                        </h2>
                        <p class="mt-1 truncate text-xs text-neutral-400" title={root.path}>
                          {root.path}
                        </p>
                      </div>

                      <div class="flex items-center gap-2">
                        <button
                          class="inline-flex h-8 items-center gap-1 rounded-md border border-neutral-700 bg-neutral-800 px-2.5 text-xs font-medium text-neutral-200 transition hover:border-neutral-600 hover:text-white"
                          onClick={() => props.openWorkspaceRoot(root.path)}
                          type="button"
                        >
                          Open
                        </button>
                        <button
                          class="inline-flex h-8 items-center gap-1 rounded-md border border-blue-600 bg-blue-600 px-2.5 text-xs font-medium text-white transition hover:border-blue-500 hover:bg-blue-500 disabled:opacity-50"
                          disabled={props.isIndexing()}
                          onClick={() => void props.reindexRoot(root.path)}
                          type="button"
                        >
                          Reindex
                        </button>
                      </div>
                    </div>

                    <div class="mt-3 grid gap-2 text-xs text-neutral-300 sm:grid-cols-2 lg:grid-cols-4">
                      <div class="rounded-md border border-neutral-800 bg-neutral-900/40 px-2.5 py-2">
                        <div class="text-[10px] uppercase tracking-wide text-neutral-500">Files</div>
                        <div class="mt-1 text-sm font-medium text-neutral-100">{root.fileCount.toLocaleString()}</div>
                      </div>
                      <div class="rounded-md border border-neutral-800 bg-neutral-900/40 px-2.5 py-2">
                        <div class="text-[10px] uppercase tracking-wide text-neutral-500">Headings</div>
                        <div class="mt-1 text-sm font-medium text-neutral-100">{root.headingCount.toLocaleString()}</div>
                      </div>
                      <div class="rounded-md border border-neutral-800 bg-neutral-900/40 px-2.5 py-2">
                        <div class="text-[10px] uppercase tracking-wide text-neutral-500">Added</div>
                        <div class="mt-1 text-sm font-medium text-neutral-100">{formatTime(root.addedAtMs)}</div>
                      </div>
                      <div class="rounded-md border border-neutral-800 bg-neutral-900/40 px-2.5 py-2">
                        <div class="text-[10px] uppercase tracking-wide text-neutral-500">Last indexed</div>
                        <div class="mt-1 text-sm font-medium text-neutral-100">{formatTime(root.lastIndexedMs)}</div>
                      </div>
                    </div>
                  </article>
                );
              }}
            </For>
          </div>
        </Show>
      </main>

      <footer class="border-t border-neutral-800/50 px-4 py-2 text-xs text-neutral-500">{props.status()}</footer>
    </div>
  );
}
