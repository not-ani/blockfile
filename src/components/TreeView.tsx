import { For, Show, type Accessor } from "solid-js";
import type { SearchHit, TreeRow } from "../lib/types";

type VirtualWindow = {
  topSpacerPx: number;
  bottomSpacerPx: number;
};

type TreeViewProps = {
  visibleTreeRows: Accessor<TreeRow[]>;
  virtualWindow: Accessor<VirtualWindow>;
  focusedNodeKey: Accessor<string>;
  searchMode: Accessor<boolean>;
  expandedFolders: Accessor<Set<string>>;
  expandedFiles: Accessor<Set<number>>;
  collapsedHeadings: Accessor<Set<string>>;
  activateRow: (row: TreeRow, fromKeyboard: boolean) => Promise<void>;
  applyPreviewFromRow: (row: TreeRow) => void;
  openSearchResult: (result: SearchHit) => Promise<void>;
  setFocusedNodeKey: (key: string) => void;
  onTreeKeyDown: (event: KeyboardEvent) => void;
  onTreeScroll: (scrollTop: number) => void;
  setTreeRef: (element: HTMLDivElement) => void;
  isLoadingSnapshot: Accessor<boolean>;
  treeRowsLength: Accessor<number>;
  selectedRootPath: Accessor<string>;
  isSearching: Accessor<boolean>;
};

const FileIcon = () => (
  <svg class="h-4 w-4 text-neutral-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
  </svg>
);

const ChevronRightIcon = () => (
  <svg class="h-3 w-3 text-neutral-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7" />
  </svg>
);

const ChevronDownIcon = () => (
  <svg class="h-3 w-3 text-neutral-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
  </svg>
);

const HeadingIcon = () => (
  <svg class="h-3.5 w-3.5 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 12h16M4 18h7" />
  </svg>
);

const F8Icon = () => (
  <span class="flex h-4 w-4 items-center justify-center rounded bg-amber-500/20 text-[9px] font-bold text-amber-400">F8</span>
);

const AuthorIcon = () => (
  <svg class="h-3.5 w-3.5 text-purple-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
  </svg>
);

export default function TreeView(props: TreeViewProps) {
  return (
    <div
      class="h-full min-h-0 overflow-auto px-3 pb-3 pt-1 outline-none focus-visible:ring-2 focus-visible:ring-blue-500/30"
      onKeyDown={props.onTreeKeyDown}
      onScroll={(event) => props.onTreeScroll(event.currentTarget.scrollTop)}
      ref={props.setTreeRef}
      tabindex={0}
    >
      <div style={{ height: `${props.virtualWindow().topSpacerPx}px` }} />

      <For each={props.visibleTreeRows()}>
        {(row) => {
          const focused = () => props.focusedNodeKey() === row.key;
          const isFolder = () => row.kind === "folder";
          const isFile = () => row.kind === "file";
          const isHeading = () => row.kind === "heading";
          const isF8 = () => row.kind === "f8";
          const isAuthor = () => row.kind === "author";

          return (
            <button
              class={`tree-row group ${
                focused() ? "tree-row-focused" : ""
              } ${row.kind === "loading" ? "cursor-default opacity-50" : ""}`}
              data-row-key={row.key}
              onClick={() => {
                if (row.kind === "loading") return;
                if (props.searchMode() && row.searchResult) {
                  void props.openSearchResult(row.searchResult);
                  return;
                }
                void props.activateRow(row, false);
              }}
              onFocus={() => {
                if (row.kind !== "loading") {
                  props.setFocusedNodeKey(row.key);
                  props.applyPreviewFromRow(row);
                }
              }}
              onMouseEnter={() => {
                if (row.kind !== "loading") {
                  props.setFocusedNodeKey(row.key);
                }
                props.applyPreviewFromRow(row);
              }}
              style={{ "padding-left": `${16 + row.depth * 22}px` }}
              type="button"
            >
              <span class="flex h-5 w-5 items-center justify-center">
                {isFolder() ? (
                  props.expandedFolders().has(row.folderPath ?? "") ? (
                    <ChevronDownIcon />
                  ) : (
                    <ChevronRightIcon />
                  )
                ) : isFile() ? (
                  props.searchMode() ? (
                    <FileIcon />
                  ) : props.expandedFiles().has(row.fileId ?? -1) ? (
                    <ChevronDownIcon />
                  ) : (
                    <ChevronRightIcon />
                  )
                ) : isHeading() ? (
                  row.hasChildren ? (
                    props.collapsedHeadings().has(row.key) ? (
                      <ChevronRightIcon />
                    ) : (
                      <ChevronDownIcon />
                    )
                  ) : (
                    <HeadingIcon />
                  )
                ) : isF8() ? (
                  <F8Icon />
                ) : isAuthor() ? (
                  <AuthorIcon />
                ) : (
                  <span class="h-1.5 w-1.5 rounded-full bg-neutral-600" />
                )}
              </span>
              
              <span class="min-w-0 flex-1">
                <p
                  class={`truncate text-sm ${
                    isHeading() || isF8() || isAuthor()
                      ? "font-medium text-neutral-200"
                      : "text-neutral-400"
                  }`}
                >
                  {row.label}
                </p>
                <Show when={row.subLabel}>
                  <p class="truncate text-xs leading-5 text-neutral-600">{row.subLabel}</p>
                </Show>
              </span>

              <Show when={focused() && (isHeading() || isF8() || isAuthor())}>
                <span class="vercel-badge opacity-0 transition-opacity group-hover:opacity-100">
                  Space to insert
                </span>
              </Show>
              <Show when={props.searchMode() && row.searchResult && row.searchResult.source !== "lexical"}>
                <span class="vercel-badge border-indigo-500/30 bg-indigo-500/10 text-indigo-300">
                  {row.searchResult?.source === "hybrid" ? "AI + Lex" : "AI Match"}
                </span>
              </Show>
            </button>
          );
        }}
      </For>

      <div style={{ height: `${props.virtualWindow().bottomSpacerPx}px` }} />

      <Show when={!props.isLoadingSnapshot() && props.treeRowsLength() === 0 && !props.selectedRootPath()}>
        <div class="flex flex-col items-center justify-center py-12 text-center">
          <div class="mb-4 flex h-12 w-12 items-center justify-center rounded-xl bg-neutral-800">
            <svg class="h-6 w-6 text-neutral-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 19a2 2 0 01-2-2V7a2 2 0 012-2h4l2 2h4a2 2 0 012 2v1M5 19h14a2 2 0 002-2v-5a2 2 0 00-2-2H9a2 2 0 00-2 2v5a2 2 0 01-2 2z" />
            </svg>
          </div>
          <p class="text-sm font-medium text-neutral-400">No folders indexed</p>
          <p class="mt-1 text-xs text-neutral-600">Add a folder to start building your document index.</p>
        </div>
      </Show>
      
      <Show when={!props.isLoadingSnapshot() && props.treeRowsLength() === 0 && props.searchMode() && !props.isSearching()}>
        <div class="flex flex-col items-center justify-center py-12 text-center">
          <div class="mb-4 flex h-12 w-12 items-center justify-center rounded-xl bg-neutral-800">
            <svg class="h-6 w-6 text-neutral-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
            </svg>
          </div>
          <p class="text-sm font-medium text-neutral-400">No results found</p>
          <p class="mt-1 text-xs text-neutral-600">Try adjusting your search query.</p>
        </div>
      </Show>
    </div>
  );
}
