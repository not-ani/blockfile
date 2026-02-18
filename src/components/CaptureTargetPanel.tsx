import { DragDropProvider, useDraggable, useDroppable } from "@dnd-kit/solid";
import { For, Show, type Accessor, type Setter } from "solid-js";
import type { CaptureTarget, CaptureTargetPreview, FileHeading } from "../lib/types";

type CaptureTargetPanelProps = {
  captureRootPath: Accessor<string>;
  isLoadingCaptureTargets: Accessor<boolean>;
  selectedCaptureTarget: Accessor<string>;
  setSelectedCaptureTarget: (value: string) => void;
  captureTargets: Accessor<CaptureTarget[]>;
  createCaptureTarget: () => Promise<void>;
  selectCaptureTargetFromFilesystem: () => Promise<void>;
  isAllRootsSelected: Accessor<boolean>;
  selectedCaptureTargetMeta: Accessor<CaptureTarget | null>;
  isLoadingCapturePreview: Accessor<boolean>;
  captureTargetPreview: Accessor<CaptureTargetPreview | null>;
  captureTargetH1ToH4: Accessor<FileHeading[]>;
  selectedCaptureHeadingOrder: Accessor<number | null>;
  setSelectedCaptureHeadingOrder: Setter<number | null>;
  deleteCaptureHeading: (headingOrder: number) => Promise<void>;
  moveCaptureHeading: (sourceHeadingOrder: number, targetHeadingOrder: number) => Promise<void>;
};

const DND_HEADING_PREFIX = "capture-preview-heading:";

const headingDndId = (headingOrder: number) => `${DND_HEADING_PREFIX}${headingOrder}`;

const parseHeadingOrderFromDnd = (value: string | number | undefined | null) => {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string" || !value.startsWith(DND_HEADING_PREFIX)) {
    return null;
  }
  const rawOrder = Number.parseInt(value.slice(DND_HEADING_PREFIX.length), 10);
  return Number.isFinite(rawOrder) ? rawOrder : null;
};

type PreviewHeadingRowProps = {
  heading: FileHeading;
  isSelected: boolean;
  isBusy: boolean;
  setSelectedCaptureHeadingOrder: Setter<number | null>;
  deleteCaptureHeading: (headingOrder: number) => Promise<void>;
};

function PreviewHeadingRow(props: PreviewHeadingRowProps) {
  const { ref: draggableRef, handleRef, isDragging } = useDraggable({
    get id() {
      return headingDndId(props.heading.order);
    },
  });
  const { ref: droppableRef, isDropTarget } = useDroppable({
    get id() {
      return headingDndId(props.heading.order);
    },
  });

  const setCombinedRef = (element: Element | undefined) => {
    draggableRef(element);
    droppableRef(element);
  };

  return (
    <div
      class={`group flex items-center gap-3 rounded-lg border px-3 py-2 transition-all ${
        isDropTarget()
          ? "border-blue-500/50 bg-blue-500/10"
          : props.isSelected
            ? "border-emerald-500/50 bg-emerald-500/10"
            : "border-transparent hover:border-neutral-700 hover:bg-neutral-800/50"
      } ${isDragging() ? "opacity-50" : ""}`}
      ref={setCombinedRef}
    >
      <button
        aria-label={`Drag ${props.heading.text}`}
        class="cursor-grab rounded p-1 text-neutral-600 transition hover:text-neutral-400 active:cursor-grabbing"
        disabled={props.isBusy}
        ref={handleRef}
        type="button"
      >
        <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 8h16M4 16h16" />
        </svg>
      </button>
      <button
        class="min-w-0 flex-1 rounded py-1 text-left"
        onClick={() => props.setSelectedCaptureHeadingOrder(props.heading.order)}
        style={{ "padding-left": `${Math.max(0, props.heading.level - 1) * 16}px` }}
        type="button"
      >
        <span class={`mr-2 inline-flex rounded px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-wider ${
          props.heading.level === 1 
            ? "bg-blue-500/20 text-blue-400" 
            : props.heading.level === 2
              ? "bg-emerald-500/20 text-emerald-400"
              : "bg-neutral-700 text-neutral-400"
        }`}>
          H{props.heading.level}
        </span>
        <span class="truncate text-sm text-neutral-300">{props.heading.text}</span>
      </button>
      <button
        class="rounded p-1.5 text-neutral-600 transition hover:bg-rose-500/20 hover:text-rose-400 disabled:opacity-40"
        disabled={props.isBusy}
        onClick={(event) => {
          event.stopPropagation();
          void props.deleteCaptureHeading(props.heading.order);
        }}
        type="button"
      >
        <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
        </svg>
      </button>
    </div>
  );
}

export default function CaptureTargetPanel(props: CaptureTargetPanelProps) {
  const selectedHeading = () =>
    props.captureTargetH1ToH4().find((heading) => heading.order === props.selectedCaptureHeadingOrder()) ?? null;
  const selectedTargetInfo = () => {
    const absolutePath = props.selectedCaptureTargetMeta()?.absolutePath;
    if (!absolutePath) return null;

    const normalized = absolutePath.replace(/\\/g, "/");
    const split = normalized.lastIndexOf("/");
    const fileName = split >= 0 ? normalized.slice(split + 1) : normalized;
    const parentPath = split > 0 ? normalized.slice(0, split) : "/";

    return { absolutePath, fileName, parentPath };
  };

  return (
    <div class="flex h-full min-h-0 flex-col gap-4">
      <div class="flex flex-wrap items-center gap-3">
        <label class="text-xs font-medium uppercase tracking-[0.1em] text-neutral-500">Insert Into</label>
        <div class="flex flex-1 items-center gap-3">
          <select
            class="vercel-select flex-1"
            disabled={!props.captureRootPath() || props.isLoadingCaptureTargets()}
            onChange={(event) => props.setSelectedCaptureTarget(event.currentTarget.value)}
            value={props.selectedCaptureTarget()}
          >
            <option value="" disabled>
              {props.isLoadingCaptureTargets() ? "Loading..." : "Select target .docx"}
            </option>
            <For each={props.captureTargets()}>
              {(target) => (
                <option value={target.relativePath}>
                  {target.relativePath} {!target.exists && "(new)"}
                </option>
              )}
            </For>
          </select>
          <button
            class="vercel-btn vercel-btn-ghost h-9 w-9 p-0"
            disabled={!props.captureRootPath()}
            onClick={() => void props.createCaptureTarget()}
            title="Create new target"
            type="button"
          >
            <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4" />
            </svg>
          </button>
          <button
            class="vercel-btn vercel-btn-ghost h-9 w-9 p-0"
            disabled={!props.captureRootPath()}
            onClick={() => void props.selectCaptureTargetFromFilesystem()}
            title="Browse for target"
            type="button"
          >
            <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z" />
            </svg>
          </button>
        </div>
      </div>

      <Show when={props.isAllRootsSelected()}>
        <div class="flex items-center gap-2 rounded-lg border border-amber-500/30 bg-amber-500/10 px-3 py-2 text-sm text-amber-400">
          <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
          </svg>
          Select an individual root to manage destination files.
        </div>
      </Show>

      <div class="flex min-h-0 flex-1 flex-col rounded-xl border border-neutral-800 bg-neutral-950/50 p-4">
        <div class="flex items-start justify-between gap-3">
          <div class="min-w-0 flex-1">
            <p class="text-[11px] uppercase tracking-[0.08em] text-neutral-500">Current target</p>
            <Show
              when={selectedTargetInfo()}
              fallback={<p class="mt-1 truncate text-sm text-neutral-500">No target selected</p>}
            >
              {(info) => (
                <>
                  <p class="mt-1 truncate text-sm font-medium text-neutral-200" title={info().absolutePath}>
                    {info().fileName}
                  </p>
                  <p class="truncate text-xs text-neutral-600" title={info().absolutePath}>
                    {info().parentPath}
                  </p>
                </>
              )}
            </Show>
          </div>

          <span class="vercel-badge shrink-0">
            {props.isLoadingCapturePreview() ? (
              <span class="flex items-center gap-1.5">
                <div class="h-3 w-3 animate-spin rounded-full border border-neutral-600 border-t-blue-500" />
                Loading...
              </span>
            ) : (
              `${props.captureTargetPreview()?.headingCount ?? 0} headings`
            )}
          </span>
        </div>

        <Show when={selectedHeading()}>
          {(heading) => (
            <div class="mt-4 flex items-center justify-between gap-4 rounded-lg border border-emerald-500/30 bg-emerald-500/10 px-3 py-2">
              <span class="truncate text-sm text-emerald-300">
                Context: H{heading().level} - {heading().text}
              </span>
              <button
                class="vercel-btn vercel-btn-ghost h-7 px-2 py-0 text-xs"
                onClick={() => props.setSelectedCaptureHeadingOrder(null)}
                type="button"
              >
                Clear
              </button>
            </div>
          )}
        </Show>

        <div class="mt-4 min-h-0 flex-1 overflow-auto pr-1">
          <Show
            when={props.captureTargetH1ToH4().length > 0}
            fallback={
              <div class="flex flex-col items-center justify-center py-6 text-center">
                <svg class="mb-2 h-8 w-8 text-neutral-700" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                <p class="text-sm text-neutral-500">No headings yet</p>
                <p class="text-xs text-neutral-600">Add content to see headings here.</p>
              </div>
            }
          >
            <DragDropProvider
              onDragEnd={(event) => {
                if (event.canceled) return;
                const sourceOrder = parseHeadingOrderFromDnd(event.operation.source?.id ?? null);
                const targetOrder = parseHeadingOrderFromDnd(event.operation.target?.id ?? null);
                if (sourceOrder === null || targetOrder === null || sourceOrder === targetOrder) return;
                void props.moveCaptureHeading(sourceOrder, targetOrder);
              }}
            >
              <div class="space-y-2">
                <For each={props.captureTargetH1ToH4()}>
                  {(heading) => (
                    <PreviewHeadingRow
                      deleteCaptureHeading={props.deleteCaptureHeading}
                      heading={heading}
                      isBusy={props.isLoadingCapturePreview()}
                      isSelected={props.selectedCaptureHeadingOrder() === heading.order}
                      setSelectedCaptureHeadingOrder={props.setSelectedCaptureHeadingOrder}
                    />
                  )}
                </For>
              </div>
            </DragDropProvider>
          </Show>
        </div>
        
        <p class="mt-4 text-xs text-neutral-600">
          Click to set insert context. Drag to reorder.
        </p>
      </div>
    </div>
  );
}
