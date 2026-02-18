import { DragDropProvider, useDraggable, useDroppable } from "@dnd-kit/solid";
import {
  For,
  Show,
  createMemo,
  createSignal,
  type Accessor,
  type Setter,
} from "solid-js";
import type {
  CaptureTarget,
  CaptureTargetPreview,
  FileHeading,
} from "../lib/types";

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
  addCaptureHeading: (headingLevel: 1 | 2 | 3 | 4, headingText: string) => Promise<boolean>;
  deleteCaptureHeading: (headingOrder: number) => Promise<void>;
  moveCaptureHeading: (
    sourceHeadingOrder: number,
    targetHeadingOrder: number,
  ) => Promise<void>;
};

const DND_HEADING_PREFIX = "capture-preview-heading:";

const headingDndId = (headingOrder: number) =>
  `${DND_HEADING_PREFIX}${headingOrder}`;

const parseHeadingOrderFromDnd = (
  value: string | number | undefined | null,
) => {
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
  canMoveUp: boolean;
  canMoveDown: boolean;
  setSelectedCaptureHeadingOrder: Setter<number | null>;
  moveHeadingUp: (headingOrder: number) => Promise<void>;
  moveHeadingDown: (headingOrder: number) => Promise<void>;
  deleteCaptureHeading: (headingOrder: number) => Promise<void>;
};

function PreviewHeadingRow(props: PreviewHeadingRowProps) {
  const {
    ref: draggableRef,
    handleRef,
    isDragging,
  } = useDraggable({
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
      class={`group flex items-center gap-2 rounded border px-2 py-1.5 transition-all ${
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
        class="cursor-grab rounded p-0.5 text-neutral-600 transition hover:text-neutral-400 active:cursor-grabbing"
        disabled={props.isBusy}
        ref={handleRef}
        type="button"
      >
        <svg
          class="h-3 w-3"
          fill="none"
          stroke="currentColor"
          viewBox="0 0 24 24"
        >
          <path
            stroke-linecap="round"
            stroke-linejoin="round"
            stroke-width="2"
            d="M4 8h16M4 16h16"
          />
        </svg>
      </button>
      <button
        class="min-w-0 flex-1 rounded py-0.5 text-left"
        onClick={() =>
          props.setSelectedCaptureHeadingOrder(props.heading.order)
        }
        style={{
          "padding-left": `${Math.max(0, props.heading.level - 1) * 12}px`,
        }}
        type="button"
      >
        <span
          class={`mr-1.5 inline-flex rounded px-1 py-0 text-[9px] font-semibold uppercase ${
            props.heading.level === 1
              ? "bg-blue-500/20 text-blue-400"
              : props.heading.level === 2
                ? "bg-emerald-500/20 text-emerald-400"
                : "bg-neutral-700 text-neutral-400"
          }`}
        >
          H{props.heading.level}
        </span>
        <span class="truncate text-xs text-neutral-300">
          {props.heading.text}
        </span>
      </button>
      <div class="flex items-center gap-1">
        <button
          class="rounded p-1 text-neutral-600 transition hover:bg-neutral-700/70 hover:text-neutral-200 disabled:opacity-35"
          disabled={props.isBusy || !props.canMoveUp}
          onClick={(event) => {
            event.stopPropagation();
            void props.moveHeadingUp(props.heading.order);
          }}
          title="Move up"
          type="button"
        >
          <svg class="h-3 w-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 15l7-7 7 7" />
          </svg>
        </button>
        <button
          class="rounded p-1 text-neutral-600 transition hover:bg-neutral-700/70 hover:text-neutral-200 disabled:opacity-35"
          disabled={props.isBusy || !props.canMoveDown}
          onClick={(event) => {
            event.stopPropagation();
            void props.moveHeadingDown(props.heading.order);
          }}
          title="Move down"
          type="button"
        >
          <svg class="h-3 w-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
          </svg>
        </button>
      </div>
      <button
        class="rounded p-1 text-neutral-600 transition hover:bg-rose-500/20 hover:text-rose-400 disabled:opacity-40"
        disabled={props.isBusy}
        onClick={(event) => {
          event.stopPropagation();
          void props.deleteCaptureHeading(props.heading.order);
        }}
        type="button"
      >
        <svg
          class="h-3 w-3"
          fill="none"
          stroke="currentColor"
          viewBox="0 0 24 24"
        >
          <path
            stroke-linecap="round"
            stroke-linejoin="round"
            stroke-width="2"
            d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"
          />
        </svg>
      </button>
    </div>
  );
}

export default function CaptureTargetPanel(props: CaptureTargetPanelProps) {
  const quickHeadingActions = [
    { level: 1 as const, name: "pocket" },
    { level: 2 as const, name: "hat" },
    { level: 3 as const, name: "block" },
    { level: 4 as const, name: "section" },
  ];
  const headingNameByLevel: Record<1 | 2 | 3 | 4, string> = {
    1: "pocket",
    2: "hat",
    3: "block",
    4: "section",
  };
  const [newHeadingLevel, setNewHeadingLevel] = createSignal<1 | 2 | 3 | 4>(1);
  const [newHeadingName, setNewHeadingName] = createSignal(headingNameByLevel[1]);

  const setHeadingLevel = (level: 1 | 2 | 3 | 4) => {
    setNewHeadingLevel(level);
    setNewHeadingName(headingNameByLevel[level]);
  };

  const submitNewHeading = async () => {
    const created = await props.addCaptureHeading(newHeadingLevel(), newHeadingName());
    if (created) {
      setNewHeadingName("");
    }
  };

  const moveHintsByOrder = createMemo(() => {
    const headings = [...props.captureTargetH1ToH4()].sort(
      (left, right) => left.order - right.order,
    );
    const levelByOrder = new Map<number, number>();
    headings.forEach((heading) => {
      levelByOrder.set(heading.order, heading.level);
    });

    const siblingBuckets = new Map<string, number[]>();
    const stack: number[] = [];
    for (const heading of headings) {
      while (stack.length > 0) {
        const top = stack[stack.length - 1];
        const topLevel = levelByOrder.get(top) ?? 0;
        if (topLevel >= heading.level) {
          stack.pop();
          continue;
        }
        break;
      }

      const parentOrder = stack.length > 0 ? stack[stack.length - 1] : 0;
      const bucketKey = `${parentOrder}:${heading.level}`;
      const bucket = siblingBuckets.get(bucketKey) ?? [];
      bucket.push(heading.order);
      siblingBuckets.set(bucketKey, bucket);
      stack.push(heading.order);
    }

    const hints = new Map<number, { up: number | null; down: number | null }>();
    siblingBuckets.forEach((orders) => {
      orders.forEach((order, index) => {
        hints.set(order, {
          up: index > 0 ? orders[index - 1] : null,
          down: index < orders.length - 1 ? orders[index + 1] : null,
        });
      });
    });

    return hints;
  });

  const moveHeadingUp = async (headingOrder: number) => {
    const targetOrder = moveHintsByOrder().get(headingOrder)?.up;
    if (!targetOrder) return;
    await props.moveCaptureHeading(targetOrder, headingOrder);
  };

  const moveHeadingDown = async (headingOrder: number) => {
    const targetOrder = moveHintsByOrder().get(headingOrder)?.down;
    if (!targetOrder) return;
    await props.moveCaptureHeading(headingOrder, targetOrder);
  };

  const moveHintForHeading = (headingOrder: number) =>
    moveHintsByOrder().get(headingOrder);

  const selectedHeading = () =>
    props
      .captureTargetH1ToH4()
      .find(
        (heading) => heading.order === props.selectedCaptureHeadingOrder(),
      ) ?? null;
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
    <div class="flex h-full min-h-0 flex-col">
      <div class="flex items-center gap-2 border-b border-neutral-800/50 bg-neutral-900/30 px-3 py-2">
        <div class="flex flex-1 items-center gap-2 overflow-x-scroll">
          <select
            class="h-7 flex-1 rounded border border-neutral-700 bg-neutral-950 px-2 text-xs text-neutral-200 outline-none transition hover:border-neutral-600 focus:border-blue-500"
            disabled={
              !props.captureRootPath() || props.isLoadingCaptureTargets()
            }
            onChange={(event) =>
              props.setSelectedCaptureTarget(event.currentTarget.value)
            }
            value={props.selectedCaptureTarget()}
          >
            <option value="" disabled>
              {props.isLoadingCaptureTargets()
                ? "Loading..."
                : "Select target .docx"}
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
            class="inline-flex h-7 w-7 items-center justify-center rounded border border-neutral-700 bg-neutral-800 text-neutral-300 transition hover:border-neutral-600 hover:bg-neutral-700 hover:text-white disabled:opacity-50"
            disabled={!props.captureRootPath()}
            onClick={() => void props.createCaptureTarget()}
            title="Create new target"
            type="button"
          >
            <svg
              class="h-3.5 w-3.5"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                stroke-linecap="round"
                stroke-linejoin="round"
                stroke-width="2"
                d="M12 4v16m8-8H4"
              />
            </svg>
          </button>
          <button
            class="inline-flex h-7 w-7 items-center justify-center rounded border border-neutral-700 bg-neutral-800 text-neutral-300 transition hover:border-neutral-600 hover:bg-neutral-700 hover:text-white disabled:opacity-50"
            disabled={!props.captureRootPath()}
            onClick={() => void props.selectCaptureTargetFromFilesystem()}
            title="Browse for target"
            type="button"
          >
            <svg
              class="h-3.5 w-3.5"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                stroke-linecap="round"
                stroke-linejoin="round"
                stroke-width="2"
                d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z"
              />
            </svg>
          </button>
        </div>
      </div>

      <Show when={props.isAllRootsSelected()}>
        <div class="flex items-center gap-2 border-b border-amber-500/20 bg-amber-500/5 px-3 py-2 text-xs text-amber-400">
          <svg
            class="h-3.5 w-3.5 shrink-0"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path
              stroke-linecap="round"
              stroke-linejoin="round"
              stroke-width="2"
              d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"
            />
          </svg>
          <span class="text-[10px]">
            Select an individual root to manage destination files
          </span>
        </div>
      </Show>

      <div class="flex min-h-0 flex-1 flex-col p-3">
        <div class="flex items-start justify-between gap-2">
          <div class="min-w-0 flex-1">
            <p class="text-[10px] uppercase tracking-wider text-neutral-500">
              Current target
            </p>
            <Show
              when={selectedTargetInfo()}
              fallback={
                <p class="mt-0.5 truncate text-xs text-neutral-500">
                  No target selected
                </p>
              }
            >
              {(info) => (
                <>
                  <p
                    class="mt-0.5 truncate text-xs font-medium text-neutral-200"
                    title={info().absolutePath}
                  >
                    {info().fileName}
                  </p>
                  <p
                    class="truncate text-[10px] text-neutral-600"
                    title={info().absolutePath}
                  >
                    {info().parentPath}
                  </p>
                </>
              )}
            </Show>
          </div>

          <span class="inline-flex shrink-0 items-center rounded-full border border-neutral-700 bg-neutral-800 px-2 py-0.5 text-[10px] text-neutral-300">
            {props.isLoadingCapturePreview() ? (
              <span class="flex items-center gap-1">
                <div class="h-2.5 w-2.5 animate-spin rounded-full border border-neutral-600 border-t-blue-500" />
              </span>
            ) : (
              `${props.captureTargetPreview()?.headingCount ?? 0} headings`
            )}
          </span>
        </div>

        <Show when={selectedHeading()}>
          {(heading) => (
            <div class="mt-2 flex items-center justify-between gap-2 rounded border border-emerald-500/30 bg-emerald-500/10 px-2 py-1.5">
              <span class="truncate text-xs text-emerald-300">
                H{heading().level}: {heading().text}
              </span>
              <button
                class="rounded px-1.5 py-0.5 text-[10px] text-neutral-400 transition hover:bg-neutral-800 hover:text-neutral-200"
                onClick={() => props.setSelectedCaptureHeadingOrder(null)}
                type="button"
              >
                Clear
              </button>
            </div>
          )}
        </Show>

        <div class="mt-2 rounded border border-neutral-800/70 bg-neutral-900/30 px-2 py-2">
          <div class="flex items-center justify-between gap-2">
            <p class="text-[10px] uppercase tracking-wider text-neutral-500">Add heading</p>
            <p class="text-[10px] text-neutral-600">Creates new heading in target doc</p>
          </div>
          <div class="mt-2 flex flex-wrap gap-1.5">
            <For each={quickHeadingActions}>
              {(action) => (
                <button
                  class="inline-flex items-center gap-1 rounded border border-neutral-700 bg-neutral-900 px-2 py-1 text-[10px] font-medium text-neutral-200 transition hover:border-neutral-600 hover:bg-neutral-800 disabled:cursor-not-allowed disabled:opacity-50"
                  disabled={!props.captureRootPath() || !props.selectedCaptureTarget() || props.isLoadingCapturePreview()}
                  onClick={() => void props.addCaptureHeading(action.level, action.name)}
                  type="button"
                >
                  <span class="rounded bg-blue-500/20 px-1 py-0 text-[9px] font-semibold uppercase text-blue-300">H{action.level}</span>
                  <span>{action.name}</span>
                </button>
              )}
            </For>
          </div>

          <div class="mt-2 flex items-center gap-1.5">
            <select
              class="h-7 rounded border border-neutral-700 bg-neutral-950 px-2 text-xs text-neutral-200 outline-none transition hover:border-neutral-600 focus:border-blue-500"
              disabled={!props.captureRootPath() || !props.selectedCaptureTarget() || props.isLoadingCapturePreview()}
              onChange={(event) => {
                const level = Number.parseInt(event.currentTarget.value, 10);
                if (level >= 1 && level <= 4) {
                  setHeadingLevel(level as 1 | 2 | 3 | 4);
                }
              }}
              value={newHeadingLevel()}
            >
              <option value="1">H1</option>
              <option value="2">H2</option>
              <option value="3">H3</option>
              <option value="4">H4</option>
            </select>

            <input
              class="h-7 min-w-0 flex-1 rounded border border-neutral-700 bg-neutral-950 px-2 text-xs text-neutral-200 outline-none transition placeholder:text-neutral-600 hover:border-neutral-600 focus:border-blue-500"
              disabled={!props.captureRootPath() || !props.selectedCaptureTarget() || props.isLoadingCapturePreview()}
              onInput={(event) => setNewHeadingName(event.currentTarget.value)}
              onKeyDown={(event) => {
                if (event.key !== "Enter") return;
                event.preventDefault();
                void submitNewHeading();
              }}
              placeholder="Heading name"
              value={newHeadingName()}
            />

            <button
              class="inline-flex h-7 items-center justify-center rounded border border-neutral-700 bg-neutral-800 px-2 text-[10px] font-medium text-neutral-200 transition hover:border-neutral-600 hover:bg-neutral-700 disabled:cursor-not-allowed disabled:opacity-50"
              disabled={!props.captureRootPath() || !props.selectedCaptureTarget() || props.isLoadingCapturePreview() || !newHeadingName().trim()}
              onClick={() => void submitNewHeading()}
              type="button"
            >
              Add
            </button>
          </div>
        </div>

        <div class="mt-2 min-h-0 flex-1 overflow-auto">
          <Show
            when={props.captureTargetH1ToH4().length > 0}
            fallback={
              <div class="flex flex-col items-center justify-center py-4 text-center">
                <svg
                  class="mb-1.5 h-6 w-6 text-neutral-700"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    stroke-linecap="round"
                    stroke-linejoin="round"
                    stroke-width="2"
                    d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                  />
                </svg>
                <p class="text-xs text-neutral-500">No headings yet</p>
              </div>
            }
          >
            <DragDropProvider
              onDragEnd={(event) => {
                if (event.canceled) return;
                const sourceOrder = parseHeadingOrderFromDnd(
                  event.operation.source?.id ?? null,
                );
                const targetOrder = parseHeadingOrderFromDnd(
                  event.operation.target?.id ?? null,
                );
                if (
                  sourceOrder === null ||
                  targetOrder === null ||
                  sourceOrder === targetOrder
                )
                  return;
                void props.moveCaptureHeading(sourceOrder, targetOrder);
              }}
            >
              <div class="space-y-1">
                <For each={props.captureTargetH1ToH4()}>
                  {(heading) => (
                    <PreviewHeadingRow
                      canMoveDown={Boolean(moveHintForHeading(heading.order)?.down)}
                      canMoveUp={Boolean(moveHintForHeading(heading.order)?.up)}
                      deleteCaptureHeading={props.deleteCaptureHeading}
                      heading={heading}
                      isBusy={props.isLoadingCapturePreview()}
                      isSelected={
                        props.selectedCaptureHeadingOrder() === heading.order
                      }
                      moveHeadingDown={moveHeadingDown}
                      moveHeadingUp={moveHeadingUp}
                      setSelectedCaptureHeadingOrder={
                        props.setSelectedCaptureHeadingOrder
                      }
                    />
                  )}
                </For>
              </div>
            </DragDropProvider>
          </Show>
        </div>

        <p class="mt-2 text-[10px] text-neutral-600">
          Click to set context · Drag to reorder
        </p>
      </div>
    </div>
  );
}
