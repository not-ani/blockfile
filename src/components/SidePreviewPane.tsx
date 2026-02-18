import { For, Show, type Accessor } from "solid-js";
import type { SidePreview } from "../lib/types";

type SidePreviewPaneProps = {
  sidePreview: Accessor<SidePreview | null>;
  width: Accessor<number>;
};

export default function SidePreviewPane(props: SidePreviewPaneProps) {
  return (
    <aside class="flex min-h-0 shrink-0 flex-col border-l border-neutral-800/50 bg-neutral-950/30" style={{ width: `${props.width()}px` }}>
      <div class="flex items-center justify-between border-b border-neutral-800/50 bg-neutral-900/50 px-3 py-2">
        <div class="flex items-center gap-2">
          <svg class="h-3.5 w-3.5 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
          </svg>
          <h2 class="text-xs font-medium text-neutral-300">Preview</h2>
        </div>
        <p class="text-[10px] text-neutral-500">
          <span class="text-blue-400">Space</span> insert · <span class="text-blue-400">C</span> copy · <span class="text-blue-400">O</span> open
        </p>
      </div>

      <div class="flex-1 overflow-hidden">
        <Show
          when={props.sidePreview()}
          fallback={
            <div class="flex h-full flex-col items-center justify-center p-6 text-center">
              <div class="mb-3 flex h-12 w-12 items-center justify-center rounded-xl bg-neutral-800/50">
                <svg class="h-6 w-6 text-neutral-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              </div>
              <p class="text-xs font-medium text-neutral-400">No content selected</p>
              <p class="mt-1 text-[10px] text-neutral-600">Hover over items in the tree to preview</p>
            </div>
          }
        >
          {(current) => (
            <div class="flex h-full flex-col">
              <div class="border-b border-neutral-800/50 bg-neutral-950/50 px-3 py-2.5">
                <p class={`text-xs font-medium text-white ${current().kind === "heading" ? `preview-title-h${current().headingLevel ?? 0}` : ""}`}>
                  {current().title}
                </p>
                <Show when={current().subTitle}>
                  <p class="mt-0.5 text-[10px] text-neutral-500">{current().subTitle}</p>
                </Show>
              </div>
              
              <div class="flex-1 overflow-auto p-3">
                <Show
                  when={(current().richHtml?.trim().length ?? 0) > 0}
                  fallback={
                    <div class="preview-content text-[13px] leading-7">
                      <For each={current().text.split(/\n\s*\n/g).filter((part) => part.trim().length > 0)}>
                        {(paragraph) => (
                          <p class="mb-3 whitespace-pre-wrap">{paragraph}</p>
                        )}
                      </For>
                    </div>
                  }
                >
                  <div class="preview-rich" innerHTML={current().richHtml ?? ""} />
                </Show>
              </div>
            </div>
          )}
        </Show>
      </div>
    </aside>
  );
}
