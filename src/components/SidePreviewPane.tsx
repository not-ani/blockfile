import { For, Show, type Accessor } from "solid-js";
import type { SidePreview } from "../lib/types";

type SidePreviewPaneProps = {
  sidePreview: Accessor<SidePreview | null>;
};

export default function SidePreviewPane(props: SidePreviewPaneProps) {
  return (
    <aside class="vercel-card flex min-h-0 flex-col overflow-hidden">
      <div class="border-b border-neutral-800 bg-neutral-900/80 px-5 py-4">
        <div class="flex items-center gap-2">
          <svg class="h-4 w-4 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
          </svg>
          <h2 class="text-sm font-medium text-neutral-200">Preview</h2>
        </div>
        <p class="mt-1 text-xs leading-5 text-neutral-500">
          Hover items to preview. <span class="text-blue-400">Space</span> to insert, <span class="text-blue-400">C</span> to copy, <span class="text-blue-400">O</span> to open.
        </p>
      </div>

      <div class="flex-1 overflow-hidden">
        <Show
          when={props.sidePreview()}
          fallback={
              <div class="flex h-full flex-col items-center justify-center p-10 text-center">
              <div class="mb-4 flex h-16 w-16 items-center justify-center rounded-2xl bg-neutral-800/50">
                <svg class="h-8 w-8 text-neutral-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              </div>
              <p class="text-sm font-medium text-neutral-400">No content selected</p>
              <p class="mt-1 text-xs text-neutral-600">Hover over a heading, F8, or author in the tree to preview.</p>
            </div>
          }
        >
          {(current) => (
            <div class="flex h-full flex-col">
              <div class="border-b border-neutral-800 bg-neutral-950/50 px-5 py-4">
                <p class={`text-sm font-medium text-white ${current().kind === "heading" ? `preview-title-h${current().headingLevel ?? 0}` : ""}`}>
                  {current().title}
                </p>
                <Show when={current().subTitle}>
                  <p class="text-xs text-neutral-500">{current().subTitle}</p>
                </Show>
              </div>
               
               <div class="flex-1 overflow-auto p-5">
                <Show
                  when={(current().richHtml?.trim().length ?? 0) > 0}
                  fallback={
                    <div class="preview-content">
                      <For each={current().text.split(/\n\s*\n/g).filter((part) => part.trim().length > 0)}>
                        {(paragraph) => (
                          <p class="mb-4 whitespace-pre-wrap">{paragraph}</p>
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
