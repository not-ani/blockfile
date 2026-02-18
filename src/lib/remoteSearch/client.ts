import { normalizeDebatifySearchResponse } from "./normalize";
import type { DebatifyTagHit } from "./types";

const DEBATIFY_SEARCH_ENDPOINT = "https://api.debatify.app/search";

type SearchDebatifyTagsOptions = {
  signal?: AbortSignal;
};

export const searchDebatifyTags = async (
  query: string,
  options: SearchDebatifyTagsOptions = {},
): Promise<DebatifyTagHit[]> => {
  const trimmedQuery = query.trim();
  if (trimmedQuery.length < 2) return [];

  const url = new URL(DEBATIFY_SEARCH_ENDPOINT);
  url.searchParams.set("q", trimmedQuery);

  const response = await fetch(url.toString(), {
    method: "GET",
    headers: {
      Accept: "application/json",
    },
    signal: options.signal,
  });

  if (!response.ok) {
    throw new Error(`Debatify API search failed (${response.status})`);
  }

  const payload: unknown = await response.json();
  return normalizeDebatifySearchResponse(payload, trimmedQuery);
};

export const isAbortError = (error: unknown) =>
  error instanceof DOMException && error.name === "AbortError";
