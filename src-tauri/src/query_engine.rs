use std::collections::{HashMap, VecDeque};
use std::time::{Duration, Instant};

use futures::future;
use tauri::AppHandle;

use crate::db::{open_database, root_id};
use crate::lexical;
use crate::search::{normalize_for_search, MAX_QUERY_CHARS};
use crate::types::SearchHit;
use crate::util::{canonicalize_folder, now_ms, path_display};
use crate::vector::{self, VECTOR_MIN_QUERY_CHARS};
use crate::CommandResult;

const DEFAULT_RESULT_LIMIT: usize = 120;
const CACHE_CAPACITY: usize = 480;
const CACHE_TTL_MS: i64 = 120_000;
const LEXICAL_SOFT_BUDGET_MS: u64 = 60;
const HYBRID_SOFT_BUDGET_MS: u64 = 180;

#[derive(Clone)]
struct CacheEntry {
    created_at_ms: i64,
    results: Vec<SearchHit>,
}

#[derive(Default)]
struct QueryCache {
    order: VecDeque<String>,
    entries: HashMap<String, CacheEntry>,
}

impl QueryCache {
    fn get(&self, key: &str) -> Option<Vec<SearchHit>> {
        let entry = self.entries.get(key)?;
        if now_ms() - entry.created_at_ms > CACHE_TTL_MS {
            return None;
        }
        Some(entry.results.clone())
    }

    fn put(&mut self, key: String, results: Vec<SearchHit>) {
        if self.entries.contains_key(&key) {
            self.order.retain(|item| item != &key);
        }
        self.order.push_back(key.clone());
        self.entries.insert(
            key,
            CacheEntry {
                created_at_ms: now_ms(),
                results,
            },
        );
        while self.order.len() > CACHE_CAPACITY {
            if let Some(oldest) = self.order.pop_front() {
                self.entries.remove(&oldest);
            }
        }
    }
}

static QUERY_CACHE: std::sync::OnceLock<std::sync::Mutex<QueryCache>> = std::sync::OnceLock::new();

fn query_cache() -> &'static std::sync::Mutex<QueryCache> {
    QUERY_CACHE.get_or_init(|| std::sync::Mutex::new(QueryCache::default()))
}

pub(crate) fn clear_query_cache() {
    if let Ok(mut cache) = query_cache().lock() {
        cache.entries.clear();
        cache.order.clear();
    }
}

fn normalize_query(query: &str) -> String {
    query
        .trim()
        .chars()
        .take(MAX_QUERY_CHARS)
        .collect::<String>()
}

fn effective_limit(limit: Option<usize>) -> usize {
    limit.unwrap_or(DEFAULT_RESULT_LIMIT).clamp(10, 400)
}

fn resolve_requested_root_id(
    app: &AppHandle,
    root_path: Option<String>,
) -> CommandResult<Option<i64>> {
    let Some(root_path) = root_path else {
        return Ok(None);
    };

    let canonical = canonicalize_folder(&root_path)
        .map(|path| path_display(&path))
        .unwrap_or(root_path);
    let connection = open_database(app)?;
    root_id(&connection, &canonical)
}

fn cache_key(mode: &str, query: &str, root_id: Option<i64>, limit: usize) -> String {
    format!(
        "{mode}|{}|{}|{}",
        normalize_for_search(query),
        root_id.unwrap_or(0),
        limit
    )
}

fn dedupe_key(hit: &SearchHit) -> String {
    format!(
        "{}:{}:{}:{}:{}",
        hit.kind,
        hit.file_id,
        hit.heading_order.unwrap_or(0),
        hit.heading_text.clone().unwrap_or_default(),
        hit.relative_path
    )
}

async fn run_lexical_search_task(
    app: AppHandle,
    query: String,
    requested_root_id: Option<i64>,
    limit: usize,
    file_name_only: bool,
) -> CommandResult<Vec<SearchHit>> {
    tauri::async_runtime::spawn_blocking(move || {
        lexical::search(&app, &query, requested_root_id, limit, file_name_only)
    })
    .await
    .map_err(|error| format!("Lexical search task failed: {error}"))?
}

fn fuse_rrf(
    lexical_hits: &[SearchHit],
    semantic_hits: &[SearchHit],
    limit: usize,
) -> Vec<SearchHit> {
    let mut scores = HashMap::<String, f64>::new();
    let mut by_key = HashMap::<String, SearchHit>::new();
    let mut seen_lexical = HashMap::<String, bool>::new();
    let mut seen_semantic = HashMap::<String, bool>::new();

    for (rank, hit) in lexical_hits.iter().enumerate() {
        let key = dedupe_key(hit);
        scores
            .entry(key.clone())
            .and_modify(|value| *value += 1.0 / (60.0 + f64::from((rank + 1) as u32)))
            .or_insert(1.0 / (60.0 + f64::from((rank + 1) as u32)));
        by_key.entry(key.clone()).or_insert_with(|| hit.clone());
        seen_lexical.insert(key, true);
    }

    for (rank, hit) in semantic_hits.iter().enumerate() {
        let key = dedupe_key(hit);
        scores
            .entry(key.clone())
            .and_modify(|value| *value += 1.0 / (60.0 + f64::from((rank + 1) as u32)))
            .or_insert(1.0 / (60.0 + f64::from((rank + 1) as u32)));

        by_key
            .entry(key.clone())
            .and_modify(|existing| {
                if existing.source == "lexical" {
                    existing.source = "hybrid".to_string();
                }
            })
            .or_insert_with(|| hit.clone());
        seen_semantic.insert(key, true);
    }

    let mut ranked = scores
        .into_iter()
        .filter_map(|(key, score)| {
            let mut hit = by_key.get(&key)?.clone();
            if seen_lexical.get(&key).copied().unwrap_or(false)
                && seen_semantic.get(&key).copied().unwrap_or(false)
            {
                hit.source = "hybrid".to_string();
            }
            hit.score = 1_000.0 - (score * 1_000.0);
            Some(hit)
        })
        .collect::<Vec<SearchHit>>();

    ranked.sort_by(|left, right| {
        left.score
            .partial_cmp(&right.score)
            .unwrap_or(std::cmp::Ordering::Equal)
            .then(left.relative_path.cmp(&right.relative_path))
            .then(
                left.heading_order
                    .unwrap_or(0)
                    .cmp(&right.heading_order.unwrap_or(0)),
            )
            .then(left.kind.cmp(&right.kind))
    });
    ranked.truncate(limit);
    ranked
}

pub(crate) fn search_lexical(
    app: &AppHandle,
    query: &str,
    root_path: Option<String>,
    limit: Option<usize>,
) -> CommandResult<Vec<SearchHit>> {
    let started = Instant::now();
    let capped_query = normalize_query(query);
    let cleaned_query = capped_query.trim();
    if cleaned_query.len() < 2 {
        return Ok(Vec::new());
    }
    if normalize_for_search(cleaned_query).is_empty() {
        return Ok(Vec::new());
    }

    let requested_root_id = resolve_requested_root_id(app, root_path)?;
    let limit = effective_limit(limit);
    let key = cache_key("lexical", cleaned_query, requested_root_id, limit);
    if let Ok(cache) = query_cache().lock() {
        if let Some(cached) = cache.get(&key) {
            return Ok(cached);
        }
    }

    let results = lexical::search(app, cleaned_query, requested_root_id, limit, false)?;
    if let Ok(mut cache) = query_cache().lock() {
        cache.put(key, results.clone());
    }

    if started.elapsed().as_millis() > u128::from(LEXICAL_SOFT_BUDGET_MS) {
        eprintln!(
            "Lexical query over budget ({}ms): '{}'",
            started.elapsed().as_millis(),
            normalize_for_search(cleaned_query)
        );
    }

    Ok(results)
}

pub(crate) async fn search_semantic(
    app: &AppHandle,
    query: &str,
    root_path: Option<String>,
    limit: Option<usize>,
) -> CommandResult<Vec<SearchHit>> {
    let capped_query = normalize_query(query);
    let cleaned_query = capped_query.trim();
    if cleaned_query.chars().count() < VECTOR_MIN_QUERY_CHARS {
        return Ok(Vec::new());
    }
    if normalize_for_search(cleaned_query).is_empty() {
        return Ok(Vec::new());
    }

    let requested_root_id = resolve_requested_root_id(app, root_path)?;
    vector::trigger_rebuild(app.clone(), false);
    vector::search(
        app,
        cleaned_query,
        requested_root_id,
        effective_limit(limit),
    )
    .await
}

pub(crate) async fn search_hybrid(
    app: &AppHandle,
    query: &str,
    root_path: Option<String>,
    limit: Option<usize>,
    file_name_only: bool,
    semantic_enabled: bool,
) -> CommandResult<Vec<SearchHit>> {
    let started = Instant::now();
    let capped_query = normalize_query(query);
    let cleaned_query = capped_query.trim();
    if cleaned_query.len() < 2 {
        return Ok(Vec::new());
    }
    if normalize_for_search(cleaned_query).is_empty() {
        return Ok(Vec::new());
    }

    let requested_root_id = resolve_requested_root_id(app, root_path)?;
    let limit = effective_limit(limit);
    let mode_key = if file_name_only {
        "hybrid_file_name_only"
    } else if semantic_enabled {
        "hybrid"
    } else {
        "lexical_only"
    };
    let key = cache_key(mode_key, cleaned_query, requested_root_id, limit);
    if let Ok(cache) = query_cache().lock() {
        if let Some(cached) = cache.get(&key) {
            return Ok(cached);
        }
    }

    if file_name_only {
        let lexical_hits = run_lexical_search_task(
            app.clone(),
            cleaned_query.to_string(),
            requested_root_id,
            limit,
            true,
        )
        .await?;
        if let Ok(mut cache) = query_cache().lock() {
            cache.put(key, lexical_hits.clone());
        }
        return Ok(lexical_hits);
    }

    if !semantic_enabled {
        let lexical_hits = run_lexical_search_task(
            app.clone(),
            cleaned_query.to_string(),
            requested_root_id,
            limit,
            false,
        )
        .await?;
        if let Ok(mut cache) = query_cache().lock() {
            cache.put(key, lexical_hits.clone());
        }
        return Ok(lexical_hits);
    }

    vector::trigger_rebuild(app.clone(), false);

    let lexical_task = run_lexical_search_task(
        app.clone(),
        cleaned_query.to_string(),
        requested_root_id,
        limit,
        false,
    );
    let semantic_task = vector::search(app, cleaned_query, requested_root_id, limit);
    let (lexical_result, semantic_result) = future::join(lexical_task, semantic_task).await;

    let lexical_hits = lexical_result?;
    let semantic_hits = semantic_result.unwrap_or_default();
    let fused = fuse_rrf(&lexical_hits, &semantic_hits, limit);

    if let Ok(mut cache) = query_cache().lock() {
        cache.put(key, fused.clone());
    }

    if started.elapsed() > Duration::from_millis(HYBRID_SOFT_BUDGET_MS) {
        eprintln!(
            "Hybrid query over budget ({}ms): '{}'",
            started.elapsed().as_millis(),
            normalize_for_search(cleaned_query)
        );
    }

    Ok(fused)
}
