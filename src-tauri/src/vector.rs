use tauri::AppHandle;

use crate::semantic::{semantic_search, trigger_semantic_rebuild, SEMANTIC_MIN_QUERY_CHARS};
use crate::types::SearchHit;
use crate::CommandResult;

pub(crate) const VECTOR_MIN_QUERY_CHARS: usize = SEMANTIC_MIN_QUERY_CHARS;

pub(crate) fn trigger_rebuild(app: AppHandle, force: bool) {
    trigger_semantic_rebuild(app, force);
}

pub(crate) async fn search(
    app: &AppHandle,
    query: &str,
    requested_root_id: Option<i64>,
    limit: usize,
) -> CommandResult<Vec<SearchHit>> {
    semantic_search(app, query, requested_root_id, limit).await
}
