mod chunking;
mod commands;
mod db;
mod docx_capture;
mod docx_parse;
mod indexer;
mod lexical;
mod preview;
mod query_engine;
mod search;
mod semantic;
mod types;
mod util;
mod vector;

pub(crate) type CommandResult<T> = Result<T, String>;

pub(crate) const DEFAULT_CAPTURE_TARGET: &str = "BlockFile-Captures.docx";

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            commands::add_root,
            commands::remove_root,
            commands::insert_capture,
            commands::list_capture_targets,
            commands::get_capture_target_preview,
            commands::add_capture_heading,
            commands::delete_capture_heading,
            commands::move_capture_heading,
            commands::list_roots,
            commands::index_root,
            commands::get_index_snapshot,
            commands::get_file_preview,
            commands::get_heading_preview_html,
            commands::search_index,
            commands::search_index_semantic,
            commands::search_index_hybrid,
            commands::benchmark_root_performance
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
