use tauri::AppHandle;

use crate::db::open_database;
use crate::lexical;
use crate::CommandResult;

pub(crate) fn rebuild_lexical_index(app: &AppHandle) -> CommandResult<()> {
    let connection = open_database(app)?;
    lexical::replace_all_documents_from_connection(app, &connection)?;
    Ok(())
}
