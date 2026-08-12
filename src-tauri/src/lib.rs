mod mailer;
mod recipients;
mod settings;

use mailer::{render_recipient, SendRequest, SendSummary};
use recipients::{BatchType, ImportResult, Recipient};
use settings::MailSettings;
use tauri::{AppHandle, Manager};

#[tauri::command]
fn import_recipients(path: String, batch_type: BatchType) -> Result<ImportResult, String> {
    recipients::read_workbook(&path, batch_type)
}

#[tauri::command]
fn render_preview(recipient: Recipient) -> Result<String, String> {
    render_recipient(&recipient)
}

#[tauri::command]
fn load_settings(app: AppHandle) -> Result<MailSettings, String> {
    let directory = app
        .path()
        .app_config_dir()
        .map_err(|error| error.to_string())?;
    settings::load(&directory)
}

#[tauri::command]
fn save_settings(app: AppHandle, settings: MailSettings) -> Result<(), String> {
    let directory = app
        .path()
        .app_config_dir()
        .map_err(|error| error.to_string())?;
    settings::save(&directory, &settings)
}

#[tauri::command]
async fn send_emails(app: AppHandle, request: SendRequest) -> Result<SendSummary, String> {
    tauri::async_runtime::spawn_blocking(move || mailer::send_all(&app, request))
        .await
        .map_err(|error| format!("发送任务异常终止：{error}"))?
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            import_recipients,
            render_preview,
            load_settings,
            save_settings,
            send_emails
        ])
        .run(tauri::generate_context!())
        .expect("failed to run mail sender");
}
