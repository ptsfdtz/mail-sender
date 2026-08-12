use serde::{Deserialize, Serialize};
use std::{fs, path::Path};

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MailSettings {
    pub smtp_host: String,
    pub smtp_port: u16,
    pub username: String,
}

impl Default for MailSettings {
    fn default() -> Self {
        Self {
            smtp_host: "mail.wzj.su".into(),
            smtp_port: 465,
            username: String::new(),
        }
    }
}

pub fn load(directory: &Path) -> Result<MailSettings, String> {
    let path = directory.join("settings.json");
    if !path.exists() {
        return Ok(MailSettings::default());
    }
    let content = fs::read_to_string(&path).map_err(|error| format!("无法读取设置：{error}"))?;
    serde_json::from_str(&content).map_err(|error| format!("设置文件格式错误：{error}"))
}

pub fn save(directory: &Path, settings: &MailSettings) -> Result<(), String> {
    fs::create_dir_all(directory).map_err(|error| format!("无法创建设置目录：{error}"))?;
    let content = serde_json::to_string_pretty(settings).map_err(|error| error.to_string())?;
    fs::write(directory.join("settings.json"), content)
        .map_err(|error| format!("无法保存设置：{error}"))
}
