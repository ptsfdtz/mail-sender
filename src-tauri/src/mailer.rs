use crate::{
    recipients::{Outcome, Recipient},
    settings::MailSettings,
};
use lettre::{
    message::{header::ContentType, Mailbox},
    transport::smtp::authentication::Credentials,
    Message, SmtpTransport, Transport,
};
use serde::{Deserialize, Serialize};
use tauri::{AppHandle, Emitter};

const INTERVIEW_TEMPLATE: &str = include_str!("../../web/面试通知.html");
const ADMITTED_TEMPLATE: &str = include_str!("../../web/录取通知.html");
const REJECTED_TEMPLATE: &str = include_str!("../../web/未录取通知.html");

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SendRequest {
    pub recipients: Vec<Recipient>,
    pub settings: MailSettings,
    pub password: String,
}

#[derive(Debug, Serialize)]
pub struct SendSummary {
    pub total: usize,
    pub sent: usize,
    pub failed: usize,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct SendProgress {
    index: usize,
    total: usize,
    name: String,
    email: String,
    success: bool,
    error: Option<String>,
}

fn escape_html(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&#39;")
}

pub fn render_recipient(recipient: &Recipient) -> Result<String, String> {
    let template = match recipient.outcome {
        Outcome::Interview => INTERVIEW_TEMPLATE,
        Outcome::Admitted => ADMITTED_TEMPLATE,
        Outcome::Rejected => REJECTED_TEMPLATE,
    };
    let html = template
        .replace("{{name}}", &escape_html(&recipient.name))
        .replace(
            "{{department}}",
            &escape_html(recipient.department.as_deref().unwrap_or("")),
        )
        .replace(
            "{{qqGroupId}}",
            &escape_html(recipient.qq_group_id.as_deref().unwrap_or("")),
        );
    Ok(html)
}

fn label(outcome: Outcome) -> &'static str {
    match outcome {
        Outcome::Interview => "面试",
        Outcome::Admitted => "录取",
        Outcome::Rejected => "未录取",
    }
}

fn build_message(sender: &str, recipient: &Recipient) -> Result<Message, String> {
    let from: Mailbox = sender
        .parse()
        .map_err(|_| "发件邮箱格式不正确".to_string())?;
    let to: Mailbox = recipient
        .email
        .parse()
        .map_err(|_| format!("收件邮箱格式不正确：{}", recipient.email))?;
    Message::builder()
        .from(from)
        .to(to)
        .subject(format!(
            "南京工业大学大学生科学技术协会{}通知",
            label(recipient.outcome)
        ))
        .header(ContentType::TEXT_HTML)
        .body(render_recipient(recipient)?)
        .map_err(|error| format!("无法生成邮件：{error}"))
}

pub fn send_all(app: &AppHandle, request: SendRequest) -> Result<SendSummary, String> {
    if request.recipients.is_empty() {
        return Err("收件名单为空".into());
    }
    if request.settings.smtp_host.trim().is_empty()
        || request.settings.username.trim().is_empty()
        || request.password.is_empty()
    {
        return Err("发件设置不完整".into());
    }
    let credentials = Credentials::new(request.settings.username.clone(), request.password);
    let transport = SmtpTransport::relay(&request.settings.smtp_host)
        .map_err(|error| format!("SMTP 服务器配置错误：{error}"))?
        .port(request.settings.smtp_port)
        .credentials(credentials)
        .build();
    let total = request.recipients.len();
    let mut sent = 0;
    for (offset, recipient) in request.recipients.iter().enumerate() {
        let result = build_message(&request.settings.username, recipient).and_then(|message| {
            transport
                .send(&message)
                .map(|_| ())
                .map_err(|error| error.to_string())
        });
        let success = result.is_ok();
        if success {
            sent += 1;
        }
        let _ = app.emit(
            "send-progress",
            SendProgress {
                index: offset + 1,
                total,
                name: recipient.name.clone(),
                email: recipient.email.clone(),
                success,
                error: result.err(),
            },
        );
    }
    Ok(SendSummary {
        total,
        sent,
        failed: total - sent,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preview_escapes_spreadsheet_values() {
        let recipient = Recipient {
            row: 2,
            name: "<小明>".into(),
            email: "x@example.com".into(),
            department: Some("技术&部".into()),
            qq_group_id: Some("123".into()),
            outcome: Outcome::Admitted,
        };
        let html = render_recipient(&recipient).unwrap();
        assert!(html.contains("&lt;小明&gt;"));
        assert!(html.contains("技术&amp;部"));
        assert!(!html.contains("{{qqGroupId}}"));
    }
}
