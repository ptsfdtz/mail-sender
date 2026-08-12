import { invoke } from "@tauri-apps/api/core";
import type { BatchType, ImportResult, MailSettings, Recipient, SendRequest, SendSummary } from "./types";

export const isTauri = () => "__TAURI_INTERNALS__" in window;

export async function importRecipients(path: string, batchType: BatchType): Promise<ImportResult> {
  return invoke("import_recipients", { path, batchType });
}

export async function loadSettings(): Promise<MailSettings> {
  return invoke("load_settings");
}

export async function saveSettings(settings: MailSettings): Promise<void> {
  return invoke("save_settings", { settings });
}

export async function renderPreview(recipient: Recipient): Promise<string> {
  return invoke("render_preview", { recipient });
}

export async function sendEmails(request: SendRequest): Promise<SendSummary> {
  return invoke("send_emails", { request });
}
