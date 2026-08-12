export type BatchType = "interview" | "admission";

export interface Recipient {
  row: number;
  name: string;
  email: string;
  department: string | null;
  qqGroupId: string | null;
  outcome: "interview" | "admitted" | "rejected";
}

export interface ImportResult {
  recipients: Recipient[];
  skippedRows: number[];
}

export interface MailSettings {
  smtpHost: string;
  smtpPort: number;
  username: string;
}

export interface SendRequest {
  recipients: Recipient[];
  settings: MailSettings;
  password: string;
}

export interface SendProgress {
  index: number;
  total: number;
  name: string;
  email: string;
  success: boolean;
  error: string | null;
}

export interface SendSummary {
  total: number;
  sent: number;
  failed: number;
}
