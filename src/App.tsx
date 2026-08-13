import { useEffect, useMemo, useState } from "react";
import { open } from "@tauri-apps/plugin-dialog";
import { listen } from "@tauri-apps/api/event";
import {
  CheckCircle2,
  CircleAlert,
  Eye,
  FileSpreadsheet,
  ClipboardCheck,
  FileCheck2,
  LoaderCircle,
  Mail,
  Send,
  Settings2,
  ShieldCheck,
  X,
} from "lucide-react";
import {
  importRecipients,
  isTauri,
  loadSettings,
  renderPreview,
  saveSettings,
  sendEmails,
} from "./tauri";
import type {
  BatchType,
  MailSettings,
  Recipient,
  SendProgress,
  SendSummary,
} from "./types";

const defaults: MailSettings = {
  smtpHost: "mail.wzj.su",
  smtpPort: 465,
  username: "",
};

function outcomeLabel(outcome: Recipient["outcome"]) {
  if (outcome === "interview") return "面试";
  return outcome === "admitted" ? "录取" : "未录取";
}

export default function App() {
  const [batchType, setBatchType] = useState<BatchType>("interview");
  const [filePath, setFilePath] = useState("");
  const [recipients, setRecipients] = useState<Recipient[]>([]);
  const [skipped, setSkipped] = useState<number[]>([]);
  const [settings, setSettings] = useState<MailSettings>(defaults);
  const [password, setPassword] = useState("");
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [preview, setPreview] = useState<{
    html: string;
    recipient: Recipient;
  } | null>(null);
  const [sending, setSending] = useState(false);
  const [progress, setProgress] = useState<SendProgress[]>([]);
  const [summary, setSummary] = useState<SendSummary | null>(null);
  const [notice, setNotice] = useState("");

  useEffect(() => {
    if (!isTauri()) return;
    loadSettings()
      .then(setSettings)
      .catch(() => undefined);
  }, []);

  useEffect(() => {
    if (!isTauri()) return;
    let disposed = false;
    let cleanup: (() => void) | undefined;
    listen<SendProgress>("send-progress", (event) => {
      setProgress((current) => [...current, event.payload]);
    }).then((unlisten) => {
      if (disposed) unlisten();
      else cleanup = unlisten;
    });
    return () => {
      disposed = true;
      cleanup?.();
    };
  }, []);

  const sent = progress.filter((item) => item.success).length;
  const percent = recipients.length
    ? Math.round((progress.length / recipients.length) * 100)
    : 0;
  const admissionCounts = useMemo(
    () => ({
      admitted: recipients.filter((r) => r.outcome === "admitted").length,
      rejected: recipients.filter((r) => r.outcome === "rejected").length,
    }),
    [recipients],
  );
  const isInterview = batchType === "interview";
  const workspace = isInterview
    ? {
        title: "面试通知",
        sendLabel: "发送面试通知",
      }
    : {
        title: "录取结果",
        sendLabel: "发送结果通知",
      };

  async function chooseFile() {
    if (!isTauri()) {
      setNotice("请在桌面应用中选择 Excel 文件");
      return;
    }
    const selected = await open({
      multiple: false,
      filters: [{ name: "Excel 名单", extensions: ["xlsx", "xls"] }],
    });
    if (!selected) return;
    setNotice("");
    setSummary(null);
    setProgress([]);
    try {
      const result = await importRecipients(selected, batchType);
      setFilePath(selected);
      setRecipients(result.recipients);
      setSkipped(result.skippedRows);
    } catch (error) {
      setNotice(String(error));
      setRecipients([]);
    }
  }

  function changeBatch(next: BatchType) {
    setBatchType(next);
    setFilePath("");
    setRecipients([]);
    setSkipped([]);
    setProgress([]);
    setSummary(null);
  }

  async function showPreview(recipient: Recipient) {
    try {
      setPreview({ html: await renderPreview(recipient), recipient });
    } catch (error) {
      setNotice(String(error));
    }
  }

  async function persistSettings() {
    if (
      !settings.smtpHost.trim() ||
      !settings.username.trim() ||
      !settings.smtpPort
    ) {
      setNotice("请完整填写 SMTP 服务器、端口和发件邮箱");
      return;
    }
    await saveSettings(settings);
    setSettingsOpen(false);
    setNotice("发件设置已保存");
  }

  async function startSending() {
    if (!recipients.length || !settings.username.trim() || !password) {
      setNotice("请先导入名单，并填写发件邮箱和授权码");
      if (!settings.username.trim()) setSettingsOpen(true);
      return;
    }
    setSending(true);
    setProgress([]);
    setSummary(null);
    setNotice("");
    try {
      await saveSettings(settings);
      setSummary(await sendEmails({ recipients, settings, password }));
    } catch (error) {
      setNotice(String(error));
    } finally {
      setSending(false);
    }
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-mark">
            <Mail size={20} />
          </span>
          <div>
            <strong>发邮件</strong>
            <small>EMAIL SENDER</small>
          </div>
        </div>
        <nav aria-label="主导航">
          <button
            className={`nav-item ${isInterview ? "active" : ""}`}
            onClick={() => changeBatch("interview")}
          >
            <ClipboardCheck size={18} />
            面试通知
          </button>
          <button
            className={`nav-item ${!isInterview ? "active" : ""}`}
            onClick={() => changeBatch("admission")}
          >
            <FileCheck2 size={18} />
            录取结果
          </button>
        </nav>
        <button
          className="account-button"
          onClick={() => setSettingsOpen(true)}
        >
          <Settings2 size={17} />
          <span>
            <strong>发件设置</strong>
            {settings.username && <small>{settings.username}</small>}
          </span>
        </button>
      </aside>

      <main>
        <header className="topbar">
          <h1>{workspace.title}</h1>
          <button className="settings-button" onClick={() => setSettingsOpen(true)} title="发件设置">
            <Settings2 size={18} />
          </button>
        </header>

        <section className="workflow list-section">
          <div className="step-header">
            <h2>
              {recipients.length ? `名单预览 · ${recipients.length} 人` : "名单"}
            </h2>
            {recipients.length > 0 && (
              <button className="secondary-button" onClick={chooseFile}>
                <FileSpreadsheet size={17} />
                更换
              </button>
            )}
          </div>
          {!recipients.length ? (
            <button className="drop-zone" onClick={chooseFile}>
              <FileSpreadsheet size={24} />
              <strong>选择名单</strong>
            </button>
          ) : (
            <div className="recipient-panel">
              <div className="file-summary">
                <div>
                  <FileSpreadsheet size={19} />
                  <span>
                    <strong>{filePath.split(/[\\/]/).pop()}</strong>
                    {skipped.length > 0 && <small>跳过 {skipped.length} 行</small>}
                  </span>
                </div>
                <div className="summary-chips">
                  {batchType === "admission" && (
                    <>
                      <span className="chip admitted">
                        录取 {admissionCounts.admitted}
                      </span>
                      <span className="chip rejected">
                        未录取 {admissionCounts.rejected}
                      </span>
                    </>
                  )}
                </div>
              </div>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>收件人</th>
                      <th>邮箱</th>
                      <th>
                        {batchType === "admission" ? "录取部门" : "邮件类型"}
                      </th>
                      <th>状态</th>
                      <th>
                        <span className="sr-only">操作</span>
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {recipients.map((recipient) => {
                      const item = progress.find(
                        (p) => p.email === recipient.email,
                      );
                      return (
                        <tr key={`${recipient.row}-${recipient.email}`}>
                          <td>
                            <strong>{recipient.name}</strong>
                          </td>
                          <td>{recipient.email}</td>
                          <td>
                            {recipient.department ||
                              outcomeLabel(recipient.outcome)}
                          </td>
                          <td>
                            <span
                              className={`status ${item ? (item.success ? "success" : "failed") : "pending"}`}
                            >
                              {item
                                ? item.success
                                  ? "已发送"
                                  : "失败"
                                : "待发送"}
                            </span>
                          </td>
                          <td>
                            <button
                              className="icon-button"
                              title="预览邮件"
                              onClick={() => showPreview(recipient)}
                            >
                              <Eye size={17} />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </section>

        <section className="send-bar">
          <div className="credential">
            <input
              id="password"
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="邮箱授权码"
              disabled={sending}
            />
          </div>
          <div className="send-status">
            {sending ? (
              <>
                <LoaderCircle className="spin" size={18} />
                <span>
                  正在发送 {progress.length}/{recipients.length}
                  <small>成功 {sent} 封</small>
                </span>
              </>
            ) : summary ? (
              <>
                <CheckCircle2 size={19} />
                <span>
                  本次发送完成
                  <small>
                    成功 {summary.sent}，失败 {summary.failed}
                  </small>
                </span>
              </>
            ) : (
              <span>
                {recipients.length ? `${recipients.length} 封` : "未选择名单"}
              </span>
            )}
          </div>
          {sending && (
            <div className="progress-track">
              <span style={{ width: `${percent}%` }} />
            </div>
          )}
          <button
            className="send-button"
            onClick={startSending}
            disabled={sending || !recipients.length}
          >
            <Send size={18} />
            {sending ? "发送中" : workspace.sendLabel}
          </button>
        </section>
        {notice && (
          <div className="toast">
            <CircleAlert size={18} />
            <span>{notice}</span>
            <button onClick={() => setNotice("")}>
              <X size={16} />
            </button>
          </div>
        )}
      </main>

      {settingsOpen && (
        <div
          className="modal-backdrop"
          onMouseDown={(e) =>
            e.target === e.currentTarget && setSettingsOpen(false)
          }
        >
          <div className="modal settings-modal">
            <div className="modal-header">
              <div>
                <p className="eyebrow">SMTP</p>
                <h2>发件设置</h2>
              </div>
              <button
                className="icon-button"
                onClick={() => setSettingsOpen(false)}
              >
                <X size={18} />
              </button>
            </div>
            <div className="form-grid">
              <label className="wide">
                发件邮箱
                <input
                  type="email"
                  value={settings.username}
                  onChange={(e) =>
                    setSettings({ ...settings, username: e.target.value })
                  }
                  placeholder="name@example.com"
                />
              </label>
              <label>
                SMTP 服务器
                <input
                  value={settings.smtpHost}
                  onChange={(e) =>
                    setSettings({ ...settings, smtpHost: e.target.value })
                  }
                />
              </label>
              <label>
                SSL 端口
                <input
                  type="number"
                  value={settings.smtpPort}
                  onChange={(e) =>
                    setSettings({
                      ...settings,
                      smtpPort: Number(e.target.value),
                    })
                  }
                />
              </label>
            </div>
            <div className="security-copy">
              <ShieldCheck size={18} />
              <span>服务器和邮箱会保存在本机。授权码不会保存。</span>
            </div>
            <div className="modal-actions">
              <button
                className="text-button"
                onClick={() => setSettingsOpen(false)}
              >
                取消
              </button>
              <button className="primary-button" onClick={persistSettings}>
                保存设置
              </button>
            </div>
          </div>
        </div>
      )}
      {preview && (
        <div className="modal-backdrop">
          <div className="modal preview-modal">
            <div className="modal-header">
              <div>
                <p className="eyebrow">
                  {outcomeLabel(preview.recipient.outcome)}通知
                </p>
                <h2>
                  {preview.recipient.name} · {preview.recipient.email}
                </h2>
              </div>
              <button className="icon-button" onClick={() => setPreview(null)}>
                <X size={18} />
              </button>
            </div>
            <iframe title="邮件预览" srcDoc={preview.html} sandbox="" />
          </div>
        </div>
      )}
    </div>
  );
}
