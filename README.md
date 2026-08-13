# 发邮件

一个用于批量发送 HTML 通知邮件的 Windows 桌面应用。名单从 Excel 导入，邮件内容继续使用项目原有的三份富文本模板。

## 功能

- 面试批次：向名单中的所有收件人发送面试通知。
- 录取批次：根据“录取部门”自动发送录取或未录取通知。
- 导入前校验 Excel 列名，跳过没有邮箱的行。
- 发送前逐封预览替换姓名、部门和群号后的 HTML 邮件。
- 展示实时发送进度、单封状态和最终成功/失败数量。
- SMTP 服务器与发件邮箱保存在本机，邮箱授权码仅在当前运行期间保留。

## Excel 格式

面试名单至少需要“姓名”和“邮箱”列。录取名单需要“姓名”、“邮箱”、“录取部门”和“群号”列。“录取部门”为空或填写“无”时，应用会发送未录取通知。

## 开发

需要 Node.js 20+、pnpm 10+、Rust 1.77+ 和 Windows WebView2。

```powershell
pnpm install
pnpm tauri dev
```

运行测试：

```powershell
pnpm build
cargo test --manifest-path src-tauri/Cargo.toml
```

## 构建 Windows 安装包

```powershell
pnpm tauri build
```

构建完成后，安装包位于 `src-tauri/target/release/bundle/nsis/`，主程序位于 `src-tauri/target/release/mail-sender.exe`。

## 发布版本

推送以 `v` 开头的版本标签会自动构建 Windows 安装包，并创建同名 GitHub Release 后上传安装包附件。

```powershell
git tag v2.1.1
git push origin v2.1.1
```

## 邮件模板

模板位于 `web/`，支持变量 `{{name}}`、`{{department}}`、`{{qqGroupId}}`。修改模板后需要重新构建应用。
