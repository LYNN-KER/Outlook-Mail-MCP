# Outlook-Mail-MCP

**家机的极简邮箱 MCP：连接一个 Outlook 邮箱，读取邮件并回复邮件。**

这个项目只做三件事：

- 📬 查看最近邮件
- 📖 阅读邮件
- ↩️ 回复邮件

不做日历、不做网盘、不做 Power Automate，也不做多账号管理。

## 使用方式

部署到 Render 后，打开服务首页，点击 **连接 Outlook 邮箱**，完成 Microsoft 登录和授权。

授权完成后，把 MCP 地址 `https://你的服务地址/mcp` 接到家机即可。

## Microsoft 配置

在 Microsoft Entra / Azure App Registration 创建一个应用，并启用以下 Delegated permissions：

- `openid`
- `profile`
- `offline_access`
- `User.Read`
- `Mail.Read`
- `Mail.ReadWrite`
- `Mail.Send`

Redirect URI 使用 `https://你的服务地址/auth/callback`，应用类型选择 **Web**。

## 环境变量

```text
MS_CLIENT_ID=你的 Application (client) ID
MS_CLIENT_SECRET=你的 Client Secret VALUE
MS_TENANT_ID=common
TOKEN_DIR=/data
```

Render 会自动提供 `RENDER_EXTERNAL_URL`，服务器会用它生成 OAuth 回调地址；也可以手动设置 `PUBLIC_URL`。

## Render

仓库已经包含 `render.yaml`，可以直接用 Blueprint 部署。

Token 存在 Render 持久化磁盘 `/data` 中，不会提交到 GitHub。

## MCP 工具

| 工具 | 作用 |
| --- | --- |
| `list_emails` | 查看最近邮件 |
| `read_email` | 阅读指定邮件 |
| `reply_email` | 回复指定邮件 |

## 本地运行

```bash
npm install
npm start
```

默认地址：`http://localhost:3000/`。

本地 OAuth 回调：`http://localhost:3000/auth/callback`。

## 安全

OAuth 密码不会交给家机。Microsoft 登录完成后，服务只保存 OAuth token；token 文件权限为仅当前进程用户可读写。
