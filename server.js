const http = require('node:http');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { URL } = require('node:url');
const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StreamableHTTPServerTransport } = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const { CallToolRequestSchema, ListToolsRequestSchema } = require('@modelcontextprotocol/sdk/types.js');

const PORT = Number(process.env.PORT || 3000);
const CLIENT_ID = process.env.MS_CLIENT_ID || '';
const CLIENT_SECRET = process.env.MS_CLIENT_SECRET || '';
const TENANT_ID = process.env.MS_TENANT_ID || 'common';
const BASE_URL = (process.env.PUBLIC_URL || process.env.RENDER_EXTERNAL_URL || `http://localhost:${PORT}`).replace(/\/$/, '');
const SCOPES = ['openid', 'profile', 'offline_access', 'User.Read', 'Mail.Read', 'Mail.ReadWrite', 'Mail.Send'];
const TOKEN_DIR = process.env.TOKEN_DIR || path.join(process.cwd(), 'data');
const TOKEN_FILE = path.join(TOKEN_DIR, 'tokens.json');
const GRAPH = 'https://graph.microsoft.com/v1.0';

let oauthState = null;
let tokens = loadTokens();
let refreshPromise = null;

function loadTokens() {
  try { return JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8')); } catch { return null; }
}
function saveTokens(value) {
  fs.mkdirSync(TOKEN_DIR, { recursive: true, mode: 0o700 });
  fs.writeFileSync(TOKEN_FILE, JSON.stringify(value, null, 2), { mode: 0o600 });
  tokens = value;
}
function redirectUri() { return `${BASE_URL}/auth/callback`; }
function authBase() { return `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0`; }

async function exchange(params) {
  const body = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, ...params });
  const response = await fetch(`${authBase()}/token`, { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error_description || `OAuth error ${response.status}`);
  return data;
}

async function accessToken() {
  if (!tokens?.access_token) throw new Error('邮箱尚未连接，请先打开 /auth/login 完成 Microsoft 授权。');
  if (tokens.expires_at && Date.now() < tokens.expires_at - 60_000) return tokens.access_token;
  if (!tokens.refresh_token) throw new Error('Microsoft 登录已失效，请重新打开 /auth/login。');
  if (!refreshPromise) {
    refreshPromise = exchange({ grant_type: 'refresh_token', refresh_token: tokens.refresh_token, scope: SCOPES.join(' ') })
      .then(next => { saveTokens({ ...tokens, ...next, refresh_token: next.refresh_token || tokens.refresh_token, expires_at: Date.now() + (next.expires_in || 3600) * 1000 }); return tokens.access_token; })
      .finally(() => { refreshPromise = null; });
  }
  return refreshPromise;
}

async function graph(pathname, options = {}) {
  const token = await accessToken();
  const response = await fetch(`${GRAPH}${pathname}`, { ...options, headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json', ...(options.headers || {}) } });
  const text = await response.text();
  let data = {};
  try { data = text ? JSON.parse(text) : {}; } catch { data = { raw: text }; }
  if (response.status === 401 && tokens?.refresh_token) {
    tokens.expires_at = 0;
    const retryToken = await accessToken();
    const retry = await fetch(`${GRAPH}${pathname}`, { ...options, headers: { authorization: `Bearer ${retryToken}`, 'content-type': 'application/json', ...(options.headers || {}) } });
    const retryText = await retry.text();
    let retryData = {}; try { retryData = retryText ? JSON.parse(retryText) : {}; } catch { retryData = { raw: retryText }; }
    if (!retry.ok) throw new Error(retryData.error?.message || `Graph error ${retry.status}`);
    return retryData;
  }
  if (!response.ok) throw new Error(data.error?.message || `Graph error ${response.status}`);
  return data;
}

const server = new Server({ name: 'outlook-mail-mcp', version: '0.1.0' }, { capabilities: { tools: {} } });
server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: [
  { name: 'list_emails', description: '读取连接的 Outlook 邮箱最近的邮件。', inputSchema: { type: 'object', properties: { top: { type: 'number', description: '最多返回多少封，默认 10。' } } } },
  { name: 'read_email', description: '读取一封邮件的完整正文和基本信息。', inputSchema: { type: 'object', required: ['message_id'], properties: { message_id: { type: 'string' } } } },
  { name: 'reply_email', description: '回复指定的 Outlook 邮件。', inputSchema: { type: 'object', required: ['message_id', 'body'], properties: { message_id: { type: 'string' }, body: { type: 'string', description: '要发送的回复正文。' } } } }
] }));
server.setRequestHandler(CallToolRequestSchema, async request => {
  const args = request.params.arguments || {};
  try {
    if (request.params.name === 'list_emails') {
      const top = Math.min(Math.max(Number(args.top || 10), 1), 25);
      const data = await graph(`/me/mailFolders/inbox/messages?$top=${top}&$orderby=receivedDateTime%20desc&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments`);
      return { content: [{ type: 'text', text: JSON.stringify(data.value || [], null, 2) }] };
    }
    if (request.params.name === 'read_email') {
      const id = encodeURIComponent(args.message_id);
      const data = await graph(`/me/messages/${id}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,isRead,hasAttachments`);
      return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
    }
    if (request.params.name === 'reply_email') {
      const id = encodeURIComponent(args.message_id);
      await graph(`/me/messages/${id}/reply`, { method: 'POST', body: JSON.stringify({ message: { body: { contentType: 'Text', content: String(args.body) } } }) });
      return { content: [{ type: 'text', text: '邮件已回复。' }] };
    }
    throw new Error(`未知工具：${request.params.name}`);
  } catch (error) {
    return { isError: true, content: [{ type: 'text', text: error.message }] };
  }
});

async function handleMcp(req, res) {
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  await server.connect(transport);
  await transport.handleRequest(req, res);
}

function send(res, status, body, contentType = 'text/html; charset=utf-8') {
  res.writeHead(status, { 'content-type': contentType }); res.end(body);
}

const app = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, BASE_URL);
    if (req.method === 'GET' && url.pathname === '/health') return send(res, 200, JSON.stringify({ ok: true, connected: Boolean(tokens?.access_token) }), 'application/json');
    if (req.method === 'GET' && url.pathname === '/auth/login') {
      if (!CLIENT_ID || !CLIENT_SECRET) return send(res, 500, '请先配置 MS_CLIENT_ID 和 MS_CLIENT_SECRET。');
      oauthState = crypto.randomBytes(24).toString('hex');
      const auth = new URL(`${authBase()}/authorize`);
      auth.searchParams.set('client_id', CLIENT_ID); auth.searchParams.set('response_type', 'code'); auth.searchParams.set('redirect_uri', redirectUri()); auth.searchParams.set('response_mode', 'query'); auth.searchParams.set('scope', SCOPES.join(' ')); auth.searchParams.set('state', oauthState);
      res.writeHead(302, { location: auth.toString() }); return res.end();
    }
    if (req.method === 'GET' && url.pathname === '/auth/callback') {
      if (!oauthState || url.searchParams.get('state') !== oauthState) return send(res, 400, 'OAuth state 校验失败，请重新连接。');
      oauthState = null;
      if (url.searchParams.get('error')) return send(res, 400, `Microsoft 授权失败：${url.searchParams.get('error_description') || url.searchParams.get('error')}`);
      const data = await exchange({ grant_type: 'authorization_code', code: url.searchParams.get('code'), redirect_uri: redirectUri(), scope: SCOPES.join(' ') });
      saveTokens({ ...data, expires_at: Date.now() + (data.expires_in || 3600) * 1000 });
      return send(res, 200, '<h1>家机邮箱已连接</h1><p>可以回到家机继续使用了。</p>');
    }
    if (url.pathname === '/mcp') return handleMcp(req, res);
    if (req.method === 'GET' && url.pathname === '/') return send(res, 200, '<h1>家机 Mail MCP</h1><p><a href="/auth/login">连接 Outlook 邮箱</a></p>');
    send(res, 404, 'Not Found');
  } catch (error) { console.error(error); send(res, 500, error.message); }
});

app.listen(PORT, '0.0.0.0', () => console.log(`家机 Mail MCP listening on ${PORT}`));
