const express = require('express');
const crypto = require('crypto');
const { StreamableHTTPServerTransport } = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const config = require('./config');
const { createMcpServer } = require('./index');
const { McpOAuth, SUPPORTED_SCOPES } = require('./auth/mcp-oauth');
const MicrosoftOAuth = require('./auth/microsoft-oauth');
const { createAdminController } = require('./admin');
const { ensureAuthenticated } = require('./auth');
const { callGraphAPIDownload } = require('./utils/graph-api');
const downloadCapabilities = require('./onedrive/capability');
const { importUrlToPath } = require('./onedrive/import-url');

const oauth = new McpOAuth(config);
const microsoftOAuth = new MicrosoftOAuth(config);
const basePath = new URL(config.PUBLIC_BASE_URL).pathname.replace(/\/$/, '') || '';
const admin = createAdminController(config, { prefix: `${basePath}/internal-admin/v1` });

function pathFor(path) {
  return `${basePath}${path}` || '/';
}

function jsonError(res, status, message) {
  return res.status(status).json({ error: 'invalid_request', error_description: message });
}

function redirectWithCode(result) {
  const location = new URL(result.redirectUri);
  location.searchParams.set('code', result.code);
  if (result.state) location.searchParams.set('state', result.state);
  return location.toString();
}

function consentPage(ticket, error = '', clientName = 'ChatGPT') {
  const escapeHtml = value => String(value).replace(/[&<>'"]/g, character => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    "'": '&#39;',
    '"': '&quot;'
  }[character]));
  const escapedTicket = escapeHtml(ticket);
  const escapedClient = escapeHtml(clientName || 'MCP client');
  const escapedError = error ? `<p class="error">${escapeHtml(error)}</p>` : '';
  const action = escapeHtml(pathFor('/oauth/consent'));
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>Authorize Outlook</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>body{font:16px system-ui,sans-serif;max-width:30rem;margin:4rem auto;padding:0 1rem;color:#202124}input{box-sizing:border-box;width:100%;padding:.7rem;margin:.4rem 0 1rem}button{padding:.7rem 1rem;background:#2457d6;color:#fff;border:0;border-radius:4px}.error{color:#b3261e}</style></head>
<body><h1>Authorize Outlook</h1>${escapedError}
<p><strong>${escapedClient}</strong> requests access to Outlook and OneDrive.</p>
<p>Enter the MCP access token stored in your configured secret manager.</p>
<form method="post" action="${action}">
<input type="hidden" name="ticket" value="${escapedTicket}">
<label for="mcp_token">MCP access token</label>
<input id="mcp_token" name="mcp_token" type="password" autocomplete="current-password" required>
<button type="submit">Authorize</button></form></body></html>`;
}

function requireBearer(req, res, next) {
  const header = req.get('authorization') || '';
  const match = /^Bearer\s+(.+)$/i.exec(header);
  const token = match && match[1];
  const payload = token && oauth.verifyAccessToken(token);
  if (!payload) {
    res.set('WWW-Authenticate', `Bearer resource_metadata="${config.PUBLIC_BASE_URL}/.well-known/oauth-protected-resource/mcp"`);
    return res.status(401).json({ error: 'unauthorized', error_description: 'Valid MCP OAuth bearer token required' });
  }
  req.mcpAuth = payload;
  next();
}

function requireTranscriptionDelivery(req, res, next) {
  const expected = config.TRANSCRIPTION_DELIVERY_TOKEN;
  const supplied = req.get('authorization') || '';
  const match = /^Bearer\s+(.+)$/i.exec(supplied);
  const token = match && match[1];
  if (!expected || !token || token.length !== expected.length || !crypto.timingSafeEqual(Buffer.from(token), Buffer.from(expected))) {
    return res.status(401).json({ error: 'unauthorized' });
  }
  return next();
}

function safeDeliveryName(value, fallback) {
  const candidate = String(value || '').split(/[\\/]/).pop().trim();
  const safe = candidate.replace(/[^A-Za-z0-9._()\- áéíóúÁÉÍÓÚñÑ]/g, '_').slice(0, 180);
  return safe || fallback;
}

function deliveryPath(folder, name) {
  const rawFolder = String(folder || '').trim();
  if (rawFolder.includes('\\') || rawFolder.split('/').some(part => part === '..')) {
    throw new Error('Invalid OneDrive destination path.');
  }
  const normalizedFolder = rawFolder.replace(/^\/+|\/+$/g, '');
  const safeName = safeDeliveryName(name, 'transcript.bin');
  return normalizedFolder ? `/${normalizedFolder}/${safeName}` : `/${safeName}`;
}

function createApp() {
  const app = express();
  app.disable('x-powered-by');
  app.use(express.urlencoded({ extended: false, limit: '32kb' }));
  app.use(express.json({ limit: '16mb' }));

  for (const route of admin.routes) {
    app[route.method](route.path, async (req, res) => {
      try {
        await admin.ready;
        await route.handler(req, res);
      } catch (error) {
        if (!res.headersSent) res.status(400).json({ error: error.message.slice(0, 240) });
      }
    });
  }

  app.get(pathFor('/health'), (_req, res) => res.json({ status: 'ok', service: config.SERVER_NAME, version: config.SERVER_VERSION }));

  app.get(pathFor('/files/:token'), async (req, res) => {
    const target = downloadCapabilities.resolve(String(req.params.token || ''));
    if (!target) return res.status(404).json({ error: 'file_not_found' });
    try {
      const accessToken = await ensureAuthenticated();
      const downloadUrl = await callGraphAPIDownload(
        accessToken,
        `me/drive/items/${encodeURIComponent(target.itemId)}/content`
      );
      res.set({
        'Cache-Control': 'private, no-store, max-age=0',
        'X-Content-Type-Options': 'nosniff',
        'Referrer-Policy': 'no-referrer',
        'Content-Disposition': `attachment; filename="${String(target.name).replace(/["\\\r\n]/g, '_')}"`
      });
      return res.redirect(302, downloadUrl);
    } catch (error) {
      if (error.message === 'Authentication required') {
        return res.status(401).json({ error: 'authentication_required' });
      }
      return res.status(502).json({ error: 'file_proxy_failed' });
    }
  });

  app.post(pathFor('/transcription/deliver'), requireTranscriptionDelivery, async (req, res) => {
    const body = req.body || {};
    const outputs = Array.isArray(body.outputs) ? body.outputs : [];
    if (!body.job_id || !body.destination_path || !outputs.length || outputs.length > config.TRANSCRIPTION_DELIVERY_MAX_OUTPUTS) {
      return res.status(400).json({ error: 'invalid_delivery_payload' });
    }
    try {
      const savedFiles = [];
      for (const output of outputs) {
        if (!output || typeof output.download_url !== 'string') {
          throw new Error('Invalid transcription output.');
        }
        const filename = safeDeliveryName(
          output.filename,
          `${String(body.base_name || 'transcript')}.${String(output.kind || 'result')}`
        );
        const saved = await importUrlToPath(
          output.download_url,
          deliveryPath(body.destination_path, filename),
          'replace'
        );
        savedFiles.push({
          job_id: String(body.job_id),
          kind: String(output.kind || 'result'),
          ...saved
        });
      }
      return res.json({ job_id: String(body.job_id), state: 'delivered', saved_files: savedFiles });
    } catch (error) {
      return res.status(502).json({ error: 'delivery_failed', message: error.message.slice(0, 240) });
    }
  });

  app.get(pathFor('/.well-known/oauth-authorization-server'), (_req, res) => {
    res.json({
      issuer: config.PUBLIC_BASE_URL,
      authorization_endpoint: `${config.PUBLIC_BASE_URL}/authorize`,
      token_endpoint: `${config.PUBLIC_BASE_URL}/token`,
      registration_endpoint: `${config.PUBLIC_BASE_URL}/register`,
      response_types_supported: ['code'],
      code_challenge_methods_supported: ['S256'],
      token_endpoint_auth_methods_supported: ['none'],
      grant_types_supported: ['authorization_code', 'refresh_token'],
      scopes_supported: SUPPORTED_SCOPES
    });
  });

  app.get(pathFor('/.well-known/oauth-protected-resource/mcp'), (_req, res) => {
    res.set('Cache-Control', 'public, max-age=3600').json({
      resource: config.MCP_RESOURCE_URL,
      authorization_servers: [config.PUBLIC_BASE_URL],
      scopes_supported: SUPPORTED_SCOPES,
      bearer_methods_supported: ['header']
    });
  });

  app.post(pathFor('/register'), (req, res) => {
    try {
      return res.status(201).json(oauth.register(req.body || {}));
    } catch (error) {
      return jsonError(res, 400, error.message);
    }
  });

  app.get(pathFor('/authorize'), (req, res) => {
    try {
      const ticket = oauth.beginAuthorization(req.query);
      return res.redirect(302, `${config.PUBLIC_BASE_URL}/oauth/consent?ticket=${encodeURIComponent(ticket)}`);
    } catch (error) {
      return jsonError(res, 400, error.message);
    }
  });

  app.get(pathFor('/oauth/consent'), (req, res) => {
    const ticket = String(req.query.ticket || '');
    const pending = oauth.getPending(ticket);
    if (!pending) return res.status(400).send('Authorization request expired');
    return res
      .status(req.query.error ? 401 : 200)
      .set('Cache-Control', 'no-store')
      .set('Content-Security-Policy', "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'")
      .type('html')
      .send(consentPage(ticket, req.query.error ? String(req.query.error) : '', pending.client.clientName));
  });

  app.post(pathFor('/oauth/consent'), async (req, res) => {
    const ticket = String(req.body.ticket || '');
    if (!oauth.getPending(ticket)) return res.status(400).send('Authorization request expired');
    try {
      if (!await microsoftOAuth.hasTokens()) {
        return res.redirect(302, microsoftOAuth.authorizationUrl(ticket));
      }
      const result = oauth.approve(ticket, req.body.mcp_token);
      if (!result) return res.status(401).type('html').send(consentPage(ticket, 'Invalid token or expired authorization request.'));
      return res.redirect(302, redirectWithCode(result));
    } catch (error) {
      return res.status(500).type('html').send(consentPage(ticket, error.message, oauth.getPending(ticket)?.client.clientName));
    }
  });

  app.get(pathFor('/microsoft/callback'), async (req, res) => {
    try {
      if (req.query.error) throw new Error('Microsoft authorization was declined');
      const ticket = await microsoftOAuth.exchange(String(req.query.code || ''), String(req.query.state || ''));
      return res.type('html').send(MicrosoftOAuth.successPage(ticket, config.PUBLIC_BASE_URL));
    } catch (error) {
      return res.status(400).send(error.message);
    }
  });

  app.post(pathFor('/token'), (req, res) => {
    try {
      const grantType = req.body.grant_type;
      const tokens = grantType === 'authorization_code'
        ? oauth.exchangeCode({
          code: req.body.code,
          clientId: req.body.client_id,
          redirectUri: req.body.redirect_uri,
          codeVerifier: req.body.code_verifier
        })
        : grantType === 'refresh_token'
          ? oauth.refresh({ refreshToken: req.body.refresh_token, clientId: req.body.client_id, scope: req.body.scope })
          : null;
      if (!tokens) return jsonError(res, 400, 'Unsupported grant_type');
      return res.set('Cache-Control', 'no-store').json(tokens);
    } catch (error) {
      return jsonError(res, 400, error.message);
    }
  });

  app.post(pathFor('/mcp'), requireBearer, async (req, res) => {
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    const server = createMcpServer({ remote: true, scopes: req.mcpAuth.scopes });
    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      if (!res.headersSent) res.status(500).json({ error: 'internal_error', error_description: 'MCP request failed' });
    } finally {
      await server.close().catch(() => {});
      await transport.close().catch(() => {});
    }
  });

  app.all(pathFor('/mcp'), (_req, res) => res.status(405).set('Allow', 'POST').end());
  return app;
}

function startHttp() {
  const app = createApp();
  const server = app.listen(config.HTTP_PORT, config.HTTP_HOST, () => {
    console.error(`${config.SERVER_NAME} HTTP server listening on ${config.HTTP_HOST}:${config.HTTP_PORT}`);
  });
  const shutdown = () => server.close(() => process.exit(0));
  process.once('SIGINT', shutdown);
  process.once('SIGTERM', shutdown);
  return server;
}

if (require.main === module) startHttp();

admin.ready = admin.initialize().catch(error => {
  console.error(`Unable to initialize admin store: ${error.message}`);
  process.exitCode = 1;
  throw error;
});

module.exports = { createApp, startHttp, oauth, microsoftOAuth, admin, pathFor };
