const crypto = require('crypto');
const os = require('os');
const path = require('path');
const request = require('supertest');

process.env.OUTLOOK_DATA_DIR = path.join(os.tmpdir(), `outlook-mcp-test-${process.pid}`);
process.env.OUTLOOK_MCP_TOKEN = 'test-bootstrap-token';
process.env.OUTLOOK_OAUTH_SECRET = 'test-oauth-secret';
process.env.MCP_ADMIN_OUTLOOK_TOKEN = 'test-admin-token';
process.env.OUTLOOK_PUBLIC_BASE_URL = 'http://127.0.0.1:8767/outlook';
process.env.USE_TEST_MODE = 'true';

const { createApp, oauth } = require('../../server');
const { createMcpServer, DESTRUCTIVE_TOOLS, requiredScope, toolPolicy } = require('../../index');

function createClient() {
  return oauth.register({
    client_name: 'ChatGPT test client',
    redirect_uris: ['https://chatgpt.com/connector/oauth/callback'],
    response_types: ['code'],
    grant_types: ['authorization_code', 'refresh_token'],
    token_endpoint_auth_method: 'none'
  });
}

function issueToken() {
  const client = createClient();
  const verifier = 'test-code-verifier-with-enough-entropy-1234567890';
  const challenge = crypto.createHash('sha256').update(verifier).digest('base64url');
  const ticket = oauth.beginAuthorization({
    client_id: client.client_id,
    redirect_uri: client.redirect_uris[0],
    code_challenge: challenge,
    code_challenge_method: 'S256',
    resource: 'http://127.0.0.1:8767/outlook/mcp',
    scope: 'outlook:read outlook:write outlook:destructive',
    state: 'test-state'
  });
  const approval = oauth.approve(ticket, 'test-bootstrap-token');
  return oauth.exchangeCode({
    code: approval.code,
    clientId: client.client_id,
    redirectUri: approval.redirectUri,
    codeVerifier: verifier
  });
}

describe('ChatGPT MCP contract', () => {
  test('exposes a reusable server with Outlook tools but no Power Automate tools', async () => {
    const server = createMcpServer({ remote: true });
    const response = await new Promise((resolve, reject) => {
      server.fallbackRequestHandler({ method: 'tools/list', params: {} }).then(resolve, reject);
    });
    const names = response.tools.map(tool => tool.name);
    expect(names).toContain('list-emails');
    expect(names).toContain('accept-event');
    expect(names).toContain('permanently-delete-email');
    expect(names).toContain('update-event');
    expect(names).toContain('list-master-categories');
    expect(names).toContain('list-calendars');
    expect(names).toContain('create-calendar');
    expect(names).toContain('copy-event');
    expect(names).toContain('migrate-events');
    expect(names).toContain('onedrive-import-url');
    expect(names).toContain('onedrive-move');
    expect(names.some(name => name.startsWith('flow-'))).toBe(false);
    expect(names).not.toContain('authenticate');
    const aboutTool = response.tools.find(tool => tool.name === 'about');
    expect(aboutTool.outputSchema).toMatchObject({
      type: 'object',
      required: ['message']
    });
    expect(response.tools.every(tool => tool.outputSchema)).toBe(true);
    expect(response.tools.find(tool => tool.name === 'onedrive-import-url')._meta.openai).toBeUndefined();
    expect(response.tools.find(tool => tool.name === 'list-emails').annotations).toMatchObject({
      readOnlyHint: true,
      destructiveHint: false
    });
    expect(response.tools.find(tool => tool.name === 'permanently-delete-email')._meta.securitySchemes[0].scopes)
      .toEqual(['outlook:destructive']);
    expect(response.tools.find(tool => tool.name === 'update-event')).toMatchObject({
      annotations: { readOnlyHint: false, destructiveHint: false },
      _meta: { securitySchemes: [{ scopes: ['outlook:write'] }] }
    });
    expect(response.tools.find(tool => tool.name === 'onedrive-move')).toMatchObject({
      annotations: { readOnlyHint: false, destructiveHint: false },
      _meta: { securitySchemes: [{ scopes: ['outlook:write'] }] }
    });
    expect(response.tools.find(tool => tool.name === 'list-master-categories')._meta.securitySchemes[0].scopes)
      .toEqual(['outlook:read']);
    expect(response.tools.find(tool => tool.name === 'accept-event')).toMatchObject({
      annotations: { readOnlyHint: false, destructiveHint: false },
      _meta: { securitySchemes: [{ scopes: ['outlook:write'] }] }
    });
    expect(response.tools.find(tool => tool.name === 'migrate-events')).toMatchObject({
      annotations: { readOnlyHint: false, destructiveHint: true },
      _meta: { securitySchemes: [{ scopes: ['outlook:destructive'] }] }
    });
  });

  test('returns structured information from about', async () => {
    const server = createMcpServer({ remote: true });
    const response = await server.fallbackRequestHandler({ method: 'tools/call', params: { name: 'about', arguments: {} } });

    expect(response.structuredContent).toMatchObject({
      message: expect.stringContaining('Outlook MCP Server'),
      data: {
        name: 'm365-assistant',
        services: ['Outlook', 'OneDrive'],
        power_automate_exposed: false
      }
    });
    expect(response.content[0].text).toContain('Outlook MCP Server');
  });

  test('serves public metadata and protects the MCP endpoint', async () => {
    const app = createApp();
    expect((await request(app).get('/outlook/health')).status).toBe(200);
    expect((await request(app).get('/outlook/.well-known/oauth-authorization-server')).body.issuer)
      .toBe('http://127.0.0.1:8767/outlook');
    const response = await request(app)
      .post('/outlook/mcp')
      .set('Accept', 'application/json, text/event-stream')
      .send({});
    expect(response.status).toBe(401);
    expect(response.headers['www-authenticate']).toContain('/outlook/.well-known/oauth-protected-resource/mcp');
  });

  test('renders the bootstrap consent page with the shared MCP form', async () => {
    const client = createClient();
    const ticket = oauth.beginAuthorization({
      client_id: client.client_id,
      redirect_uri: client.redirect_uris[0],
      code_challenge: 'test-challenge',
      code_challenge_method: 'S256',
      resource: 'http://127.0.0.1:8767/outlook/mcp',
      scope: 'outlook:read',
    });
    const response = await request(createApp()).get('/outlook/oauth/consent').query({ ticket });

    expect(response.status).toBe(200);
    expect(response.headers['cache-control']).toBe('no-store');
    expect(response.headers['content-security-policy']).toContain("form-action 'self'");
    expect(response.text).toContain('<html lang="en">');
    expect(response.text).toContain('<label for="mcp_token">MCP access token</label>');
    expect(response.text).toContain('autocomplete="current-password"');
    expect(response.text).toContain('ChatGPT test client');
    expect(response.text).not.toContain('test-bootstrap-token');
  });

  test('exposes the admin contract without returning secret values', async () => {
    const app = createApp();
    const manifestResponse = await request(app)
      .get('/outlook/internal-admin/v1/manifest')
      .set('X-MCP-Admin-Token', 'test-admin-token');
    expect(manifestResponse.status).toBe(200);
    expect(manifestResponse.body.service.id).toBe('outlook');
    expect(manifestResponse.body.config.find(section => section.id === 'security').schema.properties.mcp_token.writeOnly)
      .toBe(true);

    const configResponse = await request(app)
      .get('/outlook/internal-admin/v1/config')
      .set('X-MCP-Admin-Token', 'test-admin-token');
    expect(configResponse.status).toBe(200);
    expect(configResponse.body.sections.security.mcp_token).toEqual({ configured: true, updated_at: null });
    expect(JSON.stringify(configResponse.body)).not.toContain('test-bootstrap-token');

    const updateResponse = await request(app)
      .patch('/outlook/internal-admin/v1/config/microsoft')
      .set('X-MCP-Admin-Token', 'test-admin-token')
      .send({ ms_client_secret: 'test-secret', ms_tenant_id: 'consumers' });
    expect(updateResponse.status).toBe(200);
    expect(updateResponse.body.requires_restart).toBe(true);

    const deleteResponse = await request(app)
      .delete('/outlook/internal-admin/v1/config/microsoft/ms_tenant_id')
      .set('X-MCP-Admin-Token', 'test-admin-token');
    expect(deleteResponse.status).toBe(200);

    const deleteSecretResponse = await request(app)
      .delete('/outlook/internal-admin/v1/config/security/mcp_token')
      .set('X-MCP-Admin-Token', 'test-admin-token');
    expect(deleteSecretResponse.status).toBe(200);
    const statusResponse = await request(app)
      .get('/outlook/internal-admin/v1/status')
      .set('X-MCP-Admin-Token', 'test-admin-token');
    expect(statusResponse.body.readiness).toBe('not_ready');
  });

  test('accepts an MCP token and negotiates Streamable HTTP', async () => {
    const tokens = issueToken();
    const response = await request(createApp())
      .post('/outlook/mcp')
      .set('Authorization', `Bearer ${tokens.access_token}`)
      .set('Accept', 'application/json, text/event-stream')
      .send({
        jsonrpc: '2.0',
        id: 1,
        method: 'initialize',
        params: { protocolVersion: '2025-03-26', capabilities: {}, clientInfo: { name: 'test', version: '1' } }
      });
    expect(response.status).toBe(200);
    expect(response.text).toContain('m365-assistant');
  });

  test('enforces PKCE and rotates refresh tokens', () => {
    const client = createClient();
    const verifier = 'another-test-code-verifier-with-enough-entropy-123';
    const challenge = crypto.createHash('sha256').update(verifier).digest('base64url');
    const ticket = oauth.beginAuthorization({
      client_id: client.client_id,
      redirect_uri: client.redirect_uris[0],
      code_challenge: challenge,
      code_challenge_method: 'S256',
      resource: 'http://127.0.0.1:8767/outlook/mcp',
      scope: 'outlook:read'
    });
    const approval = oauth.approve(ticket, 'test-bootstrap-token');
    expect(() => oauth.exchangeCode({ code: approval.code, clientId: client.client_id, redirectUri: approval.redirectUri, codeVerifier: 'wrong' }))
      .toThrow('PKCE verification failed');
    const tokens = oauth.exchangeCode({ code: approval.code, clientId: client.client_id, redirectUri: approval.redirectUri, codeVerifier: verifier });
    expect(tokens.scope).toBe('outlook:read');
    const refreshed = oauth.refresh({ refreshToken: tokens.refresh_token, clientId: client.client_id });
    expect(refreshed.refresh_token).not.toBe(tokens.refresh_token);
    expect(() => oauth.refresh({ refreshToken: tokens.refresh_token, clientId: client.client_id })).toThrow('Refresh token revoked');
  });

  test('classifies destructive tools consistently', () => {
    expect(DESTRUCTIVE_TOOLS.has('permanently-delete-email')).toBe(true);
    expect(requiredScope('permanently-delete-email')).toBe('outlook:destructive');
    expect(toolPolicy({ name: 'permanently-delete-email' })).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true
    });
  });
});
