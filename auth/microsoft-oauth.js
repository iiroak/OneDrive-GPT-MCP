const fs = require('fs').promises;
const https = require('https');
const crypto = require('crypto');

function escape(value) {
  return String(value).replace(/[&<>'"]/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[character]));
}

class MicrosoftOAuth {
  constructor(config) {
    this.config = config;
    this.pending = new Map();
  }

  async hasTokens() {
    try {
      const tokens = JSON.parse(await fs.readFile(this.config.AUTH_CONFIG.tokenStorePath, 'utf8'));
      return Boolean(tokens.refresh_token || tokens.access_token);
    } catch (_) {
      return false;
    }
  }

  authorizationUrl(ticket) {
    if (!this.config.AUTH_CONFIG.clientId) throw new Error('MS_CLIENT_ID is not configured');
    const state = crypto.randomBytes(32).toString('hex');
    this.pending.set(state, { ticket, expiresAt: Date.now() + 10 * 60 * 1000 });
    const query = new URLSearchParams({
      client_id: this.config.AUTH_CONFIG.clientId,
      response_type: 'code',
      redirect_uri: this.config.AUTH_CONFIG.redirectUri,
      response_mode: 'query',
      scope: this.config.AUTH_CONFIG.scopes.join(' '),
      state
    });
    return `${this.config.AUTHORITY_URL}/${this.config.MS_TENANT_ID}/oauth2/v2.0/authorize?${query}`;
  }

  async exchange(code, state) {
    const pending = this.pending.get(state);
    this.pending.delete(state);
    if (!pending || pending.expiresAt < Date.now()) throw new Error('Invalid or expired Microsoft OAuth state');
    const body = new URLSearchParams({
      client_id: this.config.AUTH_CONFIG.clientId,
      client_secret: this.config.AUTH_CONFIG.clientSecret || '',
      code,
      redirect_uri: this.config.AUTH_CONFIG.redirectUri,
      grant_type: 'authorization_code',
      scope: this.config.AUTH_CONFIG.scopes.join(' ')
    }).toString();
    const tokens = await this.postToken(body);
    tokens.expires_at = Date.now() + Number(tokens.expires_in || 3600) * 1000;
    await fs.mkdir(this.config.DATA_DIR, { recursive: true, mode: 0o700 });
    await fs.writeFile(this.config.AUTH_CONFIG.tokenStorePath, JSON.stringify(tokens, null, 2), { mode: 0o600 });
    return pending.ticket;
  }

  postToken(body) {
    return new Promise((resolve, reject) => {
      const request = https.request(`${this.config.AUTHORITY_URL}/${this.config.MS_TENANT_ID}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) }
      }, response => {
        let data = '';
        response.setEncoding('utf8');
        response.on('data', chunk => { data += chunk; });
        response.on('end', () => {
          let parsed;
          try { parsed = JSON.parse(data); } catch (_) { parsed = null; }
          if (response.statusCode >= 200 && response.statusCode < 300 && parsed) return resolve(parsed);
          reject(new Error(`Microsoft token exchange failed (${response.statusCode})`));
        });
      });
      request.on('error', reject);
      request.write(body);
      request.end();
    });
  }

  static successPage(ticket, baseUrl) {
    return `<!doctype html><title>Microsoft connected</title><p>Microsoft account connected.</p><p><a href="${escape(`${baseUrl}/oauth/consent?ticket=${encodeURIComponent(ticket)}`)}">Return to consent</a></p>`;
  }
}

module.exports = MicrosoftOAuth;
