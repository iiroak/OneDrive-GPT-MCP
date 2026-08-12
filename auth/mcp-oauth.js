const crypto = require('crypto');

const SUPPORTED_SCOPES = ['outlook:read', 'outlook:write', 'outlook:destructive'];

function b64url(value) {
  return Buffer.from(value).toString('base64url');
}

function parseB64url(value) {
  return Buffer.from(value, 'base64url').toString('utf8');
}

function now() {
  return Math.floor(Date.now() / 1000);
}

function safeEqual(left, right) {
  const a = Buffer.from(left || '');
  const b = Buffer.from(right || '');
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

class McpOAuth {
  constructor(config) {
    this.issuer = config.PUBLIC_BASE_URL;
    this.resource = config.MCP_RESOURCE_URL;
    this.bootstrapToken = config.MCP_BOOTSTRAP_TOKEN;
    this.secret = config.MCP_OAUTH_SECRET || this.bootstrapToken;
    this.accessTtl = config.MCP_ACCESS_TOKEN_TTL;
    this.refreshTtl = config.MCP_REFRESH_TOKEN_TTL;
    this.authorizationTtl = config.MCP_AUTHORIZATION_TTL;
    this.pending = new Map();
    this.usedCodes = new Set();
    this.revoked = new Set();
  }

  sign(payload) {
    if (!this.secret) throw new Error('OUTLOOK_MCP_TOKEN or OUTLOOK_OAUTH_SECRET is required');
    const body = b64url(JSON.stringify(payload));
    const signature = crypto.createHmac('sha256', this.secret).update(body).digest('base64url');
    return `${body}.${signature}`;
  }

  decode(value) {
    if (!this.secret || typeof value !== 'string') return null;
    const separator = value.lastIndexOf('.');
    if (separator < 1) return null;
    const body = value.slice(0, separator);
    const supplied = value.slice(separator + 1);
    const expected = crypto.createHmac('sha256', this.secret).update(body).digest('base64url');
    if (!safeEqual(supplied, expected)) return null;
    try {
      return JSON.parse(parseB64url(body));
    } catch (_) {
      return null;
    }
  }

  register(metadata) {
    const redirectUris = Array.isArray(metadata.redirect_uris) ? metadata.redirect_uris : [];
    if (!redirectUris.length || !redirectUris.every(uri => this.validRedirectUri(uri))) {
      throw new Error('redirect_uris must contain HTTPS or loopback HTTP URLs');
    }
    if (metadata.token_endpoint_auth_method && metadata.token_endpoint_auth_method !== 'none') {
      throw new Error('Only public OAuth clients are supported');
    }
    if (!Array.isArray(metadata.response_types) || !metadata.response_types.includes('code')) {
      throw new Error('PKCE authorization code is required');
    }
    const requestedScopes = String(metadata.scope || SUPPORTED_SCOPES.join(' ')).split(/\s+/).filter(Boolean);
    if (requestedScopes.some(scope => !SUPPORTED_SCOPES.includes(scope))) {
      throw new Error(`Supported scopes are ${SUPPORTED_SCOPES.join(' ')}`);
    }
    const clientId = this.sign({
      kind: 'client',
      id: crypto.randomBytes(18).toString('base64url'),
      redirectUris,
      clientName: metadata.client_name || 'MCP client',
      scopes: SUPPORTED_SCOPES,
      grantTypes: ['authorization_code', 'refresh_token'],
      exp: now() + 365 * 24 * 3600
    });
    return {
      client_id: clientId,
      client_name: metadata.client_name || 'MCP client',
      redirect_uris: redirectUris,
      response_types: ['code'],
      grant_types: ['authorization_code', 'refresh_token'],
      token_endpoint_auth_method: 'none',
      scope: SUPPORTED_SCOPES.join(' ')
    };
  }

  getClient(clientId) {
    const client = this.decode(clientId);
    if (!client || client.kind !== 'client' || client.exp < now()) return null;
    return client;
  }

  validRedirectUri(value) {
    try {
      const uri = new URL(value);
      if (uri.username || uri.password || uri.hash) return false;
      return uri.protocol === 'https:' ||
        (uri.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(uri.hostname));
    } catch (_) {
      return false;
    }
  }

  beginAuthorization(params) {
    const client = this.getClient(params.client_id);
    if (!client || !client.redirectUris.includes(params.redirect_uri)) throw new Error('Invalid client or redirect URI');
    if (params.resource && params.resource !== this.resource) throw new Error('Invalid resource');
    if (!params.code_challenge || params.code_challenge_method !== 'S256') throw new Error('PKCE S256 is required');
    const scopes = String(params.scope || SUPPORTED_SCOPES.join(' ')).split(/\s+/).filter(Boolean);
    if (scopes.some(scope => !SUPPORTED_SCOPES.includes(scope))) throw new Error('Invalid scope');
    const ticket = crypto.randomBytes(24).toString('base64url');
    this.pending.set(ticket, {
      client,
      params: { ...params, scopes },
      expiresAt: Date.now() + this.authorizationTtl * 1000
    });
    return ticket;
  }

  getPending(ticket) {
    const pending = this.pending.get(ticket);
    if (!pending || pending.expiresAt < Date.now()) {
      this.pending.delete(ticket);
      return null;
    }
    return pending;
  }

  approve(ticket, suppliedBootstrapToken) {
    const pending = this.getPending(ticket);
    if (!pending || !this.bootstrapToken || !safeEqual(suppliedBootstrapToken, this.bootstrapToken)) return null;
    this.pending.delete(ticket);
    const payload = {
      kind: 'code',
      jti: crypto.randomBytes(18).toString('base64url'),
      clientId: pending.client.id,
      redirectUri: pending.params.redirect_uri,
      challenge: pending.params.code_challenge,
      state: pending.params.state,
      resource: this.resource,
      scopes: pending.params.scopes,
      exp: now() + this.authorizationTtl
    };
    return {
      code: this.sign(payload),
      redirectUri: pending.params.redirect_uri,
      state: pending.params.state
    };
  }

  exchangeCode({ code, clientId, redirectUri, codeVerifier }) {
    const payload = this.decode(code);
    const client = this.getClient(clientId);
    if (!payload || payload.kind !== 'code' || payload.exp < now() || this.usedCodes.has(payload.jti) || !client) throw new Error('Invalid authorization code');
    if (payload.clientId !== client.id || payload.redirectUri !== redirectUri || !client.redirectUris.includes(redirectUri)) throw new Error('Authorization code binding failed');
    const challenge = crypto.createHash('sha256').update(codeVerifier || '').digest('base64url');
    if (!safeEqual(challenge, payload.challenge)) throw new Error('PKCE verification failed');
    this.usedCodes.add(payload.jti);
    return this.issueTokens(client.id, payload.scopes);
  }

  refresh({ refreshToken, clientId, scope }) {
    if (this.revoked.has(refreshToken)) throw new Error('Refresh token revoked');
    const payload = this.decode(refreshToken);
    const client = this.getClient(clientId);
    if (!payload || payload.kind !== 'refresh' || payload.exp < now() || !client || payload.clientId !== client.id) throw new Error('Invalid refresh token');
    const scopes = String(scope || payload.scopes.join(' ')).split(/\s+/).filter(Boolean);
    if (scopes.some(item => !payload.scopes.includes(item))) throw new Error('Refresh scope escalation denied');
    this.revoked.add(refreshToken);
    return this.issueTokens(client.id, scopes);
  }

  issueTokens(clientId, scopes) {
    const timestamp = now();
    const accessToken = this.sign({
      kind: 'access',
      jti: crypto.randomBytes(18).toString('base64url'),
      clientId,
      scopes,
      resource: this.resource,
      exp: timestamp + this.accessTtl
    });
    const refreshToken = this.sign({
      kind: 'refresh',
      jti: crypto.randomBytes(18).toString('base64url'),
      clientId,
      scopes,
      exp: timestamp + this.refreshTtl
    });
    return { access_token: accessToken, token_type: 'Bearer', expires_in: this.accessTtl, refresh_token: refreshToken, scope: scopes.join(' ') };
  }

  verifyAccessToken(token) {
    const payload = this.decode(token);
    if (!payload || payload.kind !== 'access' || payload.exp < now() || payload.resource !== this.resource || this.revoked.has(token)) return null;
    return payload;
  }
}

module.exports = { McpOAuth, SUPPORTED_SCOPES };
