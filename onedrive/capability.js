const crypto = require('crypto');

const capabilities = new Map();
const DEFAULT_TTL_SECONDS = 30 * 60;

function issue(item, baseUrl, ttlSeconds = DEFAULT_TTL_SECONDS) {
  const token = crypto.randomBytes(32).toString('base64url');
  const expiresAt = Date.now() + Math.max(1, Number(ttlSeconds) || DEFAULT_TTL_SECONDS) * 1000;
  capabilities.set(token, {
    itemId: item.id,
    name: item.name,
    size: item.size || 0,
    mimeType: item.file?.mimeType || 'application/octet-stream',
    expiresAt
  });
  return {
    token,
    url: `${String(baseUrl).replace(/\/+$/, '')}/files/${encodeURIComponent(token)}`,
    expiresAt: new Date(expiresAt).toISOString()
  };
}

function resolve(token) {
  const value = capabilities.get(token);
  if (!value) return null;
  if (value.expiresAt <= Date.now()) {
    capabilities.delete(token);
    return null;
  }
  return { ...value };
}

function revoke(token) {
  capabilities.delete(token);
}

function clearExpired() {
  const now = Date.now();
  for (const [token, value] of capabilities) {
    if (value.expiresAt <= now) capabilities.delete(token);
  }
}

module.exports = { issue, resolve, revoke, clearExpired };
