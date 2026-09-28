const crypto = require('crypto');

const capabilities = new Map();
const DEFAULT_TTL_SECONDS = 30 * 60;
let sweepTimer = null;

/**
 * Issue a proxy capability token for a Graph byte-serving endpoint.
 * @param {object} item - driveItem-like object (needs id, name, size, file.mimeType).
 * @param {string} baseUrl - Public base URL to build the proxy link from.
 * @param {number} [ttlSeconds]
 * @param {string} [graphSuffix] - Path suffix appended after `me/drive/items/{id}`,
 *   e.g. '/content' (default), '/thumbnails/0/large/content', or
 *   '/content?format=pdf'. Lets one proxy mechanism serve downloads,
 *   thumbnails and format conversions alike, since all three follow the
 *   same Graph 302-redirect-to-a-preauthenticated-CDN-URL contract.
 * @returns {{token: string, url: string, expiresAt: string}}
 */
function issue(item, baseUrl, ttlSeconds = DEFAULT_TTL_SECONDS, graphSuffix = '/content') {
  const token = crypto.randomBytes(32).toString('base64url');
  const expiresAt = Date.now() + Math.max(1, Number(ttlSeconds) || DEFAULT_TTL_SECONDS) * 1000;
  capabilities.set(token, {
    itemId: item.id,
    name: item.name,
    size: item.size || 0,
    mimeType: item.file?.mimeType || 'application/octet-stream',
    graphSuffix,
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

/**
 * Schedule periodic cleanup of expired, never-fetched tokens.
 * Without this, a token that is issued but never used leaks for the life of
 * the process — `resolve()` only prunes lazily, on access. Safe to call more
 * than once; only the first call installs a timer.
 * @param {number} [intervalMs]
 */
function startSweeper(intervalMs = 5 * 60 * 1000) {
  if (sweepTimer) return sweepTimer;
  sweepTimer = setInterval(clearExpired, intervalMs);
  if (typeof sweepTimer.unref === 'function') sweepTimer.unref();
  return sweepTimer;
}

module.exports = { issue, resolve, revoke, clearExpired, startSweeper };
