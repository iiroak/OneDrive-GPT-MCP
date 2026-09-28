/**
 * Shared helpers for resolving a OneDrive driveItem from either an itemId
 * or a path, and for building the two Graph URL address forms used
 * throughout this module (`items/{id}` vs `root:/{path}`).
 *
 * Extracted so every new tool addresses items the same way instead of each
 * one re-implementing path normalization and existence checks slightly
 * differently.
 */
const { callGraphAPI } = require('./graph-api');

/**
 * Strip leading/trailing slashes so we can safely embed a path in the
 * `root:/{path}` Graph address form.
 * @param {string} path
 * @returns {string}
 */
function normalizePath(path) {
  return String(path || '').replace(/^\/+|\/+$/g, '');
}

/**
 * Encode a user-facing OneDrive path one segment at a time. Reject dot
 * segments and backslashes so URL parsing cannot reinterpret the path.
 * @param {string} path
 * @returns {string}
 */
function encodePath(path) {
  const normalized = normalizePath(path);
  if (!normalized) return '';
  if (normalized.includes('\\')) throw new Error('Invalid OneDrive path.');
  const segments = normalized.split('/').filter(Boolean);
  if (segments.some((segment) => segment === '.' || segment === '..')) {
    throw new Error('Invalid OneDrive path.');
  }
  return segments.map((segment) => encodeURIComponent(segment)).join('/');
}

function encodeId(id) {
  return encodeURIComponent(String(id));
}

/**
 * Build the Graph endpoint (relative, no leading slash) that addresses a
 * single driveItem by itemId or by path.
 * @param {{itemId?: string, path?: string}} ref
 * @returns {string}
 */
function itemEndpoint(ref) {
  if (ref.itemId) return `me/drive/items/${encodeId(ref.itemId)}`;
  const normalized = encodePath(ref.path);
  return normalized && normalized.toLowerCase() !== 'root'
    ? `me/drive/root:/${normalized}`
    : 'me/drive/root';
}

/**
 * Resolve a driveItem by itemId or path, throwing a descriptive error if
 * neither is provided or the item does not exist.
 * @param {string} accessToken
 * @param {{itemId?: string, path?: string, select?: string}} ref
 * @returns {Promise<object>} the driveItem
 */
async function resolveItem(accessToken, ref) {
  if (!ref.itemId && !ref.path) {
    throw new Error('Either itemId or path is required.');
  }
  const endpoint = itemEndpoint(ref);
  const queryParams = ref.select ? { $select: ref.select } : {};
  const item = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);
  if (!item || !item.id) {
    throw new Error(ref.itemId ? 'Item not found.' : `Item not found at path: ${ref.path}`);
  }
  return item;
}

module.exports = { normalizePath, encodePath, encodeId, itemEndpoint, resolveItem };
