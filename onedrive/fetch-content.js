/**
 * Fetch a OneDrive item's bytes into the local staging store.
 *
 * This is the step that was missing: `onedrive-download` resolves a
 * pre-authenticated URL, but nothing ever dereferenced it. Here the server
 * dereferences it, verifies the bytes, and mints a handle the model can consume
 * through text extraction, resources/read, or an inline EmbeddedResource.
 */
const pathModule = require('path');

const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { fetchBytes, looksLikeHtml } = require('../utils/binary-fetch');
const { getFileStore, safeFilename } = require('../utils/file-store');

const MIME_BY_EXTENSION = {
  '.pdf': 'application/pdf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.md': 'text/markdown',
  '.markdown': 'text/markdown',
  '.txt': 'text/plain',
  '.csv': 'text/csv',
  '.json': 'application/json',
  '.xml': 'application/xml',
  '.html': 'text/html',
  '.htm': 'text/html',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.m4a': 'audio/mp4',
  '.mp3': 'audio/mpeg',
  '.mp4': 'video/mp4',
  '.wav': 'audio/wav',
  '.webm': 'audio/webm'
};

/**
 * @param {string} filename
 * @returns {string}
 */
function mimeTypeFor(filename) {
  const extension = pathModule.extname(filename || '').toLowerCase();
  return MIME_BY_EXTENSION[extension] || 'application/octet-stream';
}

/**
 * Resolve a drive item's metadata by id or by path.
 * @param {string} accessToken
 * @param {{itemId?: string, path?: string}} args
 * @returns {Promise<object>}
 */
async function resolveItem(accessToken, args) {
  const endpoint = args.itemId
    ? `me/drive/items/${args.itemId}`
    : `me/drive/root:/${String(args.path).replace(/^\/+|\/+$/g, '')}`;

  const response = await callGraphAPI(accessToken, 'GET', endpoint, null, {
    $select: 'id,name,size,file,folder,@microsoft.graph.downloadUrl'
  });

  if (!response || !response.id) {
    const error = new Error(`No se encontro el archivo (${args.itemId || args.path})`);
    error.code = 'NOT_FOUND';
    throw error;
  }
  if (response.folder) {
    const error = new Error(`"${response.name}" es una carpeta y no se puede descargar`);
    error.code = 'IS_FOLDER';
    throw error;
  }
  return response;
}

/**
 * Download an item's bytes and stage them.
 * @param {{itemId?: string, path?: string, maxBytes?: number}} args
 * @returns {Promise<{stored: import('../utils/file-store').StoredFile, item: object}>}
 */
async function stageItem(args) {
  const accessToken = await ensureAuthenticated();
  const item = await resolveItem(accessToken, args);

  const filename = safeFilename(item.name || 'download.bin');
  const declaredMime = item.file?.mimeType || mimeTypeFor(filename);
  const maxBytes = Math.min(
    Number(args.maxBytes) || config.FILE_MAX_BYTES,
    config.FILE_MAX_BYTES
  );

  const downloadUrl = item['@microsoft.graph.downloadUrl'];

  // The pre-authenticated URL already carries its own credential, so sending
  // the Graph bearer token to it would leak the token to a CDN for no benefit.
  // The /content fallback goes to graph.microsoft.com and does need the token.
  const target = downloadUrl
    ? { url: downloadUrl, headers: {} }
    : {
      url: `${config.GRAPH_API_ENDPOINT}me/drive/items/${item.id}/content`,
      headers: { Authorization: `Bearer ${accessToken}` }
    };

  const fetched = await fetchBytes(target.url, {
    allowedHosts: config.FILE_DOWNLOAD_ALLOWED_HOSTS,
    maxBytes,
    headers: target.headers,
    timeoutMs: config.FILE_DOWNLOAD_TIMEOUT_MS
  });

  // A 200 carrying a login page is the classic silent failure here. Only reject
  // when the item is not itself HTML, otherwise a real .html file cannot load.
  const itemIsHtml = /\.(html?|xhtml)$/i.test(filename);
  if (!itemIsHtml && looksLikeHtml(fetched.buffer, fetched.contentType)) {
    throw new Error(
      'La descarga devolvio una pagina HTML en vez del archivo; probablemente la sesion expiro'
    );
  }

  const store = getFileStore();
  const stored = store.put(fetched.buffer, {
    filename,
    mimeType: declaredMime,
    // Graph reports the item size; a mismatch means a truncated transfer.
    expectedSize: typeof item.size === 'number' ? item.size : undefined
  });

  return { stored, item };
}

module.exports = { stageItem, resolveItem, mimeTypeFor };
