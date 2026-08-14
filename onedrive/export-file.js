/**
 * Transfer a OneDrive file's raw bytes inline as an MCP EmbeddedResource.
 *
 * Fallback channel, for clients that cannot issue a follow-up resources/read
 * and for content text extraction cannot represent (images, audio, archives).
 *
 * Base64 inflates payloads by roughly a third and the result lands directly in
 * the model context, so this path is capped far below the staging limit. Text
 * extraction via onedrive-read-file is the cheaper option whenever it applies.
 */
const config = require('../config');
const { stageItem } = require('./fetch-content');
const { FileStore, getFileStore } = require('../utils/file-store');

/**
 * @param {object} args
 * @returns {Promise<object>} MCP tool result
 */
async function handleExportFile(args = {}) {
  const { itemId, path, fileId } = args;

  if (!itemId && !path && !fileId) {
    return {
      isError: true,
      content: [{ type: 'text', text: 'Se requiere itemId, path o fileId.' }]
    };
  }

  try {
    const store = getFileStore();
    let stored;

    if (fileId) {
      stored = store.get(fileId);
      if (!stored) {
        return {
          isError: true,
          content: [{ type: 'text', text: 'El archivo no existe o ya expiro.' }]
        };
      }
    } else {
      ({ stored } = await stageItem({ itemId, path, maxBytes: config.FILE_INLINE_MAX_BYTES }));
    }

    if (stored.size_bytes > config.FILE_INLINE_MAX_BYTES) {
      const metadata = FileStore.metadata(stored);
      return {
        isError: true,
        content: [{
          type: 'text',
          text: `El archivo pesa ${stored.size_bytes} bytes y supera el limite de transferencia inline `
            + `(${config.FILE_INLINE_MAX_BYTES} bytes). Lee el recurso ${metadata.resource_uri} `
            + 'con resources/read, o usa onedrive-read-file para obtener solo el texto.'
        }],
        structuredContent: { status: 'too_large', file: metadata }
      };
    }

    const buffer = store.read(stored.file_id);
    const metadata = FileStore.metadata(stored);
    const isText = /^text\/|application\/(json|xml|javascript)|\+xml$/.test(stored.mime_type);

    // MCP resource contents are either `text` or `blob`; sending text as base64
    // forces a needless decode on the client, so branch on the MIME type.
    const resource = isText
      ? { uri: metadata.resource_uri, mimeType: stored.mime_type, text: buffer.toString('utf8') }
      : { uri: metadata.resource_uri, mimeType: stored.mime_type, blob: buffer.toString('base64') };

    return {
      content: [
        {
          type: 'text',
          text: `Transferido "${stored.filename}" (${stored.size_bytes} bytes, ${stored.mime_type}, `
            + `sha256 ${stored.sha256.slice(0, 16)}...) como ${isText ? 'texto' : 'blob Base64'} embebido.`
        },
        { type: 'resource', resource }
      ],
      structuredContent: {
        status: 'ready',
        transfer: 'embedded_resource',
        encoding: isText ? 'text' : 'base64',
        file: metadata
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required' || error.message === 'UNAUTHORIZED') {
      return {
        isError: true,
        content: [{
          type: 'text',
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    return {
      isError: true,
      content: [{ type: 'text', text: `Error exportando el archivo: ${error.message}` }]
    };
  }
}

module.exports = handleExportFile;
