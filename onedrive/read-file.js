/**
 * Read a OneDrive file's content as text.
 *
 * The primary transfer channel: the server downloads the bytes, extracts text,
 * and returns the text. No URL for the model to dereference, no base64 to
 * decode. Also emits a ResourceLink so a client that wants the raw bytes can
 * follow up with resources/read instead of downloading again.
 */
const config = require('../config');
const { stageItem } = require('./fetch-content');
const { extractText, DEFAULT_MAX_CHARS } = require('../utils/document-text');
const { FileStore, getFileStore } = require('../utils/file-store');

/**
 * @param {object} args
 * @returns {Promise<object>} MCP tool result
 */
async function handleReadFile(args = {}) {
  const { itemId, path, fileId, maxChars } = args;

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
      // Re-read an already staged file without a second network round trip.
      stored = store.get(fileId);
      if (!stored) {
        return {
          isError: true,
          content: [{
            type: 'text',
            text: 'El archivo no existe o ya expiro. Vuelve a llamar a onedrive-read-file con itemId o path.'
          }]
        };
      }
    } else {
      ({ stored } = await stageItem({ itemId, path }));
    }

    const buffer = store.read(stored.file_id);
    const limit = Math.max(1, Number(maxChars) || DEFAULT_MAX_CHARS);
    const extraction = extractText(buffer, stored.filename, limit);
    const metadata = FileStore.metadata(stored);

    const header = [
      `Archivo: ${stored.filename}`,
      `Tipo: ${stored.mime_type} | ${stored.size_bytes} bytes | sha256 ${stored.sha256.slice(0, 16)}...`,
      `Extraccion: ${extraction.format} (${extraction.status})${extraction.truncated ? ' [truncado]' : ''}`
    ].join('\n');

    const warnings = (extraction.warnings || []).length > 0
      ? `\n\nAdvertencias:\n${extraction.warnings.map((item) => `- ${item}`).join('\n')}`
      : '';

    const body = extraction.text
      ? `\n\n--- CONTENIDO ---\n${extraction.text}`
      : `\n\nNo se extrajo texto. Usa onedrive-export-file para recibir los bytes (limite ${config.FILE_INLINE_MAX_BYTES} bytes) o lee el recurso ${metadata.resource_uri}.`;

    return {
      content: [
        { type: 'text', text: header + warnings + body },
        {
          type: 'resource_link',
          uri: metadata.resource_uri,
          name: stored.filename,
          mimeType: stored.mime_type,
          description: 'Bytes originales del archivo de OneDrive'
        }
      ],
      structuredContent: {
        status: extraction.status,
        format: extraction.format,
        text: extraction.text,
        truncated: Boolean(extraction.truncated),
        warnings: extraction.warnings || [],
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
      content: [{ type: 'text', text: `Error leyendo el archivo: ${error.message}` }]
    };
  }
}

module.exports = handleReadFile;
