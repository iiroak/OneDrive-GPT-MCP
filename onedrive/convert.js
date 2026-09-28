/**
 * Convert a OneDrive file to PDF and
 * return a temporary proxied download URL for the converted bytes.
 *
 * Graph: GET /me/drive/items/{itemId}/content?format={format} -> 302
 * https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format
 *
 * Source-extension allowlists below are copied verbatim from the docs.
 * Notably absent from the pdf list: txt, csv, json, and most images (only
 * tif/tiff). Validating up front avoids an opaque Graph failure on an
 * unsupported source type.
 */
const { resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');
const capability = require('./capability');
const config = require('../config');

const PDF_SOURCE_EXTENSIONS = new Set([
  'doc', 'docx', 'dot', 'dotx', 'dotm', 'dsn', 'dwg', 'eml', 'epub', 'fluidframework',
  'form', 'htm', 'html', 'loop', 'loot', 'markdown', 'md', 'msg', 'note', 'odp', 'ods',
  'odt', 'page', 'pps', 'ppsx', 'ppt', 'pptx', 'pulse', 'rtf', 'task', 'tif', 'tiff',
  'wbtx', 'whiteboard', 'xls', 'xlsm', 'xlsx'
]);

function extensionOf(filename) {
  const match = /\.([a-z0-9]+)$/i.exec(String(filename || ''));
  return match ? match[1].toLowerCase() : '';
}

async function handleConvert(args) {
  const itemId = args.itemId;
  const path = args.path;
  const format = args.format || 'pdf';

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }
  if (format !== 'pdf') {
    return { content: [{ type: "text", text: "Only format 'pdf' is currently supported." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path, select: 'id,name,size,file,folder' });

    if (item.folder) {
      return { content: [{ type: "text", text: `"${item.name}" is a folder and cannot be converted.` }] };
    }

    const ext = extensionOf(item.name);
    if (!PDF_SOURCE_EXTENSIONS.has(ext)) {
      return {
        content: [{
          type: "text",
          text: `"${item.name}" (.${ext || 'unknown'}) cannot be converted to PDF. Supported source types: ${Array.from(PDF_SOURCE_EXTENSIONS).sort().join(', ')}.`
        }]
      };
    }

    const temporary = capability.issue(
      item, config.PUBLIC_BASE_URL, config.ONEDRIVE_DOWNLOAD_CAPABILITY_TTL,
      '/content?format=pdf'
    );

    return {
      content: [{
        type: "text",
        text: `Temporary PDF conversion URL for "${item.name}":\n\n${temporary.url}\n\nExpires: ${temporary.expiresAt}`
      }],
      structuredContent: {
        itemId: item.id,
        itemName: item.name,
        format: 'pdf',
        downloadUrl: temporary.url,
        expiresAt: temporary.expiresAt
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error converting file: ${error.message}` }] };
  }
}

module.exports = handleConvert;
