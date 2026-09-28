/**
 * Get a temporary, proxied thumbnail URL for a OneDrive file.
 *
 * Graph: GET /me/drive/items/{itemId}/thumbnails/{setId}/{size}/content -> 302
 * https://learn.microsoft.com/en-us/graph/api/driveitem-list-thumbnails
 *
 * The 302 target is a *.files.1drv.com CDN URL, so this reuses the same
 * capability-proxy mechanism as onedrive-download: the Graph bearer token
 * never reaches the browser, and the proxy link expires.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');
const capability = require('./capability');
const config = require('../config');

const VALID_SIZES = ['small', 'medium', 'large', 'smallSquare', 'mediumSquare', 'largeSquare'];

async function handleThumbnails(args) {
  const itemId = args.itemId;
  const path = args.path;
  const size = args.size || 'medium';

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }
  if (!VALID_SIZES.includes(size)) {
    return { content: [{ type: "text", text: `Invalid size "${size}". Use one of: ${VALID_SIZES.join(', ')}.` }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path, select: 'id,name,size,file,folder' });

    if (item.folder) {
      return { content: [{ type: "text", text: `"${item.name}" is a folder; thumbnails only apply to files.` }] };
    }

    const sets = await callGraphAPI(accessToken, 'GET', `me/drive/items/${encodeId(item.id)}/thumbnails`);
    const thumbnailSet = sets.value && sets.value[0];
    if (!thumbnailSet || !thumbnailSet[size]) {
      return { content: [{ type: "text", text: `No "${size}" thumbnail is available for "${item.name}" (this file type may not support thumbnails).` }] };
    }

    const temporary = capability.issue(
      item, config.PUBLIC_BASE_URL, config.ONEDRIVE_DOWNLOAD_CAPABILITY_TTL,
      `/thumbnails/${thumbnailSet.id}/${size}/content`
    );

    return {
      content: [{
        type: "text",
        text: `Temporary ${size} thumbnail URL for "${item.name}":\n\n${temporary.url}\n\nExpires: ${temporary.expiresAt}`
      }],
      structuredContent: {
        itemId: item.id,
        itemName: item.name,
        size,
        thumbnailUrl: temporary.url,
        width: thumbnailSet[size].width || null,
        height: thumbnailSet[size].height || null,
        expiresAt: temporary.expiresAt
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error getting thumbnail: ${error.message}` }] };
  }
}

module.exports = handleThumbnails;
