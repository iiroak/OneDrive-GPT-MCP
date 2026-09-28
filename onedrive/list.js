/**
 * OneDrive list files/folders functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { encodePath } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

/**
 * List files handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFiles(args) {
  const path = args.path || '';
  const requestedCount = Number(args.count);
  const count = args.count === undefined || args.count === null || !Number.isFinite(requestedCount)
    ? 25
    : Math.max(1, Math.min(50, Math.floor(requestedCount)));
  const cursor = args.cursor;

  try {
    const accessToken = await ensureAuthenticated();

    let response;
    if (cursor) {
      response = await callGraphAPI(accessToken, 'GET', cursor);
    } else {
      // Build endpoint - root or specific path
      let endpoint;
      if (!path || path === '/' || path === 'root') {
        endpoint = 'me/drive/root/children';
      } else {
        endpoint = `me/drive/root:/${encodePath(path)}:/children`;
      }

      const queryParams = {
        $top: Math.min(50, count),
        $select: config.ONEDRIVE_SELECT_FIELDS,
        $orderby: 'name'
      };

      response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);
    }

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: cursor ? "No more items." : `No files found in ${path || 'root'}.`
        }],
        structuredContent: { items: [], nextCursor: null }
      };
    }

    // Format results
    const fileList = response.value.map((item, index) => {
      const isFolder = item.folder ? '[FOLDER]' : '[FILE]';
      const size = item.size ? formatSize(item.size) : '';
      const modified = new Date(item.lastModifiedDateTime).toLocaleString();

      return `${index + 1}. ${isFolder} ${item.name}${size ? ` (${size})` : ''}\n   Modified: ${modified}\n   ID: ${item.id}`;
    }).join("\n\n");

    const nextLink = response['@odata.nextLink'] || null;

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} item(s) in ${path || 'root'}${nextLink ? ' (more items available; pass the returned cursor to see them)' : ''}:\n\n${fileList}`
      }],
      structuredContent: {
        items: response.value.map((item) => ({
          id: item.id,
          name: item.name,
          isFolder: !!item.folder,
          size: item.size ?? null,
          lastModifiedDateTime: item.lastModifiedDateTime || null
        })),
        nextCursor: nextLink
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Complete the MCP OAuth flow first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error listing files: ${error.message}`
      }]
    };
  }
}

/**
 * Format file size to human-readable string
 */
function formatSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = handleListFiles;
