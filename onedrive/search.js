/**
 * OneDrive search files functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Escape a string for embedding inside a single-quoted OData literal.
 * OData escapes an embedded single quote by doubling it ('' not \'), NOT by
 * percent-encoding it — encodeURIComponent leaves ' untouched (it is not in
 * the percent-encode set), so a query containing an apostrophe used to reach
 * Graph unescaped and break the search(q='...') parse.
 * @param {string} value
 * @returns {string}
 */
function escapeODataLiteral(value) {
  return String(value).replace(/'/g, "''");
}

/**
 * Search files handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchFiles(args) {
  const query = args.query;
  const requestedCount = Number(args.count);
  const count = args.count === undefined || args.count === null || !Number.isFinite(requestedCount)
    ? 25
    : Math.max(1, Math.min(50, Math.floor(requestedCount)));
  const cursor = args.cursor;

  if (!query && !cursor) {
    return {
      content: [{
        type: "text",
        text: "Search query is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    let response;
    if (cursor) {
      response = await callGraphAPI(accessToken, 'GET', cursor);
    } else {
      // Use the search endpoint. Percent-encode the whole literal (so
      // characters like & or # inside the query don't corrupt the URL),
      // but OData-escape the embedded quote first, since URI-encoding
      // alone does not protect the OData string literal syntax.
      const endpoint = `me/drive/search(q='${encodeURIComponent(escapeODataLiteral(query))}')`;

      const queryParams = {
        $top: Math.min(50, count),
        $select: config.ONEDRIVE_SELECT_FIELDS
      };

      response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);
    }

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: cursor ? "No more results." : `No files found matching "${query}".`
        }],
        structuredContent: { items: [], nextCursor: null }
      };
    }

    // Format results
    const fileList = response.value.map((item, index) => {
      const isFolder = item.folder ? '[FOLDER]' : '[FILE]';
      const size = item.size ? formatSize(item.size) : '';
      const modified = new Date(item.lastModifiedDateTime).toLocaleString();
      const path = item.parentReference?.path?.replace('/drive/root:', '') || '/';

      return `${index + 1}. ${isFolder} ${item.name}${size ? ` (${size})` : ''}\n   Path: ${path}\n   Modified: ${modified}\n   ID: ${item.id}`;
    }).join("\n\n");

    const nextLink = response['@odata.nextLink'] || null;

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} item(s) matching "${query || '(continued search)'}"${nextLink ? ' (more results available; pass the returned cursor to see them)' : ''}:\n\n${fileList}`
      }],
      structuredContent: {
        items: response.value.map((item) => ({
          id: item.id,
          name: item.name,
          isFolder: !!item.folder,
          size: item.size ?? null,
          lastModifiedDateTime: item.lastModifiedDateTime || null,
          path: item.parentReference?.path?.replace('/drive/root:', '') || '/'
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
        text: `Error searching files: ${error.message}`
      }]
    };
  }
}

/**
 * Format file size to human-readable string
 */
function formatSize(bytes) {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = handleSearchFiles;
