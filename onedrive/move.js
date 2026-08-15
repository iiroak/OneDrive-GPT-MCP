/**
 * OneDrive move and rename functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Move and/or rename a OneDrive item without downloading or recreating it.
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveItem(args) {
  const itemId = args.itemId;
  const path = args.path;
  const destinationPath = args.destinationPath;
  const newName = args.newName;

  if (!itemId && !path) {
    return {
      content: [{
        type: "text",
        text: "Either itemId or path is required."
      }]
    };
  }

  if (!destinationPath && newName === undefined) {
    return {
      content: [{
        type: "text",
        text: "Either destinationPath or newName is required."
      }]
    };
  }

  if (newName !== undefined && (typeof newName !== 'string' || !newName.trim())) {
    return {
      content: [{
        type: "text",
        text: "newName must be a non-empty string."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const sourceEndpoint = itemId
      ? `me/drive/items/${itemId}`
      : `me/drive/root:/${path.replace(/^\/+|\/+$/g, '')}`;
    const source = await callGraphAPI(accessToken, 'GET', sourceEndpoint);

    if (!source || !source.id) {
      return {
        content: [{
          type: "text",
          text: itemId ? "Item not found." : `Item not found at path: ${path}`
        }]
      };
    }

    const body = {};
    let destinationName;

    if (destinationPath) {
      const normalizedDestination = destinationPath.replace(/^\/+|\/+$/g, '');
      const isRootDestination = !normalizedDestination || normalizedDestination === 'root';
      const destinationEndpoint = !isRootDestination
        ? `me/drive/root:/${normalizedDestination}`
        : 'me/drive/root';
      const destination = await callGraphAPI(accessToken, 'GET', destinationEndpoint);

      if (!destination || !destination.id) {
        return {
          content: [{
            type: "text",
            text: `Destination folder not found: ${destinationPath}`
          }]
        };
      }

      if (!destination.folder) {
        return {
          content: [{
            type: "text",
            text: `Destination is not a folder: ${destinationPath}`
          }]
        };
      }

      body.parentReference = { id: destination.id };
      destinationName = destination.name || (isRootDestination ? 'root' : normalizedDestination.split('/').pop());
    }

    if (newName !== undefined) body.name = newName;

    const updated = await callGraphAPI(accessToken, 'PATCH', `me/drive/items/${source.id}`, body);
    const finalName = updated?.name || newName || source.name;
    const action = destinationPath && newName !== undefined
      ? `Moved "${source.name}" to "${destinationName}" and renamed it to "${finalName}".`
      : destinationPath
        ? `Moved "${source.name}" to "${destinationName}".`
        : `Renamed "${source.name}" to "${finalName}".`;

    return {
      content: [{
        type: "text",
        text: `${action}\n\nID: ${updated?.id || source.id}${updated?.webUrl ? `\nWeb URL: ${updated.webUrl}` : ''}`
      }],
      structuredContent: {
        id: updated?.id || source.id,
        name: finalName,
        moved: !!destinationPath,
        renamed: newName !== undefined
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
        text: `Error moving item: ${error.message}`
      }]
    };
  }
}

module.exports = handleMoveItem;
