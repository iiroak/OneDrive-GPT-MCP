/**
 * OneDrive folder operations (create/delete)
 */
const { callGraphAPI } = require('../utils/graph-api');
const { encodePath, itemEndpoint } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

/**
 * Create folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateFolder(args) {
  const path = args.path;
  const name = args.name;
  const conflictBehavior = args.conflictBehavior || 'rename';

  if (!name) {
    return {
      content: [{
        type: "text",
        text: "Folder name is required."
      }]
    };
  }

  if (!['rename', 'fail'].includes(conflictBehavior)) {
    return {
      content: [{
        type: "text",
        text: "conflictBehavior must be 'rename' or 'fail'."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Build parent folder endpoint
    let endpoint;
    if (!path || path === '/' || path === 'root') {
      endpoint = 'me/drive/root/children';
    } else {
      endpoint = `me/drive/root:/${encodePath(path)}:/children`;
    }

    const body = {
      name: name,
      folder: {},
      '@microsoft.graph.conflictBehavior': conflictBehavior
    };

    let response;
    try {
      response = await callGraphAPI(accessToken, 'POST', endpoint, body);
    } catch (error) {
      if (conflictBehavior === 'fail' && (error.code === 'nameAlreadyExists' || error.status === 409)) {
        return {
          content: [{
            type: "text",
            text: `A folder named "${name}" already exists in ${path || 'root'}.`
          }]
        };
      }
      throw error;
    }

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Failed to create folder."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Successfully created folder "${response.name}"\n\nID: ${response.id}\nWeb URL: ${response.webUrl}`
      }]
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
        text: `Error creating folder: ${error.message}`
      }]
    };
  }
}

/**
 * Delete item handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteItem(args) {
  const itemId = args.itemId;
  const path = args.path;

  if (!itemId && !path) {
    return {
      content: [{
        type: "text",
        text: "Either itemId or path is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Get item details first (to confirm existence and get name)
    const endpoint = itemEndpoint({ itemId, path });

    // Get item info first
    const itemInfo = await callGraphAPI(accessToken, 'GET', endpoint);

    if (!itemInfo || !itemInfo.id) {
      return {
        content: [{
          type: "text",
          text: "Item not found."
        }]
      };
    }

    const itemName = itemInfo.name;
    const isFolder = !!itemInfo.folder;

    // Delete the item
    const deleteEndpoint = itemEndpoint({ itemId: itemInfo.id });
    await callGraphAPI(accessToken, 'DELETE', deleteEndpoint);

    return {
      content: [{
        type: "text",
        text: `Successfully deleted ${isFolder ? 'folder' : 'file'} "${itemName}" (moved to the recycle bin; recoverable with onedrive-restore-item using this item ID).\n\nID: ${itemInfo.id}`
      }],
      // Deletion moves the item to the recycle bin (Graph DELETE, not
      // permanentDelete) and is recoverable via onedrive-restore-item, but
      // ONLY if the caller has the item ID — Graph v1.0 has no OneDrive
      // recycle-bin listing endpoint. Surfacing it here is the only chance
      // to capture it before it becomes hard to find again.
      structuredContent: {
        itemId: itemInfo.id,
        itemName,
        isFolder,
        recoverable: true,
        restoreHint: 'Call onedrive-restore-item with this itemId to recover it from the recycle bin.'
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
        text: `Error deleting item: ${error.message}`
      }]
    };
  }
}

module.exports = {
  handleCreateFolder,
  handleDeleteItem
};
