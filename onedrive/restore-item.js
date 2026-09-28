/**
 * Restore a deleted (recycle-bin) OneDrive item, by item ID.
 *
 * Graph: POST /me/drive/items/{itemId}/restore -> 200 + driveItem
 * https://learn.microsoft.com/en-us/graph/api/driveitem-restore
 *
 * Important, from the docs: "This functionality is currently only available
 * for OneDrive Personal", and it requires Files.ReadWrite.All (NOT the
 * Files.ReadWrite this server otherwise uses everywhere else) — delegated
 * work/school accounts are explicitly "Not supported".
 *
 * Also important: Graph v1.0 has no endpoint to list the recycle bin for
 * OneDrive, so this only works if the caller already has the item ID from
 * before it was deleted (onedrive-delete now returns it for exactly this
 * reason).
 */
const { callGraphAPI } = require('../utils/graph-api');
const { encodeId } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

async function handleRestoreItem(args) {
  const itemId = args.itemId;
  const destinationFolderId = args.destinationFolderId;
  const newName = args.newName;

  if (!itemId) {
    return { content: [{ type: "text", text: "itemId is required. This must be the ID the item had before it was deleted (returned by onedrive-delete)." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const body = {};
    if (destinationFolderId) body.parentReference = { id: destinationFolderId };
    if (newName) body.name = newName;

    let restored;
    try {
      restored = await callGraphAPI(accessToken, 'POST', `me/drive/items/${encodeId(itemId)}/restore`, body);
    } catch (error) {
      if (error.status === 404 || error.code === 'itemNotFound') {
        return { content: [{ type: "text", text: `No deleted item with ID "${itemId}" was found. It may have already been restored, permanently deleted, or the recycle bin retention period expired.` }] };
      }
      if (error.status === 403 || error.code === 'accessDenied') {
        return { content: [{ type: "text", text: `Restore was denied. This action requires the Files.ReadWrite.All scope and is only available on OneDrive Personal, not work/school accounts. If your app's Microsoft Graph scopes were not granted this permission, re-authenticate after adding it.` }] };
      }
      throw error;
    }

    return {
      content: [{
        type: "text",
        text: `Restored "${restored.name}" from the recycle bin.\n\nID: ${restored.id}\nWeb URL: ${restored.webUrl || 'n/a'}`
      }],
      structuredContent: {
        itemId: restored.id,
        itemName: restored.name,
        webUrl: restored.webUrl || null,
        restored: true
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error restoring item: ${error.message}` }] };
  }
}

module.exports = handleRestoreItem;
