/**
 * Permanently delete a OneDrive item, bypassing the recycle bin entirely.
 * This is IRREVERSIBLE through Graph — there is no v1.0 API to list or
 * restore from a purge (only the ordinary recycle bin, reached via
 * onedrive-delete + onedrive-restore-item, is recoverable).
 *
 * Graph: POST /drives/{driveId}/items/{itemId}/permanentDelete -> 204
 * https://learn.microsoft.com/en-us/graph/api/driveitem-permanentdelete
 *
 * Note: the docs only publish the /drives/{driveId}/... form (no
 * /me/drive/... shortcut), so we resolve the current drive id first.
 */
const { callGraphAPI, resolveDriveId } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

async function handlePermanentDelete(args) {
  const itemId = args.itemId;
  const path = args.path;
  const confirm = args.confirm === true;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path });

    if (!confirm) {
      return {
        content: [{
          type: "text",
          text: `DRY RUN: this would PERMANENTLY delete "${item.name}" (${item.folder ? 'folder' : 'file'}), bypassing the recycle bin entirely. This cannot be undone. Call again with confirm=true to proceed, or use onedrive-delete for a recoverable (recycle bin) delete instead.`
        }],
        structuredContent: { itemId: item.id, itemName: item.name, isFolder: !!item.folder, dryRun: true }
      };
    }

    const driveId = await resolveDriveId(accessToken);
    await callGraphAPI(accessToken, 'POST', `drives/${encodeId(driveId)}/items/${encodeId(item.id)}/permanentDelete`);

    return {
      content: [{
        type: "text",
        text: `Permanently deleted "${item.name}". This bypassed the recycle bin and cannot be undone.`
      }],
      structuredContent: { itemId: item.id, itemName: item.name, permanentlyDeleted: true }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error permanently deleting item: ${error.message}` }] };
  }
}

module.exports = handlePermanentDelete;
