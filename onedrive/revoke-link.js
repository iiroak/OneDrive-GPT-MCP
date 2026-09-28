/**
 * Revoke a single sharing permission (a link or a direct grant) on a
 * OneDrive item.
 *
 * Graph: DELETE /me/drive/items/{itemId}/permissions/{permissionId} -> 204
 * https://learn.microsoft.com/en-us/graph/api/permission-delete
 *
 * Constraint from the docs: only non-inherited permissions can be deleted
 * (permission.inheritedFrom must be null). We check that up front so the
 * caller gets a clear message instead of a raw Graph 400.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

async function handleRevokeLink(args) {
  const itemId = args.itemId;
  const path = args.path;
  const permissionId = args.permissionId;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }
  if (!permissionId) {
    return { content: [{ type: "text", text: "permissionId is required. Use onedrive-list-permissions to find it." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path });

    // Preflight: fetch the permission to confirm it exists and is not
    // inherited before attempting the delete.
    let permission;
    try {
      permission = await callGraphAPI(accessToken, 'GET', `me/drive/items/${encodeId(item.id)}/permissions/${encodeId(permissionId)}`);
    } catch (error) {
      if (error.code === 'itemNotFound' || (error.status === 404)) {
        return { content: [{ type: "text", text: `No permission with ID "${permissionId}" was found on "${item.name}".` }] };
      }
      throw error;
    }

    if (permission.inheritedFrom) {
      return {
        content: [{
          type: "text",
          text: `Permission "${permissionId}" on "${item.name}" is inherited from a parent folder and cannot be revoked directly. Revoke it on the parent item instead.`
        }]
      };
    }

    await callGraphAPI(accessToken, 'DELETE', `me/drive/items/${encodeId(item.id)}/permissions/${encodeId(permissionId)}`);

    const wasLink = !!permission.link;
    const label = wasLink ? `the ${permission.link.type}/${permission.link.scope} link` : 'the permission';

    return {
      content: [{
        type: "text",
        text: `Revoked ${label} (ID: ${permissionId}) on "${item.name}". It no longer grants access.`
      }],
      structuredContent: {
        itemId: item.id,
        itemName: item.name,
        permissionId,
        revoked: true,
        wasLink
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error revoking permission: ${error.message}` }] };
  }
}

module.exports = handleRevokeLink;
