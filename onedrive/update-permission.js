/**
 * Change the role(s) of an existing, non-inherited direct/invite permission.
 *
 * Graph: PATCH /me/drive/items/{itemId}/permissions/{permissionId} { roles }
 * https://learn.microsoft.com/en-us/graph/api/permission-update
 *
 * Explicitly unsupported by Graph itself, per the docs: organizational
 * sharing links and "people" sharing links cannot have their role patched.
 * Those must be deleted and recreated (onedrive-revoke-link +
 * onedrive-share/onedrive-invite).
 */
const { callGraphAPI } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

async function handleUpdatePermission(args) {
  const itemId = args.itemId;
  const path = args.path;
  const permissionId = args.permissionId;
  const roles = args.roles;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }
  if (!permissionId) {
    return { content: [{ type: "text", text: "permissionId is required. Use onedrive-list-permissions to find it." }] };
  }
  if (!Array.isArray(roles) || roles.length === 0 || !roles.every((r) => ['read', 'write', 'owner'].includes(r))) {
    return { content: [{ type: "text", text: "roles must be a non-empty array containing only 'read', 'write' and/or 'owner'." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path });

    const permissionEndpoint = `me/drive/items/${encodeId(item.id)}/permissions/${encodeId(permissionId)}`;
    const existing = await callGraphAPI(accessToken, 'GET', permissionEndpoint);
    if (existing.inheritedFrom) {
      return {
        content: [{
          type: "text",
          text: `Permission "${permissionId}" on "${item.name}" is inherited from a parent folder and cannot be changed directly. Update it on the parent item instead.`
        }]
      };
    }
    if (existing.link && (existing.link.scope === 'organization' || existing.link.scope === 'users')) {
      return {
        content: [{
          type: "text",
          text: `The role of a "${existing.link.scope}" sharing link cannot be changed. Revoke it with onedrive-revoke-link and create a new one with the desired role instead.`
        }]
      };
    }

    const updated = await callGraphAPI(accessToken, 'PATCH', permissionEndpoint, { roles });

    return {
      content: [{
        type: "text",
        text: `Updated permission "${permissionId}" on "${item.name}" to roles: ${(updated.roles || roles).join(', ')}.`
      }],
      structuredContent: {
        itemId: item.id,
        itemName: item.name,
        permissionId,
        roles: updated.roles || roles
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error updating permission: ${error.message}` }] };
  }
}

module.exports = handleUpdatePermission;
