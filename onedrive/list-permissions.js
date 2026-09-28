/**
 * List every sharing permission (links and direct invites) on a OneDrive
 * item, so a caller can see who has access before deciding to revoke or
 * change anything.
 *
 * Graph: GET /me/drive/items/{itemId}/permissions
 * https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions
 */
const { callGraphAPIPaginated } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

function summarizePermission(permission, { includeLinkUrl = true } = {}) {
  const summary = {
    id: permission.id,
    roles: permission.roles || [],
    inherited: permission.inheritedFrom != null,
    inheritedFrom: permission.inheritedFrom || null,
    expirationDateTime: permission.expirationDateTime || null,
    hasPassword: !!permission.hasPassword
  };

  if (permission.link) {
    summary.kind = 'link';
    summary.link = {
      type: permission.link.type,
      scope: permission.link.scope,
      preventsDownload: !!permission.link.preventsDownload
    };
    if (includeLinkUrl) summary.link.webUrl = permission.link.webUrl || null;
    const identities = permission.grantedToIdentitiesV2 || permission.grantedToIdentities || [];
    summary.grantedTo = identities.map((identity) => identity?.user?.displayName || identity?.user?.email || 'unknown').filter(Boolean);
  } else if (permission.invitation) {
    summary.kind = 'invitation';
    summary.invitedEmail = permission.invitation.email || null;
    summary.signInRequired = !!permission.invitation.signInRequired;
    const grantee = permission.grantedToV2 || permission.grantedTo;
    summary.grantedTo = grantee?.user?.displayName || grantee?.user?.email || null;
  } else {
    summary.kind = 'direct';
    const grantee = permission.grantedToV2 || permission.grantedTo;
    summary.grantedTo = grantee?.user?.displayName || grantee?.user?.email || null;
  }

  return summary;
}

function formatLine(summary, index) {
  const roles = summary.roles.length ? summary.roles.join('/') : 'unknown';
  const lock = summary.inherited ? ' [inherited, cannot be revoked directly]' : '';
  if (summary.kind === 'link') {
    return `${index + 1}. [LINK] ${summary.link.type}/${summary.link.scope} (${roles})${summary.hasPassword ? ' [password]' : ''}${lock}\n   ID: ${summary.id}\n   URL: ${summary.link.webUrl || 'n/a'}`;
  }
  if (summary.kind === 'invitation') {
    return `${index + 1}. [INVITE] ${summary.invitedEmail || 'unknown'} (${roles})${lock}\n   ID: ${summary.id}`;
  }
  return `${index + 1}. [DIRECT] ${summary.grantedTo || 'unknown'} (${roles})${lock}\n   ID: ${summary.id}`;
}

async function handleListPermissions(args) {
  const itemId = args.itemId;
  const path = args.path;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path });

    const response = await callGraphAPIPaginated(accessToken, 'GET', `me/drive/items/${encodeId(item.id)}/permissions`);
    const permissions = (response && response.value) || [];

    if (permissions.length === 0) {
      return {
        content: [{ type: "text", text: `"${item.name}" has no sharing permissions. It is private.` }],
        structuredContent: { itemId: item.id, itemName: item.name, permissions: [] }
      };
    }

    const summaries = permissions.map(summarizePermission);
    const lines = summaries.map(formatLine).join('\n\n');

    return {
      content: [{
        type: "text",
        text: `Permissions on "${item.name}" (${summaries.length}):\n\n${lines}`
      }],
      structuredContent: {
        itemId: item.id,
        itemName: item.name,
        permissions: summaries
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error listing permissions: ${error.message}` }] };
  }
}

module.exports = { handleListPermissions, summarizePermission };
