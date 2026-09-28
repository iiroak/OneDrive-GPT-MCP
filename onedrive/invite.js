/**
 * Share a OneDrive item with specific people by email, without creating a
 * public link. On OneDrive Personal this is the only supported way to
 * restrict access to named people (Graph's `createLink` scope="users" is
 * OneDrive for Business/SharePoint only).
 *
 * Graph: POST /me/drive/items/{itemId}/invite
 * https://learn.microsoft.com/en-us/graph/api/driveitem-invite
 */
const { callGraphAPI } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

async function handleInvite(args) {
  const itemId = args.itemId;
  const path = args.path;
  const recipients = args.recipients;
  const role = args.role || 'read'; // read | write
  const message = args.message;
  for (const key of ['requireSignIn', 'sendInvitation']) {
    if (args[key] !== undefined && typeof args[key] !== 'boolean') {
      return { content: [{ type: "text", text: `${key} must be a boolean.` }] };
    }
  }
  const requireSignIn = args.requireSignIn !== false;
  const sendInvitation = args.sendInvitation !== false;
  const password = args.password;
  const expirationDateTime = args.expirationDateTime;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }
  if (!Array.isArray(recipients) || recipients.length === 0) {
    return { content: [{ type: "text", text: "recipients must be a non-empty array of email addresses." }] };
  }
  const invalid = recipients.filter((email) => typeof email !== 'string' || !EMAIL_RE.test(email));
  if (invalid.length > 0) {
    return { content: [{ type: "text", text: `Invalid email address(es): ${invalid.join(', ')}` }] };
  }
  if (!['read', 'write'].includes(role)) {
    return { content: [{ type: "text", text: "role must be 'read' or 'write'." }] };
  }
  if (message && message.length > 2000) {
    return { content: [{ type: "text", text: "message must be 2000 characters or fewer." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path });

    const body = {
      recipients: recipients.map((email) => ({ email })),
      roles: [role],
      requireSignIn,
      sendInvitation
    };
    if (message) body.message = message;
    if (password) body.password = password;
    if (expirationDateTime) body.expirationDateTime = expirationDateTime;

    const response = await callGraphAPI(accessToken, 'POST', `me/drive/items/${encodeId(item.id)}/invite`, body);
    const granted = (response && response.value) || [];

    const succeeded = granted.filter((p) => !p.error);
    const failed = granted.filter((p) => p.error);

    const lines = granted.map((p) => {
      const email = p.invitation?.email || p.grantedTo?.user?.displayName || 'unknown';
      if (p.error) {
        return `- ${email}: FAILED (${p.error.code || 'error'}: ${p.error.message || 'unknown error'})`;
      }
      return `- ${email}: granted ${(p.roles || [role]).join('/')} (permission ID: ${p.id})`;
    }).join('\n');

    return {
      content: [{
        type: "text",
        text: `Shared "${item.name}" with ${recipients.length} recipient(s):\n\n${lines}${failed.length ? `\n\n${failed.length} invitation(s) failed to send; the file may still be shared with an existing user among them. Check individual results above.` : ''}`
      }],
      structuredContent: {
        itemId: item.id,
        itemName: item.name,
        role,
        succeeded: succeeded.map((p) => ({ email: p.invitation?.email || null, permissionId: p.id, roles: p.roles || [] })),
        failed: failed.map((p) => ({ email: p.invitation?.email || null, code: p.error?.code || null, message: p.error?.message || null }))
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    if (error.code === 'invalidRequest' && String(error.message || '').includes('root')) {
      return { content: [{ type: "text", text: "Permissions can't be created on the OneDrive root itself. Share a specific file or folder instead." }] };
    }
    return { content: [{ type: "text", text: `Error inviting recipients: ${error.message}` }] };
  }
}

module.exports = handleInvite;
