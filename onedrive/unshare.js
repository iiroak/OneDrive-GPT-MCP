/**
 * Make a OneDrive item private again: revoke every non-inherited sharing
 * permission (all links and all direct invites) on it in one call.
 *
 * There is no Graph "make private" scope or action; this is implemented as
 * list permissions -> DELETE each non-inherited one. Defaults to a dry run
 * (confirm=false) so a caller sees exactly what would be revoked before
 * committing, since this cannot be undone through this MCP (Graph gives no
 * way to recreate a link with the same URL or re-invite silently).
 */
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { summarizePermission } = require('./list-permissions');
const { ensureAuthenticated } = require('../auth');

async function handleUnshare(args) {
  const itemId = args.itemId;
  const path = args.path;
  const confirm = args.confirm === true;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path });

    const response = await callGraphAPIPaginated(accessToken, 'GET', `me/drive/items/${encodeId(item.id)}/permissions`);
    const permissions = (response && response.value) || [];

    const revocable = permissions.filter((p) => !p.inheritedFrom);
    const inherited = permissions.filter((p) => p.inheritedFrom);

    if (revocable.length === 0) {
      return {
        content: [{
          type: "text",
          text: inherited.length
            ? `"${item.name}" has no revocable permissions of its own (${inherited.length} inherited from a parent folder remain and must be revoked there).`
            : `"${item.name}" is already private.`
        }],
        structuredContent: { itemId: item.id, itemName: item.name, revoked: [], skippedInherited: inherited.length }
      };
    }

    const summaries = revocable.map((permission) => summarizePermission(permission, { includeLinkUrl: false }));

    if (!confirm) {
      const lines = summaries.map((s, i) => `${i + 1}. ${s.kind.toUpperCase()} ${s.link ? `${s.link.type}/${s.link.scope}` : s.grantedTo || s.invitedEmail || ''} (ID: ${s.id})`).join('\n');
      return {
        content: [{
          type: "text",
          text: `DRY RUN: revoking these ${summaries.length} permission(s) on "${item.name}" would make it private:\n\n${lines}${inherited.length ? `\n\n(${inherited.length} inherited permission(s) would remain, from a parent folder.)` : ''}\n\nCall again with confirm=true to actually revoke them.`
        }],
        structuredContent: { itemId: item.id, itemName: item.name, dryRun: true, wouldRevoke: summaries, skippedInherited: inherited.length }
      };
    }

    const results = [];
    for (const permission of revocable) {
      try {
        await callGraphAPI(accessToken, 'DELETE', `me/drive/items/${encodeId(item.id)}/permissions/${encodeId(permission.id)}`);
        results.push({ id: permission.id, revoked: true });
      } catch (error) {
        results.push({ id: permission.id, revoked: false, error: error.message });
      }
    }

    const succeeded = results.filter((r) => r.revoked);
    const failed = results.filter((r) => !r.revoked);

    return {
      content: [{
        type: "text",
        text: `"${item.name}" is now private: revoked ${succeeded.length}/${revocable.length} permission(s).${failed.length ? `\n\n${failed.length} failed:\n${failed.map((f) => `- ${f.id}: ${f.error}`).join('\n')}` : ''}${inherited.length ? `\n\n(${inherited.length} inherited permission(s) remain, from a parent folder.)` : ''}`
      }],
      structuredContent: {
        itemId: item.id,
        itemName: item.name,
        revoked: succeeded.map((r) => r.id),
        failed,
        skippedInherited: inherited.length
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error making item private: ${error.message}` }] };
  }
}

module.exports = handleUnshare;
