/**
 * List and restore prior versions of a OneDrive file.
 *
 * Graph:
 *   GET  /me/drive/items/{itemId}/versions
 *   POST /me/drive/items/{itemId}/versions/{versionId}/restoreVersion -> 204
 * https://learn.microsoft.com/en-us/graph/api/driveitem-list-versions
 * https://learn.microsoft.com/en-us/graph/api/driveitemversion-restoreversion
 *
 * Note: OneDrive does not preserve complete metadata for previous versions,
 * and $orderby is not supported on this collection (Graph returns it
 * newest-first already).
 */
const { callGraphAPI } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

async function handleListVersions(args) {
  const itemId = args.itemId;
  const path = args.path;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path });

    if (item.folder) {
      return { content: [{ type: "text", text: `"${item.name}" is a folder; versions only apply to files.` }] };
    }

    const response = await callGraphAPI(accessToken, 'GET', `me/drive/items/${encodeId(item.id)}/versions`);
    const versions = (response && response.value) || [];

    if (versions.length === 0) {
      return {
        content: [{ type: "text", text: `"${item.name}" has no prior versions.` }],
        structuredContent: { itemId: item.id, itemName: item.name, versions: [] }
      };
    }

    const lines = versions.map((v, i) => {
      const modified = v.lastModifiedDateTime ? new Date(v.lastModifiedDateTime).toLocaleString() : 'unknown';
      const by = v.lastModifiedBy?.user?.displayName || 'unknown';
      return `${i + 1}. Version ${v.id} - ${modified} by ${by}${v.size != null ? ` (${v.size} bytes)` : ''}`;
    }).join('\n');

    return {
      content: [{
        type: "text",
        text: `Versions of "${item.name}" (newest first):\n\n${lines}\n\nUse onedrive-restore-version with the version ID to restore one.`
      }],
      structuredContent: {
        itemId: item.id,
        itemName: item.name,
        versions: versions.map((v) => ({
          id: v.id,
          lastModifiedDateTime: v.lastModifiedDateTime || null,
          lastModifiedBy: v.lastModifiedBy?.user?.displayName || null,
          size: v.size ?? null
        }))
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error listing versions: ${error.message}` }] };
  }
}

async function handleRestoreVersion(args) {
  const itemId = args.itemId;
  const path = args.path;
  const versionId = args.versionId;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }
  if (!versionId) {
    return { content: [{ type: "text", text: "versionId is required. Use onedrive-list-versions to find it." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path });

    // restoreVersion requires Content-Type: application/json with no body.
    await callGraphAPI(accessToken, 'POST', `me/drive/items/${encodeId(item.id)}/versions/${encodeId(versionId)}/restoreVersion`);

    return {
      content: [{
        type: "text",
        text: `Restored "${item.name}" to version ${versionId}. This created a new current version; the version history (including the one you just restored from) was preserved.`
      }],
      structuredContent: { itemId: item.id, itemName: item.name, restoredFromVersion: versionId }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    if (error.code === 'itemNotFound' || error.status === 404) {
      return { content: [{ type: "text", text: `Version "${versionId}" was not found. Use onedrive-list-versions to see available versions.` }] };
    }
    return { content: [{ type: "text", text: `Error restoring version: ${error.message}` }] };
  }
}

module.exports = { handleListVersions, handleRestoreVersion };
