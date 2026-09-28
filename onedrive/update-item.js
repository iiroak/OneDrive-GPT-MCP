/**
 * Update writable metadata on a OneDrive item: description and/or
 * fileSystemInfo timestamps. Renaming/moving remain in onedrive-move; this
 * tool is for properties that PATCH accepts but move.js never touched.
 *
 * Graph: PATCH /me/drive/items/{itemId} { description, fileSystemInfo }
 * https://learn.microsoft.com/en-us/graph/api/driveitem-update
 * https://learn.microsoft.com/en-us/graph/api/resources/driveitem
 *
 * `description` is documented as "Read-write. Only on OneDrive Personal."
 * There is no separate flag to detect driveType cheaply here, so we send it
 * and surface Graph's own rejection (if any) rather than guessing.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

async function handleUpdateItem(args) {
  const itemId = args.itemId;
  const path = args.path;
  const description = args.description;
  const createdDateTime = args.createdDateTime;
  const lastModifiedDateTime = args.lastModifiedDateTime;
  const ifMatch = args.ifMatch;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }
  if (description === undefined && createdDateTime === undefined && lastModifiedDateTime === undefined) {
    return { content: [{ type: "text", text: "Provide at least one of: description, createdDateTime, lastModifiedDateTime." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const item = await resolveItem(accessToken, { itemId, path });

    const body = {};
    if (description !== undefined) body.description = description;
    if (createdDateTime || lastModifiedDateTime) {
      body.fileSystemInfo = {};
      if (createdDateTime) body.fileSystemInfo.createdDateTime = createdDateTime;
      if (lastModifiedDateTime) body.fileSystemInfo.lastModifiedDateTime = lastModifiedDateTime;
    }

    const options = ifMatch ? { headers: { 'If-Match': ifMatch } } : {};

    let updated;
    try {
      updated = await callGraphAPI(accessToken, 'PATCH', `me/drive/items/${encodeId(item.id)}`, body, {}, options);
    } catch (error) {
      if (error.status === 412) {
        return { content: [{ type: "text", text: `Update rejected: "${item.name}" was modified since ifMatch was captured. Re-fetch the item and retry.` }] };
      }
      if (description !== undefined && error.status === 400 && /description/i.test(error.message)) {
        return { content: [{ type: "text", text: `Could not set description on "${item.name}": ${error.message}. The description property is only writable on OneDrive Personal.` }] };
      }
      throw error;
    }

    const changed = [];
    if (description !== undefined) changed.push('description');
    if (createdDateTime) changed.push('createdDateTime');
    if (lastModifiedDateTime) changed.push('lastModifiedDateTime');

    return {
      content: [{
        type: "text",
        text: `Updated "${item.name}": ${changed.join(', ')}.`
      }],
      structuredContent: {
        itemId: updated.id || item.id,
        itemName: updated.name || item.name,
        description: updated.description ?? description ?? null,
        eTag: updated.eTag || null
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error updating item: ${error.message}` }] };
  }
}

module.exports = handleUpdateItem;
