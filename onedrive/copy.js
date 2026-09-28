/**
 * Copy a OneDrive item (file or folder) to a destination folder, optionally
 * renaming it. This is an asynchronous Graph operation.
 *
 * Graph: POST /me/drive/items/{itemId}/copy -> 202 Accepted + Location header
 * https://learn.microsoft.com/en-us/graph/api/driveitem-copy
 *
 * The monitor URL in Location is NOT on graph.microsoft.com (commonly
 * *.sharepoint.com or api.onedrive.com) and requires no Authorization header
 * per Microsoft's docs; see utils/operation-monitor.js for why we never
 * attach the Graph bearer token to it.
 *
 * Note (OneDrive Personal): the `@microsoft.graph.conflictBehavior` query
 * parameter is documented as NOT supported for OneDrive Consumer, so a name
 * collision at the destination is reported by the async monitor as a
 * failure rather than resolved automatically.
 */
const { callGraphAPI, resolveDriveId } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { pollOperation } = require('../utils/operation-monitor');
const { ensureAuthenticated } = require('../auth');

async function handleCopy(args) {
  const itemId = args.itemId;
  const path = args.path;
  const destinationPath = args.destinationPath;
  const newName = args.newName;
  const wait = args.wait !== false;

  if (!itemId && !path) {
    return { content: [{ type: "text", text: "Either itemId or path is required." }] };
  }
  if (!destinationPath) {
    return { content: [{ type: "text", text: "destinationPath is required." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const source = await resolveItem(accessToken, { itemId, path });

    const destination = await resolveItem(accessToken, { path: destinationPath });
    if (!destination.folder) {
      return { content: [{ type: "text", text: `Destination is not a folder: ${destinationPath}` }] };
    }

    const driveId = await resolveDriveId(accessToken);
    const body = { parentReference: { driveId, id: destination.id } };
    if (newName) body.name = newName;

    const response = await callGraphAPI(
      accessToken, 'POST', `me/drive/items/${encodeId(source.id)}/copy`, body, {}, { returnHeaders: true }
    );

    const location = response.headers && response.headers.location;
    if (!location) {
      return { content: [{ type: "text", text: "Copy was accepted but no monitor URL was returned; check the destination folder shortly to confirm it completed." }] };
    }

    if (!wait) {
      return {
        content: [{
          type: "text",
          text: `Copy of "${source.name}" to "${destinationPath}" started (async). Monitor URL captured; the copy typically finishes within seconds to a few minutes for small items.`
        }],
        structuredContent: { sourceId: source.id, sourceName: source.name, destinationPath, status: 'pending' }
      };
    }

    const result = await pollOperation(location, { maxWaitMs: 30000, pollIntervalMs: 1000 });

    if (!result.settled) {
      return {
        content: [{
          type: "text",
          text: `Copy of "${source.name}" is still in progress after 30s (percent complete: ${result.percentageComplete ?? 'unknown'}). It will likely finish shortly; check the destination folder.`
        }],
        structuredContent: { sourceId: source.id, sourceName: source.name, destinationPath, status: 'pending', percentageComplete: result.percentageComplete }
      };
    }

    if (result.status === 'failed') {
      const errMsg = result.error?.message || 'unknown error';
      return {
        content: [{ type: "text", text: `Copy of "${source.name}" failed: ${errMsg}` }],
        structuredContent: { sourceId: source.id, sourceName: source.name, destinationPath, status: 'failed', error: result.error }
      };
    }

    return {
      content: [{
        type: "text",
        text: `Copied "${source.name}" to "${destinationPath}"${newName ? ` as "${newName}"` : ''}.${result.resourceId ? `\n\nNew item ID: ${result.resourceId}` : ''}`
      }],
      structuredContent: {
        sourceId: source.id,
        sourceName: source.name,
        destinationPath,
        newItemId: result.resourceId || null,
        status: 'completed'
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error copying item: ${error.message}` }] };
  }
}

module.exports = handleCopy;
