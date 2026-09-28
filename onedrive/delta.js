/**
 * Track changes to the entire OneDrive (root-scoped only; Graph v1.0
 * documents no per-folder delta endpoint despite the prose suggesting one).
 *
 * Graph: GET /me/drive/root/delta[?token=...]
 * https://learn.microsoft.com/en-us/graph/api/driveitem-delta
 *
 * Token contract: pass no cursor for a first full sync. Each page returns
 * either @odata.nextLink (keep paging) or @odata.deltaLink (sync complete;
 * save this and pass it back next time as `cursor` to see only what changed
 * since). A 410 Gone mid-sync means the token expired; restart without a
 * cursor.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

async function handleDelta(args) {
  const cursor = args.cursor;
  if (cursor !== undefined && (typeof cursor !== 'string' || cursor.length > 32768)) {
    return { content: [{ type: "text", text: "cursor must be a Graph delta URL no longer than 32768 characters." }] };
  }
  const requestedMax = Number(args.maxItems);
  const maxItems = args.maxItems === undefined || args.maxItems === null || !Number.isFinite(requestedMax) || requestedMax < 1
    ? 200
    : Math.min(Math.floor(requestedMax), 1000);

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = cursor || 'me/drive/root/delta';
    let response;
    try {
      response = await callGraphAPI(accessToken, 'GET', endpoint, null, cursor ? {} : { $top: maxItems });
    } catch (error) {
      if (error.status === 410) {
        return {
          content: [{
            type: "text",
            text: "The delta cursor expired (410 Gone). Call onedrive-delta again without a cursor to restart a full sync."
          }],
          structuredContent: { expired: true, changes: [], nextCursor: null, deltaCursor: null }
        };
      }
      throw error;
    }

    const items = (response.value || []).slice(0, maxItems);
    const nextLink = response['@odata.nextLink'] || null;
    const deltaLink = response['@odata.deltaLink'] || null;

    const lines = items.map((item, i) => {
      if (item.deleted) return `${i + 1}. [DELETED] ${item.name || item.id} (ID: ${item.id})`;
      const kind = item.folder ? 'FOLDER' : 'FILE';
      return `${i + 1}. [${kind}] ${item.name} (ID: ${item.id})`;
    }).join('\n');

    return {
      content: [{
        type: "text",
        text: items.length === 0
          ? 'No changes since the last sync.'
          : `${items.length} change(s):\n\n${lines}${nextLink ? '\n\n(more changes available; pass the returned nextCursor to continue this sync)' : deltaLink ? '\n\n(sync complete; save deltaCursor and pass it next time to see only future changes)' : ''}`
      }],
      structuredContent: {
        changes: items.map((item) => ({
          id: item.id,
          name: item.name || null,
          isFolder: !!item.folder,
          deleted: !!item.deleted
        })),
        nextCursor: nextLink,
        deltaCursor: deltaLink
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error tracking changes: ${error.message}` }] };
  }
}

module.exports = handleDelta;
