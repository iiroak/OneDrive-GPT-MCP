/**
 * Report the current OneDrive storage quota.
 *
 * Graph: GET /me/drive
 * https://learn.microsoft.com/en-us/graph/api/drive-get
 * https://learn.microsoft.com/en-us/graph/api/resources/quota
 *
 * Note: on OneDrive Personal, quota is unified across several Microsoft
 * services, not OneDrive-only, and the docs' own example response omits
 * `used` while including the other fields — so we render every field
 * defensively as "unknown" rather than assuming presence.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

function formatBytes(bytes) {
  if (bytes == null) return 'unknown';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  if (bytes === 0) return '0 B';
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(2))} ${sizes[i]}`;
}

async function handleQuota() {
  try {
    const accessToken = await ensureAuthenticated();
    const drive = await callGraphAPI(accessToken, 'GET', 'me/drive', null, { $select: 'id,driveType,quota' });
    const quota = drive.quota || {};

    const percentUsed = quota.total && quota.used != null
      ? ((quota.used / quota.total) * 100).toFixed(1)
      : null;

    return {
      content: [{
        type: "text",
        text: `OneDrive quota (${drive.driveType || 'unknown'} drive):\n\nTotal: ${formatBytes(quota.total)}\nUsed: ${formatBytes(quota.used)}${percentUsed ? ` (${percentUsed}%)` : ''}\nRemaining: ${formatBytes(quota.remaining)}\nRecycle bin: ${formatBytes(quota.deleted)}\nState: ${quota.state || 'unknown'}`
      }],
      structuredContent: {
        driveType: drive.driveType || null,
        total: quota.total ?? null,
        used: quota.used ?? null,
        remaining: quota.remaining ?? null,
        deleted: quota.deleted ?? null,
        state: quota.state || null
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error getting quota: ${error.message}` }] };
  }
}

module.exports = handleQuota;
