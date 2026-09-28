/**
 * Resolve a OneDrive/SharePoint sharing URL back to the driveItem it points
 * to (metadata only; does not redeem/accept the share by default).
 *
 * Graph: GET /shares/{encoded-sharing-url}
 * https://learn.microsoft.com/en-us/graph/api/shares-get
 *
 * Encoding per the docs: base64 the URL, convert to unpadded base64url
 * (strip '=', '/'->'_', '+'->'-'), prefix with 'u!'.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

function encodeShareUrl(url) {
  const base64 = Buffer.from(String(url), 'utf8').toString('base64');
  const base64url = base64.replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  return `u!${base64url}`;
}

async function handleResolveLink(args) {
  const url = args.url;
  const redeem = args.redeem === true;

  if (!url) {
    return { content: [{ type: "text", text: "url is required." }] };
  }
  let parsedUrl;
  try {
    parsedUrl = new URL(url);
  } catch {
    return { content: [{ type: "text", text: "url must be an absolute HTTPS sharing URL." }] };
  }
  if (parsedUrl.protocol !== 'https:' || parsedUrl.username || parsedUrl.password || String(url).length > 8192) {
    return { content: [{ type: "text", text: "url must be an HTTPS sharing URL without embedded credentials and be at most 8192 characters." }] };
  }
  if (args.redeem !== undefined && typeof args.redeem !== 'boolean') {
    return { content: [{ type: "text", text: "redeem must be a boolean." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const encoded = encodeShareUrl(url);

    const headers = redeem
      ? { Prefer: 'redeemSharingLink' }
      : { Prefer: 'redeemSharingLinkIfNecessary' };

    const shared = await callGraphAPI(accessToken, 'GET', `shares/${encoded}/driveItem`, null, {}, { headers });

    if (!shared || !shared.id) {
      return { content: [{ type: "text", text: "The link could not be resolved to an item." }] };
    }

    return {
      content: [{
        type: "text",
        text: `Link resolves to "${shared.name}" (${shared.folder ? 'folder' : 'file'}).\n\nItem ID: ${shared.id}\nWeb URL: ${shared.webUrl || 'n/a'}\nSize: ${shared.size ?? 'n/a'} bytes${redeem ? '\n\nThis request redeemed the link, granting durable access to your account.' : ''}`
      }],
      structuredContent: {
        itemId: shared.id,
        name: shared.name,
        isFolder: !!shared.folder,
        webUrl: shared.webUrl || null,
        size: shared.size ?? null,
        redeemed: redeem
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Complete the MCP OAuth flow first." }] };
    }
    return { content: [{ type: "text", text: `Error resolving link: ${error.message}` }] };
  }
}

module.exports = { handleResolveLink, encodeShareUrl };
