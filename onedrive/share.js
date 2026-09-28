/**
 * OneDrive create sharing link functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { encodeId, resolveItem } = require('../utils/onedrive-resolve');
const { ensureAuthenticated } = require('../auth');

const VALID_TYPES = ['view', 'edit', 'embed'];
const VALID_SCOPES = ['anonymous', 'organization', 'users'];

/**
 * Create sharing link handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleShare(args) {
  const itemId = args.itemId;
  const path = args.path;
  const type = args.type || 'view'; // view, edit, embed
  // NOTE: 'anonymous' matches both the documented schema default
  // (onedrive/index.js) and the actual Microsoft Graph default when `scope`
  // is omitted on OneDrive Personal. A previous version of this handler
  // defaulted to 'organization', which is not usable on a personal
  // (consumers) tenant and contradicted the schema's own description.
  const scope = args.scope || 'anonymous'; // anonymous, organization, users
  const password = args.password;
  const expirationDateTime = args.expirationDateTime;

  if (!itemId && !path) {
    return {
      content: [{
        type: "text",
        text: "Either itemId or path is required."
      }]
    };
  }

  if (!VALID_TYPES.includes(type)) {
    return {
      content: [{
        type: "text",
        text: `Invalid type "${type}". Use one of: ${VALID_TYPES.join(', ')}.`
      }]
    };
  }

  if (!VALID_SCOPES.includes(scope)) {
    return {
      content: [{
        type: "text",
        text: `Invalid scope "${scope}". Use one of: ${VALID_SCOPES.join(', ')}.`
      }]
    };
  }

  if (scope !== 'anonymous' && password) {
    return {
      content: [{
        type: "text",
        text: "password is only supported for scope 'anonymous' (OneDrive Personal)."
      }]
    };
  }

  if (scope === 'users') {
    return {
      content: [{
        type: "text",
        text: "scope 'users' is only available on OneDrive for Business/SharePoint, not on OneDrive Personal. Use onedrive-invite to share with specific people on a personal account."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const item = await resolveItem(accessToken, { itemId, path });
    const resolvedItemId = item.id;
    const itemName = item.name;

    // Create the sharing link
    const endpoint = `me/drive/items/${encodeId(resolvedItemId)}/createLink`;
    const body = { type, scope };
    if (password) body.password = password;
    if (expirationDateTime) body.expirationDateTime = expirationDateTime;

    const response = await callGraphAPI(accessToken, 'POST', endpoint, body);

    if (!response || !response.link) {
      return {
        content: [{
          type: "text",
          text: "Failed to create sharing link."
        }]
      };
    }

    const linkInfo = response.link;
    const shareText = itemName
      ? `Sharing link created for "${itemName}":`
      : `Sharing link created:`;

    const scopeNote = linkInfo.scope === 'anonymous'
      ? 'Anyone with this link can access the file.'
      : linkInfo.scope === 'organization'
        ? 'Only people in your organization can access.'
        : 'Access is restricted per the configured scope.';

    return {
      content: [{
        type: "text",
        text: `${shareText}\n\nLink: ${linkInfo.webUrl}\nType: ${linkInfo.type}\nScope: ${linkInfo.scope}\nPermission ID: ${response.id}\n${response.hasPassword ? 'Password protected: yes\n' : ''}${linkInfo.webHtml ? `Embed HTML: ${linkInfo.webHtml}\n` : ''}\nNote: ${scopeNote}\n\nTo revoke this link later, call onedrive-revoke-link with itemId="${resolvedItemId}" and permissionId="${response.id}".`
      }],
      // The permission id is the ONLY way to revoke this link later
      // (DELETE /permissions/{id}). Dropping it here, as the previous
      // implementation did, made every link created through this tool
      // permanent from the MCP's point of view.
      structuredContent: {
        itemId: resolvedItemId,
        itemName,
        permissionId: response.id,
        roles: response.roles || [],
        hasPassword: !!response.hasPassword,
        expirationDateTime: response.expirationDateTime || null,
        link: {
          type: linkInfo.type,
          scope: linkInfo.scope,
          webUrl: linkInfo.webUrl,
          webHtml: linkInfo.webHtml || null
        }
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Complete the MCP OAuth flow first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error creating sharing link: ${error.message}`
      }]
    };
  }
}

module.exports = handleShare;
