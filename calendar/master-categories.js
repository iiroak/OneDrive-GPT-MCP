/**
 * Outlook master categories functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List the categories configured in the user's Outlook profile.
 * @returns {object} - MCP response
 */
async function handleListMasterCategories() {
  try {
    const accessToken = await ensureAuthenticated();
    const response = await callGraphAPI(accessToken, 'GET', 'me/outlook/masterCategories');
    const categories = Array.isArray(response.value) ? response.value : [];

    if (!categories.length) {
      return {
        content: [{ type: 'text', text: 'No Outlook master categories found.' }],
        structuredContent: { categories: [] }
      };
    }

    const summary = categories
      .map(category => `${category.displayName} (${category.color || 'no color'})`)
      .join('\n');

    return {
      content: [{
        type: 'text',
        text: `Found ${categories.length} Outlook master categories:\n\n${summary}`
      }],
      structuredContent: { categories }
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: 'text',
          text: "Authentication required. Complete the MCP OAuth flow first."
        }]
      };
    }

    return {
      content: [{
        type: 'text',
        text: `Error listing Outlook master categories: ${error.message}`
      }]
    };
  }
}

module.exports = handleListMasterCategories;
