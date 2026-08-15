const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { masterCategoryPath } = require('./paths');

function errorResponse(action, error) {
  if (error.message === 'Authentication required') return { content: [{ type: 'text', text: "Authentication required. Complete the MCP OAuth flow first." }] };
  return { content: [{ type: 'text', text: `Error ${action}: ${error.message}` }] };
}

async function handleCreateMasterCategory(args = {}) {
  if (!args.displayName || !args.color) return { content: [{ type: 'text', text: 'displayName and color are required to create a master category.' }] };
  try {
    const accessToken = await ensureAuthenticated();
    const category = await callGraphAPI(accessToken, 'POST', 'me/outlook/masterCategories', {
      displayName: args.displayName,
      color: args.color
    });
    return {
      content: [{ type: 'text', text: `Master category '${args.displayName}' has been successfully created.` }],
      structuredContent: { category }
    };
  } catch (error) {
    return errorResponse('creating master category', error);
  }
}

async function handleUpdateMasterCategory(args = {}) {
  if (!args.categoryId || !args.color) return { content: [{ type: 'text', text: 'categoryId and color are required to update a master category.' }] };
  try {
    const accessToken = await ensureAuthenticated();
    const category = await callGraphAPI(accessToken, 'PATCH', masterCategoryPath(args.categoryId), { color: args.color });
    return {
      content: [{ type: 'text', text: `Master category with ID ${args.categoryId} has been successfully updated.` }],
      structuredContent: { category }
    };
  } catch (error) {
    return errorResponse('updating master category', error);
  }
}

async function handleDeleteMasterCategory(args = {}) {
  if (!args.categoryId) return { content: [{ type: 'text', text: 'Category ID is required to delete a master category.' }] };
  try {
    const accessToken = await ensureAuthenticated();
    await callGraphAPI(accessToken, 'DELETE', masterCategoryPath(args.categoryId));
    return {
      content: [{ type: 'text', text: `Master category with ID ${args.categoryId} has been successfully deleted.` }],
      structuredContent: { categoryId: args.categoryId, deleted: true }
    };
  } catch (error) {
    return errorResponse('deleting master category', error);
  }
}

module.exports = {
  handleCreateMasterCategory,
  handleUpdateMasterCategory,
  handleDeleteMasterCategory
};
