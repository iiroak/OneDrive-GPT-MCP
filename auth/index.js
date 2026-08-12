/**
 * Authentication module for Outlook MCP server
 */
const tokenManager = require('./token-manager');
const tokenStorage = require('./token-storage-instance');
const { authTools } = require('./tools');

/**
 * Ensures the user is authenticated and returns an access token.
 * Automatically refreshes expired tokens using the refresh_token grant.
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    throw new Error('Authentication required');
  }

  // Use TokenStorage which handles automatic refresh
  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

module.exports = {
  tokenManager,
  tokenStorage,
  authTools,
  ensureAuthenticated
};
