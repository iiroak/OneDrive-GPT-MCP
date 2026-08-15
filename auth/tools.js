/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenStorage = require('./token-storage-instance');

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  const about = {
    name: config.SERVER_NAME,
    version: config.SERVER_VERSION,
    description: 'Access to Outlook mail, calendar, folders, inbox rules, and OneDrive through Microsoft Graph.',
    services: ['Outlook', 'OneDrive'],
    authentication: 'Microsoft Graph delegated OAuth',
    remote_transport: 'Streamable HTTP'
  };

  return {
    content: [{
      type: "text",
      text: `Outlook MCP Server v${about.version}\n\n${about.description}`
    }],
    structuredContent: about
  };
}

/**
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');

  const accessToken = await tokenStorage.getValidAccessToken();

  console.error(`[CHECK-AUTH-STATUS] Access token available: ${accessToken ? 'YES' : 'NO'}`);

  if (!accessToken) {
    console.error('[CHECK-AUTH-STATUS] No valid access token found');
    return {
      content: [{ type: "text", text: "Not authenticated" }]
    };
  }

  console.error('[CHECK-AUTH-STATUS] Access token present');
  console.error(`[CHECK-AUTH-STATUS] Token expires at: ${tokenStorage.getExpiryTime()}`);
  console.error(`[CHECK-AUTH-STATUS] Current time: ${Date.now()}`);
  
  return {
    content: [{ type: "text", text: "Authenticated and ready" }]
  };
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this M365 Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleCheckAuthStatus
};
