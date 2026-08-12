#!/usr/bin/env node
/**
 * M365 Assistant MCP Server - Main entry point
 *
 * A Model Context Protocol server that provides access to
 * Microsoft 365 services (Outlook, OneDrive, Power Automate)
 * through the Microsoft Graph API and Flow API.
 */
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const config = require('./config');

// Import module tools
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { onedriveTools } = require('./onedrive');
const { TOOL_OUTPUT_SCHEMA, toStructuredContent } = require('./utils/mcp-output');
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
  ...onedriveTools
];

const REMOTE_EXCLUDED_TOOLS = new Set(['authenticate']);
const WRITE_TOOLS = new Set([
  'draft-email', 'send-email', 'mark-as-read', 'trash-email', 'permanently-delete-email',
  'accept-event', 'decline-event', 'create-event', 'update-event', 'cancel-event', 'delete-event',
  'create-calendar', 'update-calendar', 'copy-event', 'migrate-events', 'delete-calendar',
  'create-master-category', 'update-master-category', 'delete-master-category',
  'create-folder', 'move-emails', 'create-rule', 'edit-rule-sequence',
  'onedrive-upload', 'onedrive-upload-large', 'onedrive-import-url', 'onedrive-share',
  'onedrive-create-folder', 'onedrive-move', 'onedrive-delete'
]);
const DESTRUCTIVE_TOOLS = new Set([
  'permanently-delete-email', 'cancel-event', 'delete-event', 'delete-calendar', 'migrate-events', 'delete-master-category', 'onedrive-delete'
]);

function toolPolicy(tool) {
  const destructive = DESTRUCTIVE_TOOLS.has(tool.name);
  const write = WRITE_TOOLS.has(tool.name);
  return {
    readOnlyHint: !write,
    destructiveHint: destructive,
    openWorldHint: write,
    ...(tool.name === 'mark-as-read' ? { idempotentHint: true } : {})
  };
}

function requiredScope(toolName) {
  if (DESTRUCTIVE_TOOLS.has(toolName)) return 'outlook:destructive';
  if (WRITE_TOOLS.has(toolName)) return 'outlook:write';
  return 'outlook:read';
}

function visibleTools(remote) {
  return remote ? TOOLS.filter(tool => !REMOTE_EXCLUDED_TOOLS.has(tool.name)) : TOOLS;
}

function createMcpServer({ remote = false, scopes = ['outlook:read', 'outlook:write', 'outlook:destructive'] } = {}) {
  const tools = visibleTools(remote);
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    { capabilities: { tools: {} } }
  );

  server.fallbackRequestHandler = async (request) => {
    const { method, params } = request;

    if (method === 'tools/list') {
      return {
        tools: tools.map(tool => ({
          name: tool.name,
          title: tool.title || tool.name,
           description: tool.description,
           inputSchema: tool.inputSchema,
           outputSchema: TOOL_OUTPUT_SCHEMA,
           annotations: toolPolicy(tool),
           _meta: {
             ...(tool.meta || {}),
             securitySchemes: [{ type: 'oauth2', scopes: [requiredScope(tool.name)] }]
           }
        }))
      };
    }

    if (method === 'resources/list') return { resources: [] };
    if (method === 'prompts/list') return { prompts: [] };

    if (method === 'tools/call') {
      const { name, arguments: args = {} } = params || {};
      const tool = tools.find(candidate => candidate.name === name);
      if (!tool || typeof tool.handler !== 'function') {
        return { isError: true, content: [{ type: 'text', text: `Tool not found: ${name}` }] };
      }
      const scope = requiredScope(name);
      if (remote && !scopes.includes(scope)) {
        return {
          isError: true,
          content: [{ type: 'text', text: `This action requires the ${scope} scope.` }]
        };
      }
      try {
        const result = await tool.handler(args);
        return { ...result, structuredContent: toStructuredContent(result) };
      } catch (error) {
        console.error(`Error in tool ${name}: ${error.message}`);
        return {
          isError: true,
          content: [{ type: 'text', text: `Tool failed: ${error.message}` }]
        };
      }
    }

    return { error: { code: -32601, message: `Method not found: ${method}` } };
  };
  return server;
}

async function startStdio() {
  console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
  console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);
  const server = createMcpServer();
  await server.connect(new StdioServerTransport());
  console.error(`${config.SERVER_NAME} connected and listening`);
}

if (require.main === module) {
  startStdio().catch(error => {
    console.error(`Connection error: ${error.message}`);
    process.exit(1);
  });
}

module.exports = {
  TOOLS,
  WRITE_TOOLS,
  DESTRUCTIVE_TOOLS,
  createMcpServer,
  requiredScope,
  toolPolicy,
  visibleTools,
  startStdio
};
