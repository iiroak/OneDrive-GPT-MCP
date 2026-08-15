# AGENTS.md

## Project

OneDrive-GPT-MCP is a Node.js Model Context Protocol server for ChatGPT. It provides
Outlook mail, calendars, folders and rules, plus OneDrive file operations through
Microsoft Graph.

## Commands

- `npm install` - Install dependencies
- `npm run start:http` - Start the Streamable HTTP server
- `npm start` - Start the local stdio MCP server
- `npm run inspect` - Inspect the stdio server with MCP Inspector

## Architecture

- `server.js` - Streamable HTTP transport, remote OAuth and administration routes
- `index.js` - MCP server and tool registration
- `config.js` - Runtime configuration and Microsoft Graph settings
- `auth/` - Microsoft Graph token storage and remote MCP OAuth
- `calendar/`, `email/`, `folder/`, `rules/` - Outlook tools
- `onedrive/` - Files, folders and transfer support
- `utils/` - Graph API, file staging, download and document extraction helpers
- `deploy/` - Production systemd, nginx and admin configuration examples

## Configuration

Use `.env.example` as the local configuration template. The deployed service reads
its configuration from the MCP Admin data directory and does not commit secrets.
Never commit Microsoft credentials, OAuth keys, tokens or production environment files.

Required Microsoft Graph permissions include `User.Read`, `offline_access`,
`Mail.ReadWrite`, `Mail.Send`, `MailboxSettings.ReadWrite`, `Calendars.ReadWrite`
and `Files.ReadWrite`.

## Verification

There is no automated test suite in this standalone deployment. At minimum, verify
that `index.js` and `server.js` load successfully and that the HTTP health endpoint
responds after starting the server.

## Change Guidelines

- Keep the remote Streamable HTTP contract stable.
- Preserve explicit read, write and destructive tool policies in `index.js`.
- Do not expose secrets through logs, resources or tool responses.
- Keep production deployment paths and service configuration changes separate from
  application-only changes unless the deployment is explicitly requested.

## Production deployment

This repository intentionally does not document a specific host, LXC, network,
reverse proxy, secret store or persistent production state. Review the operator's
private infrastructure documentation before deploying. Keep production changes
separate from application changes and never copy a local workspace to a server.
