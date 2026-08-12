[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/ryaker-outlook-mcp-badge.png)](https://mseep.ai/app/ryaker-outlook-mcp)

# M365 Assistant MCP Server

This fork exposes a Streamable HTTP MCP endpoint for ChatGPT at
`https://mcp.iroak.dev/outlook/mcp`. Its administrative adapter follows the
`iroak.mcp-admin/v1` contract used by the MCP Admin panel.

The remote service provides Outlook and OneDrive through Microsoft Graph with
delegated OAuth. The repository also retains Power Automate modules for legacy
or local use, but they are not registered in the remote ChatGPT tool set.

## Supported Services

- **Outlook** - Email, calendar, folders, and rules
- **OneDrive** - Files, folders, search, and sharing

## Directory Structure

```
├── index.js                 # Main entry point
├── config.js                # Configuration settings
├── auth/                    # Authentication modules
│   ├── index.js             # Authentication exports
│   ├── token-manager.js     # Legacy token storage and refresh
│   └── tools.js             # Auth-related tools
├── calendar/                # Calendar functionality
│   ├── index.js             # Calendar exports
│   ├── list.js              # List events
│   ├── create.js            # Create event
│   ├── calendars.js         # Calendar container CRUD
│   ├── get.js               # Read a complete event
│   ├── copy.js              # Copy an event without deleting the source
│   ├── migrate.js           # Copy, verify, and remove source events
│   ├── categories.js        # Master category mutations
│   ├── paths.js             # Encoded Graph calendar paths
│   ├── event-payload.js     # Event payload normalization
│   ├── update.js            # Update event fields and categories
│   ├── master-categories.js # List Outlook master categories
│   ├── delete.js            # Delete event
│   ├── cancel.js            # Cancel event
│   ├── accept.js            # Accept event
│   └── decline.js           # Decline event
├── email/                   # Email functionality
│   ├── index.js             # Email exports
│   ├── list.js              # List emails
│   ├── search.js            # Search emails
│   ├── read.js              # Read email
│   ├── send.js              # Send email
│   └── mark-as-read.js      # Mark email read/unread
├── folder/                  # Folder functionality
│   ├── index.js             # Folder exports
│   ├── list.js              # List folders
│   ├── create.js            # Create folder
│   └── move.js              # Move emails
├── rules/                   # Email rules functionality
│   ├── index.js             # Rules exports
│   ├── list.js              # List rules
│   └── create.js            # Create rule
├── onedrive/                # OneDrive functionality
│   ├── index.js             # OneDrive exports
│   ├── list.js              # List files/folders
│   ├── search.js            # Search files
│   ├── download.js          # Get download URL
│   ├── capability.js        # Short-lived server-side download capabilities
│   ├── import-url.js        # Stream a capability URL into OneDrive
│   ├── upload-file.js       # File-path based Graph upload session
│   ├── upload.js            # Simple upload (<4MB)
│   ├── upload-large.js      # Chunked upload (>4MB)
│   ├── share.js             # Create sharing link
│   └── folder.js            # Create/delete folders
├── power-automate/          # Retained Power Automate modules (not remote)
│   ├── index.js             # Power Automate exports
│   ├── flow-api.js          # Flow API client
│   ├── list-environments.js # List environments
│   ├── list-flows.js        # List flows
│   ├── run-flow.js          # Trigger flow
│   ├── list-runs.js         # Run history
│   └── toggle-flow.js       # Enable/disable flow
└── utils/                   # Utility functions
    ├── graph-api.js         # Microsoft Graph API helper
    ├── odata-helpers.js     # OData query building
    └── mock-data.js         # Test mode data
```

## Features

- **Authentication**: Microsoft OAuth for Graph and OAuth/PKCE for the remote MCP client
- **Email Management**: List, search, read, send, and organize emails
- **Calendar Management**: Calendar CRUD, event CRUD, copy/migrate, categories, accept, decline, and delete
- **OneDrive Integration**: List, search, upload, download, and share files
- **Modular Structure**: Clean separation of concerns for maintainability
- **Test Mode**: Simulated responses for testing without real API calls

## Available Tools

### Structured output

Every exposed tool publishes an `outputSchema` and returns both the existing
human-readable `content` text and machine-readable `structuredContent`:

```json
{
  "message": "Human-readable result of the tool call.",
  "data": {}
}
```

`message` is always present. `data` is optional and is populated when a tool
already has an additional structured payload, such as `about`. This preserves
the existing text response while giving clients a stable envelope for every
tool. Clients should use the declared `outputSchema` rather than infer a
schema from the text response.

### Outlook (Email & Calendar)
| Tool | Description |
|------|-------------|
| `list-emails` | List recent emails from inbox |
| `search-emails` | Search emails with filters |
| `read-email` | Read email content |
| `send-email` | Send a new email |
| `mark-as-read` | Mark email as read/unread |
| `list-events` | List calendar events |
| `list-calendars` | List Outlook calendars |
| `create-calendar` | Create a blank Outlook calendar |
| `update-calendar` | Rename or recolor a calendar |
| `delete-calendar` | Delete a non-default calendar |
| `get-event` | Read a complete event |
| `create-event` | Create calendar event |
| `update-event` | Update an existing event, including its categories |
| `copy-event` | Copy an event to another calendar without deleting the source |
| `migrate-events` | Copy, verify, and delete source events with explicit confirmation |
| `list-master-categories` | List Outlook category names and colors |
| `create-master-category` | Create an Outlook master category |
| `update-master-category` | Change a master category color |
| `delete-master-category` | Delete an Outlook master category |
| `accept-event` | Accept event invitation |
| `decline-event` | Decline event invitation |
| `delete-event` | Delete calendar event |
| `list-folders` | List mail folders |
| `create-folder` | Create mail folder |
| `move-emails` | Move emails between folders |
| `list-rules` | List inbox rules |
| `create-rule` | Create inbox rule |

### OneDrive
| Tool | Description |
|------|-------------|
| `onedrive-list` | List files in a path |
| `onedrive-search` | Search files by query |
| `onedrive-download` | Get download URL |
| `onedrive-upload` | Upload small file (<4MB) |
| `onedrive-upload-large` | Chunked upload (>4MB) |
| `onedrive-import-url` | Server-side import of an approved HTTPS capability URL |
| `onedrive-share` | Create sharing link |
| `onedrive-create-folder` | Create folder |
| `onedrive-delete` | Delete file or folder |

## Quick Start

1. **Install dependencies**: `npm install`
2. **Azure setup**: Register app in Azure Portal (see detailed steps below)
3. **Configure environment**: Copy `.env.example` to `.env` and add your Azure credentials
4. **Run locally**: Use `npm start:http` for the remote-compatible HTTP server or `npm start` for `stdio`
5. **Authenticate**: Use the remote OAuth flow or the local authentication tool, depending on the transport

## Installation

### Prerequisites
- Node.js 14.0.0 or higher
- npm or yarn package manager
- Azure account for app registration

### Install Dependencies

```bash
npm install
```

## Azure App Registration & Configuration

### App Registration

1. Open [Azure Portal](https://portal.azure.com/)
2. Search for "App registrations"
3. Click "New registration"
4. Name: "M365 MCP Server"
5. Account type: "Accounts in any organizational directory and personal Microsoft accounts"
6. Redirect URI: Web -> `https://mcp.iroak.dev/outlook/microsoft/callback` for the deployed remote service
7. Click "Register"
8. Copy the "Application (client) ID" for your `.env` file

### App Permissions

1. Go to "API permissions" under Manage
2. Click "Add a permission" → "Microsoft Graph" → "Delegated permissions"
3. Add these permissions:
   - `offline_access`
   - `User.Read`
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
    - `Calendars.Read`, `Calendars.ReadWrite`
    - `MailboxSettings.Read`, `MailboxSettings.ReadWrite`
   - `Files.Read`, `Files.ReadWrite`
4. Click "Add permissions"

### Client Secret

1. Go to "Certificates & secrets" → "Client secrets"
2. Click "New client secret"
3. Add description and select expiration
4. **Copy the VALUE** (not the Secret ID)

## Configuration

### 1. Environment Variables

```bash
cp .env.example .env
```

Edit `.env`:
```bash
# Get these values from Azure Portal > App Registrations > Your App
MS_CLIENT_ID=your-application-client-id-here
MS_CLIENT_SECRET=your-client-secret-VALUE-here
# Use "consumers" for a personal Microsoft account.
MS_TENANT_ID=consumers
USE_TEST_MODE=false
```

**Important Notes:**
- Use `MS_CLIENT_ID` and `MS_CLIENT_SECRET` in the `.env` file
- The deployed service is configured through the MCP Admin panel; its secrets are not committed to this repository.
- Always use the client secret **VALUE**, never the Secret ID

### 2. Local stdio Configuration

For a local MCP client, configure the stdio entry point directly:

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": {
        "USE_TEST_MODE": "false",
        "MS_CLIENT_ID": "your-client-id",
        "MS_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## Authentication

### Graph API (Outlook + OneDrive)

For the deployed endpoint, open the MCP Admin panel at
`https://mcp.iroak.dev/admin`, configure the Microsoft application, and use the
remote consent flow from ChatGPT. Microsoft tokens are stored in the configured
service data directory and are never returned by the admin API.

For local `stdio` development, the `authenticate` tool can start the local OAuth
flow. The legacy auth server is available with `npm run auth-server` when that
local workflow is required.

## Troubleshooting

### Common Issues

**"Cannot find module"**
```bash
npm install
```

**"Port 3333 in use"**
```bash
npx kill-port 3333
npm run auth-server
```

**"Invalid client secret" (AADSTS7000215)**
- Use the secret **VALUE**, not the Secret ID

**"Authentication required"**
- Re-authenticate through the remote consent flow or remove the local token store configured by `MS_TOKEN_STORE_PATH`.

## Testing

```bash
# Run with MCP Inspector
npm run inspect

# Run in test mode (mock data)
npm run test-mode

# Run Jest tests
npm test
```

## Extending the Server

1. Create new module directory
2. Implement tool handlers in separate files
3. Export tool definitions from module index
4. Import and add to `TOOLS` array in `index.js`
