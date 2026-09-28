/**
 * OneDrive module for Outlook MCP server
 */
const handleListFiles = require('./list');
const handleSearchFiles = require('./search');
const handleDownload = require('./download');
const handleReadFile = require('./read-file');
const handleExportFile = require('./export-file');
const handleUpload = require('./upload');
const handleUploadLarge = require('./upload-large');
const handleImportUrl = require('./import-url');
const handleShare = require('./share');
const handleMoveItem = require('./move');
const { handleCreateFolder, handleDeleteItem } = require('./folder');
const { handleListPermissions } = require('./list-permissions');
const handleRevokeLink = require('./revoke-link');
const handleUnshare = require('./unshare');
const handleInvite = require('./invite');
const handleUpdatePermission = require('./update-permission');
const { handleResolveLink } = require('./resolve-link');
const handleCopy = require('./copy');
const handleUpdateItem = require('./update-item');
const { handleListVersions, handleRestoreVersion } = require('./versions');
const handlePermanentDelete = require('./permanent-delete');
const handleRestoreItem = require('./restore-item');
const handleQuota = require('./quota');
const handleDelta = require('./delta');
const handleThumbnails = require('./thumbnails');
const handleConvert = require('./convert');

const CHATGPT_FILE_SCHEMA = {
  type: "object",
  properties: {
    download_url: {
      type: "string",
      description: "ChatGPT-provided temporary download URL. Pass it unchanged; never invent or replace it."
    },
    file_id: {
      type: "string",
      description: "ChatGPT-provided file ID. Pass it unchanged."
    },
    mime_type: {
      type: "string",
      description: "Optional MIME type supplied by ChatGPT."
    },
    file_name: {
      type: "string",
      description: "Optional original filename supplied by ChatGPT."
    }
  },
  required: ["download_url", "file_id"],
  additionalProperties: false
};

// OneDrive tool definitions
const onedriveTools = [
  {
    name: "onedrive-list",
    description: "List files and folders in OneDrive at a specific path",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Path to list (e.g., '/Documents', '/Photos'). Defaults to root."
        },
        count: {
          type: "number",
          description: "Number of items to retrieve per page (default: 25, max: 50)"
        },
        cursor: {
          type: "string",
          description: "Opaque nextCursor from a previous onedrive-list call, to fetch the next page. When provided, path and count are ignored."
        }
      },
      required: []
    },
    handler: handleListFiles
  },
  {
    name: "onedrive-search",
    description: "Search for files in OneDrive by name or content",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query to find files"
        },
        count: {
          type: "number",
          description: "Number of results per page (default: 25, max: 50)"
        },
        cursor: {
          type: "string",
          description: "Opaque nextCursor from a previous onedrive-search call, to fetch the next page. When provided, query and count are ignored."
        }
      },
      required: []
    },
    handler: handleSearchFiles
  },
  {
    name: "onedrive-download",
    description: "Get a download URL for a file in OneDrive. Either 'itemId' or 'path' must be provided.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "ID of the item to download"
        },
        path: {
          type: "string",
          description: "Path to the file (alternative to itemId)"
        }
      },
      required: []
    },
    handler: handleDownload
  },
  {
    name: "onedrive-read-file",
    description: "Read the CONTENT of a OneDrive file as text. The server downloads the bytes and extracts the text itself, so this returns readable content directly, not a URL. Supports pdf, docx, pptx, xlsx, html and plain-text formats. Always check status: complete means everything was extracted, partial means real content is missing from text (images, charts or scanned pages), and failed means extraction did not work and you should use onedrive-export-file. Provide itemId, path, or a fileId returned by a previous call.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item to read" },
        path: { type: "string", description: "Path to the file, e.g. '/Documents/notes.md'" },
        fileId: { type: "string", description: "file_id from a previous call" },
        maxChars: { type: "number", description: "Maximum characters to return (default 50000, max 200000)" }
      },
      required: []
    },
    handler: handleReadFile
  },
  {
    name: "onedrive-export-file",
    description: "Transfer a OneDrive file's raw bytes as an embedded MCP resource. Use for images, audio, archives, or when onedrive-read-file returns failed. Binary content is Base64-encoded and capped by OUTLOOK_FILE_INLINE_MAX_BYTES; prefer onedrive-read-file whenever text is enough. Provide itemId, path, or a fileId from a previous call.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item to export" },
        path: { type: "string", description: "Path to the file" },
        fileId: { type: "string", description: "file_id from a previous call" }
      },
      required: []
    },
    handler: handleExportFile
  },
  {
    name: "onedrive-upload",
    description: "Legacy upload for UTF-8 text or explicitly supplied Base64 content to OneDrive. This tool is NOT for files attached in ChatGPT. For every ChatGPT attachment, small or large, use onedrive-upload-large with the top-level 'file' field instead. Do not use a /mnt/data path or onedrive-import-url.",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Destination path including filename (e.g., '/Documents/myfile.txt')"
        },
        content: {
          type: "string",
          description: "UTF-8 text content to upload. Use contentBase64 for binary files."
        },
        contentBase64: {
          type: "string",
          description: "Standard Base64-encoded bytes to upload, for binary files such as PDF. Provide this or content, not both."
        },
        conflictBehavior: {
          type: "string",
          description: "Behavior when file exists: 'rename' (default), 'replace', or 'fail'",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["path"],
      anyOf: [{ required: ["content"] }, { required: ["contentBase64"] }]
    },
    handler: handleUpload
  },
  {
    name: "onedrive-upload-large",
    description: "Upload any file to OneDrive using a reliable upload session. THIS IS THE ONLY CORRECT TOOL FOR A CHATGPT ATTACHMENT, REGARDLESS OF SIZE: pass the attachment unchanged in the top-level 'file' field. NEVER use content or contentBase64, a /mnt/data path, onedrive-import-url, or a manual chunk/session tool for a ChatGPT attachment. Use 'path' for the OneDrive destination; ask for it if the user did not provide one. ChatGPT supplies file.download_url and file.file_id automatically.",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Destination path including filename (e.g., '/Documents/largefile.zip')"
        },
        content: {
          type: "string",
          description: "UTF-8 text content to upload. Use contentBase64 for binary files."
        },
        contentBase64: {
          type: "string",
          description: "Standard Base64-encoded bytes to upload, for binary files such as PDF. Provide this or content, not both."
        },
        file: CHATGPT_FILE_SCHEMA,
        conflictBehavior: {
          type: "string",
          description: "Behavior when file exists: 'rename' (default), 'replace', or 'fail'",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["path"],
      anyOf: [{ required: ["content"] }, { required: ["contentBase64"] }, { required: ["file"] }]
    },
    meta: { "openai/fileParams": ["file"] },
    handler: handleUploadLarge
  },
  {
    name: "onedrive-import-url",
    description: "Download a file from an approved HTTPS capability URL on the server and upload it to OneDrive without placing its bytes in the MCP request.",
    inputSchema: {
      type: "object",
      properties: {
        sourceUrl: {
          type: "string",
          description: "Short-lived HTTPS URL returned by a trusted service"
        },
        path: {
          type: "string",
          description: "Destination path including filename"
        },
        conflictBehavior: {
          type: "string",
          description: "Behavior when file exists: 'rename' (default), 'replace', or 'fail'",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["sourceUrl", "path"]
    },
    handler: handleImportUrl
  },
  {
    name: "onedrive-share",
    description: "Create a public or organization-wide sharing link for a file or folder. For sharing with specific named people instead, use onedrive-invite. To later revoke a link created here, use onedrive-revoke-link with the returned permissionId.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "ID of the item to share"
        },
        path: {
          type: "string",
          description: "Path to the item (alternative to itemId)"
        },
        type: {
          type: "string",
          description: "Link type: 'view' (default), 'edit', or 'embed' (embed is OneDrive Personal only)",
          enum: ["view", "edit", "embed"]
        },
        scope: {
          type: "string",
          description: "Link scope: 'anonymous' (default; anyone with the link) or 'organization' (OneDrive for Business/SharePoint only). 'users' is not supported on OneDrive Personal; use onedrive-invite instead.",
          enum: ["anonymous", "organization"]
        },
        password: {
          type: "string",
          description: "Optional password required to open the link. OneDrive Personal only."
        },
        expirationDateTime: {
          type: "string",
          description: "Optional ISO 8601 date-time (yyyy-MM-ddTHH:mm:ssZ) after which the link stops working."
        },
      },
      required: []
    },
    handler: handleShare
  },
  {
    name: "onedrive-list-permissions",
    description: "List every sharing permission (links and direct invites) on a OneDrive file or folder, including who has access and each permission's ID. Use this before onedrive-revoke-link or onedrive-unshare.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" }
      },
      required: []
    },
    handler: handleListPermissions
  },
  {
    name: "onedrive-revoke-link",
    description: "Revoke one specific sharing permission (a link or a direct grant) by its permission ID, obtained from onedrive-list-permissions or from the permissionId returned by onedrive-share.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        permissionId: { type: "string", description: "The permission ID to revoke" }
      },
      required: ["permissionId"]
    },
    handler: handleRevokeLink
  },
  {
    name: "onedrive-unshare",
    description: "Make a file or folder private again by revoking every non-inherited sharing permission on it (all links and all invited people) in one call. Defaults to a dry run that lists what would be revoked; pass confirm=true to actually revoke.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        confirm: { type: "boolean", description: "Set true to actually revoke. Defaults to false (dry run)." }
      },
      required: []
    },
    handler: handleUnshare
  },
  {
    name: "onedrive-invite",
    description: "Share a OneDrive file or folder with specific people by email, without creating a public link. This is the correct tool for 'share with only these people' on OneDrive Personal, since scope='users' on onedrive-share is not available there.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        recipients: {
          type: "array",
          items: { type: "string" },
          description: "Email addresses to grant access to"
        },
        role: {
          type: "string",
          description: "Access level to grant: 'read' (default) or 'write'",
          enum: ["read", "write"]
        },
        message: { type: "string", description: "Optional message included in the invitation email (max 2000 chars)" },
        requireSignIn: { type: "boolean", description: "Require the recipient to sign in to view the item. Defaults to true." },
        sendInvitation: { type: "boolean", description: "Email the recipients a notification. Defaults to true. If false, access is granted silently." },
        password: { type: "string", description: "Optional password. OneDrive Personal only." },
        expirationDateTime: { type: "string", description: "Optional ISO 8601 date-time after which access expires." }
      },
      required: ["recipients"]
    },
    handler: handleInvite
  },
  {
    name: "onedrive-update-permission",
    description: "Change the role (read/write/owner) of an existing direct or invited permission. Cannot change the role of an organization-wide or specific-people link; revoke and recreate those instead.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        permissionId: { type: "string", description: "The permission ID to update" },
        roles: {
          type: "array",
          items: { type: "string", enum: ["read", "write", "owner"] },
          description: "New role(s) for this permission"
        }
      },
      required: ["permissionId", "roles"]
    },
    handler: handleUpdatePermission
  },
  {
    name: "onedrive-resolve-link",
    description: "Resolve a OneDrive/SharePoint sharing URL to the item it points to (name, ID, size). This may redeem the link if necessary and requires outlook:write.",
    inputSchema: {
      type: "object",
      properties: {
        url: { type: "string", description: "The sharing URL to resolve" },
        redeem: { type: "boolean", description: "If true, redeem the link for durable access to your account instead of just peeking at metadata. Defaults to false." }
      },
      required: ["url"]
    },
    handler: handleResolveLink
  },
  {
    name: "onedrive-create-folder",
    description: "Create a new folder in OneDrive",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Parent folder path (e.g., '/Documents'). Defaults to root."
        },
        name: {
          type: "string",
          description: "Name of the new folder"
        },
        conflictBehavior: {
          type: "string",
          description: "Behavior when a folder with this name already exists: 'rename' (default, e.g. 'Docs 1') or 'fail'.",
          enum: ["rename", "fail"]
        }
      },
      required: ["name"]
    },
    handler: handleCreateFolder
  },
  {
    name: "onedrive-move",
    description: "Move and/or rename an existing file or folder in OneDrive without downloading it. Either 'itemId' or 'path' must be provided, and at least one of 'destinationPath' or 'newName' is required.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "ID of the item to move or rename"
        },
        path: {
          type: "string",
          description: "Path to the item (alternative to itemId)"
        },
        destinationPath: {
          type: "string",
          description: "Destination folder path. Use '/' or 'root' for the OneDrive root."
        },
        newName: {
          type: "string",
          description: "New name for the item"
        }
      },
      required: []
    },
    handler: handleMoveItem
  },
  {
    name: "onedrive-copy",
    description: "Copy a file or folder to a destination folder, optionally renaming it. This is an asynchronous Graph operation; by default this tool waits up to 30s for it to finish. The copy does not retain the source permissions; it inherits the destination folder's permissions. On OneDrive Personal, name collisions are reported as failures.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item to copy" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        destinationPath: { type: "string", description: "Destination folder path. Use '/' or 'root' for the OneDrive root." },
        newName: { type: "string", description: "Optional new name for the copy" },
        wait: { type: "boolean", description: "Wait for the copy to finish before returning (default true). If false, returns immediately once Graph accepts the request." }
      },
      required: ["destinationPath"]
    },
    handler: handleCopy
  },
  {
    name: "onedrive-update-item",
    description: "Update writable metadata on a file or folder: description (OneDrive Personal only) and/or file system timestamps. Use onedrive-move for renaming or moving.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        description: { type: "string", description: "New description. OneDrive Personal only." },
        createdDateTime: { type: "string", description: "ISO 8601 date-time to set as the created timestamp." },
        lastModifiedDateTime: { type: "string", description: "ISO 8601 date-time to set as the last-modified timestamp." },
        ifMatch: { type: "string", description: "Optional eTag/cTag to require the item be unchanged since it was read (optimistic concurrency)." }
      },
      required: []
    },
    handler: handleUpdateItem
  },
  {
    name: "onedrive-list-versions",
    description: "List prior versions of a OneDrive file, newest first.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" }
      },
      required: []
    },
    handler: handleListVersions
  },
  {
    name: "onedrive-restore-version",
    description: "Restore a file to a prior version (creates a new current version; existing version history is preserved).",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        versionId: { type: "string", description: "The version ID to restore, from onedrive-list-versions" }
      },
      required: ["versionId"]
    },
    handler: handleRestoreVersion
  },
  {
    name: "onedrive-restore-item",
    description: "Restore a deleted item from the recycle bin by its item ID (the ID it had before deletion, returned by onedrive-delete). OneDrive Personal only; Graph provides no way to list the recycle bin, so this only works if you already have the item ID.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "The item's ID before it was deleted" },
        destinationFolderId: { type: "string", description: "Optional folder ID to restore into, instead of its original location" },
        newName: { type: "string", description: "Optional new name for the restored item" }
      },
      required: ["itemId"]
    },
    handler: handleRestoreItem
  },
  {
    name: "onedrive-quota",
    description: "Get the current OneDrive storage quota: total, used, remaining, recycle bin usage, and state.",
    inputSchema: { type: "object", properties: {}, required: [] },
    handler: handleQuota
  },
  {
    name: "onedrive-delta",
    description: "Track changes (created, modified, deleted items) across the whole OneDrive since a previous sync. Call with no cursor to start a full sync; save the returned deltaCursor and pass it as cursor next time to see only what changed since.",
    inputSchema: {
      type: "object",
      properties: {
        cursor: { type: "string", description: "nextCursor or deltaCursor from a previous onedrive-delta call. Omit to start a full sync." },
        maxItems: { type: "number", description: "Maximum changes to return in this call (default 200, max 1000)" }
      },
      required: []
    },
    handler: handleDelta
  },
  {
    name: "onedrive-thumbnails",
    description: "Get a temporary URL for a thumbnail image of a OneDrive file.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        size: { type: "string", description: "Thumbnail size (default 'medium')", enum: ["small", "medium", "large", "smallSquare", "mediumSquare", "largeSquare"] }
      },
      required: []
    },
    handler: handleThumbnails
  },
  {
    name: "onedrive-convert",
    description: "Convert a supported OneDrive file (Office documents, HTML, Markdown, RTF, and a few others; NOT txt/csv/json/most images) to PDF and get a temporary download URL for the result.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        format: { type: "string", description: "Target format. Only 'pdf' is currently supported.", enum: ["pdf"] }
      },
      required: []
    },
    handler: handleConvert
  },
  {
    name: "onedrive-delete",
    description: "Delete a file or folder from OneDrive (moves it to the recycle bin; recoverable with onedrive-restore-item using the returned item ID). For an irreversible purge that bypasses the recycle bin, use onedrive-permanent-delete instead.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "ID of the item to delete"
        },
        path: {
          type: "string",
          description: "Path to the item (alternative to itemId)"
        }
      },
      required: []
    },
    handler: handleDeleteItem
  },
  {
    name: "onedrive-permanent-delete",
    description: "PERMANENTLY delete a file or folder, bypassing the recycle bin entirely. This cannot be undone through this MCP (Graph v1.0 has no way to list or recover a permanently-deleted item). Defaults to a dry run; pass confirm=true to actually delete.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "ID of the item to permanently delete" },
        path: { type: "string", description: "Path to the item (alternative to itemId)" },
        confirm: { type: "boolean", description: "Set true to actually delete. Defaults to false (dry run)." }
      },
      required: []
    },
    handler: handlePermanentDelete
  }
];

module.exports = {
  onedriveTools,
  handleListFiles,
  handleSearchFiles,
  handleDownload,
  handleReadFile,
  handleExportFile,
  handleUpload,
  handleUploadLarge,
  handleImportUrl,
  handleShare,
  handleMoveItem,
  handleCreateFolder,
  handleDeleteItem,
  handleListPermissions,
  handleRevokeLink,
  handleUnshare,
  handleInvite,
  handleUpdatePermission,
  handleResolveLink,
  handleCopy,
  handleUpdateItem,
  handleListVersions,
  handleRestoreVersion,
  handlePermanentDelete,
  handleRestoreItem,
  handleQuota,
  handleDelta,
  handleThumbnails,
  handleConvert
};
