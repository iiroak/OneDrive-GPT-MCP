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
const {
  startUploadSession,
  appendUploadSession,
  finishUploadSession,
  abortUploadSession
} = require('./upload-session');
const handleShare = require('./share');
const handleMoveItem = require('./move');
const { handleCreateFolder, handleDeleteItem } = require('./folder');

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
          description: "Number of items to retrieve (default: 25, max: 50)"
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
          description: "Number of results to return (default: 25, max: 50)"
        }
      },
      required: ["query"]
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
    description: "Upload a small file (< 4MB) to OneDrive",
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
    description: "Upload a large file (> 4MB) to OneDrive using chunked upload",
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
        conflictBehavior: {
          type: "string",
          description: "Behavior when file exists: 'rename' (default), 'replace', or 'fail'",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["path"],
      anyOf: [{ required: ["content"] }, { required: ["contentBase64"] }]
    },
    handler: handleUploadLarge
  },
  {
    name: "onedrive-upload-session-start",
    description: "Start a resumable binary upload. Send the file in sequential Base64 chunks, then call complete.",
    inputSchema: {
      type: "object",
      properties: {
        path: { type: "string", description: "Destination path including filename" },
        totalBytes: { type: "integer", minimum: 1, description: "Exact decoded byte length of the complete file" },
        conflictBehavior: {
          type: "string",
          description: "Behavior when file exists: 'rename' (default), 'replace', or 'fail'",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["path", "totalBytes"]
    },
    handler: async (args) => {
      try {
        const result = await startUploadSession(args);
        return {
          content: [{ type: 'text', text: `Upload session ${result.uploadId} started. Send chunks of at most ${result.chunkBytes} bytes.` }],
          structuredContent: result
        };
      } catch (error) {
        return { content: [{ type: 'text', text: `Unable to start upload session: ${error.message}` }] };
      }
    }
  },
  {
    name: "onedrive-upload-session-chunk",
    description: "Append one sequential Base64 chunk to a resumable OneDrive upload session.",
    inputSchema: {
      type: "object",
      properties: {
        uploadId: { type: "string", description: "uploadId returned by onedrive-upload-session-start" },
        offset: { type: "integer", minimum: 0, description: "Byte offset; must equal bytesReceived from the previous response" },
        chunkBase64: { type: "string", description: "Standard Base64 bytes for this chunk" }
      },
      required: ["uploadId", "offset", "chunkBase64"]
    },
    handler: async (args) => {
      try {
        const result = await appendUploadSession(args);
        return {
          content: [{ type: 'text', text: `Received ${result.bytesReceived}/${result.totalBytes} bytes.` }],
          structuredContent: result
        };
      } catch (error) {
        return { content: [{ type: 'text', text: `Unable to append upload chunk: ${error.message}` }] };
      }
    }
  },
  {
    name: "onedrive-upload-session-complete",
    description: "Upload a completed resumable file to OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        uploadId: { type: "string", description: "uploadId returned by onedrive-upload-session-start" }
      },
      required: ["uploadId"]
    },
    handler: async (args) => {
      try {
        const result = await finishUploadSession(args);
        return {
          content: [{ type: 'text', text: `Successfully uploaded "${result.name}" (${result.size} bytes).` }],
          structuredContent: result
        };
      } catch (error) {
        return { content: [{ type: 'text', text: `Unable to complete upload session: ${error.message}` }] };
      }
    }
  },
  {
    name: "onedrive-upload-session-abort",
    description: "Discard a resumable OneDrive upload session and its temporary bytes.",
    inputSchema: {
      type: "object",
      properties: {
        uploadId: { type: "string", description: "uploadId returned by onedrive-upload-session-start" }
      },
      required: ["uploadId"]
    },
    handler: async (args) => {
      try {
        const result = await abortUploadSession(args);
        return {
          content: [{ type: 'text', text: 'Upload session aborted.' }],
          structuredContent: result
        };
      } catch (error) {
        return { content: [{ type: 'text', text: `Unable to abort upload session: ${error.message}` }] };
      }
    }
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
    description: "Create a sharing link for a file or folder in OneDrive",
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
          description: "Link type: 'view' (default), 'edit', or 'embed'",
          enum: ["view", "edit", "embed"]
        },
        scope: {
          type: "string",
          description: "Link scope: 'anonymous' (default) or 'organization'",
          enum: ["anonymous", "organization"]
        }
      },
      required: []
    },
    handler: handleShare
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
    name: "onedrive-delete",
    description: "Delete a file or folder from OneDrive",
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
  startUploadSession,
  appendUploadSession,
  finishUploadSession,
  abortUploadSession,
  handleImportUrl,
  handleShare,
  handleMoveItem,
  handleCreateFolder,
  handleDeleteItem
};
