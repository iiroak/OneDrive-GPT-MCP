/**
 * OneDrive simple upload functionality (files < 4MB)
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { decodeUploadContent } = require('./upload-content');
const { withChatGPTFile } = require('./chatgpt-file');
const fs = require('fs').promises;

/**
 * Simple upload handler (for files < 4MB)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpload(args) {
  const path = args.path;
  const conflictBehavior = args.conflictBehavior || 'rename'; // rename, replace, fail

  if (!path) {
    return {
      content: [{
        type: "text",
        text: "Path is required (e.g., '/Documents/myfile.txt')."
      }]
    };
  }

  if (args.file) {
    try {
      return await withChatGPTFile(args.file, async (filePath, downloaded) => {
        if (downloaded.bytes > config.ONEDRIVE_UPLOAD_THRESHOLD) {
          return {
            content: [{
              type: "text",
              text: `File is too large for simple upload (${formatSize(downloaded.bytes)}). Use onedrive-upload-large for files over 4MB.`
            }]
          };
        }
        const contentBuffer = await fs.readFile(filePath);
        const response = await uploadBuffer(path, contentBuffer, conflictBehavior);
        return uploadResponse(response);
      });
    } catch (error) {
      return uploadError(error);
    }
  }

  let contentBuffer;
  try {
    contentBuffer = decodeUploadContent(args);
  } catch (error) {
    return { content: [{ type: "text", text: error.message }] };
  }

  // Check size - this is for simple upload only.
  const contentSize = contentBuffer.length;
  if (contentSize > config.ONEDRIVE_UPLOAD_THRESHOLD) {
    return {
      content: [{
        type: "text",
        text: `File is too large for simple upload (${formatSize(contentSize)}). Use onedrive-upload-large for files over 4MB.`
      }]
    };
  }

  try {
    return uploadResponse(await uploadBuffer(path, contentBuffer, conflictBehavior));
  } catch (error) {
    return uploadError(error);
  }
}

async function uploadBuffer(path, contentBuffer, conflictBehavior) {
  const accessToken = await ensureAuthenticated();
  const normalizedPath = path.replace(/^\/+|\/+$/g, '');
  const endpoint = `me/drive/root:/${normalizedPath}:/content`;
  const queryParams = {
    '@microsoft.graph.conflictBehavior': conflictBehavior
  };
  const response = await callGraphAPI(accessToken, 'PUT', endpoint, contentBuffer, queryParams);
  if (!response || !response.id) throw new Error('Upload failed - no response from server.');
  return response;
}

function uploadResponse(response) {
  return {
    content: [{
      type: "text",
      text: `Successfully uploaded "${response.name}" (${formatSize(response.size)})\n\nID: ${response.id}\nWeb URL: ${response.webUrl}`
    }]
  };
}

function uploadError(error) {
  if (error.message === 'Authentication required') {
    return {
      content: [{
        type: "text",
        text: "Authentication required. Complete the MCP OAuth flow first."
      }]
    };
  }
  return { content: [{ type: "text", text: `Error uploading file: ${error.message}` }] };
}

/**
 * Format file size to human-readable string
 */
function formatSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = handleUpload;
