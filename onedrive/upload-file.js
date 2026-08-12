const fs = require('fs').promises;
const https = require('https');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

const CHUNK_SIZE = 320 * 1024 * 10;

async function uploadFilePath(filePath, destinationPath, conflictBehavior = 'rename') {
  if (!filePath || !destinationPath) throw new Error('File path and destination path are required.');

  const accessToken = await ensureAuthenticated();
  const stats = await fs.stat(filePath);
  const normalizedPath = destinationPath.replace(/^\/+|\/+$/g, '');
  const sessionEndpoint = `me/drive/root:/${normalizedPath}:/createUploadSession`;
  const sessionResponse = await callGraphAPI(accessToken, 'POST', sessionEndpoint, {
    item: { '@microsoft.graph.conflictBehavior': conflictBehavior }
  });

  if (!sessionResponse || !sessionResponse.uploadUrl) {
    throw new Error('Failed to create upload session.');
  }

  if (stats.size === 0) {
    const response = await callGraphAPI(
      accessToken,
      'PUT',
      `me/drive/root:/${normalizedPath}:/content`,
      Buffer.alloc(0)
    );
    if (!response || !response.id) throw new Error('Empty file upload failed.');
    return response;
  }

  const file = await fs.open(filePath, 'r');
  let offset = 0;
  let response = null;
  try {
    while (offset < stats.size) {
      const buffer = Buffer.alloc(Math.min(CHUNK_SIZE, stats.size - offset));
      const read = await file.read(buffer, 0, buffer.length, offset);
      if (!read.bytesRead) throw new Error('Unexpected end of imported file.');
      const chunk = buffer.subarray(0, read.bytesRead);
      response = await uploadChunk(
        sessionResponse.uploadUrl,
        chunk,
        offset,
        offset + read.bytesRead - 1,
        stats.size
      );
      if (response.error) throw new Error(response.error);
      offset += read.bytesRead;
    }
  } finally {
    await file.close();
  }

  if (!response || !response.id) throw new Error('Upload completed without file information.');
  return response;
}

function uploadChunk(uploadUrl, chunk, start, end, totalSize) {
  return new Promise((resolve, reject) => {
    const request = https.request(uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': chunk.length,
        'Content-Range': `bytes ${start}-${end}/${totalSize}`
      }
    }, response => {
      let body = '';
      response.on('data', data => { body += data; });
      response.on('end', () => {
        if (response.statusCode < 200 || response.statusCode >= 300) {
          resolve({ error: `Status ${response.statusCode}: ${body.slice(0, 240)}` });
          return;
        }
        try {
          resolve(JSON.parse(body || '{}'));
        } catch (error) {
          resolve({ error: `Invalid upload response: ${error.message}` });
        }
      });
    });
    request.on('error', reject);
    request.end(chunk);
  });
}

module.exports = { uploadFilePath, uploadChunk, CHUNK_SIZE };
