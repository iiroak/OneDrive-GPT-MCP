const fs = require('fs').promises;
const fsc = require('fs');
const os = require('os');
const path = require('path');
const https = require('https');
const { Transform } = require('stream');
const { pipeline } = require('stream/promises');
const config = require('../config');
const { uploadFilePath } = require('./upload-file');

function hostAllowed(hostname, patterns) {
  if (!hostname) return false;
  const host = hostname.toLowerCase().replace(/\.+$/, '');
  return patterns.some(pattern => {
    const value = String(pattern).toLowerCase().replace(/\.+$/, '');
    if (value.startsWith('*.')) return host.endsWith(value.slice(1)) && host !== value.slice(2);
    return host === value;
  });
}

function validateUrl(value, allowedHosts = config.ONEDRIVE_IMPORT_ALLOWED_HOSTS) {
  let parsed;
  try {
    parsed = new URL(value);
  } catch (error) {
    throw new Error('sourceUrl must be a valid URL.');
  }
  if (parsed.protocol !== 'https:' || parsed.username || parsed.password || !parsed.hostname) {
    throw new Error('sourceUrl must be an HTTPS URL without embedded credentials.');
  }
  if (!hostAllowed(parsed.hostname, allowedHosts)) {
    throw new Error('sourceUrl host is not allowed for server-side import.');
  }
  return parsed;
}

function downloadToFile(
  sourceUrl,
  destination,
  maxBytes,
  redirects = 0,
  allowedHosts = config.ONEDRIVE_IMPORT_ALLOWED_HOSTS
) {
  const parsed = validateUrl(sourceUrl, allowedHosts);
  return new Promise((resolve, reject) => {
    const request = https.get(parsed, response => {
      if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        response.resume();
        if (redirects >= 4) {
          reject(new Error('Too many redirects while importing file.'));
          return;
        }
        let next;
        try {
          next = new URL(response.headers.location, parsed).toString();
        } catch (error) {
          reject(new Error('Invalid redirect while importing file.'));
          return;
        }
        downloadToFile(next, destination, maxBytes, redirects + 1, allowedHosts).then(resolve, reject);
        return;
      }
      if (response.statusCode < 200 || response.statusCode >= 300) {
        response.resume();
        reject(new Error(`Source download failed with status ${response.statusCode}.`));
        return;
      }

      const declared = Number.parseInt(response.headers['content-length'] || '', 10);
      if (Number.isFinite(declared) && declared > maxBytes) {
        response.resume();
        reject(new Error('Source file exceeds the configured import limit.'));
        return;
      }

      let bytes = 0;
      const counter = new Transform({
        transform(chunk, encoding, callback) {
          bytes += chunk.length;
          if (bytes > maxBytes) {
            callback(new Error('Source file exceeds the configured import limit.'));
            return;
          }
          callback(null, chunk);
        }
      });
      const output = fsc.createWriteStream(destination, { mode: 0o600 });
      pipeline(response, counter, output)
        .then(() => resolve({
          bytes,
          contentType: String(response.headers['content-type'] || '').split(';')[0] || 'application/octet-stream'
        }))
        .catch(reject);
    });
    request.on('error', reject);
  });
}

async function importUrlToPath(sourceUrl, destinationPath, conflictBehavior = 'rename') {
  if (!sourceUrl) {
    throw new Error('sourceUrl is required.');
  }
  if (!destinationPath) {
    throw new Error('Destination path is required.');
  }

  const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-import-'));
  const temporaryFile = path.join(temporaryDirectory, 'payload.bin');
  try {
    const imported = await downloadToFile(
      sourceUrl,
      temporaryFile,
      config.ONEDRIVE_IMPORT_MAX_BYTES
    );
    const uploaded = await uploadFilePath(temporaryFile, destinationPath, conflictBehavior);
    return {
      id: uploaded.id,
      name: uploaded.name,
      size: uploaded.size || imported.bytes,
      webUrl: uploaded.webUrl,
      imported: true
    };
  } finally {
    await fs.rm(temporaryDirectory, { recursive: true, force: true });
  }
}

async function handleImportUrl(args) {
  try {
    const result = await importUrlToPath(
      args.sourceUrl,
      args.path,
      args.conflictBehavior || 'rename'
    );
    return {
      content: [{
        type: 'text',
        text: `Successfully imported "${result.name}" (${formatSize(result.size)})\n\nID: ${result.id}${result.webUrl ? `\nWeb URL: ${result.webUrl}` : ''}`
      }],
      structuredContent: result
    };
  } catch (error) {
    return { content: [{ type: 'text', text: `Error importing file: ${error.message}` }] };
  }
}

function formatSize(bytes) {
  if (!bytes) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const index = Math.floor(Math.log(bytes) / Math.log(1024));
  return `${parseFloat((bytes / Math.pow(1024, index)).toFixed(2))} ${units[index]}`;
}

module.exports = handleImportUrl;
module.exports.downloadToFile = downloadToFile;
module.exports.hostAllowed = hostAllowed;
module.exports.importUrlToPath = importUrlToPath;
