const fs = require('fs').promises;
const os = require('os');
const path = require('path');
const config = require('../config');
const { downloadToFile } = require('./import-url');

async function withChatGPTFile(file, handler) {
  if (!file || typeof file !== 'object'
    || typeof file.file_id !== 'string' || !file.file_id
    || typeof file.download_url !== 'string' || !file.download_url) {
    throw new Error('The ChatGPT file must include file_id and download_url.');
  }

  const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-chatgpt-'));
  const temporaryFile = path.join(temporaryDirectory, 'payload.bin');
  try {
    const downloaded = await downloadToFile(
      file.download_url,
      temporaryFile,
      config.ONEDRIVE_IMPORT_MAX_BYTES,
      0,
      config.CHATGPT_FILE_ALLOWED_HOSTS
    );
    return await handler(temporaryFile, downloaded);
  } finally {
    await fs.rm(temporaryDirectory, { recursive: true, force: true });
  }
}

module.exports = { withChatGPTFile };
