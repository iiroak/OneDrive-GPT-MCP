const crypto = require('crypto');
const fs = require('fs').promises;
const os = require('os');
const path = require('path');
const config = require('../config');
const { uploadFilePath } = require('./upload-file');
const { decodeUploadContent } = require('./upload-content');

const sessions = new Map();

function validateDestinationPath(value) {
  const raw = String(value || '').trim();
  if (!raw || raw.includes('\\') || raw.split('/').some(part => part === '..')) {
    throw new Error('Invalid OneDrive destination path.');
  }
  return raw.replace(/^\/+|\/+$/g, '');
}

function getSession(uploadId) {
  const session = sessions.get(String(uploadId || ''));
  if (!session) throw new Error('Upload session does not exist or has expired.');
  if (Date.now() - session.updatedAt > config.ONEDRIVE_UPLOAD_SESSION_TTL_MS) {
    sessions.delete(session.id);
    fs.rm(session.directory, { recursive: true, force: true }).catch(() => {});
    throw new Error('Upload session does not exist or has expired.');
  }
  session.updatedAt = Date.now();
  return session;
}

async function startUploadSession(args) {
  const destinationPath = validateDestinationPath(args.path);
  const totalBytes = Number(args.totalBytes);
  if (!Number.isSafeInteger(totalBytes) || totalBytes <= 0) {
    throw new Error('totalBytes must be a positive safe integer.');
  }
  if (totalBytes > config.ONEDRIVE_UPLOAD_SESSION_MAX_BYTES) {
    throw new Error('totalBytes exceeds the configured upload limit.');
  }
  const conflictBehavior = args.conflictBehavior || 'rename';
  if (!['rename', 'replace', 'fail'].includes(conflictBehavior)) {
    throw new Error('conflictBehavior must be rename, replace or fail.');
  }

  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-upload-'));
  const id = crypto.randomBytes(18).toString('base64url');
  const filePath = path.join(directory, 'payload.bin');
  await fs.writeFile(filePath, '', { mode: 0o600 });
  sessions.set(id, {
    id,
    directory,
    filePath,
    destinationPath,
    totalBytes,
    bytesReceived: 0,
    conflictBehavior,
    busy: false,
    updatedAt: Date.now()
  });

  return {
    uploadId: id,
    chunkBytes: config.ONEDRIVE_UPLOAD_SESSION_CHUNK_BYTES,
    totalBytes,
    bytesReceived: 0
  };
}

async function appendUploadSession(args) {
  const session = getSession(args.uploadId);
  const offset = Number(args.offset);
  if (!Number.isSafeInteger(offset) || offset !== session.bytesReceived) {
    throw new Error(`offset must equal bytesReceived (${session.bytesReceived}).`);
  }

  const chunk = decodeUploadContent({ contentBase64: args.chunkBase64 });
  if (chunk.length > config.ONEDRIVE_UPLOAD_SESSION_CHUNK_BYTES) {
    throw new Error('chunkBase64 exceeds the configured chunk size.');
  }
  if (session.bytesReceived + chunk.length > session.totalBytes) {
    throw new Error('chunkBase64 exceeds totalBytes.');
  }

  if (session.busy) throw new Error('Upload session is busy; retry this chunk.');
  session.busy = true;
  try {
    await fs.appendFile(session.filePath, chunk, { mode: 0o600 });
    session.bytesReceived += chunk.length;
  } finally {
    session.busy = false;
  }
  return {
    uploadId: session.id,
    totalBytes: session.totalBytes,
    bytesReceived: session.bytesReceived,
    complete: session.bytesReceived === session.totalBytes
  };
}

async function finishUploadSession(args) {
  const session = getSession(args.uploadId);
  if (session.busy) throw new Error('Upload session is busy; retry completion.');
  if (session.bytesReceived !== session.totalBytes) {
    throw new Error(`Upload is incomplete: ${session.bytesReceived}/${session.totalBytes} bytes.`);
  }

  let uploaded;
  try {
    uploaded = await uploadFilePath(
      session.filePath,
      session.destinationPath,
      session.conflictBehavior
    );
  } catch (error) {
    session.updatedAt = Date.now();
    throw error;
  }

  try {
    sessions.delete(session.id);
    await fs.rm(session.directory, { recursive: true, force: true });
  } catch (error) {
    console.error(`Unable to clean completed upload session ${session.id}: ${error.message}`);
  }
  return uploaded;
}

async function abortUploadSession(args) {
  const session = getSession(args.uploadId);
  if (session.busy) throw new Error('Upload session is busy; retry abort.');
  sessions.delete(session.id);
  await fs.rm(session.directory, { recursive: true, force: true });
  return { uploadId: session.id, aborted: true };
}

const cleanupTimer = setInterval(() => {
  const now = Date.now();
  for (const session of sessions.values()) {
    if (now - session.updatedAt <= config.ONEDRIVE_UPLOAD_SESSION_TTL_MS) continue;
    sessions.delete(session.id);
    fs.rm(session.directory, { recursive: true, force: true }).catch(() => {});
  }
}, Math.min(config.ONEDRIVE_UPLOAD_SESSION_TTL_MS, 60 * 1000));
cleanupTimer.unref();

module.exports = {
  startUploadSession,
  appendUploadSession,
  finishUploadSession,
  abortUploadSession,
  validateDestinationPath,
  sessions
};
