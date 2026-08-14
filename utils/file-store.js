/**
 * Private staging store for bytes fetched from Microsoft 365.
 *
 * Design mirrors the umayor-mcp FileStore invariants:
 * - Bytes land in a private directory (0700) with an opaque storage name that is
 *   NOT the user-visible filename, so a hostile filename cannot control the path.
 * - Writes are atomic: content goes to a `.part` file and is renamed into place
 *   only after the size and SHA-256 are verified.
 * - Every entry expires. Reads re-validate existence and expiry, so a stale
 *   file_id fails closed instead of serving a deleted or replaced file.
 * - A total-bytes quota bounds how much disk a session can consume.
 *
 * This module is the reason a tool can hand the model actual content instead of
 * a URL the model cannot dereference.
 */
const crypto = require('crypto');
const fs = require('fs');
const os = require('os');
const path = require('path');

const config = require('../config');

const FILE_ID_PATTERN = /^[0-9a-f]{32}$/;

/**
 * @typedef {object} StoredFile
 * @property {string} file_id      Opaque 32-hex identifier
 * @property {string} storagePath  Absolute path on disk (never exposed to the model)
 * @property {string} filename     Display filename, safe for a client to write
 * @property {string} mime_type    Resolved MIME type
 * @property {number} size_bytes   Verified byte length
 * @property {string} sha256       Verified digest
 * @property {number} createdAt    Epoch ms
 * @property {number} expiresAt    Epoch ms
 */

class FileStoreError extends Error {
  constructor(message) {
    super(message);
    this.name = 'FileStoreError';
  }
}

/**
 * Strip a filename down to something safe to hand back to a client.
 * Keeps the extension because downstream text extraction dispatches on it.
 * @param {string} value
 * @param {string} fallback
 * @returns {string}
 */
function safeFilename(value, fallback = 'download.bin') {
  const base = path.basename(String(value || '')).trim();
  if (!base || base === '.' || base === '..') return fallback;
  const cleaned = base.replace(/[^A-Za-z0-9._()\-\sáéíóúÁÉÍÓÚñÑüÜ]/g, '_').trim();
  return (cleaned || fallback).slice(0, 180);
}

/**
 * Derive a storage suffix we are willing to put on disk.
 * Rejects anything that is not a short alphanumeric extension.
 * @param {string} filename
 * @returns {string}
 */
function storageSuffix(filename) {
  const suffix = path.extname(filename || '').toLowerCase();
  if (!suffix || suffix.length > 16) return '.bin';
  if (!/^\.[a-z0-9-]+$/.test(suffix)) return '.bin';
  return suffix;
}

class FileStore {
  /**
   * @param {object} [options]
   * @param {string} [options.root]            Staging directory
   * @param {number} [options.retentionMs]     Time a staged file stays readable
   * @param {number} [options.maxFileBytes]    Per-file ceiling
   * @param {number} [options.maxTotalBytes]   Aggregate ceiling across live files
   */
  constructor(options = {}) {
    this.root = options.root
      || process.env.OUTLOOK_FILE_STORE_DIR
      || process.env.M365_FILE_STORE_DIR
      || path.join(config.DATA_DIR || os.tmpdir(), 'files');
    this.retentionMs = options.retentionMs ?? config.FILE_RETENTION_MS;
    this.maxFileBytes = options.maxFileBytes ?? config.FILE_MAX_BYTES;
    this.maxTotalBytes = options.maxTotalBytes ?? config.FILE_MAX_TOTAL_BYTES;

    /** @type {Map<string, StoredFile>} */
    this.entries = new Map();

    fs.mkdirSync(this.root, { recursive: true, mode: 0o700 });
    try {
      fs.chmodSync(this.root, 0o700);
    } catch {
      // Best effort on filesystems without POSIX modes.
    }
  }

  /**
   * Drop expired entries and their bytes.
   * @param {number} [now]
   * @returns {number} number of entries removed
   */
  cleanup(now = Date.now()) {
    let removed = 0;
    for (const [fileId, entry] of this.entries) {
      if (entry.expiresAt > now) continue;
      this.entries.delete(fileId);
      try {
        fs.rmSync(entry.storagePath, { force: true });
      } catch {
        // Losing the unlink race is not fatal; the entry is already gone.
      }
      removed += 1;
    }
    return removed;
  }

  /**
   * Total bytes currently held by live entries.
   * @returns {number}
   */
  totalBytes() {
    let total = 0;
    for (const entry of this.entries.values()) total += entry.size_bytes;
    return total;
  }

  /**
   * Persist a buffer and return its opaque handle.
   * @param {Buffer} buffer
   * @param {object} meta
   * @param {string} meta.filename
   * @param {string} meta.mimeType
   * @param {string} [meta.expectedSha256]  Verified when provided
   * @param {number} [meta.expectedSize]    Verified when provided
   * @returns {StoredFile}
   */
  put(buffer, meta) {
    if (!Buffer.isBuffer(buffer)) {
      throw new FileStoreError('El contenido a almacenar debe ser un Buffer');
    }
    if (buffer.length > this.maxFileBytes) {
      throw new FileStoreError(
        `El archivo supera el limite permitido (${buffer.length} > ${this.maxFileBytes} bytes)`
      );
    }

    this.cleanup();

    if (this.totalBytes() + buffer.length > this.maxTotalBytes) {
      throw new FileStoreError('La cuota de archivos temporales esta agotada');
    }

    const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');
    if (meta.expectedSha256 && meta.expectedSha256 !== sha256) {
      throw new FileStoreError('El digest del archivo descargado no coincide');
    }
    if (typeof meta.expectedSize === 'number' && meta.expectedSize !== buffer.length) {
      throw new FileStoreError(
        `El tamano descargado no coincide (${buffer.length} de ${meta.expectedSize} bytes)`
      );
    }

    const filename = safeFilename(meta.filename);
    const fileId = crypto.randomBytes(16).toString('hex');
    const storageName = `${fileId}${storageSuffix(filename)}`;
    const storagePath = path.join(this.root, storageName);
    const temporaryPath = path.join(this.root, `.${fileId}.part`);

    // Atomic publish: a reader never observes a partially written file.
    try {
      fs.writeFileSync(temporaryPath, buffer, { mode: 0o600, flag: 'wx' });
      fs.renameSync(temporaryPath, storagePath);
    } catch (error) {
      try {
        fs.rmSync(temporaryPath, { force: true });
      } catch {
        // Nothing else to do; the publish already failed.
      }
      throw new FileStoreError(`No se pudo almacenar el archivo: ${error.message}`);
    }

    const now = Date.now();
    /** @type {StoredFile} */
    const entry = {
      file_id: fileId,
      storagePath,
      filename,
      mime_type: meta.mimeType || 'application/octet-stream',
      size_bytes: buffer.length,
      sha256,
      createdAt: now,
      expiresAt: now + this.retentionMs
    };
    this.entries.set(fileId, entry);
    return entry;
  }

  /**
   * Look up a live entry. Returns null for unknown, expired, or vanished files.
   * @param {string} fileId
   * @returns {StoredFile|null}
   */
  get(fileId) {
    if (typeof fileId !== 'string' || !FILE_ID_PATTERN.test(fileId)) return null;
    const entry = this.entries.get(fileId);
    if (!entry) return null;
    if (entry.expiresAt <= Date.now()) {
      this.cleanup();
      return null;
    }
    if (!fs.existsSync(entry.storagePath)) {
      this.entries.delete(fileId);
      return null;
    }
    return entry;
  }

  /**
   * Read the bytes of a live entry.
   * @param {string} fileId
   * @returns {Buffer}
   */
  read(fileId) {
    const entry = this.get(fileId);
    if (!entry) throw new FileStoreError('El archivo no existe o ya expiro');
    return fs.readFileSync(entry.storagePath);
  }

  /**
   * Extend an entry's lifetime, for a slow consumer that still needs it.
   * @param {string} fileId
   * @param {number} extraMs
   * @returns {StoredFile}
   */
  retain(fileId, extraMs) {
    const entry = this.get(fileId);
    if (!entry) throw new FileStoreError('El archivo no existe o ya expiro');
    entry.expiresAt = Math.max(entry.expiresAt, Date.now() + Math.max(1, extraMs));
    return entry;
  }

  /**
   * Live entries, newest first. Used to answer resources/list.
   * @returns {StoredFile[]}
   */
  list() {
    this.cleanup();
    return [...this.entries.values()].sort((a, b) => b.createdAt - a.createdAt);
  }

  /**
   * The MCP resource URI for a staged file.
   * @param {string} fileId
   * @returns {string}
   */
  static resourceUri(fileId) {
    return `m365-file:///${fileId}`;
  }

  /**
   * Parse a resource URI back into a file_id. Returns null when it is not ours.
   * @param {string} uri
   * @returns {string|null}
   */
  static parseResourceUri(uri) {
    const match = /^m365-file:\/\/\/([0-9a-f]{32})$/.exec(String(uri || ''));
    return match ? match[1] : null;
  }

  /**
   * Client-facing metadata. Deliberately omits storagePath.
   * @param {StoredFile} entry
   * @returns {object}
   */
  static metadata(entry) {
    return {
      file_id: entry.file_id,
      filename: entry.filename,
      mime_type: entry.mime_type,
      size_bytes: entry.size_bytes,
      sha256: entry.sha256,
      resource_uri: FileStore.resourceUri(entry.file_id),
      created_at: new Date(entry.createdAt).toISOString(),
      expires_at: new Date(entry.expiresAt).toISOString()
    };
  }
}

// Process-wide store: resources/read must resolve handles minted by tools/call.
let sharedStore = null;

/**
 * @returns {FileStore}
 */
function getFileStore() {
  if (!sharedStore) sharedStore = new FileStore();
  return sharedStore;
}

module.exports = {
  FileStore,
  FileStoreError,
  getFileStore,
  safeFilename,
  storageSuffix
};
