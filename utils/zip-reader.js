/**
 * Minimal ZIP reader for OOXML containers (.docx, .pptx, .xlsx).
 *
 * These formats are ZIP archives, and Node ships the only primitive that is
 * actually hard to write by hand (`zlib.inflateRawSync`). Reading the central
 * directory ourselves avoids adding a dependency just to open a Word file.
 *
 * Scope is deliberately narrow: stored (method 0) and deflate (method 8)
 * entries in a non-ZIP64 archive, which is what Office produces for documents
 * of the size this server is willing to stage. Anything else throws a clear
 * error rather than returning silently wrong bytes.
 */
const zlib = require('zlib');

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const LOCAL_SIGNATURE = 0x04034b50;
const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;

// EOCD is 22 bytes plus an optional comment capped at 65535 by the format.
const MAX_EOCD_SCAN = 22 + 0xffff;

class ZipError extends Error {
  constructor(message) {
    super(message);
    this.name = 'ZipError';
  }
}

/**
 * Locate the End Of Central Directory record by scanning backwards.
 * @param {Buffer} buffer
 * @returns {number} offset of the EOCD signature
 */
function findEndOfCentralDirectory(buffer) {
  const floor = Math.max(0, buffer.length - MAX_EOCD_SCAN);
  for (let offset = buffer.length - 22; offset >= floor; offset -= 1) {
    if (buffer.readUInt32LE(offset) === EOCD_SIGNATURE) return offset;
  }
  throw new ZipError('El archivo no es un contenedor ZIP valido');
}

/**
 * Parse the central directory into an index of entries.
 * @param {Buffer} buffer
 * @returns {Map<string, {compressionMethod: number, compressedSize: number, uncompressedSize: number, localHeaderOffset: number}>}
 */
function readCentralDirectory(buffer) {
  const eocd = findEndOfCentralDirectory(buffer);

  // A ZIP64 locator immediately precedes the EOCD when the archive is ZIP64.
  if (eocd >= 20 && buffer.readUInt32LE(eocd - 20) === ZIP64_LOCATOR_SIGNATURE) {
    throw new ZipError('Los contenedores ZIP64 no estan soportados');
  }

  const entryCount = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);

  const entries = new Map();
  for (let index = 0; index < entryCount; index += 1) {
    if (offset + 46 > buffer.length || buffer.readUInt32LE(offset) !== CENTRAL_SIGNATURE) {
      throw new ZipError('El directorio central del ZIP esta corrupto');
    }

    const compressionMethod = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const uncompressedSize = buffer.readUInt32LE(offset + 24);
    const nameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.toString('utf8', offset + 46, offset + 46 + nameLength);

    entries.set(name, { compressionMethod, compressedSize, uncompressedSize, localHeaderOffset });
    offset += 46 + nameLength + extraLength + commentLength;
  }

  return entries;
}

/**
 * Read and decompress a single entry.
 * @param {Buffer} buffer
 * @param {{compressionMethod: number, compressedSize: number, localHeaderOffset: number}} entry
 * @returns {Buffer}
 */
function readEntry(buffer, entry) {
  const start = entry.localHeaderOffset;
  if (start + 30 > buffer.length || buffer.readUInt32LE(start) !== LOCAL_SIGNATURE) {
    throw new ZipError('La cabecera local del ZIP esta corrupta');
  }

  // Local header name/extra lengths can differ from the central directory's.
  const nameLength = buffer.readUInt16LE(start + 26);
  const extraLength = buffer.readUInt16LE(start + 28);
  const dataStart = start + 30 + nameLength + extraLength;
  const data = buffer.subarray(dataStart, dataStart + entry.compressedSize);

  if (entry.compressionMethod === 0) return Buffer.from(data);
  if (entry.compressionMethod === 8) {
    try {
      return zlib.inflateRawSync(data);
    } catch (error) {
      throw new ZipError(`No se pudo descomprimir una entrada del ZIP: ${error.message}`);
    }
  }
  throw new ZipError(`Metodo de compresion ZIP no soportado: ${entry.compressionMethod}`);
}

/**
 * Open an OOXML container.
 * @param {Buffer} buffer
 * @returns {{names: string[], has: (name: string) => boolean, read: (name: string) => Buffer|null}}
 */
function openZip(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 22) {
    throw new ZipError('El archivo es demasiado pequeno para ser un ZIP');
  }
  // Local file header magic: "PK\x03\x04".
  if (buffer.readUInt32LE(0) !== LOCAL_SIGNATURE) {
    throw new ZipError('El archivo no comienza con la firma de un ZIP');
  }

  const entries = readCentralDirectory(buffer);

  return {
    names: [...entries.keys()],
    has: (name) => entries.has(name),
    read: (name) => {
      const entry = entries.get(name);
      return entry ? readEntry(buffer, entry) : null;
    }
  };
}

module.exports = { openZip, ZipError };
