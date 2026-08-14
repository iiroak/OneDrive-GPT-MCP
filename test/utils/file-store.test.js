const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

const { FileStore, FileStoreError, safeFilename, storageSuffix } = require('../../utils/file-store');

describe('FileStore', () => {
  let root;
  let store;

  beforeEach(() => {
    root = fs.mkdtempSync(path.join(os.tmpdir(), 'file-store-test-'));
    store = new FileStore({ root, retentionMs: 60000, maxFileBytes: 1024, maxTotalBytes: 4096 });
  });

  afterEach(() => {
    fs.rmSync(root, { recursive: true, force: true });
  });

  test('stages bytes under an opaque name distinct from the display filename', () => {
    const entry = store.put(Buffer.from('hola'), { filename: 'Instrucciones Resumen.md', mimeType: 'text/markdown' });

    expect(entry.filename).toBe('Instrucciones Resumen.md');
    expect(path.basename(entry.storagePath)).not.toBe(entry.filename);
    expect(path.basename(entry.storagePath)).toBe(`${entry.file_id}.md`);
    expect(entry.sha256).toBe(crypto.createHash('sha256').update('hola').digest('hex'));
    expect(store.read(entry.file_id).toString()).toBe('hola');
  });

  test('stores staged files with owner-only permissions', () => {
    const entry = store.put(Buffer.from('secreto'), { filename: 'a.txt', mimeType: 'text/plain' });
    const mode = fs.statSync(entry.storagePath).mode & 0o777;
    expect(mode).toBe(0o600);
  });

  test('rejects a size mismatch, catching a truncated transfer', () => {
    expect(() => store.put(Buffer.from('12345'), {
      filename: 'a.txt',
      mimeType: 'text/plain',
      expectedSize: 99
    })).toThrow(FileStoreError);
  });

  test('rejects a digest mismatch', () => {
    expect(() => store.put(Buffer.from('12345'), {
      filename: 'a.txt',
      mimeType: 'text/plain',
      expectedSha256: 'deadbeef'
    })).toThrow(/digest/i);
  });

  test('enforces the per-file ceiling', () => {
    expect(() => store.put(Buffer.alloc(2048), { filename: 'big.bin', mimeType: 'application/octet-stream' }))
      .toThrow(/limite permitido/);
  });

  test('enforces the aggregate quota', () => {
    for (let i = 0; i < 4; i += 1) {
      store.put(Buffer.alloc(1024), { filename: `f${i}.bin`, mimeType: 'application/octet-stream' });
    }
    expect(() => store.put(Buffer.alloc(1024), { filename: 'overflow.bin', mimeType: 'application/octet-stream' }))
      .toThrow(/cuota/);
  });

  test('expired handles fail closed and their bytes are removed', () => {
    const shortLived = new FileStore({ root, retentionMs: -1 });
    const entry = shortLived.put(Buffer.from('x'), { filename: 'a.txt', mimeType: 'text/plain' });

    expect(shortLived.get(entry.file_id)).toBeNull();
    expect(fs.existsSync(entry.storagePath)).toBe(false);
    expect(() => shortLived.read(entry.file_id)).toThrow(/expiro/);
  });

  test('rejects malformed file ids instead of touching the filesystem', () => {
    expect(store.get('../../etc/passwd')).toBeNull();
    expect(store.get('')).toBeNull();
    expect(store.get('abc')).toBeNull();
  });

  test('metadata never leaks the on-disk path', () => {
    const entry = store.put(Buffer.from('x'), { filename: 'a.txt', mimeType: 'text/plain' });
    const metadata = FileStore.metadata(entry);

    expect(metadata.storagePath).toBeUndefined();
    expect(JSON.stringify(metadata)).not.toContain(root);
    expect(metadata.resource_uri).toBe(`m365-file:///${entry.file_id}`);
  });

  test('parses only its own resource URIs', () => {
    const fileId = 'a'.repeat(32);
    expect(FileStore.parseResourceUri(`m365-file:///${fileId}`)).toBe(fileId);
    expect(FileStore.parseResourceUri('file:///etc/passwd')).toBeNull();
    expect(FileStore.parseResourceUri('m365-file:///nope')).toBeNull();
    expect(FileStore.parseResourceUri('')).toBeNull();
  });

  test('retain extends the lifetime of a live handle', () => {
    const entry = store.put(Buffer.from('x'), { filename: 'a.txt', mimeType: 'text/plain' });
    const before = entry.expiresAt;
    store.retain(entry.file_id, 600000);
    expect(store.get(entry.file_id).expiresAt).toBeGreaterThan(before);
  });

  test('leaves no .part file behind after a successful write', () => {
    store.put(Buffer.from('x'), { filename: 'a.txt', mimeType: 'text/plain' });
    expect(fs.readdirSync(root).filter((name) => name.endsWith('.part'))).toEqual([]);
  });
});

describe('safeFilename', () => {
  test('strips directory traversal', () => {
    expect(safeFilename('../../etc/passwd')).toBe('passwd');
    expect(safeFilename('/absolute/path/file.txt')).toBe('file.txt');
  });

  test('keeps accented characters used in Spanish filenames', () => {
    expect(safeFilename('Instrucciones Resumen ñ.md')).toBe('Instrucciones Resumen ñ.md');
  });

  test('replaces shell and control metacharacters', () => {
    expect(safeFilename('a;rm -rf $HOME.txt')).not.toContain(';');
    expect(safeFilename('a;rm -rf $HOME.txt')).not.toContain('$');
  });

  test('falls back for empty and dot-only names', () => {
    expect(safeFilename('')).toBe('download.bin');
    expect(safeFilename('..')).toBe('download.bin');
  });
});

describe('storageSuffix', () => {
  test('keeps a normal extension', () => {
    expect(storageSuffix('a.PDF')).toBe('.pdf');
  });

  test('falls back to .bin for missing or hostile extensions', () => {
    expect(storageSuffix('noext')).toBe('.bin');
    expect(storageSuffix('a.this-extension-is-far-too-long')).toBe('.bin');
    expect(storageSuffix('a.tar gz')).toBe('.bin');
  });
});
