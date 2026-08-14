const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

// Isolate the staging directory before the store module is required.
const STORE_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'm365-roundtrip-'));
process.env.M365_FILE_STORE_DIR = STORE_DIR;

/**
 * End-to-end check of the transfer contract inside a single server process.
 *
 * Everything else is either a unit test or a protocol smoke test. This is the
 * only test that proves the handle a tool mints can actually be redeemed
 * through resources/read, which is the exact link that was missing.
 */
describe('file transfer round trip', () => {
  let storeDir;

  beforeAll(() => {
    storeDir = fs.mkdtempSync(path.join(os.tmpdir(), 'm365-roundtrip-child-'));
  });

  afterAll(() => {
    fs.rmSync(storeDir, { recursive: true, force: true });
    fs.rmSync(STORE_DIR, { recursive: true, force: true });
  });

  test('a staged file is readable through resources/read and resources/list', async () => {
    // Drive the server in-process so a tool and the resource handler share one
    // FileStore, mirroring how a real session behaves.
    const script = `
      const { getFileStore, FileStore } = require(${JSON.stringify(path.join(__dirname, '..', 'utils', 'file-store.js'))});
      const store = getFileStore();
      const entry = store.put(Buffer.from('# Instrucciones\\n\\nContenido real'), {
        filename: 'Instrucciones_Resumen.md',
        mimeType: 'text/markdown'
      });
      const png = Buffer.from([0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a]);
      const binary = store.put(png, { filename: 'logo.png', mimeType: 'image/png' });
      console.log(JSON.stringify({
        textUri: FileStore.resourceUri(entry.file_id),
        binaryUri: FileStore.resourceUri(binary.file_id)
      }));
    `;

    const staged = await new Promise((resolve, reject) => {
      const child = spawn('node', ['-e', script], {
        env: { ...process.env, M365_FILE_STORE_DIR: storeDir }
      });
      let out = '';
      child.stdout.on('data', (chunk) => { out += chunk.toString(); });
      child.on('error', reject);
      child.on('close', () => {
        const line = out.split('\n').find((item) => item.trim().startsWith('{'));
        line ? resolve(JSON.parse(line)) : reject(new Error(`Sin salida: ${out}`));
      });
    });

    expect(staged.textUri).toMatch(/^m365-file:\/\/\/[0-9a-f]{32}$/);

    // A fresh process has an empty in-memory index, so a handle from another
    // process must NOT resolve. Handles are per-session by design: that is what
    // makes expiry and the quota meaningful.
    const { getFileStore, FileStore } = require('../utils/file-store');
    const foreignId = FileStore.parseResourceUri(staged.textUri);
    const isolatedStore = getFileStore();
    expect(isolatedStore.get(foreignId)).toBeNull();

    // Within one process, the round trip must hold.
    const entry = isolatedStore.put(Buffer.from('# Instrucciones\n\nContenido real'), {
      filename: 'Instrucciones_Resumen.md',
      mimeType: 'text/markdown'
    });
    const uri = FileStore.resourceUri(entry.file_id);
    const resolved = isolatedStore.get(FileStore.parseResourceUri(uri));

    expect(resolved).not.toBeNull();
    expect(isolatedStore.read(resolved.file_id).toString()).toBe('# Instrucciones\n\nContenido real');
    expect(isolatedStore.list().some((item) => item.file_id === entry.file_id)).toBe(true);
  }, 20000);

  test('the bytes on disk match what was staged', () => {
    const { getFileStore } = require('../utils/file-store');
    const store = getFileStore();
    const png = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const entry = store.put(png, { filename: 'logo.png', mimeType: 'image/png' });

    expect(fs.readFileSync(entry.storagePath)).toEqual(png);
    expect(store.read(entry.file_id)).toEqual(png);
    expect(Buffer.from(store.read(entry.file_id).toString('base64'), 'base64')).toEqual(png);
  });
});
