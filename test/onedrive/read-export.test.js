const fs = require('fs');
const os = require('os');
const path = require('path');

// Point the staging store at a throwaway directory before anything requires it,
// so the suite does not leave files in the real temp dir.
const STORE_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'm365-read-export-'));
process.env.M365_FILE_STORE_DIR = STORE_DIR;

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { fetchBytes } = require('../../utils/binary-fetch');
const { FileStore, getFileStore } = require('../../utils/file-store');

jest.mock('../../utils/binary-fetch', () => {
  const actual = jest.requireActual('../../utils/binary-fetch');
  return { ...actual, fetchBytes: jest.fn() };
});

const handleReadFile = require('../../onedrive/read-file');
const handleExportFile = require('../../onedrive/export-file');

afterAll(() => {
  fs.rmSync(STORE_DIR, { recursive: true, force: true });
});

/**
 * @param {Buffer|string} content
 * @param {string} [contentType]
 */
function mockFetched(content, contentType = 'application/octet-stream') {
  const buffer = Buffer.isBuffer(content) ? content : Buffer.from(content);
  fetchBytes.mockResolvedValue({
    buffer,
    sha256: require('crypto').createHash('sha256').update(buffer).digest('hex'),
    bytes: buffer.length,
    contentType,
    filename: '',
    finalUrl: 'https://tenant.sharepoint.com/final'
  });
  return buffer;
}

describe('onedrive-read-file', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('access-token');
    // Isolate the store per test so quotas and ids never leak between them.
    getFileStore().entries.clear();
  });

  test('returns the file text, not a URL', async () => {
    const body = '# Resumen\n\nContenido';
    callGraphAPI.mockResolvedValue({
      id: 'item-1',
      name: 'Instrucciones_Resumen.md',
      size: Buffer.byteLength(body),
      file: { mimeType: 'text/markdown' },
      '@microsoft.graph.downloadUrl': 'https://tenant.sharepoint.com/tmp'
    });
    mockFetched(body);

    const result = await handleReadFile({ path: '/Instrucciones_Resumen.md' });

    expect(result.structuredContent.status).toBe('complete');
    expect(result.structuredContent.text).toContain('# Resumen');
    expect(result.content[0].text).toContain('Contenido');
    // The whole point: no pre-authenticated URL is handed to the model.
    expect(JSON.stringify(result)).not.toContain('sharepoint.com/tmp');
  });

  test('does not send the Graph bearer token to the CDN download URL', async () => {
    callGraphAPI.mockResolvedValue({
      id: 'item-1',
      name: 'a.md',
      size: 2,
      file: { mimeType: 'text/markdown' },
      '@microsoft.graph.downloadUrl': 'https://tenant.sharepoint.com/tmp'
    });
    mockFetched('hi');

    await handleReadFile({ itemId: 'item-1' });

    expect(fetchBytes).toHaveBeenCalledWith('https://tenant.sharepoint.com/tmp', expect.objectContaining({
      headers: {}
    }));
  });

  test('sends the bearer token on the graph /content fallback', async () => {
    callGraphAPI.mockResolvedValue({ id: 'item-2', name: 'a.md', size: 2, file: {} });
    mockFetched('hi');

    await handleReadFile({ itemId: 'item-2' });

    const [url, options] = fetchBytes.mock.calls[0];
    expect(url).toContain('graph.microsoft.com');
    expect(url).toContain('me/drive/items/item-2/content');
    expect(options.headers.Authorization).toBe('Bearer access-token');
  });

  test('emits a resource_link a client can redeem', async () => {
    callGraphAPI.mockResolvedValue({ id: 'item-1', name: 'a.md', size: 2, file: {} });
    mockFetched('hi');

    const result = await handleReadFile({ itemId: 'item-1' });
    const link = result.content.find((item) => item.type === 'resource_link');

    expect(link.uri).toMatch(/^m365-file:\/\/\/[0-9a-f]{32}$/);
    expect(link.uri).toBe(result.structuredContent.file.resource_uri);
    expect(getFileStore().get(result.structuredContent.file.file_id)).not.toBeNull();
  });

  test('re-reads a staged file by fileId without a second download', async () => {
    callGraphAPI.mockResolvedValue({ id: 'item-1', name: 'a.md', size: 5, file: {} });
    mockFetched('hello');

    const first = await handleReadFile({ itemId: 'item-1' });
    fetchBytes.mockClear();
    callGraphAPI.mockClear();

    const second = await handleReadFile({ fileId: first.structuredContent.file.file_id });

    expect(fetchBytes).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(second.structuredContent.text).toBe('hello');
  });

  test('reports a size mismatch instead of returning a truncated file', async () => {
    callGraphAPI.mockResolvedValue({ id: 'item-1', name: 'a.md', size: 999, file: {} });
    mockFetched('short');

    const result = await handleReadFile({ itemId: 'item-1' });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/no coincide/);
  });

  test('rejects an HTML login page served as a document', async () => {
    callGraphAPI.mockResolvedValue({ id: 'item-1', name: 'a.pdf', size: 30, file: {} });
    mockFetched('<!DOCTYPE html><html>Sign in</html>', 'text/html');

    const result = await handleReadFile({ itemId: 'item-1' });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/HTML/);
  });

  test('still loads a genuine .html file', async () => {
    const body = '<p>Contenido real</p>';
    callGraphAPI.mockResolvedValue({ id: 'item-1', name: 'page.html', size: body.length, file: { mimeType: 'text/html' } });
    mockFetched(body, 'text/html');

    const result = await handleReadFile({ itemId: 'item-1' });

    expect(result.isError).toBeUndefined();
    expect(result.structuredContent.text).toContain('Contenido real');
  });

  test('refuses a folder', async () => {
    callGraphAPI.mockResolvedValue({ id: 'folder-1', name: 'Docs', folder: { childCount: 3 } });

    const result = await handleReadFile({ path: '/Docs' });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/carpeta/);
  });

  test('reports a missing item', async () => {
    callGraphAPI.mockResolvedValue(null);
    const result = await handleReadFile({ path: '/nope.txt' });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/No se encontro/);
  });

  test('requires an identifier', async () => {
    const result = await handleReadFile({});
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/itemId, path o fileId/);
  });

  test('surfaces the auth prompt', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));
    const result = await handleReadFile({ itemId: 'item-1' });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/authenticate/i);
  });

  test('reports an expired fileId rather than failing opaquely', async () => {
    const result = await handleReadFile({ fileId: 'f'.repeat(32) });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/expiro/);
  });
});

describe('onedrive-export-file', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('access-token');
    getFileStore().entries.clear();
  });

  test('embeds binary content as a base64 blob', async () => {
    const png = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    callGraphAPI.mockResolvedValue({ id: 'item-1', name: 'logo.png', size: png.length, file: { mimeType: 'image/png' } });
    mockFetched(png, 'image/png');

    const result = await handleExportFile({ itemId: 'item-1' });
    const resource = result.content.find((item) => item.type === 'resource').resource;

    expect(result.structuredContent.transfer).toBe('embedded_resource');
    expect(result.structuredContent.encoding).toBe('base64');
    expect(Buffer.from(resource.blob, 'base64')).toEqual(png);
    expect(resource.text).toBeUndefined();
  });

  test('embeds text content as text, sparing the client a decode', async () => {
    callGraphAPI.mockResolvedValue({ id: 'item-1', name: 'a.md', size: 5, file: { mimeType: 'text/markdown' } });
    mockFetched('hello', 'text/markdown');

    const result = await handleExportFile({ itemId: 'item-1' });
    const resource = result.content.find((item) => item.type === 'resource').resource;

    expect(result.structuredContent.encoding).toBe('text');
    expect(resource.text).toBe('hello');
    expect(resource.blob).toBeUndefined();
  });

  test('refuses to inline an oversized file and names the alternative', async () => {
    const config = require('../../config');
    const stored = getFileStore().put(Buffer.from('x'), { filename: 'big.bin', mimeType: 'application/octet-stream' });
    stored.size_bytes = config.FILE_INLINE_MAX_BYTES + 1;

    const result = await handleExportFile({ fileId: stored.file_id });

    expect(result.isError).toBe(true);
    expect(result.structuredContent.status).toBe('too_large');
    expect(result.content[0].text).toMatch(/resources\/read|onedrive-read-file/);
  });

  test('requires an identifier', async () => {
    const result = await handleExportFile({});
    expect(result.isError).toBe(true);
  });
});

describe('resource URI round trip', () => {
  test('a handle minted by a tool resolves back to the same bytes', () => {
    const store = getFileStore();
    const entry = store.put(Buffer.from('payload'), { filename: 'a.txt', mimeType: 'text/plain' });
    const uri = FileStore.resourceUri(entry.file_id);

    expect(FileStore.parseResourceUri(uri)).toBe(entry.file_id);
    expect(store.read(FileStore.parseResourceUri(uri)).toString()).toBe('payload');
  });
});
