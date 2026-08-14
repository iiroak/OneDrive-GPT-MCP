const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

// The spawned server creates its staging directory on startup; keep that out of
// the shared temp dir.
const STORE_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'm365-protocol-'));

afterAll(() => {
  fs.rmSync(STORE_DIR, { recursive: true, force: true });
});

/**
 * Drive the real server over stdio and collect responses.
 *
 * Unit tests cannot prove that resources/read is actually wired into the
 * request handler, only that the store works. This exercises the protocol
 * surface a client sees.
 *
 * @param {object[]} requests
 * @returns {Promise<object[]>}
 */
function callServer(requests) {
  return new Promise((resolve, reject) => {
    const child = spawn('node', [path.join(__dirname, '..', 'index.js')], {
      env: { ...process.env, USE_TEST_MODE: 'true', M365_FILE_STORE_DIR: STORE_DIR },
      stdio: ['pipe', 'pipe', 'pipe']
    });

    let stdout = '';
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`El servidor no respondio a tiempo. Salida: ${stdout}`));
    }, 20000);

    child.stdout.on('data', (chunk) => { stdout += chunk.toString(); });
    child.stderr.on('data', () => {}); // The server logs diagnostics to stderr.

    child.on('error', (error) => {
      clearTimeout(timer);
      reject(error);
    });

    child.on('close', () => {
      clearTimeout(timer);
      const messages = stdout
        .split('\n')
        .filter((line) => line.trim().startsWith('{'))
        .map((line) => {
          try {
            return JSON.parse(line);
          } catch {
            return null;
          }
        })
        .filter(Boolean);
      resolve(messages);
    });

    for (const request of requests) {
      child.stdin.write(`${JSON.stringify(request)}\n`);
    }
    // Give the server a moment to answer before closing the pipe.
    setTimeout(() => child.stdin.end(), 2500);
  });
}

describe('MCP protocol surface', () => {
  let messages;

  beforeAll(async () => {
    messages = await callServer([
      {
        jsonrpc: '2.0',
        id: 1,
        method: 'initialize',
        params: {
          protocolVersion: '2025-11-25',
          capabilities: {},
          clientInfo: { name: 'test', version: '1.0.0' }
        }
      },
      { jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} },
      { jsonrpc: '2.0', id: 3, method: 'resources/list', params: {} },
      { jsonrpc: '2.0', id: 4, method: 'resources/read', params: { uri: 'm365-file:///' + 'a'.repeat(32) } },
      { jsonrpc: '2.0', id: 5, method: 'resources/read', params: { uri: 'file:///etc/passwd' } }
    ]);
  }, 30000);

  /**
   * @param {number} id
   * @returns {object|undefined}
   */
  const byId = (id) => messages.find((message) => message.id === id);

  test('advertises the resources capability during initialize', () => {
    const result = byId(1)?.result;
    expect(result).toBeDefined();
    // Without this, a client never attempts resources/read and the
    // ResourceLinks the file tools return are dead ends.
    expect(result.capabilities.resources).toBeDefined();
    expect(result.capabilities.tools).toBeDefined();
  });

  test('exposes the file reading tools with output schemas', () => {
    const tools = byId(2)?.result?.tools || [];
    const names = tools.map((tool) => tool.name);

    expect(names).toContain('onedrive-read-file');
    expect(names).toContain('onedrive-export-file');
    expect(names).toContain('onedrive-download');

    const read = tools.find((tool) => tool.name === 'onedrive-read-file');
    expect(read.outputSchema.properties.message).toBeDefined();
    expect(read.outputSchema.properties.data).toBeDefined();
  });

  test('answers resources/list', () => {
    expect(Array.isArray(byId(3)?.result?.resources)).toBe(true);
  });

  test('resources/read rejects an unknown handle instead of hanging', () => {
    const response = byId(4);
    expect(response).toBeDefined();
    // The error must sit at the top level. Nested inside `result` a client
    // reads it as a successful read.
    expect(response.result).toBeUndefined();
    expect(response.error?.code).not.toBe(-32601);
    expect(response.error?.message).toMatch(/does not exist|expired/);
  });

  test('resources/read refuses a foreign URI scheme', () => {
    const response = byId(5);
    expect(response.result).toBeUndefined();
    expect(response.error?.code).toBe(-32602);
    expect(response.error?.message).toMatch(/m365-file/);
  });
});
