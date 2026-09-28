const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const https = require('node:https');
const { afterEach, test } = require('node:test');

process.env.OUTLOOK_DATA_DIR = '/tmp/opencode/onedrive-regression-test-data';
process.env.MS_CLIENT_ID = 'local-test-client-id';
process.env.MS_CLIENT_SECRET = 'local-test-client-secret';

const graphApi = require('../utils/graph-api');
const operationMonitor = require('../utils/operation-monitor');
const { encodeId, encodePath, itemEndpoint } = require('../utils/onedrive-resolve');
const { TOOLS, requiredScope, toolPolicy } = require('../index');

const originalRequest = https.request;

afterEach(() => {
  https.request = originalRequest;
});

function mockGraphResponses(responses) {
  const requests = [];
  https.request = (url, options, onResponse) => {
    const request = new EventEmitter();
    let body = '';
    requests.push({ url: String(url), options, body: () => body });
    request.write = (chunk) => { body += Buffer.isBuffer(chunk) ? chunk.toString('utf8') : String(chunk); };
    request.setTimeout = () => {};
    request.end = () => {
      const result = responses.shift();
      if (!result) throw new Error('Unexpected Graph request in test');
      const response = new EventEmitter();
      response.statusCode = result.status || 200;
      response.headers = result.headers || {};
      process.nextTick(() => {
        onResponse(response);
        if (result.body !== undefined) {
          const data = typeof result.body === 'string' ? result.body : JSON.stringify(result.body);
          response.emit('data', Buffer.from(data));
        }
        response.emit('end');
      });
    };
    return request;
  };
  return requests;
}

test('OneDrive paths encode each segment and reject traversal', () => {
  assert.equal(encodePath('/My Folder/100% #1'), 'My%20Folder/100%25%20%231');
  assert.equal(itemEndpoint({ path: '/Documents/My File.txt' }), 'me/drive/root:/Documents/My%20File.txt');
  assert.equal(itemEndpoint({ path: '/' }), 'me/drive/root');
  assert.equal(itemEndpoint({ path: 'root' }), 'me/drive/root');
  assert.equal(itemEndpoint({ itemId: 'opaque/id' }), 'me/drive/items/opaque%2Fid');
  assert.equal(encodeId('permission/id'), 'permission%2Fid');
  assert.throws(() => encodePath('../Documents'), /Invalid OneDrive path/);
  assert.throws(() => encodePath('Documents\\..\\secret'), /Invalid OneDrive path/);
});

test('Graph requests reject origins and paths outside the configured v1 API prefix', async () => {
  const requests = mockGraphResponses([]);
  await assert.rejects(
    graphApi.callGraphAPI('test-token', 'GET', 'https://graph.microsoft.com.evil/v1.0/me/drive'),
    /Invalid Graph URL/
  );
  await assert.rejects(
    graphApi.callGraphAPI('test-token', 'GET', '../beta/me/drive'),
    /Invalid Graph URL/
  );
  assert.equal(requests.length, 0);
});

test('Graph pagination preserves a continuation URL when the count cap is met exactly', async () => {
  const nextLink = 'https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=page2';
  const requests = mockGraphResponses([{
    body: { value: [{ id: 'item-1' }], '@odata.nextLink': nextLink }
  }]);

  const result = await graphApi.callGraphAPIPaginated(
    'test-token', 'GET', 'me/drive/root/children', {}, 1
  );

  assert.deepEqual(result.value, [{ id: 'item-1' }]);
  assert.equal(result['@odata.nextLink'], nextLink);
  assert.equal(requests.length, 1);
});

test('Graph pagination follows absolute continuation URLs and returns all pages', async () => {
  const nextLink = 'https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=page2';
  const requests = mockGraphResponses([
    { body: { value: [{ id: 'item-1' }], '@odata.nextLink': nextLink } },
    { body: { value: [{ id: 'item-2' }], '@odata.count': 1 } }
  ]);

  const result = await graphApi.callGraphAPIPaginated(
    'test-token', 'GET', 'me/drive/root/children'
  );

  assert.deepEqual(result.value.map((item) => item.id), ['item-1', 'item-2']);
  assert.equal(result['@odata.nextLink'], null);
  assert.equal(requests.length, 2);
  assert.equal(requests[1].url, nextLink);
});

test('Graph API retries throttling but does not automatically replay ambiguous mutations', async () => {
  const requests = mockGraphResponses([
    { status: 429, headers: { 'retry-after': '0' }, body: { error: { code: 'tooManyRequests' } } },
    { body: { value: [] } }
  ]);
  const result = await graphApi.callGraphAPI('test-token', 'GET', 'me/drive');
  assert.deepEqual(result, { value: [] });
  assert.equal(requests.length, 2);

  const mutationRequests = mockGraphResponses([
    { status: 503, body: { error: { code: 'serviceUnavailable' } } },
    { status: 204 }
  ]);
  await assert.rejects(
    graphApi.callGraphAPI('test-token', 'POST', 'me/drive/items/item/copy', {}),
    /serviceUnavailable/
  );
  assert.equal(mutationRequests.length, 1);
});

test('long-running-operation monitor requires HTTPS and never forwards the Graph bearer token', async () => {
  const requests = mockGraphResponses([{ body: { status: 'completed', resourceId: 'copy-id' } }]);
  const { pollOperation } = require('../utils/operation-monitor');
  const result = await pollOperation('https://tenant.sharepoint.com/_api/v2.0/monitor/op', {
    maxWaitMs: 100,
    pollIntervalMs: 1
  });
  assert.equal(result.status, 'completed');
  assert.equal(requests.length, 1);
  assert.equal(requests[0].options.headers.Authorization, undefined);

  await assert.rejects(pollOperation('http://127.0.0.1/private', { maxWaitMs: 10 }), /HTTPS/);
});

test('new OneDrive tools are registered uniquely and carry the intended OAuth scopes', () => {
  const names = TOOLS.map((tool) => tool.name);
  assert.equal(new Set(names).size, names.length);
  for (const name of [
    'onedrive-list-permissions', 'onedrive-revoke-link', 'onedrive-unshare',
    'onedrive-invite', 'onedrive-update-permission', 'onedrive-resolve-link',
    'onedrive-copy', 'onedrive-update-item', 'onedrive-list-versions',
    'onedrive-restore-version', 'onedrive-restore-item', 'onedrive-quota',
    'onedrive-delta', 'onedrive-thumbnails', 'onedrive-convert',
    'onedrive-permanent-delete'
  ]) {
    assert.ok(names.includes(name), `${name} is registered`);
  }
  assert.equal(requiredScope('onedrive-list-permissions'), 'outlook:read');
  assert.equal(requiredScope('onedrive-resolve-link'), 'outlook:write');
  assert.equal(requiredScope('onedrive-update-permission'), 'outlook:destructive');
  assert.equal(requiredScope('onedrive-permanent-delete'), 'outlook:destructive');
  assert.equal(toolPolicy(TOOLS.find((tool) => tool.name === 'onedrive-permanent-delete')).destructiveHint, true);
});

async function withHandlerMocks(relativePath, mocks, run) {
  const graphApi = require('../utils/graph-api');
  const operationMonitor = require('../utils/operation-monitor');
  const resolver = require('../utils/onedrive-resolve');
  const auth = require('../auth');
  const saved = { graphApi: {}, operationMonitor: {}, resolver: {}, auth: {} };
  for (const [key, value] of Object.entries(mocks.graphApi || {})) {
    saved.graphApi[key] = graphApi[key];
    graphApi[key] = value;
  }
  for (const [key, value] of Object.entries(mocks.resolver || {})) {
    saved.resolver[key] = resolver[key];
    resolver[key] = value;
  }
  for (const [key, value] of Object.entries(mocks.operationMonitor || {})) {
    saved.operationMonitor[key] = operationMonitor[key];
    operationMonitor[key] = value;
  }
  for (const [key, value] of Object.entries(mocks.auth || {})) {
    saved.auth[key] = auth[key];
    auth[key] = value;
  }

  const modulePath = require.resolve(relativePath);
  const previousModule = require.cache[modulePath];
  delete require.cache[modulePath];
  const handler = require(relativePath);
  try {
    await run(handler);
  } finally {
    delete require.cache[modulePath];
    if (previousModule) require.cache[modulePath] = previousModule;
    for (const [key, value] of Object.entries(saved.graphApi)) graphApi[key] = value;
    for (const [key, value] of Object.entries(saved.operationMonitor)) operationMonitor[key] = value;
    for (const [key, value] of Object.entries(saved.resolver)) resolver[key] = value;
    for (const [key, value] of Object.entries(saved.auth)) auth[key] = value;
  }
}

test('permanent delete is dry-run unless confirm is the boolean true', async () => {
  const calls = [];
  await withHandlerMocks('../onedrive/permanent-delete', {
    graphApi: {
      callGraphAPI: async (...args) => { calls.push(args); return {}; },
      resolveDriveId: async () => 'drive-id'
    },
    resolver: { resolveItem: async () => ({ id: 'item-id', name: 'report.pdf', file: {} }) },
    auth: { ensureAuthenticated: async () => 'test-access-token' }
  }, async (handler) => {
    const implicit = await handler({ itemId: 'item-id' });
    const stringValue = await handler({ itemId: 'item-id', confirm: 'true' });
    assert.match(implicit.content[0].text, /DRY RUN/);
    assert.match(stringValue.content[0].text, /DRY RUN/);
    assert.equal(calls.length, 0);

    await handler({ itemId: 'item-id', confirm: true });
    assert.equal(calls.length, 1);
    assert.equal(calls[0][1], 'POST');
    assert.equal(calls[0][2], 'drives/drive-id/items/item-id/permanentDelete');
    assert.equal(calls[0].length, 3, 'permanentDelete sends no request body');
  });
});

test('copy requests include the destination drive and folder IDs', async () => {
  let copyRequest;
  await withHandlerMocks('../onedrive/copy', {
    graphApi: {
      callGraphAPI: async (...args) => {
        copyRequest = args;
        return { status: 202, headers: { location: 'https://tenant.sharepoint.com/monitor/op' } };
      },
      resolveDriveId: async () => 'drive-id'
    },
    resolver: {
      resolveItem: async (_token, ref) => ref.itemId
        ? { id: 'source-id', name: 'report.pdf', file: {} }
        : { id: 'destination-id', name: 'Documents', folder: {} }
    },
    operationMonitor: {
      pollOperation: async () => ({ settled: true, status: 'completed', resourceId: 'copy-id' })
    },
    auth: { ensureAuthenticated: async () => 'test-access-token' }
  }, async (handler) => {
    await handler({ itemId: 'source-id', destinationPath: '/Documents' });
    assert.equal(copyRequest[1], 'POST');
    assert.equal(copyRequest[2], 'me/drive/items/source-id/copy');
    assert.deepEqual(copyRequest[3].parentReference, { driveId: 'drive-id', id: 'destination-id' });
  });
});

test('folder creation refuses replacement of an existing item', async () => {
  let graphCalls = 0;
  await withHandlerMocks('../onedrive/folder', {
    graphApi: { callGraphAPI: async () => { graphCalls += 1; return {}; } },
    auth: { ensureAuthenticated: async () => 'test-access-token' }
  }, async ({ handleCreateFolder }) => {
    const result = await handleCreateFolder({ name: 'Reports', conflictBehavior: 'replace' });
    assert.match(result.content[0].text, /must be 'rename' or 'fail'/);
    assert.equal(graphCalls, 0);
  });
});

test('unshare dry-run does not return sharing-link capabilities and confirmation only revokes direct grants', async () => {
  const deletions = [];
  const permissions = [
    { id: 'direct-link', roles: ['read'], link: { type: 'view', scope: 'anonymous', webUrl: 'https://1drv.ms/secret-capability' } },
    { id: 'inherited-link', roles: ['read'], inheritedFrom: { id: 'parent-id' }, link: { type: 'view', scope: 'anonymous' } }
  ];
  await withHandlerMocks('../onedrive/unshare', {
    graphApi: {
      callGraphAPI: async (...args) => { deletions.push(args); return {}; },
      callGraphAPIPaginated: async () => ({ value: permissions })
    },
    resolver: { resolveItem: async () => ({ id: 'item-id', name: 'report.pdf', file: {} }) },
    auth: { ensureAuthenticated: async () => 'test-access-token' }
  }, async (handler) => {
    const dryRun = await handler({ itemId: 'item-id' });
    assert.match(dryRun.content[0].text, /DRY RUN/);
    assert.doesNotMatch(JSON.stringify(dryRun), /secret-capability/);
    assert.equal(deletions.length, 0);

    const confirmed = await handler({ itemId: 'item-id', confirm: true });
    assert.match(confirmed.content[0].text, /revoked 1\/1/);
    assert.equal(deletions.length, 1);
    assert.match(deletions[0][2], /direct-link$/);
  });
});
