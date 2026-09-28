/**
 * Poll a Microsoft Graph long-running-operation "Location" monitor URL.
 *
 * Endpoints like `driveItem: copy` return `202 Accepted` with a `Location`
 * header pointing at a monitor resource. Per Microsoft's own docs
 * (https://learn.microsoft.com/en-us/graph/long-running-actions-overview):
 *
 *   - "The location URL returned might not be on the Microsoft Graph API
 *     endpoint" (it is commonly a *.sharepoint.com or api.onedrive.com host,
 *     and the exact host varies by tenant/geo, so it cannot be allowlisted
 *     ahead of time the way CDN download hosts are).
 *   - "This request doesn't require authentication, because the URL is
 *     short-lived and unique to the original caller."
 *
 * So the invariant here is NOT a host allowlist (there is none to pin) — it
 * is: HTTPS only, no query-string credentials, and the Graph bearer token is
 * NEVER attached to this request. Leaking the Graph Authorization header to
 * an arbitrary SharePoint/OneDrive CDN host would be a token exfiltration
 * bug, not a feature.
 */
const https = require('https');
const { URL } = require('url');

const DEFAULT_POLL_INTERVAL_MS = 1000;
const DEFAULT_MAX_WAIT_MS = 60000;
const MONITOR_REQUEST_TIMEOUT_MS = 10000;
const MAX_MONITOR_BODY_BYTES = 64 * 1024; // monitor bodies are tiny JSON; bound defensively

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Fetch the monitor resource once. No Authorization header is sent.
 * @param {string} url
 * @returns {Promise<{status: number, body: any}>}
 */
function fetchMonitorOnce(url) {
  return new Promise((resolve, reject) => {
    let parsed;
    try {
      parsed = new URL(url);
    } catch {
      reject(new Error('Operation monitor URL is not valid'));
      return;
    }
    if (parsed.protocol !== 'https:') {
      reject(new Error('Operation monitor URL must be HTTPS'));
      return;
    }
    if (parsed.username || parsed.password) {
      reject(new Error('Operation monitor URL must not carry embedded credentials'));
      return;
    }

    const req = https.request(parsed, { method: 'GET', headers: { Accept: 'application/json' } }, (res) => {
      const chunks = [];
      let total = 0;
      res.on('data', (chunk) => {
        total += chunk.length;
        if (total > MAX_MONITOR_BODY_BYTES) {
          res.destroy();
          reject(new Error('Operation monitor response exceeded the expected size'));
          return;
        }
        chunks.push(chunk);
      });
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        let body = null;
        try {
          body = text ? JSON.parse(text) : null;
        } catch {
          body = null;
        }
        resolve({ status: res.statusCode || 0, body });
      });
      res.on('error', (error) => reject(new Error(`Error reading operation monitor: ${error.message}`)));
    });

    req.on('error', (error) => reject(new Error(`Network error polling operation monitor: ${error.message}`)));
    req.setTimeout(MONITOR_REQUEST_TIMEOUT_MS, () => {
      req.destroy(new Error('Timed out polling operation monitor'));
    });
    req.end();
  });
}

/**
 * Poll a long-running-operation monitor URL until it reports a terminal
 * state, or until maxWaitMs elapses.
 *
 * @param {string} locationUrl - The Location header from a 202 response.
 * @param {object} [opts]
 * @param {number} [opts.pollIntervalMs=1000]
 * @param {number} [opts.maxWaitMs=60000]
 * @returns {Promise<{
 *   settled: boolean,          // true if a terminal status was observed
 *   status: string|null,       // 'completed' | 'failed' | Graph-specific enum, or null if never observed
 *   percentageComplete: number|null,
 *   resourceId: string|null,
 *   resourceLocation: string|null,
 *   error: object|null,
 *   raw: object|null            // last raw monitor body
 * }>}
 */
async function pollOperation(locationUrl, opts = {}) {
  const pollIntervalMs = opts.pollIntervalMs || DEFAULT_POLL_INTERVAL_MS;
  const maxWaitMs = opts.maxWaitMs || DEFAULT_MAX_WAIT_MS;
  const deadline = Date.now() + maxWaitMs;

  let lastBody = null;

  while (Date.now() < deadline) {
    const { body } = await fetchMonitorOnce(locationUrl);
    lastBody = body;

    const status = body && (body.status || null);
    const isTerminal = status === 'completed' || status === 'failed'
      || (body && body.error && !status);

    if (isTerminal) {
      return {
        settled: true,
        status: status || (body && body.error ? 'failed' : null),
        percentageComplete: body && (body.percentageComplete ?? body.percentComplete ?? null),
        resourceId: body && (body.resourceId || null),
        resourceLocation: body && (body.resourceLocation || null),
        error: body && body.error || null,
        raw: body
      };
    }

    await sleep(pollIntervalMs);
  }

  return {
    settled: false,
    status: lastBody && lastBody.status || null,
    percentageComplete: lastBody && (lastBody.percentageComplete ?? lastBody.percentComplete ?? null),
    resourceId: lastBody && lastBody.resourceId || null,
    resourceLocation: lastBody && lastBody.resourceLocation || null,
    error: null,
    raw: lastBody
  };
}

module.exports = { pollOperation };
