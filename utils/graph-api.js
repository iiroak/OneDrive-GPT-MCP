/**
 * Microsoft Graph API helper functions
 */
const https = require('https');
const config = require('../config');

const DEFAULT_MAX_RETRIES = 3;
const DEFAULT_RETRY_BASE_MS = 500;
const RETRYABLE_STATUS = new Set([429, 502, 503, 504]);
const SAFE_RETRY_METHODS = new Set(['GET', 'HEAD', 'OPTIONS']);

function graphBaseUrl() {
  return new URL(config.GRAPH_API_ENDPOINT);
}

function buildGraphUrl(path, queryParams = {}) {
  const base = graphBaseUrl();
  let finalUrl;

  try {
    if (/^https?:\/\//i.test(path)) {
      finalUrl = new URL(path);
      if (finalUrl.protocol !== 'https:' || finalUrl.origin !== base.origin) {
        throw new Error('Graph URL is outside the approved Microsoft Graph origin');
      }
    } else {
      finalUrl = new URL(String(path).replace(/^\/+/, ''), base);
    }

    const basePath = base.pathname.endsWith('/') ? base.pathname : `${base.pathname}/`;
    if (finalUrl.pathname !== basePath.slice(0, -1) && !finalUrl.pathname.startsWith(basePath)) {
      throw new Error('Graph URL is outside the configured API version');
    }
  } catch (error) {
    throw new Error(`Invalid Graph URL: ${error.message}`);
  }

  for (const [key, value] of Object.entries(queryParams || {})) {
    if (value !== undefined && value !== null) finalUrl.searchParams.set(key, String(value));
  }

  return finalUrl;
}

/**
 * Parse a Retry-After header value (seconds, or an HTTP-date) into a ms delay.
 * @param {string|undefined} value
 * @returns {number|null}
 */
function parseRetryAfterMs(value) {
  if (!value) return null;
  const seconds = Number(value);
  if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
  const when = Date.parse(value);
  if (Number.isFinite(when)) return Math.max(0, when - Date.now());
  return null;
}

/**
 * Parse a Microsoft Graph JSON error body into a normalized shape.
 * Graph errors look like: { error: { code, message, innerError: {...} } }
 * @param {string} raw
 * @returns {{code: string|null, message: string|null, innerError: object|null}}
 */
function parseGraphErrorBody(raw) {
  try {
    const parsed = JSON.parse(raw);
    const err = parsed && parsed.error;
    if (err && typeof err === 'object') {
      return {
        code: err.code || null,
        message: err.message || null,
        innerError: err.innerError || err.innererror || null
      };
    }
  } catch {
    // Not JSON, or not the Graph error envelope. Leave fields null.
  }
  return { code: null, message: null, innerError: null };
}

/**
 * A Graph API error carrying the parsed error.code, HTTP status and raw body
 * so callers can branch on machine-readable state instead of string-matching
 * Error.message.
 */
class GraphApiError extends Error {
  constructor(message, { status = null, code = null, body = null, retryable = false } = {}) {
    super(message);
    this.name = 'GraphApiError';
    this.status = status;
    this.code = code;
    this.body = body;
    this.retryable = retryable;
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Makes a single (non-retried) request to the Microsoft Graph API.
 * @param {string} accessToken
 * @param {string} method
 * @param {string} path
 * @param {object|Buffer|null} data
 * @param {object} queryParams
 * @param {object} options
 * @param {object} [options.headers] - Extra headers merged over the defaults (e.g. If-Match, Prefer).
 * @param {boolean} [options.raw] - If true, resolve with the raw Buffer body instead of JSON.parse-ing it.
 * @param {boolean} [options.returnHeaders] - If true, resolve with { status, headers, body } instead of just body.
 * @returns {Promise<any>}
 */
function callGraphAPIOnce(accessToken, method, path, data, queryParams, options) {
  const { headers: extraHeaders = {}, raw = false, returnHeaders = false } = options || {};

  const finalUrl = buildGraphUrl(path, queryParams);

  return new Promise((resolve, reject) => {
    const isBinaryContentPut = method === 'PUT' && String(path).endsWith('/content');
    const baseHeaders = {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': Buffer.isBuffer(data) || isBinaryContentPut
        ? 'application/octet-stream'
        : 'application/json'
    };

    const reqOptions = {
      method,
      headers: { ...baseHeaders, ...extraHeaders }
    };

    const req = https.request(finalUrl, reqOptions, (res) => {
      /** @type {Buffer[]} */
      const chunks = [];

      res.on('data', (chunk) => {
        chunks.push(chunk);
      });

      res.on('end', () => {
        const status = res.statusCode || 0;
        const buffer = Buffer.concat(chunks);

        if (status >= 200 && status < 300) {
          try {
            const body = raw ? buffer : parseJsonBody(buffer);
            resolve(returnHeaders ? { status, headers: res.headers, body } : body);
          } catch (error) {
            reject(error);
          }
          return;
        }

        const bodyText = buffer.toString('utf8');
        const parsedError = parseGraphErrorBody(bodyText);
        const retryable = RETRYABLE_STATUS.has(status);
        const retryAfterMs = parseRetryAfterMs(res.headers['retry-after']);

        if (status === 401) {
          reject(new GraphApiError('UNAUTHORIZED', { status, code: parsedError.code, body: bodyText }));
          return;
        }

        const message = parsedError.code
          ? `API call failed with status ${status}: ${parsedError.code}${parsedError.message ? ` - ${parsedError.message}` : ''}`
          : `API call failed with status ${status}`;

        const error = new GraphApiError(message, {
          status,
          code: parsedError.code,
          body: bodyText,
          retryable
        });
        error.retryAfterMs = retryAfterMs;
        reject(error);
      });

      res.on('error', (error) => {
        reject(new GraphApiError(`Error reading API response: ${error.message}`, { retryable: true }));
      });
    });

    req.on('error', (error) => {
      reject(new GraphApiError(`Network error during API call: ${error.message}`, { retryable: true }));
    });

    if (data !== null && data !== undefined && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
      const payload = Buffer.isBuffer(data) || isBinaryContentPut
        ? data
        : JSON.stringify(data);
      req.write(payload);
    }

    req.end();
  });
}

function parseJsonBody(buffer) {
  const text = buffer.length ? buffer.toString('utf8') : '{}';
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new GraphApiError(`Error parsing API response: ${error.message}`);
  }
}

/**
 * Makes a request to the Microsoft Graph API, with automatic retry on
 * throttling (429) and transient upstream failures (502/503/504).
 *
 * Backward compatible: existing callers using the 5-argument form get the
 * exact same return shape as before (parsed JSON body, or throw).
 *
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET, POST, PATCH, PUT, DELETE, ...)
 * @param {string} path - API endpoint path
 * @param {object|Buffer|null} data - Data to send for POST/PATCH/PUT requests
 * @param {object} queryParams - Query parameters
 * @param {object} [options] - Optional extras.
 * @param {object} [options.headers] - Extra headers merged over the defaults.
 * @param {boolean} [options.raw] - Resolve with the raw Buffer body, skip JSON.parse.
 * @param {boolean} [options.returnHeaders] - Resolve with { status, headers, body }.
 * @param {number} [options.maxRetries] - Max retry attempts on throttling/5xx (default 3, 0 disables).
 * @param {boolean} [options.retryUnsafe] - Also retry methods other than GET/HEAD/OPTIONS after ambiguous transient failures.
 * @returns {Promise<object|Buffer|{status:number, headers:object, body:any}>}
 */
async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}, options = {}) {
  const maxRetries = Number.isFinite(options.maxRetries) ? options.maxRetries : DEFAULT_MAX_RETRIES;
  const normalizedMethod = String(method || 'GET').toUpperCase();
  let attempt = 0;

  // eslint-disable-next-line no-constant-condition
  while (true) {
    try {
      return await callGraphAPIOnce(accessToken, normalizedMethod, path, data, queryParams, options);
    } catch (error) {
      const isRetryable = error instanceof GraphApiError && error.retryable;
      // A 429 explicitly means Graph throttled the request before processing
      // it, so retrying it is safe for all verbs. A 5xx or network failure on
      // POST/PATCH/DELETE is ambiguous: the mutation may have succeeded even
      // though the response was lost. Never replay those implicitly (copy,
      // invite and uploads can have side effects); callers must explicitly
      // opt in with retryUnsafe if their operation is known to be idempotent.
      const safeToRetry = error.status === 429
        || SAFE_RETRY_METHODS.has(normalizedMethod)
        || options.retryUnsafe === true;
      if (!isRetryable || !safeToRetry || attempt >= maxRetries) {
        if (error instanceof GraphApiError) throw error;
        console.error('Error calling Graph API request.');
        throw error;
      }
      const backoff = error.retryAfterMs != null
        ? error.retryAfterMs
        : DEFAULT_RETRY_BASE_MS * Math.pow(2, attempt);
      console.error(`Graph API throttled/unavailable (status ${error.status}), retrying in ${backoff}ms (attempt ${attempt + 1}/${maxRetries})`);
      await sleep(backoff);
      attempt += 1;
    }
  }
}

/**
 * Calls Graph API with pagination support to retrieve all results up to maxCount.
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET only for pagination)
 * @param {string} path - API endpoint path
 * @param {object} queryParams - Initial query parameters
 * @param {number} maxCount - Maximum number of items to retrieve (0 = all)
 * @returns {Promise<{value: any[], '@odata.count': number, '@odata.nextLink': string|null}>}
 */
async function callGraphAPIPaginated(accessToken, method, path, queryParams = {}, maxCount = 0) {
  if (method !== 'GET') {
    throw new Error('Pagination only supports GET requests');
  }

  const allItems = [];
  let nextLink = null;
  let currentUrl = path;
  let currentParams = queryParams;

  try {
    do {
      // Make API call
      const response = await callGraphAPI(accessToken, method, currentUrl, null, currentParams);

      // Add items from this page
      if (response.value && Array.isArray(response.value)) {
        allItems.push(...response.value);
        console.error(`Pagination: Retrieved ${response.value.length} items, total so far: ${allItems.length}`);
      }

      // Check if we've reached the desired count
      if (maxCount > 0 && allItems.length >= maxCount) {
        console.error(`Pagination: Reached max count of ${maxCount}, stopping`);
        nextLink = response['@odata.nextLink'] || null;
        break;
      }

      // Get next page URL
      nextLink = response['@odata.nextLink'];

      if (nextLink) {
        // Pass the full nextLink URL directly to callGraphAPI
        buildGraphUrl(nextLink);
        currentUrl = nextLink;
        currentParams = {}; // nextLink already contains all params
        console.error(`Pagination: Following nextLink, ${allItems.length} items so far`);
      }
    } while (nextLink);

    // Trim to exact count if needed
    const finalItems = maxCount > 0 ? allItems.slice(0, maxCount) : allItems;
    const hasMore = maxCount > 0 && finalItems.length >= maxCount && Boolean(nextLink);

    console.error(`Pagination complete: Retrieved ${finalItems.length} total items`);

    return {
      value: finalItems,
      '@odata.count': finalItems.length,
      // Only meaningful when we stopped early because of maxCount; null once
      // the collection is genuinely exhausted, so callers can tell "there is
      // more" from "that's everything".
      '@odata.nextLink': hasMore ? nextLink : null
    };
  } catch (error) {
    console.error(`Error during Graph API pagination (status ${error.status || 'unknown'}).`);
    throw error;
  }
}

/**
 * Makes a request to the Microsoft Graph API that returns a download URL (302 redirect)
 * Used for OneDrive file downloads which return a pre-authenticated download URL
 * @param {string} accessToken - The access token for authentication
 * @param {string} path - API endpoint path
 * @returns {Promise<string>} - The download URL from the redirect
 */
async function callGraphAPIDownload(accessToken, path) {
  return new Promise((resolve, reject) => {
    const fullUrl = buildGraphUrl(path);
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    };

    const req = https.request(fullUrl, options, (res) => {
      // Graph API returns 302 with Location header containing the download URL
      if (res.statusCode === 302 && res.headers.location) {
        res.resume();
        resolve(res.headers.location);
      } else if (res.statusCode >= 200 && res.statusCode < 300) {
        // Some endpoints might return the URL in the body instead
        let responseData = '';
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        res.on('end', () => {
          try {
            const jsonResponse = JSON.parse(responseData);
            if (jsonResponse['@microsoft.graph.downloadUrl']) {
              resolve(jsonResponse['@microsoft.graph.downloadUrl']);
            } else {
              reject(new Error('No download URL found in response'));
            }
          } catch (error) {
            reject(new Error(`Error parsing download response: ${error.message}`));
          }
        });
      } else if (res.statusCode === 401) {
        reject(new Error('UNAUTHORIZED'));
      } else {
        let responseData = '';
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        res.on('end', () => {
          reject(new Error(`Download request failed with status ${res.statusCode}: ${responseData}`));
        });
      }
    });

    req.on('error', (error) => {
      reject(new Error(`Network error during download request: ${error.message}`));
    });

    req.end();
  });
}

/**
 * Resolve the driveId for the signed-in user's OneDrive.
 * A handful of driveItem operations (e.g. permanentDelete) only expose a
 * `/drives/{driveId}/items/{itemId}/...` form and have no `/me/drive/...`
 * shortcut, so callers need the drive's own id first. Cached in-process
 * since it does not change for a given signed-in account.
 * @param {string} accessToken
 * @returns {Promise<string>}
 */
let cachedDriveId = null;
async function resolveDriveId(accessToken) {
  if (cachedDriveId) return cachedDriveId;
  const drive = await callGraphAPI(accessToken, 'GET', 'me/drive', null, { $select: 'id' });
  if (!drive || !drive.id) {
    throw new GraphApiError('Unable to resolve the current drive id');
  }
  cachedDriveId = drive.id;
  return cachedDriveId;
}

module.exports = {
  callGraphAPI,
  callGraphAPIPaginated,
  callGraphAPIDownload,
  resolveDriveId,
  GraphApiError,
  parseRetryAfterMs
};
