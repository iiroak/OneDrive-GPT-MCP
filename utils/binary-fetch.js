/**
 * Bounded HTTPS fetch for file bytes.
 *
 * Ported from the umayor-mcp `HttpClient.stream_to_file` contract because the
 * pre-authenticated OneDrive download URL is an opaque redirect chain into CDN
 * hosts, and following it naively is what turns a file transfer into an SSRF.
 *
 * Guarantees:
 * - HTTPS only. No http:// hop is ever followed.
 * - Every hop, including the first, is checked against a host allowlist.
 * - Redirects are followed manually and bounded, so a redirect loop terminates.
 * - The Authorization header is dropped when a redirect crosses to another host,
 *   so a Graph token is never handed to a CDN.
 * - Byte budget enforced both from Content-Length and while streaming, so a
 *   lying or absent Content-Length cannot exhaust memory.
 */
const https = require('https');
const crypto = require('crypto');
const { URL } = require('url');

/**
 * @typedef {object} FetchedBytes
 * @property {Buffer} buffer
 * @property {string} sha256
 * @property {number} bytes
 * @property {string} contentType   Raw content-type header, may be ''
 * @property {string} filename      Filename parsed from Content-Disposition, may be ''
 * @property {string} finalUrl      URL that actually served the bytes
 */

/**
 * Does a hostname satisfy the allowlist? Supports leading-wildcard patterns
 * such as `*.sharepoint.com`, which must not match the bare apex.
 * @param {string|null} hostname
 * @param {string[]} allowedHosts
 * @returns {boolean}
 */
function isHostAllowed(hostname, allowedHosts) {
  if (!hostname) return false;
  const host = hostname.toLowerCase().replace(/\.+$/, '');
  return allowedHosts.some((raw) => {
    const pattern = String(raw || '').toLowerCase().replace(/\.+$/, '');
    if (!pattern) return false;
    if (pattern.startsWith('*.')) {
      const suffix = pattern.slice(1); // ".sharepoint.com"
      return host.endsWith(suffix) && host !== suffix.slice(1);
    }
    return host === pattern;
  });
}

/**
 * Extract a filename from a Content-Disposition header.
 * Prefers RFC 5987 `filename*` over the plain form.
 * @param {string} value
 * @returns {string}
 */
function filenameFromContentDisposition(value) {
  const header = String(value || '');
  if (!header) return '';

  const encoded = /filename\*\s*=\s*UTF-8''([^;]+)/i.exec(header);
  if (encoded) {
    try {
      return decodeURIComponent(encoded[1].trim().replace(/^"|"$/g, ''));
    } catch {
      // Malformed percent-encoding: fall through to the plain form.
    }
  }

  const plain = /filename\s*=\s*(?:"([^"]+)"|([^;]+))/i.exec(header);
  if (plain) return (plain[1] || plain[2] || '').trim();
  return '';
}

/**
 * Heuristic: did we get an HTML error/login page instead of the file?
 * A 200 with a login page is the most common silent failure on these endpoints.
 * @param {Buffer} buffer
 * @param {string} contentType
 * @returns {boolean}
 */
function looksLikeHtml(buffer, contentType) {
  const type = String(contentType || '').split(';')[0].trim().toLowerCase();
  if (type === 'text/html' || type === 'application/xhtml+xml') return true;
  const prefix = buffer.subarray(0, 512).toString('latin1').trimStart().toLowerCase();
  return prefix.startsWith('<!doctype html')
    || prefix.startsWith('<html')
    || prefix.startsWith('<head')
    || prefix.startsWith('<body');
}

/**
 * Perform one request without following redirects.
 * @param {string} url
 * @param {object} headers
 * @param {number} maxBytes
 * @param {number} timeoutMs
 * @returns {Promise<{redirectTo: string|null, buffer: Buffer|null, headers: object, statusCode: number}>}
 */
function requestOnce(url, headers, maxBytes, timeoutMs) {
  return new Promise((resolve, reject) => {
    const request = https.request(url, { method: 'GET', headers }, (response) => {
      const status = response.statusCode || 0;

      if (status >= 300 && status < 400 && response.headers.location) {
        response.resume(); // Discard the body; we only need the Location.
        resolve({
          redirectTo: new URL(response.headers.location, url).toString(),
          buffer: null,
          headers: response.headers,
          statusCode: status
        });
        return;
      }

      const declaredLength = Number(response.headers['content-length']);
      if (Number.isFinite(declaredLength) && declaredLength > maxBytes) {
        response.destroy();
        reject(new Error(
          `El archivo remoto es demasiado grande (${declaredLength} > ${maxBytes} bytes)`
        ));
        return;
      }

      /** @type {Buffer[]} */
      const chunks = [];
      let total = 0;

      response.on('data', (chunk) => {
        total += chunk.length;
        if (total > maxBytes) {
          response.destroy();
          reject(new Error(`El archivo remoto excede el limite de ${maxBytes} bytes`));
          return;
        }
        chunks.push(chunk);
      });

      response.on('end', () => {
        if (status < 200 || status >= 300) {
          const preview = Buffer.concat(chunks).subarray(0, 240).toString('utf8').replace(/\n/g, ' ');
          if (status === 401 || status === 403) {
            reject(new Error('UNAUTHORIZED'));
          } else {
            reject(new Error(`La descarga fallo con estado ${status}: ${preview}`));
          }
          return;
        }
        resolve({
          redirectTo: null,
          buffer: Buffer.concat(chunks),
          headers: response.headers,
          statusCode: status
        });
      });

      response.on('error', (error) => reject(new Error(`Error leyendo la respuesta: ${error.message}`)));
    });

    request.setTimeout(timeoutMs, () => {
      request.destroy();
      reject(new Error(`La descarga excedio el tiempo limite de ${timeoutMs} ms`));
    });

    request.on('error', (error) => reject(new Error(`Error de red durante la descarga: ${error.message}`)));
    request.end();
  });
}

/**
 * Fetch bytes from an HTTPS URL under an explicit host allowlist.
 * @param {string} url
 * @param {object} options
 * @param {string[]} options.allowedHosts
 * @param {number} options.maxBytes
 * @param {object} [options.headers]
 * @param {number} [options.maxRedirects]
 * @param {number} [options.timeoutMs]
 * @returns {Promise<FetchedBytes>}
 */
async function fetchBytes(url, options) {
  const {
    allowedHosts,
    maxBytes,
    headers = {},
    maxRedirects = 5,
    timeoutMs = 120000
  } = options;

  if (!Array.isArray(allowedHosts) || allowedHosts.length === 0) {
    throw new Error('fetchBytes requiere una lista de hosts permitidos');
  }

  let currentUrl = String(url || '');
  let currentHeaders = { ...headers };

  for (let hop = 0; hop <= maxRedirects; hop += 1) {
    let parsed;
    try {
      parsed = new URL(currentUrl);
    } catch {
      throw new Error('La URL de descarga no es valida');
    }

    if (parsed.protocol !== 'https:') {
      throw new Error('Solo se permiten descargas por HTTPS');
    }
    if (parsed.username || parsed.password) {
      throw new Error('La URL de descarga no puede incluir credenciales');
    }
    if (!isHostAllowed(parsed.hostname, allowedHosts)) {
      throw new Error(`La descarga apunta a un host no permitido: ${parsed.hostname}`);
    }

    const result = await requestOnce(currentUrl, currentHeaders, maxBytes, timeoutMs);

    if (result.redirectTo) {
      const next = new URL(result.redirectTo);
      // Never leak an Authorization header across a host boundary.
      if (next.hostname.toLowerCase() !== parsed.hostname.toLowerCase()) {
        currentHeaders = {};
      }
      currentUrl = result.redirectTo;
      continue;
    }

    const buffer = result.buffer || Buffer.alloc(0);
    return {
      buffer,
      sha256: crypto.createHash('sha256').update(buffer).digest('hex'),
      bytes: buffer.length,
      contentType: String(result.headers['content-type'] || ''),
      filename: filenameFromContentDisposition(result.headers['content-disposition']),
      finalUrl: currentUrl
    };
  }

  throw new Error('Demasiadas redirecciones durante la descarga');
}

module.exports = {
  fetchBytes,
  isHostAllowed,
  filenameFromContentDisposition,
  looksLikeHtml
};
