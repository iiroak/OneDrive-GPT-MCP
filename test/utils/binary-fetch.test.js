const { isHostAllowed, filenameFromContentDisposition, looksLikeHtml, fetchBytes } = require('../../utils/binary-fetch');

describe('isHostAllowed', () => {
  const allowed = ['graph.microsoft.com', '*.sharepoint.com', '*.1drv.com'];

  test('accepts an exact host', () => {
    expect(isHostAllowed('graph.microsoft.com', allowed)).toBe(true);
  });

  test('accepts a subdomain of a wildcard pattern', () => {
    expect(isHostAllowed('tenant.sharepoint.com', allowed)).toBe(true);
    expect(isHostAllowed('a.b.sharepoint.com', allowed)).toBe(true);
  });

  test('rejects the bare apex of a wildcard pattern', () => {
    expect(isHostAllowed('sharepoint.com', allowed)).toBe(false);
  });

  test('rejects a suffix-confusion host', () => {
    expect(isHostAllowed('evilsharepoint.com', allowed)).toBe(false);
    expect(isHostAllowed('graph.microsoft.com.evil.com', allowed)).toBe(false);
  });

  test('rejects internal and metadata targets', () => {
    expect(isHostAllowed('localhost', allowed)).toBe(false);
    expect(isHostAllowed('127.0.0.1', allowed)).toBe(false);
    expect(isHostAllowed('169.254.169.254', allowed)).toBe(false);
    expect(isHostAllowed('10.0.0.5', allowed)).toBe(false);
  });

  test('is case and trailing-dot insensitive', () => {
    expect(isHostAllowed('GRAPH.microsoft.COM', allowed)).toBe(true);
    expect(isHostAllowed('graph.microsoft.com.', allowed)).toBe(true);
  });

  test('rejects everything when the allowlist is empty', () => {
    expect(isHostAllowed('graph.microsoft.com', [])).toBe(false);
  });
});

describe('fetchBytes input validation', () => {
  const options = { allowedHosts: ['graph.microsoft.com'], maxBytes: 1024 };

  test('refuses plain HTTP', async () => {
    await expect(fetchBytes('http://graph.microsoft.com/x', options)).rejects.toThrow(/HTTPS/);
  });

  test('refuses a non-HTTP scheme', async () => {
    await expect(fetchBytes('file:///etc/passwd', options)).rejects.toThrow(/HTTPS/);
  });

  test('refuses an unlisted host before opening a socket', async () => {
    await expect(fetchBytes('https://evil.example.com/x', options)).rejects.toThrow(/host no permitido/);
  });

  test('refuses embedded credentials', async () => {
    await expect(fetchBytes('https://user:pass@graph.microsoft.com/x', options))
      .rejects.toThrow(/credenciales/);
  });

  test('refuses a malformed URL', async () => {
    await expect(fetchBytes('not a url', options)).rejects.toThrow(/no es valida/);
  });

  test('requires an explicit allowlist', async () => {
    await expect(fetchBytes('https://graph.microsoft.com/x', { maxBytes: 10 }))
      .rejects.toThrow(/hosts permitidos/);
  });
});

describe('filenameFromContentDisposition', () => {
  test('prefers the RFC 5987 encoded form', () => {
    const header = "attachment; filename=\"fallback.txt\"; filename*=UTF-8''Informe%20final.pdf";
    expect(filenameFromContentDisposition(header)).toBe('Informe final.pdf');
  });

  test('reads the quoted plain form', () => {
    expect(filenameFromContentDisposition('attachment; filename="a b.txt"')).toBe('a b.txt');
  });

  test('reads the unquoted plain form', () => {
    expect(filenameFromContentDisposition('attachment; filename=a.txt')).toBe('a.txt');
  });

  test('returns empty for a missing or useless header', () => {
    expect(filenameFromContentDisposition('')).toBe('');
    expect(filenameFromContentDisposition(undefined)).toBe('');
    expect(filenameFromContentDisposition('inline')).toBe('');
  });

  test('survives malformed percent-encoding by falling back', () => {
    const header = "attachment; filename=\"ok.txt\"; filename*=UTF-8''%ZZ";
    expect(filenameFromContentDisposition(header)).toBe('ok.txt');
  });
});

describe('looksLikeHtml', () => {
  test('detects an HTML content type', () => {
    expect(looksLikeHtml(Buffer.from('anything'), 'text/html; charset=utf-8')).toBe(true);
  });

  test('detects an HTML body served under a lying content type', () => {
    expect(looksLikeHtml(Buffer.from('<!DOCTYPE html><html>'), 'application/octet-stream')).toBe(true);
    expect(looksLikeHtml(Buffer.from('\n  <html lang="es">'), 'application/pdf')).toBe(true);
  });

  test('does not flag real binary content', () => {
    expect(looksLikeHtml(Buffer.from('%PDF-1.4'), 'application/pdf')).toBe(false);
    expect(looksLikeHtml(Buffer.from('PK\u0003\u0004'), 'application/zip')).toBe(false);
  });
});
