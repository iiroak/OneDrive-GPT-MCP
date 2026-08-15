function decodeUploadContent(args) {
  const hasBase64 = typeof args.contentBase64 === 'string';
  const hasText = typeof args.content === 'string';
  if (hasBase64 && hasText) throw new Error('Provide content or contentBase64, not both.');

  if (hasBase64) {
    const encoded = args.contentBase64.replace(/\s/g, '');
    if (!encoded || encoded.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(encoded)) {
      throw new Error('contentBase64 must be a valid standard Base64 string.');
    }

    const buffer = Buffer.from(encoded, 'base64');
    if (buffer.length === 0) throw new Error('contentBase64 must not be empty.');
    return buffer;
  }

  if (hasText && args.content.length > 0) {
    return Buffer.from(args.content, 'utf8');
  }

  throw new Error('Provide non-empty content or contentBase64.');
}

module.exports = { decodeUploadContent };
