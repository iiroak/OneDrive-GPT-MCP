/**
 * Server-side text extraction for staged files.
 *
 * This is the third and most useful transfer channel: instead of handing the
 * model a URL or a base64 blob it must decode, the server turns the document
 * into text and returns the text. Mirrors `document_processing.py` in
 * umayor-mcp, including its honesty contract:
 *
 *   status: 'complete' -> everything meaningful was extracted as text
 *   status: 'partial'  -> real content exists that is NOT in `text`
 *                         (images, charts, scanned pages); `warnings` says which
 *   status: 'failed'   -> extraction did not work; use another channel
 *
 * A caller that ignores `status` will silently mistake a scanned PDF for an
 * empty document, which is the failure mode this field exists to prevent.
 */
const { openZip } = require('./zip-reader');
const { sanitizeHtmlToText } = require('./html-sanitizer');

const MAX_RETURNED_CHARS = 200000;
const DEFAULT_MAX_CHARS = 50000;

const TEXT_EXTENSIONS = new Set([
  '.txt', '.md', '.markdown', '.csv', '.tsv', '.json', '.xml', '.yml', '.yaml',
  '.log', '.ini', '.cfg', '.conf', '.rst', '.tex', '.srt', '.vtt',
  '.js', '.ts', '.jsx', '.tsx', '.py', '.rb', '.go', '.rs', '.java', '.c',
  '.h', '.cpp', '.cs', '.sh', '.bash', '.zsh', '.sql', '.css', '.scss'
]);

/**
 * @param {string} text
 * @param {number} maxChars
 * @returns {{text: string, truncated: boolean}}
 */
function boundText(text, maxChars) {
  const limit = Math.max(1, Math.min(Number(maxChars) || DEFAULT_MAX_CHARS, MAX_RETURNED_CHARS));
  return { text: text.slice(0, limit), truncated: text.length > limit };
}

/**
 * Decode the five XML predefined entities plus numeric references.
 * @param {string} value
 * @returns {string}
 */
function decodeXmlEntities(value) {
  return value
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => {
      const code = parseInt(hex, 16);
      return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : '';
    })
    .replace(/&#(\d+);/g, (_, dec) => {
      const code = parseInt(dec, 10);
      return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : '';
    })
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

/**
 * Collect the text content of every occurrence of an XML element.
 * @param {string} xml
 * @param {string} tag  Namespaced tag, e.g. 'w:t'
 * @returns {string[]}
 */
function collectElementText(xml, tag) {
  const pattern = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>|<${tag}(?:\\s[^>]*)?/>`, 'g');
  const values = [];
  let match;
  while ((match = pattern.exec(xml)) !== null) {
    values.push(match[1] === undefined ? '' : decodeXmlEntities(match[1]));
  }
  return values;
}

/**
 * Sort OOXML part names by their trailing number (slide2 before slide10).
 * @param {string[]} names
 * @returns {string[]}
 */
function sortByOrdinal(names) {
  return names.slice().sort((a, b) => {
    const left = Number(/(\d+)\.xml$/.exec(a)?.[1] ?? 0);
    const right = Number(/(\d+)\.xml$/.exec(b)?.[1] ?? 0);
    return left - right;
  });
}

/**
 * Extract paragraph text from a WordprocessingML body.
 * @param {string} xml
 * @returns {string}
 */
function wordParagraphs(xml) {
  const paragraphs = [];
  const pattern = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
  let match;
  while ((match = pattern.exec(xml)) !== null) {
    const body = match[1];
    // w:tab and w:br carry layout that matters for readability.
    const withBreaks = body
      .replace(/<w:tab(?:\s[^>]*)?\/>/g, '\t')
      .replace(/<w:br(?:\s[^>]*)?\/>/g, '\n');
    const text = collectElementText(withBreaks, 'w:t').join('');
    paragraphs.push(text);
  }
  return paragraphs.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

/**
 * @param {Buffer} buffer
 * @param {number} maxChars
 * @returns {object}
 */
function processDocx(buffer, maxChars) {
  const zip = openZip(buffer);
  const documentXml = zip.read('word/document.xml');
  if (!documentXml) {
    return {
      format: 'docx',
      status: 'failed',
      text: '',
      truncated: false,
      warnings: ['El DOCX no contiene word/document.xml']
    };
  }

  const xml = documentXml.toString('utf8');
  const body = wordParagraphs(xml);

  // Headers/footers live in separate parts and often hold real content.
  const extraParts = zip.names.filter((name) => /^word\/(header|footer)\d+\.xml$/.test(name));
  const extras = sortByOrdinal(extraParts)
    .map((name) => wordParagraphs(zip.read(name).toString('utf8')))
    .filter(Boolean);

  const embedded = zip.names.filter((name) => /^word\/media\//.test(name));
  const warnings = [];
  if (embedded.length > 0) {
    warnings.push(
      `El DOCX contiene ${embedded.length} objeto(s) incrustado(s) (imagenes o graficos) no representados como texto`
    );
  }

  const combined = [body, ...extras].filter(Boolean).join('\n\n');
  const bounded = boundText(combined, maxChars);
  return {
    format: 'docx',
    status: embedded.length > 0 ? 'partial' : 'complete',
    ...bounded,
    embedded_objects: embedded.length,
    warnings
  };
}

/**
 * @param {Buffer} buffer
 * @param {number} maxChars
 * @returns {object}
 */
function processPptx(buffer, maxChars) {
  const zip = openZip(buffer);
  const slideNames = sortByOrdinal(zip.names.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)));

  if (slideNames.length === 0) {
    return {
      format: 'pptx',
      status: 'failed',
      text: '',
      truncated: false,
      slides: [],
      visual_pages: [],
      warnings: ['El PPTX no contiene diapositivas legibles']
    };
  }

  const slides = [];
  const parts = [];
  const visualPages = [];
  const warnings = [];

  slideNames.forEach((name, index) => {
    const number = index + 1;
    const xml = zip.read(name).toString('utf8');
    const text = collectElementText(xml, 'a:t')
      .map((value) => value.trim())
      .filter(Boolean)
      .join('\n');

    // Speaker notes are separate parts, matched by slide ordinal.
    const notesName = `ppt/notesSlides/notesSlide${number}.xml`;
    let notes = '';
    if (zip.has(notesName)) {
      notes = collectElementText(zip.read(notesName).toString('utf8'), 'a:t')
        .map((value) => value.trim())
        .filter(Boolean)
        .join('\n');
    }

    // Pictures, charts and diagrams carry information text extraction misses.
    const visualCount = (xml.match(/<p:pic(?:\s|>)/g) || []).length
      + (xml.match(/<c:chart(?:\s|\/|>)/g) || []).length
      + (xml.match(/<dgm:relIds(?:\s|\/|>)/g) || []).length
      + (xml.match(/<p:graphicFrame(?:\s|>)/g) || []).length;

    const slideText = [text, notes ? `Notas: ${notes}` : ''].filter(Boolean).join('\n');
    slides.push({ number, text: slideText, visual_objects: visualCount, visual: visualCount > 0 });
    parts.push(`[Diapositiva ${number}]\n${slideText}`.trimEnd());

    if (visualCount > 0) {
      visualPages.push(number);
      warnings.push(
        `La diapositiva ${number} contiene ${visualCount} objeto(s) visual(es) no representados como texto`
      );
    }
  });

  const bounded = boundText(parts.join('\n\n'), maxChars);
  return {
    format: 'pptx',
    status: visualPages.length > 0 ? 'partial' : 'complete',
    ...bounded,
    slides,
    visual_pages: visualPages,
    warnings
  };
}

/**
 * @param {Buffer} buffer
 * @param {number} maxChars
 * @returns {object}
 */
function processXlsx(buffer, maxChars) {
  const zip = openZip(buffer);

  // Cell values are indices into a shared string table when t="s".
  let sharedStrings = [];
  if (zip.has('xl/sharedStrings.xml')) {
    const xml = zip.read('xl/sharedStrings.xml').toString('utf8');
    sharedStrings = (xml.match(/<si(?:\s[^>]*)?>[\s\S]*?<\/si>/g) || []).map((item) =>
      collectElementText(item, 't').join('')
    );
  }

  const sheetNames = sortByOrdinal(zip.names.filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name)));
  if (sheetNames.length === 0) {
    return {
      format: 'xlsx',
      status: 'failed',
      text: '',
      truncated: false,
      warnings: ['El XLSX no contiene hojas legibles']
    };
  }

  const parts = [];
  sheetNames.forEach((name, index) => {
    const xml = zip.read(name).toString('utf8');
    const rows = [];
    const rowPattern = /<row(?:\s[^>]*)?>([\s\S]*?)<\/row>/g;
    let rowMatch;
    while ((rowMatch = rowPattern.exec(xml)) !== null) {
      const cells = [];
      const cellPattern = /<c(?:\s([^>]*))?>([\s\S]*?)<\/c>|<c(?:\s([^>]*))?\/>/g;
      let cellMatch;
      while ((cellMatch = cellPattern.exec(rowMatch[1])) !== null) {
        const attributes = cellMatch[1] || cellMatch[3] || '';
        const body = cellMatch[2] || '';
        const type = /\bt="([^"]+)"/.exec(attributes)?.[1] || 'n';
        if (type === 's') {
          const pointer = Number(collectElementText(body, 'v').join(''));
          cells.push(sharedStrings[pointer] ?? '');
        } else if (type === 'inlineStr') {
          cells.push(collectElementText(body, 't').join(''));
        } else {
          cells.push(collectElementText(body, 'v').join(''));
        }
      }
      if (cells.some((value) => value !== '')) rows.push(cells.join('\t'));
    }
    if (rows.length > 0) parts.push(`[Hoja ${index + 1}]\n${rows.join('\n')}`);
  });

  const bounded = boundText(parts.join('\n\n'), maxChars);
  return {
    format: 'xlsx',
    status: 'complete',
    ...bounded,
    sheets: sheetNames.length,
    warnings: []
  };
}

/**
 * Decode a PDF literal string, resolving backslash and octal escapes.
 * @param {string} raw
 * @returns {string}
 */
function decodePdfLiteral(raw) {
  let output = '';
  for (let index = 0; index < raw.length; index += 1) {
    const char = raw[index];
    if (char !== '\\') {
      output += char;
      continue;
    }
    const next = raw[index + 1];
    if (next === undefined) break;
    if (next >= '0' && next <= '7') {
      const octal = /^[0-7]{1,3}/.exec(raw.slice(index + 1))[0];
      output += String.fromCharCode(parseInt(octal, 8));
      index += octal.length;
      continue;
    }
    const simple = { n: '\n', r: '\r', t: '\t', b: '\b', f: '\f', '(': '(', ')': ')', '\\': '\\' };
    if (next in simple) {
      output += simple[next];
      index += 1;
      continue;
    }
    if (next === '\n') {
      index += 1; // Line continuation inside a literal.
      continue;
    }
    output += next;
    index += 1;
  }
  return output;
}

/**
 * Pull text out of a PDF content stream by walking the text-showing operators.
 *
 * This handles the common case of Tj/TJ/'/" with literal strings. It does NOT
 * resolve font encoding differences (CID fonts, custom /Differences maps), so
 * some PDFs yield mojibake or nothing. That is why the caller reports 'partial'
 * whenever the yield looks poor, instead of claiming the PDF is empty.
 *
 * @param {string} stream  Content stream decoded as latin1
 * @returns {string}
 */
function pdfStreamText(stream) {
  const lines = [];
  const blockPattern = /BT([\s\S]*?)ET/g;
  let block;
  while ((block = blockPattern.exec(stream)) !== null) {
    const body = block[1];
    const parts = [];
    const showPattern = /\((?:\\.|[^\\()])*\)\s*(?:Tj|TJ|'|")|\[((?:[^\][\\]|\\.)*)\]\s*TJ|T\*|Td|TD/g;
    let token;
    while ((token = showPattern.exec(body)) !== null) {
      const raw = token[0];
      if (raw === 'T*' || raw === 'Td' || raw === 'TD') {
        parts.push('\n');
        continue;
      }
      if (token[1] !== undefined) {
        // TJ array: concatenate its literal strings, ignore kerning numbers.
        const literals = token[1].match(/\((?:\\.|[^\\()])*\)/g) || [];
        parts.push(literals.map((item) => decodePdfLiteral(item.slice(1, -1))).join(''));
        continue;
      }
      const literal = /\((?:\\.|[^\\()])*\)/.exec(raw);
      if (literal) parts.push(decodePdfLiteral(literal[0].slice(1, -1)));
    }
    const text = parts.join('').replace(/\n{2,}/g, '\n').trim();
    if (text) lines.push(text);
  }
  return lines.join('\n');
}

/**
 * @param {Buffer} buffer
 * @param {number} maxChars
 * @returns {object}
 */
function processPdf(buffer, maxChars) {
  const zlib = require('zlib');
  const latin = buffer.toString('latin1');

  if (!latin.startsWith('%PDF-')) {
    return {
      format: 'pdf',
      status: 'failed',
      text: '',
      truncated: false,
      warnings: ['El archivo no tiene una cabecera PDF valida']
    };
  }

  const chunks = [];
  let streamCount = 0;
  let inflateFailures = 0;

  // Walk every `stream ... endstream` pair; decompress the Flate ones.
  const pattern = /stream\r?\n?/g;
  let match;
  while ((match = pattern.exec(latin)) !== null) {
    const start = match.index + match[0].length;
    const end = latin.indexOf('endstream', start);
    if (end === -1) break;
    pattern.lastIndex = end;
    streamCount += 1;

    const dictionaryStart = Math.max(0, latin.lastIndexOf('<<', match.index));
    const dictionary = latin.slice(dictionaryStart, match.index);
    const raw = Buffer.from(latin.slice(start, end), 'latin1');

    let decoded = null;
    if (/\/FlateDecode/.test(dictionary)) {
      try {
        decoded = zlib.inflateSync(raw);
      } catch {
        try {
          decoded = zlib.inflateRawSync(raw);
        } catch {
          inflateFailures += 1;
        }
      }
    } else if (!/\/(?:DCTDecode|JPXDecode|CCITTFaxDecode|JBIG2Decode|RunLengthDecode|LZWDecode)/.test(dictionary)) {
      decoded = raw;
    }

    if (!decoded) continue;
    const text = pdfStreamText(decoded.toString('latin1'));
    if (text) chunks.push(text);
  }

  const combined = chunks.join('\n\n').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
  const bounded = boundText(combined, maxChars);
  const warnings = [];

  if (!combined) {
    warnings.push(
      'No se extrajo texto del PDF. Puede ser un PDF escaneado o usar codificacion de fuente no soportada; usa onedrive-export-file para obtener los bytes'
    );
  } else {
    warnings.push(
      'La extraccion de PDF no resuelve codificaciones de fuente personalizadas ni imagenes; verifica el texto antes de citarlo'
    );
  }
  if (inflateFailures > 0) {
    warnings.push(`${inflateFailures} flujo(s) del PDF no se pudieron descomprimir`);
  }

  return {
    format: 'pdf',
    status: combined ? 'partial' : 'failed',
    ...bounded,
    streams: streamCount,
    warnings
  };
}

/**
 * Turn a staged buffer into text, dispatching on the filename extension.
 * @param {Buffer} buffer
 * @param {string} filename
 * @param {number} [maxChars]
 * @returns {object}
 */
function extractText(buffer, filename, maxChars = DEFAULT_MAX_CHARS) {
  const extension = (/\.[^.]+$/.exec(String(filename || '').toLowerCase()) || [''])[0];

  try {
    if (extension === '.docx') return processDocx(buffer, maxChars);
    if (extension === '.pptx') return processPptx(buffer, maxChars);
    if (extension === '.xlsx') return processXlsx(buffer, maxChars);
    if (extension === '.pdf') return processPdf(buffer, maxChars);

    if (extension === '.html' || extension === '.htm') {
      const bounded = boundText(sanitizeHtmlToText(buffer.toString('utf8')), maxChars);
      return { format: extension.slice(1), status: 'complete', ...bounded, warnings: [] };
    }

    if (extension === '.doc' || extension === '.ppt' || extension === '.xls') {
      return {
        format: extension.slice(1),
        status: 'failed',
        text: '',
        truncated: false,
        warnings: [
          `El formato binario ${extension} de Office antiguo no se puede extraer; convierte a ${extension}x o usa onedrive-export-file`
        ]
      };
    }

    if (TEXT_EXTENSIONS.has(extension) || extension === '') {
      const text = buffer.toString('utf8');
      // A NUL byte means we decoded binary as text; do not hand that to a model.
      if (text.includes('\u0000')) {
        return {
          format: 'binary',
          status: 'failed',
          text: '',
          truncated: false,
          warnings: ['El archivo parece binario; usa onedrive-export-file para transferir los bytes']
        };
      }
      const bounded = boundText(text, maxChars);
      return { format: extension.slice(1) || 'text', status: 'complete', ...bounded, warnings: [] };
    }

    return {
      format: extension.slice(1) || 'unknown',
      status: 'failed',
      text: '',
      truncated: false,
      warnings: [
        `La extraccion de texto no soporta ${extension || 'este tipo'}; usa onedrive-export-file para transferir los bytes`
      ]
    };
  } catch (error) {
    return {
      format: extension.slice(1) || 'unknown',
      status: 'failed',
      text: '',
      truncated: false,
      warnings: [`No se pudo procesar el archivo (${error.name}: ${error.message})`]
    };
  }
}

module.exports = {
  extractText,
  boundText,
  decodeXmlEntities,
  collectElementText,
  DEFAULT_MAX_CHARS,
  MAX_RETURNED_CHARS,
  TEXT_EXTENSIONS
};
