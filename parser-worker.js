// A blob Worker inherits the page's CSP, including connect-src 'none'.
// A separate worker URL on GitHub Pages would not inherit that policy.
function createDictionaryParserWorker() {
  const libraryUrl = new URL('vendor/xlsx-0.20.3.full.min.js', document.baseURI).href;
  const source = `(${dictionaryParserWorker.toString()})(${JSON.stringify(libraryUrl)})`;
  const url = URL.createObjectURL(new Blob([source], {type: 'text/javascript'}));
  try { return new Worker(url); }
  finally { URL.revokeObjectURL(url); }
}

function dictionaryParserWorker(libraryUrl) {
  'use strict';
  importScripts(libraryUrl);
  const MAX_ROWS = 250001; // Header plus 250,000 data rows.
  const MAX_COLUMNS = 256;
  const MAX_CELLS = 2000000;
  const MAX_EXPANDED_BYTES = 100 * 1024 * 1024;

  function checkZipSize(buffer) {
    const bytes = new DataView(buffer);
    if (buffer.byteLength < 4 || bytes.getUint32(0, true) !== 0x04034b50) return;
    let end = -1;
    for (let i = buffer.byteLength - 22; i >= Math.max(0, buffer.byteLength - 65557); i--) {
      if (bytes.getUint32(i, true) === 0x06054b50 && i + 22 + bytes.getUint16(i + 20, true) === buffer.byteLength) { end = i; break; }
    }
    if (end < 0) throw new Error('Invalid ZIP workbook directory.');
    const count = bytes.getUint16(end + 10, true);
    let pos = bytes.getUint32(end + 16, true), expanded = 0;
    if (!count || count > 4096 || pos === 0xffffffff || pos >= end || bytes.getUint16(end + 4, true) || bytes.getUint16(end + 6, true)) throw new Error('Workbook ZIP is too large or uses an unsupported archive format.');
    for (let i = 0; i < count; i++) {
      if (pos + 46 > end || bytes.getUint32(pos, true) !== 0x02014b50) throw new Error('Invalid ZIP workbook directory.');
      const size = bytes.getUint32(pos + 24, true);
      if (size === 0xffffffff || (expanded += size) > MAX_EXPANDED_BYTES) throw new Error('Workbook expands beyond the 100 MB safety limit.');
      pos += 46 + bytes.getUint16(pos + 28, true) + bytes.getUint16(pos + 30, true) + bytes.getUint16(pos + 32, true);
    }
  }

  self.onmessage = event => {
    try {
      const buffer = event.data;
      checkZipSize(buffer);
      const wb = XLSX.read(buffer, {type: 'array', raw: true, cellDates: false, cellHTML: false, cellFormula: false, bookVBA: false, sheetRows: MAX_ROWS + 1});
      if (!wb.SheetNames.length) throw new Error('No sheets found.');
      if (wb.SheetNames.length > 50) throw new Error('Workbook exceeds the 50-tab safety limit.');
      let rowCount = 0, cellCount = 0;
      const sheets = wb.SheetNames.map(name => {
        const sheet = wb.Sheets[name];
        const ref = sheet['!fullref'] || sheet['!ref'];
        if (ref) {
          const range = XLSX.utils.decode_range(ref);
          const rows = range.e.r + 1, columns = range.e.c + 1;
          if (!Number.isSafeInteger(rows) || !Number.isSafeInteger(columns) || rows > MAX_ROWS || columns > MAX_COLUMNS) throw new Error(`Tab ${name} exceeds 250,000 data rows or 256 columns.`);
          cellCount += rows * columns;
          if (cellCount > MAX_CELLS) throw new Error('Workbook exceeds the 2 million cell safety limit.');
        }
        const rows = XLSX.utils.sheet_to_json(sheet, {header: 1, raw: false, defval: ''});
        rowCount += rows.length;
        if (rowCount > MAX_ROWS) throw new Error('Workbook exceeds the 250,000-row safety limit across all tabs.');
        return {name, rows};
      });
      self.postMessage({sheets, rowCount});
    } catch (error) {
      self.postMessage({error: error && error.message ? error.message : 'Could not read this workbook.'});
    }
  };
}
