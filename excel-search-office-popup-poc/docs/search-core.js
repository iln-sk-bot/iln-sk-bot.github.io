/* Shared search logic: no Office.js or DOM dependencies. */
(function (root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.SearchCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  function columnName(index) {
    if (!Number.isInteger(index) || index < 0) throw new Error('Invalid column index');
    let i = index + 1;
    let result = '';
    while (i > 0) {
      i -= 1;
      result = String.fromCharCode(65 + i % 26) + result;
      i = Math.floor(i / 26);
    }
    return result;
  }
  function columnIndex(name) {
    if (!/^[A-Za-z]{1,3}$/.test(name)) throw new Error(`列指定が不正です: ${name}`);
    let value = 0;
    for (const char of name.toUpperCase()) value = value * 26 + char.charCodeAt(0) - 64;
    const index = value - 1;
    if (index > 16383) throw new Error(`列指定が範囲外です: ${name}`);
    return index;
  }
  function parseColumnFilter(input) {
    const trimmed = String(input || '').trim();
    if (!trimmed) return null;
    const columns = new Set();
    for (const chunk of trimmed.split(',')) {
      const name = chunk.trim();
      if (!name) throw new Error('対象列は A,C のように入力してください。');
      columns.add(columnIndex(name));
    }
    return columns;
  }
  function compileMatcher(query, { regex = false, caseSensitive = false } = {}) {
    const source = String(query || '');
    if (!source) return () => true;
    if (regex) {
      let expression;
      try { expression = new RegExp(source, caseSensitive ? 'u' : 'iu'); }
      catch (err) { throw new Error(`正規表現が不正です: ${err.message}`); }
      return (text) => expression.test(String(text));
    }
    const normalized = caseSensitive ? source : source.toLocaleLowerCase();
    return (text) => (caseSensitive ? String(text) : String(text).toLocaleLowerCase()).includes(normalized);
  }
  function search(datasets, query, options = {}) {
    const match = compileMatcher(query, options);
    const columns = parseColumnFilter(options.columns);
    const limit = Number.isInteger(options.limit) ? Math.max(0, options.limit) : 300;
    const skipHeader = options.skipHeader !== false;
    const results = [];
    let scanned = 0;
    let total = 0;
    const available = new Set();
    for (const sheet of datasets || []) {
      if (!sheet || !Array.isArray(sheet.rows)) continue;
      const startColumn = sheet.startColumn || 0;
      for (let i = skipHeader ? 1 : 0; i < sheet.rows.length; i++) {
        const row = sheet.rows[i];
        scanned++;
        const matchColumns = [];
        for (let j = 0; j < row.cells.length; j++) {
          const absoluteColumn = startColumn + j;
          available.add(absoluteColumn);
          if (columns && !columns.has(absoluteColumn)) continue;
          if (match(row.cells[j])) matchColumns.push(absoluteColumn);
        }
        if (!matchColumns.length) continue;
        total++;
        if (results.length < limit) {
          results.push({
            sheetId: sheet.id, sheetName: sheet.name,
            row: row.row, startColumn, cells: row.cells,
            matchColumns,
          });
        }
      }
    }
    const displayedColumns = [...available].sort((a, b) => a - b).slice(0, 40);
    return { results, scanned, total, truncated: total > results.length, displayedColumns };
  }
  function cellAt(row, column) {
    return row && column >= row.startColumn && column < row.startColumn + row.cells.length
      ? row.cells[column - row.startColumn] : '';
  }
  return { columnName, columnIndex, parseColumnFilter, compileMatcher, search, cellAt };
});
