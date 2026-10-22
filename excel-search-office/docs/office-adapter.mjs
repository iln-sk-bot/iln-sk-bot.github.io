/**
 * Office.js workbook adapter for Excel Search v0.2.0.
 * Read-only. It does not change the workbook while indexing data.
 * Compatible with the cache's absolute 1-based row / 0-based column layout.
 */
export const MAX_EXCEL_ROWS = 1048576;
export const MAX_EXCEL_COLUMNS = 16384;

export function columnLetter(index) {
  if (!Number.isInteger(index) || index < 0 || index >= MAX_EXCEL_COLUMNS) throw new RangeError(`Invalid column index: ${index}`);
  let name = '';
  for (let n = index + 1; n; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name;
  return name;
}

export function a1Range(rowIndex, columnIndex, rowCount, columnCount) {
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex + rowCount > MAX_EXCEL_ROWS ||
      !Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex + columnCount > MAX_EXCEL_COLUMNS ||
      !Number.isInteger(rowCount) || rowCount <= 0 || !Number.isInteger(columnCount) || columnCount <= 0) {
    throw new RangeError('Invalid Excel range geometry');
  }
  return `${columnLetter(columnIndex)}${rowIndex + 1}:${columnLetter(columnIndex + columnCount - 1)}${rowIndex + rowCount}`;
}

export function convertOfficeBlock(values, text, rowIndex, columnIndex) {
  const records = [];
  for (let r = 0; r < values.length; r++) {
    const sourceRow = values[r] || [];
    const formattedRow = text?.[r] || [];
    for (let c = 0; c < sourceRow.length; c++) {
      const raw = sourceRow[c];
      const display = formattedRow[c] == null ? String(raw ?? '') : String(formattedRow[c]);
      if ((raw === '' || raw == null) && !display) continue;
      const type = typeof raw === 'number' ? 'n' : typeof raw === 'boolean' ? 'b' : 's';
      // Unlike the original file loader, dates remain Excel serial values in v0.2.0.
      records.push([rowIndex + r + 1, columnIndex + c, display, raw, type]);
    }
  }
  return records;
}

export function tablesFromGeometry(raw, cells) {
  const byCell = new Map(cells.map(([r,c,text]) => [`${r}:${c}`, text]));
  return raw.map(info => {
    const rowIndex = info.rowIndex, columnIndex = info.columnIndex;
    const headerRowCount = info.showHeaders ? 1 : 0;
    const columns = Array.from({length:info.columnCount}, (_, i) => {
      const label = headerRowCount ? byCell.get(`${rowIndex + 1}:${columnIndex + i}`) : '';
      return String(label || columnLetter(columnIndex + i));
    });
    return {
      name: info.name,
      displayName: info.name,
      ref: a1Range(rowIndex, columnIndex, info.rowCount, info.columnCount),
      headerRowCount,
      totalsRowCount: info.showTotals ? 1 : 0,
      columns
    };
  });
}

/**
 * @param {object} excel Excel global from Office.js (or a test double)
 * @param {object} options onProgress, maxCells (combined rectangular size), batchCells
 * @returns {Promise<Array>} Cached worksheet descriptors for the original v0.5 search engine.
 */
export async function readWorkbookSnapshot(excel, options = {}) {
  if (!excel || typeof excel.run !== 'function') throw new Error('Excel.run is unavailable');
  const maxCells = options.maxCells ?? 2_000_000;
  const batchCells = options.batchCells ?? 8_000;
  if (maxCells < 1 || batchCells < 1) throw new Error('Invalid read limit');
  return excel.run(async context => {
    const sheetsCollection = context.workbook.worksheets;
    sheetsCollection.load('items/name');
    await context.sync();
    const descriptions = sheetsCollection.items.map(sheet => {
      const used = sheet.getUsedRange(true);
      used.load('rowIndex,columnIndex,rowCount,columnCount');
      return {sheet, used};
    });
    await context.sync();
    let cellCount = 0;
    for (const {sheet,used} of descriptions) {
      const count = used.rowCount * used.columnCount;
      cellCount += count;
      if (cellCount > maxCells) {
        throw new Error(`読み込み範囲が ${maxCells.toLocaleString()} セルを超えています。` +
          `シート「${sheet.name}」までで ${cellCount.toLocaleString()} セルです。` +
          '未使用セルの書式が大きく広がっていないか確認してください。');
      }
    }
    const results = [];
    for (const {sheet,used} of descriptions) {
      const cells = [];
      const chunkRows = Math.max(1,Math.floor(batchCells / Math.max(used.columnCount,1)));
      // All Office API calls stay in the originating Excel.run context.
      for (let offset = 0; offset < used.rowCount; offset += chunkRows) {
        const count = Math.min(chunkRows, used.rowCount - offset);
        const chunk = sheet.getRangeByIndexes(used.rowIndex + offset, used.columnIndex, count, used.columnCount);
        chunk.load('values,text');
        await context.sync();
        cells.push(...convertOfficeBlock(chunk.values, chunk.text, used.rowIndex + offset, used.columnIndex));
        options.onProgress?.({sheet:sheet.name, completedRows:offset + count, rowCount:used.rowCount});
      }
      // Read the table metadata only after the values have been retrieved.
      const tableCollection = sheet.tables;
      tableCollection.load('items/name,items/showHeaders,items/showTotals');
      await context.sync();
      const tables = tableCollection.items.map(table => {
        const range = table.getRange();
        range.load('rowIndex,columnIndex,rowCount,columnCount');
        return {table, range};
      });
      if (tables.length) await context.sync();
      const descriptors = tables.map(({table,range}) => ({
        name:table.name,
        showHeaders:table.showHeaders,
        showTotals:table.showTotals,
        rowIndex:range.rowIndex,
        columnIndex:range.columnIndex,
        rowCount:range.rowCount,
        columnCount:range.columnCount
      }));
      results.push({
        name: sheet.name,
        minRow: used.rowIndex + 1,
        maxRow: used.rowIndex + used.rowCount,
        minCol: used.columnIndex,
        maxCol: used.columnIndex + used.columnCount - 1,
        cells,
        tables: tablesFromGeometry(descriptors,cells)
      });
    }
    return results;
  });
}

export async function selectWorkbookCell(excel, sheetName, a1) {
  if (!sheetName || !/^\$?[A-Z]{1,3}\$?[1-9][0-9]*$/i.test(a1)) throw new Error('Invalid Excel cell address');
  return excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.activate();
    sheet.getRange(a1).select();
    await context.sync();
  });
}
