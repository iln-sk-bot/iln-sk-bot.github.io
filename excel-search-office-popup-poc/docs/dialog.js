/* global Office, SearchCore */
(() => {
  'use strict';
  const MAX_VISIBLE = 300;
  const state = {
    connected: false,
    demo: new URLSearchParams(window.location.search).get('demo') === '1',
    availableSheets: [],
    activeId: null,
    limits: { sheets: 5, rowsPerSheet: 1500, columnsPerSheet: 32 },
    datasets: [],
    activeLoad: null,
    loaded: false,
    selectedCell: null,
    displayedRows: [],
    displayedColumns: [],
    requestNo: 0,
    pending: new Map(),
  };
  const $ = (id) => document.getElementById(id);
  const els = {
    connection: $('connectionLabel'),
    checks: $('sheetChecks'),
    load: $('loadSheets'),
    refresh: $('refreshSheetList'),
    search: $('searchInput'),
    regex: $('regexMode'),
    caseSensitive: $('caseSensitive'),
    skipHeader: $('skipHeader'),
    columns: $('columnFilter'),
    notice: $('notice'),
    warning: $('warning'),
    resultCount: $('resultCount'),
    resultDescription: $('resultDescription'),
    resultsHead: $('resultsHead'),
    resultsBody: $('resultsBody'),
    loadSummary: $('loadSummary'),
    selectionStatus: $('selectionStatus'),
    copyTsv: $('copyTsv'),
    copyCell: $('copyCell'),
  };
  function requestId() { state.requestNo += 1; return `req${state.requestNo}`; }
  function inform(message, isError = false) {
    els.notice.textContent = message;
    els.notice.classList.toggle('has-error', isError);
  }
  function warn(message) {
    els.warning.hidden = !message;
    els.warning.textContent = message || '';
  }
  function send(message) {
    if (state.demo) {
      if (message.type === 'listSheets' || message.type === 'hello') renderDemoSheets();
      else if (message.type === 'loadSheets') fillDemoData();
      else if (message.type === 'selectCell') {
        state.selectedCell = { ...message, text: state.selectedCell?.text || '' };
        els.selectionStatus.textContent = `デモ: ${message.row + 1}行 ${SearchCore.columnName(message.column)}列`;
      } else if (message.type === 'close') window.close();
      return;
    }
    Office.context.ui.messageParent(JSON.stringify(message));
  }
  function acceptMessage(arg) {
    let msg;
    try { msg = JSON.parse(arg.message); }
    catch (error) { inform('親ランタイムから不正なJSONを受信しました。', true); return; }
    if (!msg || !msg.type) return;
    switch (msg.type) {
      case 'sheets':
        state.availableSheets = msg.sheets;
        state.activeId = msg.activeId;
        state.limits = msg.limits || state.limits;
        state.connected = true;
        renderSheetChecks();
        inform(`${msg.sheets.length}シートを取得しました。対象を選択して読み込んでください。`);
        els.connection.textContent = 'Excelに接続中';
        break;
      case 'loadStarted':
        if (state.activeLoad !== msg.requestId) return;
        state.datasets = [];
        state.loaded = false;
        state.selectedCell = null;
        els.copyCell.disabled = true;
        els.load.disabled = true;
        warn('');
        inform(`${msg.count}シートを読み込み中…`);
        els.loadSummary.textContent = '';
        renderResults();
        break;
      case 'sheetStart':
        if (state.activeLoad !== msg.requestId) return;
        state.datasets.push({ ...msg.sheet, rows: [] });
        inform(`読み込み中: ${msg.sheet.name} (${msg.sheet.rowCount}行 × ${msg.sheet.columnCount}列)`);
        break;
      case 'chunk':
        // Always ACK every frame: do not leave the sender waiting on a stale request.
        if (state.activeLoad === msg.requestId) {
          const target = state.datasets.find((s) => s.id === msg.sheetId);
          if (target) target.rows.push(...msg.rows);
        }
        send({ type: 'ack', requestId: msg.requestId, seq: msg.seq });
        break;
      case 'sheetDone':
        break;
      case 'loadDone':
        if (state.activeLoad !== msg.requestId) return;
        state.activeLoad = null;
        state.loaded = true;
        els.load.disabled = false;
        inform(`${state.datasets.length}シートの読み込みが完了しました。検索できます。`);
        updateLoadSummary();
        renderResults();
        break;
      case 'selected':
        els.selectionStatus.textContent = `Excelで ${state.availableSheets.find(s => s.id === msg.sheetId)?.name || ''}!${SearchCore.columnName(msg.column)}${msg.row + 1} を選択`;
        break;
      case 'error':
        if (state.activeLoad === msg.requestId) state.activeLoad = null;
        els.load.disabled = false;
        inform(msg.message || 'Excel操作に失敗しました。', true);
        break;
      case 'notice': inform(msg.message); break;
      default: break;
    }
  }
  function renderSheetChecks() {
    els.checks.replaceChildren();
    if (!state.availableSheets.length) {
      els.checks.textContent = '表示可能なシートがありません。';
      els.load.disabled = true;
      return;
    }
    for (const sheet of state.availableSheets) {
      const label = document.createElement('label');
      label.className = 'sheet-pill';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = sheet.id;
      cb.checked = sheet.id === state.activeId;
      cb.addEventListener('change', updateLoadButton);
      label.append(cb, document.createTextNode(sheet.name));
      els.checks.append(label);
    }
    updateLoadButton();
  }
  function updateLoadButton() {
    const count = selectedIds().length;
    els.load.disabled = !state.connected || count === 0 || count > state.limits.sheets || !!state.activeLoad;
    if (count > state.limits.sheets) {
      warn(`検証版では最大${state.limits.sheets}シートまで選択できます。`);
    } else if (!state.loaded) {
      warn('');
    }
  }
  function selectedIds() {
    return [...els.checks.querySelectorAll('input[type="checkbox"]:checked')].map((x) => x.value);
  }
  function updateLoadSummary() {
    const rows = state.datasets.reduce((sum, s) => sum + s.rows.length, 0);
    els.loadSummary.textContent = `${state.datasets.length}シート / ${rows.toLocaleString()}行 キャッシュ済み`;
    const notes = [];
    for (const s of state.datasets) {
      if (s.originalRows > s.rowCount || s.originalColumns > s.columnCount) {
        notes.push(`${s.name}: ${s.originalRows}行×${s.originalColumns}列中、${s.rowCount}行×${s.columnCount}列だけ読み込み`);
      }
    }
    warn(notes.join(' ／ '));
  }
  function startLoad() {
    const ids = selectedIds();
    if (!ids.length) return;
    if (ids.length > state.limits.sheets) return;
    const id = requestId();
    state.activeLoad = id;
    els.load.disabled = true;
    inform('Excelへの読み込み要求を送信しています…');
    send({ type: 'loadSheets', requestId: id, ids });
  }
  function headerForColumn(col) {
    const letter = SearchCore.columnName(col);
    if (!els.skipHeader.checked || state.datasets.length !== 1) return letter;
    const sheet = state.datasets[0];
    if (!sheet.rows.length || col < sheet.startColumn || col >= sheet.startColumn + sheet.columnCount) return letter;
    const title = String(sheet.rows[0].cells[col - sheet.startColumn] || '').trim();
    return title ? `${letter} · ${title.slice(0, 30)}` : letter;
  }
  function renderResults() {
    els.resultsHead.replaceChildren();
    els.resultsBody.replaceChildren();
    els.copyTsv.disabled = true;
    state.displayedRows = [];
    state.displayedColumns = [];
    if (!state.loaded) {
      els.resultCount.textContent = '検索結果';
      els.resultDescription.textContent = state.activeLoad ? '読み込み中…' : '未読み込み';
      showEmpty('シートを選んで［読み込む］を押してください。');
      return;
    }
    let matches;
    try {
      matches = SearchCore.search(state.datasets, els.search.value, {
        regex: els.regex.checked,
        caseSensitive: els.caseSensitive.checked,
        skipHeader: els.skipHeader.checked,
        columns: els.columns.value,
        limit: MAX_VISIBLE,
      });
    } catch (error) {
      els.resultCount.textContent = '検索条件エラー';
      els.resultDescription.textContent = error.message;
      showEmpty('検索条件を修正してください。');
      return;
    }
    state.displayedRows = matches.results;
    state.displayedColumns = matches.displayedColumns;
    els.resultCount.textContent = `${matches.total.toLocaleString()} 件`;
    els.resultDescription.textContent = `${matches.scanned.toLocaleString()}行を検索 / 表示 ${matches.results.length.toLocaleString()}件${matches.truncated ? '（先頭300件まで）' : ''}`;
    els.copyTsv.disabled = matches.results.length === 0;
    const headRow = document.createElement('tr');
    headRow.append(makeHead('シート', 'sheet-col'), makeHead('行', 'row-col'));
    for (const col of matches.displayedColumns) headRow.append(makeHead(headerForColumn(col), 'data-col'));
    els.resultsHead.append(headRow);
    if (!matches.results.length) {
      showEmpty('一致する行がありません。');
      return;
    }
    const fragment = document.createDocumentFragment();
    for (const [index, row] of matches.results.entries()) {
      const tr = document.createElement('tr');
      tr.className = index % 2 === 0 ? 'even-row' : '';
      const sheetCell = document.createElement('td');
      sheetCell.className = 'sheet-col';
      sheetCell.textContent = row.sheetName;
      sheetCell.title = row.sheetName;
      tr.append(sheetCell);
      const number = document.createElement('td');
      number.className = 'row-col';
      const rowBtn = document.createElement('button');
      rowBtn.type = 'button';
      rowBtn.className = 'cell-button row-button';
      rowBtn.textContent = String(row.row + 1);
      rowBtn.title = `${row.sheetName} の ${row.row + 1} 行に移動`;
      rowBtn.addEventListener('click', () => selectCell(row, row.startColumn));
      number.append(rowBtn);
      tr.append(number);
      for (const col of matches.displayedColumns) {
        const td = document.createElement('td');
        td.className = 'data-col';
        const val = SearchCore.cellAt(row, col);
        if (col >= row.startColumn && col < row.startColumn + row.cells.length) {
          const btn = document.createElement('button');
          btn.className = 'cell-button';
          if (row.matchColumns.includes(col) && els.search.value) btn.classList.add('matched-cell');
          btn.title = `${row.sheetName}!${SearchCore.columnName(col)}${row.row + 1}: ${val.slice(0, 100)}`;
          btn.textContent = val || '\u00a0';
          btn.addEventListener('click', () => selectCell(row, col));
          td.append(btn);
        }
        tr.append(td);
      }
      fragment.append(tr);
    }
    els.resultsBody.append(fragment);
  }
  function makeHead(text, cls) {
    const th = document.createElement('th');
    th.className = cls;
    th.textContent = text;
    th.title = text;
    return th;
  }
  function showEmpty(text) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.className = 'empty-placeholder';
    td.textContent = text;
    tr.append(td);
    els.resultsBody.append(tr);
  }
  function selectCell(row, col) {
    const text = SearchCore.cellAt(row, col);
    state.selectedCell = { sheetId: row.sheetId, row: row.row, column: col, text };
    els.copyCell.disabled = false;
    els.selectionStatus.textContent = `選択要求: ${row.sheetName}!${SearchCore.columnName(col)}${row.row + 1}`;
    send({ type: 'selectCell', requestId: requestId(), sheetId: row.sheetId, row: row.row, column: col });
  }
  async function copyText(text) {
    if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
      try { await navigator.clipboard.writeText(text); return; }
      catch (_ignored) { /* Try classic clipboard fallback. */ }
    }
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.cssText = 'position:fixed;left:-10000px;top:0;opacity:0';
    document.body.append(ta);
    ta.focus();
    ta.select();
    const result = document.execCommand('copy');
    ta.remove();
    if (!result) throw new Error('Clipboard APIと従来のコピー処理が拒否されました。');
  }
  function asTsvField(value) {
    return String(value == null ? '' : value).replace(/[\r\n\t]+/g, ' ');
  }
  function makeTsv() {
    const cols = state.displayedColumns;
    const header = ['シート', '行', ...cols.map((i) => SearchCore.columnName(i))].join('\t');
    const body = state.displayedRows.map((row) => [
      row.sheetName, row.row + 1,
      ...cols.map((col) => SearchCore.cellAt(row, col)),
    ].map(asTsvField).join('\t'));
    return [header, ...body].join('\r\n');
  }
  async function handleCopy(text, label) {
    try { await copyText(text); inform(`${label}をコピーしました。`); }
    catch (error) { inform(`コピーできません: ${error.message}`, true); }
  }
  function renderDemoSheets() {
    state.availableSheets = [{ id: 'demo-orders', name: '受注データ' }, { id: 'demo-parts', name: '部品マスタ' }];
    state.activeId = 'demo-orders';
    state.connected = true;
    els.connection.textContent = 'デモモード（Excelには接続していません）';
    renderSheetChecks();
    inform('デモデータを使って画面操作を確認できます。');
  }
  function fillDemoData() {
    const chosen = new Set(selectedIds());
    const build = (id, name, rows) => ({
      id, name, startRow: 0, startColumn: 0, originalRows: rows.length,
      originalColumns: rows[0].length, rowCount: rows.length, columnCount: rows[0].length,
      rows: rows.map((cells, index) => ({ row: index, cells })),
    });
    state.datasets = [
      build('demo-orders', '受注データ', [
        ['注文番号', '顧客', '品番', '数量', '状態', '備考'],
        ['O-001', '東京商事', 'AX-120', '12', '処理中', '早期納品希望'],
        ['O-002', '関西工業', 'BZ-900', '5', '完了', '検収済'],
        ['O-003', '東京商事', 'CX-021', '50', '処理中', '納期調整'],
        ['O-004', '名古屋精機', 'AX-120', '8', '未着手', '優先対応'],
        ['O-005', '札幌部品', 'RT-800', '20', '完了', ''],
      ]),
      build('demo-parts', '部品マスタ', [
        ['品番', '品名', '分類', '原価', '在庫'],
        ['AX-120', 'ベアリング', '機械', '350', '120'],
        ['BZ-900', 'モーター', '電装', '18500', '15'],
        ['CX-021', 'シャフト', '機械', '1200', '81'],
        ['RT-800', 'スイッチ', '電装', '860', '205'],
      ]),
    ].filter((x) => chosen.has(x.id));
    state.activeLoad = null;
    state.loaded = true;
    els.load.disabled = false;
    updateLoadSummary();
    inform('デモデータを読み込みました。');
    renderResults();
  }
  function bindEvents() {
    els.load.addEventListener('click', startLoad);
    els.refresh.addEventListener('click', () => send({ type: 'listSheets', requestId: requestId() }));
    for (const el of [els.search, els.columns]) el.addEventListener('input', renderResults);
    for (const el of [els.regex, els.caseSensitive, els.skipHeader]) el.addEventListener('change', renderResults);
    $('clearSearch').addEventListener('click', () => {
      els.search.value = '';
      renderResults();
      els.search.focus();
    });
    els.copyTsv.addEventListener('click', () => void handleCopy(makeTsv(), '表示結果のTSV'));
    els.copyCell.addEventListener('click', () => {
      if (state.selectedCell) void handleCopy(state.selectedCell.text, '選択セルの文字列');
    });
    $('closeDialog').addEventListener('click', () => send({ type: 'close' }));
  }
  bindEvents();
  if (state.demo) {
    renderDemoSheets();
    fillDemoData();
  } else {
    Office.onReady(() => {
      Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        acceptMessage,
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            inform('親ランタイムとのメッセージ受信登録に失敗しました: ' + (result.error?.message || ''), true);
            return;
          }
          els.connection.textContent = '親ランタイムへの接続を確認中…';
          send({ type: 'hello', requestId: requestId() });
        }
      );
    }).catch((error) => inform(`Office初期化エラー: ${error.message}`, true));
  }
})();
