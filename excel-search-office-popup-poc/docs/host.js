/* global Office, Excel */
(() => {
  'use strict';

  const LIMITS = Object.freeze({
    sheets: 5,
    rowsPerSheet: 1500,
    columnsPerSheet: 32,
    chunkRows: 40,
    cellChars: 500,
    messageChars: 60000,
    ackTimeoutMs: 18000,
  });
  const dialogUrl = new URL('dialog.html', window.location.href).href;
  const state = {
    dialog: null,
    opening: false,
    loading: false,
    sheets: [],
    ackWaiters: new Map(),
  };
  const ready = new Promise((resolve) => Office.onReady(resolve));

  function log(message) {
    const line = `${new Date().toLocaleTimeString()}  ${message}`;
    console.log('[Excel Search]', line);
    const el = document.getElementById('diagnosticLog');
    if (el) {
      const lines = [line, ...el.textContent.split('\n').filter(Boolean)];
      el.textContent = lines.slice(0, 30).join('\n');
    }
  }
  function setStatus(message, error = false) {
    const el = document.getElementById('hostStatus');
    if (el) {
      el.textContent = message;
      el.classList.toggle('error', error);
    }
    log(message);
  }
  function stringifyError(err) {
    return (err && (err.message || err.code)) ? `${err.code ? err.code + ': ' : ''}${err.message || err}` : String(err);
  }
  function send(message) {
    if (!state.dialog) throw new Error('ダイアログとの接続がありません。');
    state.dialog.messageChild(JSON.stringify(message));
  }
  function cancelAcks(reason) {
    for (const item of state.ackWaiters.values()) {
      clearTimeout(item.timer);
      item.reject(new Error(reason));
    }
    state.ackWaiters.clear();
  }
  async function sendAndWaitAck(message) {
    const key = `${message.requestId}:${message.seq}`;
    const accepted = new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        state.ackWaiters.delete(key);
        reject(new Error('チャンク受信確認がタイムアウトしました。ウィンドウを閉じて開き直してください。'));
      }, LIMITS.ackTimeoutMs);
      state.ackWaiters.set(key, { resolve, reject, timer });
    });
    try {
      send(message);
    } catch (error) {
      const item = state.ackWaiters.get(key);
      if (item) clearTimeout(item.timer);
      state.ackWaiters.delete(key);
      throw error;
    }
    await accepted;
  }

  async function openPopup({ restart = false } = {}) {
    await ready;
    if (!Office.context || !Office.context.ui) throw new Error('Office UI APIを利用できません。');
    if (!Office.context.requirements.isSetSupported('DialogApi', '1.2')) {
      throw new Error('このExcelは DialogApi 1.2 をサポートしていません。');
    }
    if (state.opening) return;
    if (state.dialog) {
      if (!restart) {
        send({ type: 'notice', message: '検索ウィンドウは既に開いています。' });
        setStatus('検索ウィンドウは開いています。');
        return;
      }
      state.dialog.close();
      state.dialog = null;
      cancelAcks('ウィンドウの再起動');
      // Office Dialog.close() is asynchronous with no completion callback.
      // Avoid opening a second dialog in the same synchronous turn.
      await new Promise((resolve) => setTimeout(resolve, 350));
    }
    state.opening = true;
    setStatus('ポップアップを起動しています…');
    try {
      await new Promise((resolve, reject) => {
        Office.context.ui.displayDialogAsync(
          dialogUrl,
          { height: 72, width: 78, displayInIframe: false },
          (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              reject(new Error(result.error && result.error.message || 'Officeダイアログを開けません。'));
              return;
            }
            state.dialog = result.value;
            state.dialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
            state.dialog.addEventHandler(Office.EventType.DialogEventReceived, (event) => {
              log(`ウィンドウを閉じました (code=${event.error || 'normal'})`);
              state.dialog = null;
              state.loading = false;
              cancelAcks('ウィンドウが閉じられました。');
              setStatus('ポップアップは閉じています。');
            });
            resolve();
          }
        );
      });
      setStatus('ポップアップ起動。シート一覧の要求を待っています。');
    } catch (error) {
      setStatus(`ポップアップ起動失敗: ${stringifyError(error)}`, true);
      throw error;
    } finally {
      state.opening = false;
    }
  }

  function onDialogMessage(args) {
    let msg;
    try { msg = JSON.parse(args.message); }
    catch (error) { log('受信メッセージのJSON解析に失敗'); return; }
    if (!msg || typeof msg.type !== 'string') return;
    if (msg.type === 'ack') {
      const key = `${msg.requestId}:${msg.seq}`;
      const item = state.ackWaiters.get(key);
      if (item) {
        clearTimeout(item.timer);
        state.ackWaiters.delete(key);
        item.resolve();
      }
      return;
    }
    if (msg.type === 'hello' || msg.type === 'listSheets') {
      void listSheets(msg.requestId || 'initial');
    } else if (msg.type === 'loadSheets') {
      void loadSheets(msg);
    } else if (msg.type === 'selectCell') {
      void selectCell(msg);
    } else if (msg.type === 'close') {
      if (state.dialog) state.dialog.close();
    } else {
      log(`未定義の要求: ${msg.type}`);
    }
  }

  async function listSheets(requestId) {
    try {
      const result = await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        const activeSheet = sheets.getActiveWorksheet();
        sheets.load('items/id,items/name,items/visibility');
        activeSheet.load('id');
        await context.sync();
        return {
          sheets: sheets.items.filter((s) => s.visibility === 'Visible').map((s) => ({ id: s.id, name: s.name })),
          activeId: activeSheet.id,
        };
      });
      state.sheets = result.sheets;
      send({ type: 'sheets', requestId, ...result, limits: LIMITS });
      log(`シート一覧: ${result.sheets.length}件`);
    } catch (error) {
      reportError(requestId, `シート一覧取得失敗: ${stringifyError(error)}`);
    }
  }

  function splitIntoMessages(base, rows) {
    const frames = [];
    let segment = [];
    for (const row of rows) {
      const maybe = [...segment, row];
      const next = { ...base, rows: maybe };
      if (segment.length > 0 && JSON.stringify(next).length > LIMITS.messageChars) {
        frames.push(segment);
        segment = [row];
      } else {
        segment = maybe;
      }
    }
    if (segment.length) frames.push(segment);
    return frames;
  }

  async function readMetadata(sheetId) {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetId);
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load('rowIndex,columnIndex,rowCount,columnCount');
      await context.sync();
      if (used.isNullObject) return { rowIndex: 0, columnIndex: 0, rowCount: 0, columnCount: 0 };
      return {
        rowIndex: used.rowIndex,
        columnIndex: used.columnIndex,
        rowCount: used.rowCount,
        columnCount: used.columnCount,
      };
    });
  }

  async function readChunk(sheetId, startRow, startColumn, count, columns) {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetId);
      const range = sheet.getRangeByIndexes(startRow, startColumn, count, columns);
      range.load('text');
      await context.sync();
      return range.text.map((cells, offset) => ({
        row: startRow + offset,
        cells: cells.map((v) => String(v == null ? '' : v).slice(0, LIMITS.cellChars)),
      }));
    });
  }

  async function loadSheets(msg) {
    const requestId = String(msg.requestId || 'load');
    if (state.loading) {
      reportError(requestId, '別の読み込みが進行中です。完了してから再実行してください。');
      return;
    }
    const ids = Array.isArray(msg.ids) ? msg.ids : [];
    const picked = [...new Set(ids.filter((id) => state.sheets.some((sheet) => sheet.id === id)))];
    if (picked.length === 0 || picked.length > LIMITS.sheets) {
      reportError(requestId, `読み込むシートは1〜${LIMITS.sheets}枚を選択してください。`);
      return;
    }
    state.loading = true;
    let seq = 0;
    try {
      send({ type: 'loadStarted', requestId, count: picked.length });
      for (const id of picked) {
        const sheet = state.sheets.find((s) => s.id === id);
        const meta = await readMetadata(id);
        const rowCount = Math.min(meta.rowCount, LIMITS.rowsPerSheet);
        const colCount = Math.min(meta.columnCount, LIMITS.columnsPerSheet);
        send({
          type: 'sheetStart', requestId,
          sheet: { id, name: sheet.name, startRow: meta.rowIndex, startColumn: meta.columnIndex,
            rowCount, columnCount: colCount, originalRows: meta.rowCount, originalColumns: meta.columnCount },
        });
        for (let offset = 0; offset < rowCount; offset += LIMITS.chunkRows) {
          const count = Math.min(LIMITS.chunkRows, rowCount - offset);
          const rows = await readChunk(id, meta.rowIndex + offset, meta.columnIndex, count, colCount);
          for (const frameRows of splitIntoMessages({type: 'chunk', requestId, sheetId: id}, rows)) {
            seq += 1;
            await sendAndWaitAck({ type: 'chunk', requestId, sheetId: id, seq, rows: frameRows });
          }
        }
        send({ type: 'sheetDone', requestId, sheetId: id });
        log(`読込完了: ${sheet.name} (${rowCount}行 x ${colCount}列)`);
      }
      send({ type: 'loadDone', requestId });
      setStatus(`読み込み完了: ${picked.length}シート`);
    } catch (error) {
      reportError(requestId, `読み込みに失敗: ${stringifyError(error)}`);
    } finally {
      state.loading = false;
    }
  }

  async function selectCell(msg) {
    const requestId = String(msg.requestId || 'select');
    const sheetId = msg.sheetId;
    const row = msg.row;
    const col = msg.column;
    if (!state.sheets.some((sheet) => sheet.id === sheetId) ||
        !Number.isSafeInteger(row) || row < 0 || row > 1048575 ||
        !Number.isSafeInteger(col) || col < 0 || col > 16383) {
      reportError(requestId, '選択セルの指定が不正です。');
      return;
    }
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetId);
        sheet.activate();
        sheet.getCell(row, col).select();
        await context.sync();
      });
      send({ type: 'selected', requestId, sheetId, row, column: col });
    } catch (error) {
      reportError(requestId, `セルの選択に失敗: ${stringifyError(error)}`);
    }
  }
  function reportError(requestId, message) {
    setStatus(message, true);
    try { send({ type: 'error', requestId, message }); }
    catch (_ignored) { /* Dialog may have closed. */ }
  }

  async function ribbonCommand(event) {
    try {
      await openPopup();
    } catch (error) {
      console.error('openSearchPopup failed:', error);
    } finally {
      // Required to release the ribbon command queue; shared runtime stays alive.
      if (event && typeof event.completed === 'function') event.completed();
    }
  }

  // Register synchronously, to be available before a task pane is ever shown.
  Office.actions.associate('openSearchPopup', ribbonCommand);
  void ready.then((info) => {
    if (info.host !== Office.HostType.Excel) {
      setStatus('Excel以外のOfficeホストは対象外です。', true);
      return;
    }
    setStatus('Excelと接続しました。［検索ポップアップ］で起動できます。');
    const open = document.getElementById('openPopup');
    const reset = document.getElementById('resetPopup');
    const hide = document.getElementById('hidePane');
    if (open) open.addEventListener('click', () => void openPopup().catch(console.error));
    if (reset) reset.addEventListener('click', () => void openPopup({ restart: true }).catch(console.error));
    if (hide) hide.addEventListener('click', () => {
      if (Office.addin && typeof Office.addin.hide === 'function') {
        Office.addin.hide().catch((error) => setStatus(stringifyError(error), true));
      }
    });
  }).catch((error) => setStatus(`Office初期化失敗: ${stringifyError(error)}`, true));
})();
