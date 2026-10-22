"""UI/parent-protocol smoke test: does not require Excel or network access."""
import re
from pathlib import Path
from playwright.sync_api import sync_playwright

root = Path(__file__).parents[1]
docs = root / 'docs'
html = (docs / 'dialog.html').read_text(encoding='utf-8')
css = (docs / 'styles.css').read_text(encoding='utf-8')
html = re.sub(r'<link rel="stylesheet" href="styles.css"\s*/>', '<style>'+css+'</style>', html)
html = re.sub(r'<script src="[^"]+"(?: defer)?></script>', '', html)
MOCK_OFFICE = r'''
window.Office = {
  AsyncResultStatus: {Succeeded:'succeeded'},
  EventType: {DialogParentMessageReceived:'parent'},
  onReady(callback) { return Promise.resolve().then(callback); },
  context: { ui: {
    addHandlerAsync(_eventType, handler, cb) { window._parentMessage = handler; cb({status:'succeeded'}); },
    messageParent(payload) {
      const msg = JSON.parse(payload);
      window._messages.push(msg);
      const reply = m => window._parentMessage({message:JSON.stringify(m)});
      if (msg.type === 'hello' || msg.type === 'listSheets') {
        reply({type:'sheets',requestId:msg.requestId,activeId:'s1',sheets:[{id:'s1',name:'受注データ'},{id:'s2',name:'部品マスタ'}],limits:{sheets:5,rowsPerSheet:1500,columnsPerSheet:32}});
      } else if (msg.type === 'loadSheets') {
        reply({type:'loadStarted',requestId:msg.requestId,count:msg.ids.length});
        for(const id of msg.ids) {
          const rows = id === 's1' ? [['注文番号','顧客','品番'],['O-101','東京商事','AX-120'],['O-102','関西工業','BZ-900'],['O-103','東京商事','AX-120']] : [['品番','名称'],['AX-120','ベアリング']];
          reply({type:'sheetStart',requestId:msg.requestId,sheet:{id,name:id==='s1'?'受注データ':'部品マスタ',startRow:0,startColumn:0,originalRows:rows.length,originalColumns:rows[0].length,rowCount:rows.length,columnCount:rows[0].length}});
          reply({type:'chunk',requestId:msg.requestId,sheetId:id,seq:1,rows:rows.map((cells,row)=>({row,cells}))});
        }
        reply({type:'loadDone',requestId:msg.requestId});
      } else if (msg.type === 'selectCell') {
        reply({type:'selected',requestId:msg.requestId,sheetId:msg.sheetId,row:msg.row,column:msg.column});
      }
    }
  }}
};
window._messages=[];
'''
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True, executable_path='/usr/bin/chromium', args=['--no-sandbox'])
    page = browser.new_page(viewport={'width':1200,'height':760},device_scale_factor=1)
    errors=[]
    page.on('pageerror', lambda err: errors.append(str(err)))
    page.set_content(html)
    page.add_script_tag(content=MOCK_OFFICE)
    page.add_script_tag(content=(docs/'search-core.js').read_text(encoding='utf-8'))
    page.add_script_tag(content=(docs/'dialog.js').read_text(encoding='utf-8'))
    page.wait_for_function("document.querySelectorAll('#sheetChecks input').length === 2")
    page.locator('#loadSheets').click()
    page.wait_for_function("document.querySelector('#resultCount').textContent.includes('3 件')")
    page.locator('#searchInput').fill('AX-120')
    assert page.locator('#resultCount').inner_text() == '2 件'
    page.locator('#regexMode').check()
    page.locator('#searchInput').fill('^AX-')
    assert page.locator('#resultCount').inner_text() == '2 件'
    page.locator('#regexMode').uncheck()
    page.locator('#searchInput').fill('東京商事')
    assert page.locator('#resultCount').inner_text() == '2 件'
    page.locator('td.data-col .cell-button').first.click()
    assert page.evaluate("window._messages.some(m=>m.type==='selectCell')")
    page.locator('#searchInput').fill('AX-120')
    page.locator('#sheetChecks input[value="s2"]').check()
    page.locator('#loadSheets').click()
    page.wait_for_function("document.querySelector('#resultCount').textContent.includes('3 件')")
    page.locator('#searchInput').fill('[')
    page.locator('#regexMode').check()
    assert page.locator('#resultCount').inner_text() == '検索条件エラー'
    page.locator('#regexMode').uncheck()
    page.locator('#searchInput').fill('AX-120')
    page.screenshot(path=str(root/'tests'/'demo-screenshot.png'),full_page=True)
    print('Browser UI smoke: OK')
    print('Office message calls:',len(page.evaluate('window._messages')))
    print('Console errors:',errors)
    assert not errors, errors
    browser.close()
