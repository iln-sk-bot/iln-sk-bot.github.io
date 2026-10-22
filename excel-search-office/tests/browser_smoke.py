#!/usr/bin/env python3
"""Chromium smoke test; fakes Office.js and Excel.run without modifying the add-in."""
from pathlib import Path
from playwright.sync_api import sync_playwright

DOCS=Path(__file__).resolve().parent.parent/'docs'

MOCK="""
window.Office={HostType:{Excel:'Excel'},context:{host:'Excel',document:{url:'file:///C:/Demo/Inventory.xlsx',name:'Inventory.xlsx'}},onReady:async()=>({host:'Excel'})};
window.testSelect=[];
window.testSync=0;
const data=[['品名','数量','状態'],['widget',5,'OK'],['gadget',8,'NG'],['other',0,'OK']];
const words=[['品名','数量','状態'],['widget','5個','OK'],['gadget','8個','NG'],['other','0個','OK']];
const sheetTable={name:'InventoryTable',showHeaders:true,showTotals:false,getRange(){return {load(){},rowIndex:0,columnIndex:0,rowCount:4,columnCount:3}}};
const sheet={
  name:'Inventory',
  getUsedRange:()=>({rowIndex:0,columnIndex:0,rowCount:4,columnCount:3,load(){}}),
  getRangeByIndexes:(r,c,h,w)=>({load(){},values:data.slice(r,r+h).map(x=>x.slice(c,c+w)),text:words.slice(r,r+h).map(x=>x.slice(c,c+w))}),
  tables:{items:[sheetTable],load(){}},
  activate(){window.testSelect.push('activate');},
  getRange:(a)=>({select(){window.testSelect.push(a);}})
};
window.Excel={run:async function(fn){return fn({workbook:{worksheets:{items:[sheet],load(){},getItem(n){if(n!=='Inventory')throw Error(n);return sheet;}}},sync:async()=>{window.testSync++;}});}};
"""

import re
import base64
html=(DOCS/'taskpane.html').read_text('utf8')
module=re.search(r'<script type="module" data-office-entry>([\s\S]*?)</script>',html).group(1)
adapter=(DOCS/'office-adapter.mjs').read_text('utf8').replace('export const ', 'const ').replace('export function ', 'function ').replace('export async function ','async function ')
module=module.replace('import {readWorkbookSnapshot, selectWorkbookCell} from "./office-adapter.mjs";',adapter)
for name in ['jexl-6.0.0.mjs','date-and-time-4.6.0.mjs']:
    payload=base64.b64encode((DOCS/'vendor'/name).read_bytes()).decode()
    module=module.replace('./vendor/'+name,'data:text/javascript;base64,'+payload)
html=re.sub(r'<script type="module" data-office-entry>[\s\S]*?</script>','',html)
html=re.sub(r'<script src="[^"]+"></script>','',html)
MOCK_IDB="""
(() => {
 const stores=new Map();
 const objectStoreNames={contains:name=>stores.has(name)};
 const database={objectStoreNames,createObjectStore(name,options={}) {stores.set(name,{data:new Map(),keyPath:options.keyPath});},
   transaction(name) {return {objectStore() {const store=stores.get(name);return {
     get(key){return request(()=>store.data.get(key));},
     put(value,key){return request(()=>{store.data.set(key ?? (store.keyPath ? value[store.keyPath] : null),value);return key;});},
     delete(key){return request(()=>{store.data.delete(key);});}
   };}};}};
 function request(fn){const req={};queueMicrotask(()=>{try{req.result=fn();req.onsuccess?.();}catch(e){req.error=e;req.onerror?.();}});return req;}
 Object.defineProperty(window,'indexedDB',{configurable:true,value:{open(){const req={result:database};queueMicrotask(()=>{req.onupgradeneeded?.();req.onsuccess?.();});return req;}}});
})();
"""

errors=[]
with sync_playwright() as p:
    browser=p.chromium.launch(headless=True,executable_path='/usr/bin/chromium',args=['--no-sandbox'])
    page=browser.new_page(viewport={'width':1250,'height':850},device_scale_factor=1)
    page.on('pageerror', lambda error: errors.append(str(error)))
    page.set_content(html)
    page.add_script_tag(content=MOCK)
    page.add_script_tag(content=MOCK_IDB)
    page.add_script_tag(content=(DOCS/'xlsx.full.min.js').read_text('utf8'))
    page.add_script_tag(type='module', content=module)
    page.wait_for_function("document.querySelector('#summaryText')?.textContent?.includes('件')",timeout=25000)
    assert page.locator('#results .sheet-card').count() > 0, 'No rendered Read Unit'
    assert 'widget' in page.locator('#results').inner_text()
    assert page.locator('#workbookList').count() == 1
    page.locator('#searchInput').fill('widget')
    page.wait_for_function("document.querySelector('#summaryText')?.textContent?.includes('1件')",timeout=10000)
    rows=page.locator('.sheet-card tbody tr')
    assert rows.count() == 1, rows.count()
    assert 'widget' in rows.first.inner_text()
    rows.first.locator('td.grid-cell').first.click()
    page.wait_for_function("window.testSelect.includes('A2')")
    page.locator('#searchMode').select_option('expression')
    page.locator('#searchInput').fill('names["数量"] >= 8')
    page.wait_for_function("document.querySelector('#summaryText')?.textContent?.includes('1件')",timeout=12000)
    assert 'gadget' in page.locator('#results').inner_text()
    assert 'widget' not in page.locator('#results').inner_text()
    # Read Unit settings include a table and allow adding a manual range.
    page.screenshot(path=str(Path(__file__).resolve().parent/'taskpane-wide-preview.png'),full_page=False)
    page.locator('#targetsButton').click()
    panel=page.locator('#targetPanel')
    assert 'InventoryTable' in panel.inner_text()
    panel.locator('button[title="手動の読み込み範囲を追加"]').first.click()
    assert page.locator('#sheetDialog').evaluate('(e)=>e.open')
    page.locator('#cancelSheetDialog').click()
    page.locator('#closeTargetsButton').click()
    before=page.evaluate('window.testSync')
    page.locator('#refreshButton').click()
    page.wait_for_function('(previous)=>window.testSync > previous + 2',arg=before,timeout=20000)
    page.set_viewport_size({'width':355,'height':800})
    page.screenshot(path=str(Path(__file__).resolve().parent/'taskpane-preview.png'),full_page=False)
    body_width=page.evaluate('document.body.scrollWidth')
    assert body_width <= 375, f'narrow viewport overflow: {body_width}'
    assert not errors, errors
    print('Browser smoke passed: workbook read, search, Jexl, Excel select, narrow layout')
    browser.close()
