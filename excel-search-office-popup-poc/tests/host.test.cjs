const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function createHostFixture() {
  const sent = [];
  const listeners = {};
  const cells = [
    ['品番','名称','状態'],
    ['AX-10','軸受','有効'],
    ['BC-20','モーター','廃止'],
  ];
  const selected = [];
  const associated = {};
  const dialog = {
    handlers:{},
    addEventHandler(type,handler) { this.handlers[type]=handler; },
    messageChild(text) {
      const msg=JSON.parse(text);
      sent.push(msg);
      if (msg.type === 'chunk') setImmediate(() => {
        dialog.handlers.dialogMessage({message:JSON.stringify({type:'ack', requestId:msg.requestId, seq:msg.seq})});
      });
    },
    close() {},
  };
  const sheet = {
    id:'sheet-one',name:'部品表',visibility:'Visible',
    load() {},
    getUsedRangeOrNullObject() {
      return {isNullObject:false,rowIndex:0,columnIndex:0,rowCount:cells.length,columnCount:3,load() {}};
    },
    getRangeByIndexes(row,column,count,cols) {
      return { load() {},get text(){return cells.slice(row,row+count).map(r=>r.slice(column,column+cols));}};
    },
    activate() {},
    getCell(row,column) { return {select(){selected.push({row,column})}}; },
  };
  const sheets = {
    items:[sheet],load() {},getActiveWorksheet(){return sheet;},getItem(id){if(id!==sheet.id) throw Error('invalid sheet');return sheet;},
  };
  const Excel = { run: fn=>fn({workbook:{worksheets:sheets},sync:async()=>{}}) };
  const Office = {
    HostType:{Excel:'Excel'},
    AsyncResultStatus:{Succeeded:'success'},
    EventType:{DialogMessageReceived:'dialogMessage',DialogEventReceived:'dialogEvent'},
    onReady(callback) { return Promise.resolve({host:'Excel'}).then(callback); },
    actions:{associate(name,func){associated[name]=func}},
    context:{
      requirements:{isSetSupported(){return true}},
      ui:{displayDialogAsync(_url,_opts,callback){callback({status:'success',value:dialog})}},
    },
  };
  const controls = new Map();
  const document={getElementById(id){
    if (!controls.has(id)) controls.set(id,{textContent:'',classList:{toggle(){}},addEventListener(name,fn){listeners[id+':'+name]=fn}});
    return controls.get(id);
  }};
  const logs=[];
  const context={document,window:{location:{href:'https://demo.github.io/addin/host.html'}},Office,Excel,URL,console:{log(...x){logs.push(x)},error(){}},setTimeout,clearTimeout,setImmediate};
  vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../docs/host.js'),'utf-8'),context);
  return {dialog, sent, selected, associated, controls, logs};
}
async function waitUntil(check,timeout=2000){
  const end=Date.now()+timeout;
  while(Date.now()<end){
    if(check())return;
    await new Promise(r=>setTimeout(r,5));
  }
  throw new Error('Timed out');
}
test('Ribbon open, sheet data chunk/ACK and Excel cell selection', async () => {
  const fixture = createHostFixture();
  assert.equal(typeof fixture.associated.openSearchPopup,'function');
  let completed=0;
  await fixture.associated.openSearchPopup({completed(){completed++}});
  assert.equal(completed,1);
  const emit = msg => fixture.dialog.handlers.dialogMessage({message:JSON.stringify(msg)});
  emit({type:'hello',requestId:'init'});
  await waitUntil(()=>fixture.sent.some(x=>x.type==='sheets'));
  const first=fixture.sent.find(x=>x.type==='sheets');
  assert.equal(first.sheets[0].name,'部品表');
  emit({type:'loadSheets',requestId:'load1',ids:['sheet-one']});
  await waitUntil(()=>fixture.sent.some(x=>x.type==='loadDone'));
  const frames = fixture.sent.filter(x=>x.type==='chunk');
  assert.equal(frames.length,1);
  assert.equal(frames[0].rows.length,3);
  assert.equal(frames[0].rows[1].row,1);
  assert.equal(frames[0].rows[1].cells[0],'AX-10');
  emit({type:'selectCell',requestId:'select1',sheetId:'sheet-one',row:1,column:2});
  await waitUntil(()=>fixture.sent.some(x=>x.type==='selected'));
  assert.deepEqual(fixture.selected,[{row:1,column:2}]);
});
