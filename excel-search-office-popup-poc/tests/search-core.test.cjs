const assert = require('node:assert/strict');
const test = require('node:test');
const core = require('../docs/search-core.js');
const datasets = [{
  id:'s1', name:'Sheet1', startColumn:2, startRow:4,
  rows:[
    {row:4,cells:['品番','名称','状態']},
    {row:5,cells:['A-100','ベアリング','処理中']},
    {row:6,cells:['C-200','モーター','完了']},
    {row:7,cells:['A-300','スイッチ','完了']},
  ]
}];
test('Excel column index conversion', () => {
  assert.equal(core.columnName(0),'A');
  assert.equal(core.columnName(25),'Z');
  assert.equal(core.columnName(26),'AA');
  assert.equal(core.columnName(701),'ZZ');
  assert.equal(core.columnIndex('ABC'),730);
  assert.throws(() => core.columnIndex('XFE'), /範囲外/);
});
test('Incremental contains returns exact Excel coordinates', () => {
  const value = core.search(datasets,'ベア',{limit:100});
  assert.equal(value.total,1);
  assert.equal(value.results[0].row,5);
  assert.equal(value.results[0].startColumn,2);
  assert.deepEqual(value.results[0].matchColumns,[3]);
  assert.equal(core.cellAt(value.results[0],3),'ベアリング');
});
test('Regexp, case-sensitive, and absolute column filter', () => {
  assert.equal(core.search(datasets,'^A-',{regex:true}).total,2);
  assert.equal(core.search(datasets,'^A-',{regex:true,columns:'C'}).total,2);
  assert.equal(core.search(datasets,'^A-',{regex:true,columns:'D'}).total,0);
  assert.equal(core.search(datasets,'a-',{caseSensitive:false}).total,2);
  assert.equal(core.search(datasets,'a-',{caseSensitive:true}).total,0);
  assert.throws(() => core.search(datasets,'[',{regex:true}), /正規表現/);
});
test('Skip only actual first used row per sheet', () => {
  assert.equal(core.search(datasets,'品番',{skipHeader:true}).total,0);
  assert.equal(core.search(datasets,'品番',{skipHeader:false}).total,1);
  assert.equal(core.search(datasets,'',{skipHeader:true}).total,3);
});
test('Result limit affects displayed rows, not total count', () => {
  const result = core.search(datasets,'',{limit:1});
  assert.equal(result.total,3);
  assert.equal(result.results.length,1);
  assert.equal(result.truncated,true);
  assert.deepEqual(result.displayedColumns,[2,3,4]);
});
