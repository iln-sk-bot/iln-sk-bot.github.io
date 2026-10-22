import assert from 'node:assert/strict';
import {columnLetter,a1Range,convertOfficeBlock,tablesFromGeometry,readWorkbookSnapshot,selectWorkbookCell} from '../docs/office-adapter.mjs';

assert.equal(columnLetter(0),'A');
assert.equal(columnLetter(701),'ZZ');
assert.equal(a1Range(3,2,2,3),'C4:E5');
assert.deepEqual(convertOfficeBlock([[1,false,''],['ID',3,'B']], [['1.0','FALSE',''],['ID','3','B']],2,1),[
  [3,1,'1.0',1,'n'],[3,2,'FALSE',false,'b'],[4,1,'ID','ID','s'],[4,2,'3',3,'n'],[4,3,'B','B','s']
]);
const defs=tablesFromGeometry([{name:'DataTable',showHeaders:true,showTotals:false,rowIndex:0,columnIndex:0,rowCount:3,columnCount:2}], [[1,0,'品名'],[1,1,'数量']]);
assert.equal(defs[0].ref,'A1:B3');
assert.deepEqual(defs[0].columns,['品名','数量']);

const callLog=[];
const grid=[['品名','数量'],['widget',5],['gadget',8]];
const texts=[['品名','数量'],['widget','5個'],['gadget','8個']];
const table={name:'DataTable',showHeaders:true,showTotals:false,getRange(){return Object.assign({load(){}},{rowIndex:0,columnIndex:0,rowCount:3,columnCount:2});}};
const sheet={
 name:'Inventory',
 getUsedRange(){return Object.assign({load(){}},{rowIndex:0,columnIndex:0,rowCount:3,columnCount:2});},
 getRangeByIndexes(r,c,h,w){return {load(){}, values:grid.slice(r,r+h).map(x=>x.slice(c,c+w)),text:texts.slice(r,r+h).map(x=>x.slice(c,c+w))};},
 tables:{load(){},items:[table]},
 activate(){callLog.push('activate');},
 getRange(a1){return {select(){callLog.push(a1);}}}
};
const excel={
 run:async fn=>fn({workbook:{worksheets:{load(){},items:[sheet],getItem(name){assert.equal(name,'Inventory');return sheet;}}}, sync:async()=>{callLog.push('sync');}})
};
const progress=[];
const snapshots=await readWorkbookSnapshot(excel,{batchCells:2,onProgress:x=>progress.push(x)});
assert.equal(progress.length,3);
assert.equal(snapshots.length,1);
assert.equal(snapshots[0].cells.length,6);
assert.equal(snapshots[0].minRow,1);
assert.equal(snapshots[0].maxCol,1);
assert.equal(snapshots[0].cells.find(x=>x[0]===2 && x[1]===1)[2],'5個');
assert.equal(snapshots[0].cells.find(x=>x[0]===2 && x[1]===1)[3],5);
assert.equal(snapshots[0].tables[0].name,'DataTable');
await selectWorkbookCell(excel,'Inventory','B2');
assert.deepEqual(callLog.slice(-3),['activate','B2','sync']);
await assert.rejects(readWorkbookSnapshot(excel,{maxCells:5}),/5.*セル/);
await assert.rejects(selectWorkbookCell(excel,'Inventory','A0'),/Invalid/);
console.log('adapter tests passed');
