'use strict';
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const root = __dirname + '/..';
const ctx={window:{},console};
vm.createContext(ctx);
for(const f of ['engine-database.js','spec-conditional.js','source-engine.js','legacy-step1.js','conditional-check.js','source-check.js']){
  vm.runInContext(fs.readFileSync(root+'/'+f,'utf8'),ctx,{filename:f});
}
const db=ctx.window.SPEC_ENGINE_DATABASE.MRL;
assert.strictEqual(db.sheets.DataTable.J366,'A','AECC AI-2200C must be A');
const Engine=ctx.window.SpecExcelSourceEngine.SourceEngine;
for(const [operation,qty] of [['1C-2BC',1],['2C-2BC',2],['3C-SAI-2200C',3],['4C-SAI-2200C',4]]){
  const changes={D22:operation,D11:qty,A165:1};
  const engine=new Engine(db,changes);
  assert.strictEqual(engine.safe('DataTable','K366').value,'A',operation+' DataTable K366');
  assert.strictEqual(engine.safe('SpecForm','E165').value,'A',operation+' SpecForm E165');
  const check=ctx.window.SPEC_SOURCE_CHECK('MRL',changes,db);
  assert.strictEqual(check.status,'PASS',operation+' Check Spec: '+JSON.stringify(check.errors));
  console.log('PASS:',operation,'AECC=A, Check Spec PASS');
}
// Select another forbidden feature to ensure unsupported-feature detection remains active.
const prohibited=ctx.window.SPEC_SOURCE_CHECK('MRL',{D22:'3C-SAI-2200C',D11:3,A186:1},db);
assert.strictEqual(prohibited.status,'FAIL');
console.log('PASS: Other prohibited features remain blocked');
