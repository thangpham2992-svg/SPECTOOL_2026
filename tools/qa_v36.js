'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const root=__dirname+'/..';const ctx={window:{},console};vm.createContext(ctx);
for(const f of ['engine-database.js','spec-conditional.js','source-engine.js','legacy-step1.js','conditional-check.js','source-check.js'])vm.runInContext(fs.readFileSync(root+'/'+f,'utf8'),ctx,{filename:f});
const db=ctx.window.SPEC_ENGINE_DATABASE.MRL, E=ctx.window.SpecExcelSourceEngine.SourceEngine;
const row=db.sheets.DataTable,cell=x=>row[x]&&typeof row[x]==='object'?row[x].v:row[x];
const configs=new Map();for(let r=10;r<=26;r++){const kg=+cell('A'+r),code=cell('B'+r);if(!configs.has(kg))configs.set(kg,code);}
let tests=0,supported=0,unsupported=0;
for(const speed of [60,96,105,120,150])for(const kg of [450,630,825,1000,1275,1350,1600,1800])for(const [low,high] of [[0,30],[30,60],[60,80],[80,105]]){
 const r=5+tests, expected={speed:cell('DA'+r),kg:cell('DB'+r),low:cell('DC'+r),high:cell('DD'+r)};
 assert.deepStrictEqual(expected,{speed,kg,low,high},'Incorrect matrix row '+r);
 const name=configs.get(kg);assert(name,'Missing capacity '+kg);
 const changes={D13:name,D17:speed,D35:Math.round((low+high)/2*1000),D26:'Mitsubishi Standard'};
 const engine=new E(db,changes);const pd=engine.safe('SpecForm','F36'),oh=engine.safe('SpecForm','F37');
 assert(pd.ok&&oh.ok,JSON.stringify({speed,kg,low,high,pd,oh}));
 if(cell('DK'+r)==='OK'){
  assert.strictEqual(pd.value,cell('DG'+r),`PD ${kg}@${speed} TR${low}-${high}`);
  assert.strictEqual(oh.value,cell('DE'+r),`OH ${kg}@${speed} TR${low}-${high}`);
  supported++;
  const en=new E(db,{...changes,D26:'EN81:20/50'}).safe('SpecForm','F37');
  assert(en.ok,JSON.stringify(en));assert.strictEqual(en.value,cell('DF'+r),`EN OH ${kg}@${speed} TR${low}-${high}`);
 }else{assert.strictEqual(pd.value,'CHECK LAYOUT');assert.strictEqual(oh.value,'CHECK LAYOUT');unsupported++;}
 tests++;
}
// exact inclusivity boundary tests: lower range wins at 30/60/80 m
for(const [travel,mr] of [[30,0],[60,1],[80,2],[105,3]]){
 const r=37+mr; // speed 96, capacity 450 (rows 37 to 40)
 const expected=cell('DE'+r);const out=new E(db,{D13:configs.get(450),D17:96,D35:travel*1000}).safe('SpecForm','F37');
 assert(out.ok);assert.strictEqual(out.value,expected,'TR boundary '+travel);
}
for(const model of ['MR','MRL']){
 const sf=model==='MR'?'D17':'D19';const baseline=ctx.window.SPEC_SOURCE_CHECK(model,{},ctx.window.SPEC_ENGINE_DATABASE[model]);
 for(const value of ['', 'B3; Tầng mái; 1-20', 'Miêu tả tùy ý, không bằng Stops']){
  const res=ctx.window.SPEC_SOURCE_CHECK(model,{[sf]:value},ctx.window.SPEC_ENGINE_DATABASE[model]);
  assert.strictEqual(res.status,baseline.status,model+' Service floor must be unrestricted '+value);
  assert(!res.errors.some(e=>e.cell===sf),model+' Service floor must not be a validation target');
 }
}
console.log('PASS v3.6:',tests,'OH/PD mapped scenarios',supported,'supported',unsupported,'CHECK LAYOUT; 4 TR boundaries and 6 Service floor tests.');
