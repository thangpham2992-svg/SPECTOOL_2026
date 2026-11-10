const fs=require('fs'),vm=require('vm'),assert=require('assert');
const root=__dirname+'/..',cx={window:{},console};vm.createContext(cx);
for(const f of ['engine-database.js','spec-conditional.js','source-engine.js','legacy-step1.js','conditional-check.js','source-check.js'])vm.runInContext(fs.readFileSync(root+'/'+f,'utf8'),cx,{filename:f});
const db=cx.window.SPEC_ENGINE_DATABASE,Engine=cx.window.SpecExcelSourceEngine.SourceEngine,check=cx.window.SPEC_SOURCE_CHECK;
let passed=0;
for(const model of ['MR','MRL']){
 const floor=model==='MR'?'D17':'D19';const stops=model==='MR'?'D16':'D18';
 for(const value of ['B3,B2,1,3,10,20,RF','Tầng kỹ thuật + Tầng mái / bất kỳ','Just descriptions without numbers']){
  const change={[floor]:value};
  const result=check(model,change,db[model]);
  const baseline=check(model,{},db[model]);
  const floorIssues=result.errors.filter(x=>String(x.cell)===floor||/floor count|floor names|K10_SFloor/i.test(x.message));
  assert.strictEqual(floorIssues.length,0,model+' unexpected floor issues '+JSON.stringify(floorIssues));
  assert.strictEqual(result.status,baseline.status,'Service floor free text changed status for '+model+': '+value);
  passed++;
 }
}
const capEntries=Object.entries(db.MRL.sheets.DataTable).filter(([addr])=>/^B(?:1[0-9]|2[0-6])$/.test(addr)).map(([addr,name])=>({name,kg:Number(String(name).match(/\((\d+) KG\)/)?.[1])}));
const data=db.MRL.sheets.DataTable;
const get=addr=>{const v=data[addr];return typeof v==='object'?v.v:v};
for(const kg of [630,825,1000,1275,1600,1800])for(const speed of [96,105,120,150]){
 const cap=capEntries.find(x=>x.kg===kg);assert(cap,`missing capacity ${kg}`);
 const changed={D13:cap.name,D17:speed,D35:22000};
 const eng=new Engine(db.MRL,changed);
 const oh=eng.safe('SpecForm','F37');assert(oh.ok,`${kg}@${speed} formula error ${oh.error}`);
 const matching=[];
 for(let row=8;row<=43;row++)if(Number(get('BG'+row))===speed&&Number(get('BH'+row))<=kg&&kg<=Number(get('BI'+row))&&Number(get('BK'+row))<22&&22<=Number(get('BL'+row)))matching.push(Number(get('BN'+row)));
 if(matching.length===1){assert.strictEqual(oh.value,matching[0],`F37 mismatch ${kg}@${speed}`)}
 else{assert.strictEqual(matching.length,0,`multiple layouts ${kg}@${speed}`);assert.strictEqual(oh.value,'CHECK LAYOUT',`unsupported layout ${kg}@${speed}`)}
 passed++;
}
console.log('v3.5 QA passed',passed,'tests: 6 arbitrary service-floor cases + 24 OH matrix cases');
