'use strict';
const fs=require('fs'),vm=require('vm'),path=require('path'),root=path.resolve(__dirname,'..');
const ctx={window:{},console};vm.createContext(ctx);
for(const f of ['engine-database.js','spec-conditional.js','source-engine.js','legacy-step1.js','conditional-check.js','source-check.js'])vm.runInContext(fs.readFileSync(path.join(root,f),'utf8'),ctx,{filename:f});
const output=[];let tests=0,unknown=0;
for(const model of ['MR','MRL']){
 const db=ctx.window.SPEC_ENGINE_DATABASE[model],t=db.tables.TblK5_Capacity;
 const match=/^[A-Z]+(\d+):[A-Z]+(\d+)$/.exec(t.ref),start=+match[1]+1,end=+match[2];
 const rowVal=(r,col)=>{const x=db.sheets.DataTable[col+r];return x&&typeof x==='object'?x.v:x};
 for(let r=start;r<=end;r++){
  const cap=rowVal(r,'B');if(!cap)continue;
  const changes={D13:cap};
  if(model==='MR')Object.assign(changes,{D14:rowVal(r,'D'),D29:rowVal(r,'E'),D30:rowVal(r,'F'),D32:rowVal(r,'G'),D34:rowVal(r,'I'),D35:rowVal(r,'J')});
  else Object.assign(changes,{D15:rowVal(r,'D'),D28:rowVal(r,'E'),D29:rowVal(r,'F'),D31:rowVal(r,'G')});
  const result=ctx.window.SPEC_SOURCE_CHECK(model,changes,db);
  tests++;if(result.unverified.length)unknown++;
  const info={model,capacity:cap,result:result.status,errors:result.errors.map(x=>({cell:x.cell,id:x.id,message:x.message})),unknown:result.unverified.slice(0,2)};
  output.push(info);
  if(result.errors.length)console.log(model,cap,'FAIL',result.errors.slice(0,3).map(e=>e.cell+':'+e.message.slice(0,50)).join(' | '));
 }
}
// Targeted boundary conditions with actual source engine formulas
const mr=ctx.window.SPEC_ENGINE_DATABASE.MR;
const P10={D13:'P10-CO (700 KG)-CWT Rear',D14:'CO',D29:1400,D30:1250,D32:800,D34:1750,D35:1810};
for(const [w,expected] of [[10,'PASS'],[25,'FAIL']]){
 let res=ctx.window.SPEC_SOURCE_CHECK('MR',{...P10,D61:w},mr);tests++;
 if(res.status!==expected)throw Error('P10 weight='+w+' expected '+expected+' got '+res.status);
 if(w===25&&!res.errors.some(x=>x.cell==='D61'))throw Error('P10 overweight error did not point D61');
}
const mrl=ctx.window.SPEC_ENGINE_DATABASE.MRL;
for(const [cap,jj,expected] of [
 ['P08-CO (630 KG)',800,'PASS'],['P08-CO (630 KG)',900,'PASS'],
 ['P08-CO (630 KG)',700,'FAIL'],['P08-2S (630 KG)',800,'FAIL']]){
 const res=ctx.window.SPEC_SOURCE_CHECK('MRL',{D13:cap,D31:jj},mrl);tests++;
 // This test asserts the dedicated JJ condition, other configuration conditions may reject before separate inputs change.
 const problem=res.errors.some(x=>x.id==='SpecForm!D31');
 if(problem!==(expected==='FAIL'))throw Error(cap+': JJ='+jj+' invalid flag expected '+expected+' got '+problem);
}
const filename=path.join(root,'QA_capacity_matrix_v3.3.json');fs.writeFileSync(filename,JSON.stringify(output,null,2));
console.log('Checked capacity rows',output.length,'targeted checks',6,'source statuses',output.reduce((o,x)=>(o[x.result]=(o[x.result]||0)+1,o),{}),'unverified',unknown,'total',tests);
