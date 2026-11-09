const fs=require('fs'),vm=require('vm');let cx={window:{}};vm.createContext(cx);for(const f of ['engine-database.js','source-engine.js'])vm.runInContext(fs.readFileSync(__dirname+'/../'+f,'utf8'),cx);const Engine=cx.window.SpecExcelSourceEngine.SourceEngine;
for(const m of ['MR','MRL']){
 const en=new Engine(cx.window.SPEC_ENGINE_DATABASE[m]);let mismatches=[],errors=[],tot=0;
 for(const [a,c] of Object.entries(en.d.sheets.DataTable)){if(!c?.f)continue;tot++;let v=en.safe('DataTable',a);if(!v.ok){errors.push([a,v.error.slice(0,85)]);continue;}
  if(c.v!==undefined && c.v!=='' && (typeof c.v==='number'||typeof c.v==='boolean'||typeof c.v==='string')){const actual=v.value,cache=c.v;if(typeof actual==='number'&&typeof cache==='number'?Math.abs(actual-cache)>0.00001:String(actual)!==String(cache))mismatches.push([a,String(actual).slice(0,80),String(cache).slice(0,80)]);}
 }
 console.log(m,'formulas',tot,'errors',errors.length,'mismatches',mismatches.length,'calls',en.callCount);
 console.log('Errors samples',errors.slice(0,15));console.log('Mismatch samples',mismatches.slice(0,12));
}
