const fs=require('fs'),vm=require('vm');let cx={window:{},console};vm.createContext(cx);for(const f of ['engine-database.js','source-engine.js'])vm.runInContext(fs.readFileSync(__dirname+'/../'+f,'utf8'),cx);const Eng=cx.window.SpecExcelSourceEngine.SourceEngine;
for(const m of ['MR','MRL']){
 const en=new Eng(cx.window.SPEC_ENGINE_DATABASE[m]);console.log('\n'+m);let list=[];
 for(let i=m==='MR'?482:490;i<= (m==='MR'?518:527);i++)list.push(['DataTable','D'+i]);
 if(m==='MR')for(let i=1;i<=10;i++)list.push(['DataTable','AP'+i]);
 for(const cell of m==='MR'?['B207','B208']:['R13','R14','R23','R46','R73','R74','F33','F34','F36','F37','B206','B209','B210'])list.push(['SpecForm',cell]);
 for(const [sheet,addr] of list){const cell=en.d.sheets[sheet][addr],out=en.safe(sheet,addr);const cached=cell?.v;if(!out.ok){console.log('ERROR',sheet,addr,out.error,'cached=',cached);continue}let value=out.value;if(cached!==undefined && ((typeof value==='number' && typeof cached==='number')?Math.abs(value-cached)>1e-6:String(value)!==String(cached)))console.log('MISMATCH',sheet,addr,'computed=',JSON.stringify(value),'cached=',JSON.stringify(cached));}
 console.log('Calls',en.callCount,'AST',en.parsed.size)
}
