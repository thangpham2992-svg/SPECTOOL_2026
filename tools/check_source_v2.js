const fs=require('fs'),vm=require('vm');let cx={window:{},console};vm.createContext(cx);const prefix=__dirname+'/../';for(const f of ['engine-database.js','spec-conditional.js','source-engine.js','legacy-step1.js','conditional-check.js','source-check.js'])vm.runInContext(fs.readFileSync(prefix+f,'utf8'),cx,{filename:f});
const D=cx.window.SPEC_ENGINE_DATABASE,check=cx.window.SPEC_SOURCE_CHECK;
for(const [m,label,changes] of [['MR','Default',{}],['MR','SDE both',{A116:1,A117:1}],['MR','OH below',{D41:4000}],['MR','Paint value when SUS-HL',{D51:'RAL 9003'}],['MR','Blank required',{D15:''}],['MRL','Default',{}],['MRL','SDE both',{A118:1,A119:1}],['MRL','OH below',{D37:2000}],['MRL','SUS with paint',{D47:'RAL 9003'}],['MRL','Blank required',{D17:''}]]){
 let t=Date.now(),r=check(m,changes,D[m]);console.log(m,label,r.status,JSON.stringify(r.conditions),'errors',r.errors.length,'unknown',r.unverified.length,'ms',Date.now()-t);
 console.log(' ',r.errors.slice(0,6).map(x=>x.id+': '+x.message.slice(0,65)));
 if(r.unverified.length)console.log(' UNKNOWN:',r.unverified.slice(0,5))
}
