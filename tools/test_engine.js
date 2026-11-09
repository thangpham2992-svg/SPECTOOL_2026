'use strict';const fs=require('fs'),vm=require('vm'),assert=require('assert');const root=__dirname+'/..';const sandbox={window:{},console};vm.createContext(sandbox);for(const f of ['engine-database.js','source-engine.js'])vm.runInContext(fs.readFileSync(root+'/'+f,'utf8'),sandbox,{filename:f});
const Engine=sandbox.window.SpecExcelSourceEngine.SourceEngine;
for(const model of ['MR','MRL']){
 const en=new Engine(sandbox.window.SPEC_ENGINE_DATABASE[model]);const tests=model==='MR'?['DataTable!D482','DataTable!D485','DataTable!D489','DataTable!D512','DataTable!D517','DataTable!AP10','SpecForm!B207','SpecForm!B208']:['DataTable!D490','DataTable!D493','DataTable!D498','DataTable!D512','DataTable!D520','DataTable!D523','DataTable!D527','SpecForm!F36','SpecForm!F37','SpecForm!B206','SpecForm!B209','SpecForm!B210'];
 console.log('\n'+model);for(const t of tests){const [sheet,a]=t.split('!');const val=en.safe(sheet,a),cached=en.d.sheets[sheet][a]?.v;console.log(t,JSON.stringify(val).slice(0,160),'cached=',cached)}console.log('calls',en.callCount,'AST',en.parsed.size)
}
