const fs=require('fs'),vm=require('vm');let cx={window:{},console};vm.createContext(cx);for(const f of ['engine-database.js','source-engine.js'])vm.runInContext(fs.readFileSync(__dirname+'/../'+f,'utf8'),cx);const Eng=cx.window.SpecExcelSourceEngine.SourceEngine,D=cx.window.SPEC_ENGINE_DATABASE;
const cases=[
 ['MR','Base',{},['D489','D512','AP10']],
 ['MR','One and both SDE',{A116:1,A117:1},['D489','D512','AP10']],
 ['MR','CO DSJ1 + One SDE',{D50:'SUS-HL with DSJ1',D14:'CO',A116:1},['AP10','D512']],
 ['MRL','Base',{},['D498','D512','D527','D523']],
 ['MRL','One and both SDE',{A118:1,A119:1},['D498','D512','D527','D523']],
 ['MRL','OH 2000',{D37:2000},['D498','D512','D527','D523']],
 ['MRL','PIT 1000',{D36:1000},['D498','D512','D527','D523']],
 ['MRL','DSJ1 SUS-HL wrong SUS',{D46:'SUS-HL with DSJ1',D40:'SUS Standard'},['D498','D512','D527','D523']],
 ['MRL','NO SETS-L high travel',{D17:150,D35:70000,D13:'P17D-CO (1275 KG)',A157:0},['D520','D512','D527','D523']]
];
for(const [model,name,changes,rs] of cases){let en=new Eng(D[model],changes);console.log(model,name,rs.map(r=>r+':'+JSON.stringify(en.safe('DataTable',r))).join(' | '),'B209',model==='MRL'?JSON.stringify(en.safe('SpecForm','B209')):'');}
