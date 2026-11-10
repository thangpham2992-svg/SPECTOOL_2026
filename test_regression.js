const fs=require('fs'),vm=require('vm');const path=__dirname;
const ctx={window:{},console,structuredClone,Set,Map};vm.createContext(ctx);
for(const f of ['engine-database.js','ceiling-finishes.js','cwt-dimensions.js','source-engine.js','legacy-step1.js','conditional-check.js','source-check.js'])vm.runInContext(fs.readFileSync(path+'/'+f,'utf8'),ctx,{filename:f});
const eng=(model,changes={})=>new ctx.window.SpecExcelSourceEngine.SourceEngine(ctx.window.SPEC_ENGINE_DATABASE[model],changes);
for(const m of ['MR','MRL']){
 let r=ctx.window.SPEC_SOURCE_CHECK(m,{},ctx.window.SPEC_ENGINE_DATABASE[m]);console.log('MODEL',m,'default STATUS',r.status,'ERRORS',r.errors.slice(0,9),'UNKNOWN',r.unverified.slice(0,9),'CONDS',r.conditions);
}
for(const [name,changes] of [
 ['P08 standard 620D',{D13:'P08-CO (630 KG)',D17:'60',D31:'800',D16:'1D-1G',D84:'Not supplied',D23:'Bên hông phòng thang (Side drop) - BWH-620D'}],
 ['P08 alt 620E',{D13:'P08-CO (630 KG)',D17:'60',D31:'800',D16:'1D-1G',D84:'Not supplied',D23:'Bên hông phòng thang (Side drop) - BWH-620E'}],
 ['P11 high alt',{D13:'P11-2S (825 KG)',D17:'120',D31:'900',D16:'1D-2G',D84:'Not supplied',D23:'Bên hông phòng thang (Side drop) - BWH-620E'}],
]){
 const e=eng('MRL',changes);const l=ctx.window.SPEC_CWT_ENGINE.lookupEngine(e);
 console.log('TEST',name,l,'F33',e.safe('SpecForm','F33'),'F34',e.safe('SpecForm','F34'));
}
for(const [name,changes] of [['MR SUS',{D45:'L210S',D46:'SUS-HL'}],['MR SUS wrong paint',{D45:'L210S',D46:'Y033'}],['MRL DL6 SUS',{D41:'DL6-C',D42:'SUS-HL',D30:2160}],['MRL DL6 paint invalid',{D41:'DL6-C',D42:'Y135D',D30:2160}]]){
 const model=name.startsWith('MRL')?'MRL':'MR',e=eng(model,changes),addr=model==='MR'?'D509':'D508';console.log(name,addr,e.safe('DataTable',addr));
}
