const fs=require('fs'),vm=require('vm'),ctx={window:{},console};vm.createContext(ctx);const base='/mnt/data/specweb_v47/';
for(const f of ['engine-database.js','embedded-data.js','ceiling-finishes.js','cwt-dimensions.js','source-engine.js','legacy-step1.js','conditional-check.js','source-check.js'])vm.runInContext(fs.readFileSync(base+f,'utf8'),ctx,{filename:f});
const engine=(model,changes={})=>new ctx.window.SpecExcelSourceEngine.SourceEngine(ctx.window.SPEC_ENGINE_DATABASE[model],changes);
for(const model of ['MR','MRL']){
 const fireCell=model==='MR'?'D83':'D84';const outCell=model==='MR'?'C81':'C85';const e=ctx.window.SPEC_EMBEDDED[model].fields.find(x=>x.name==='K36_FireproofDoor');console.log('OPTIONS',model,e?.options);
 for (const choice of ['Not supplied','GHCL E30','EN81-58(II)']){
  let en=engine(model,{[fireCell]:choice});let out=en.safe('SpecFormExport',outCell), min=model==='MRL'?{F33:en.safe('SpecForm','F33'),F34:en.safe('SpecForm','F34'),cwt:ctx.window.SPEC_CWT_ENGINE.lookupEngine(en)}:{Q10:en.safe('DataTable','Q10'),R10:en.safe('DataTable','R10'),S10:en.safe('DataTable','S10')};
  console.log(model,choice,'EXPORT',out,'DIM',min);
 }
}
for (const option of ['PIV1-A1010N','PIV1-C1010N']) {
 const e=engine('MR',{E81:option,E78:'PIV1-C1010N'});
 console.log('OTHER FLOOR',option,e.safe('SpecFormExport','C79'),'MAIN',e.safe('SpecFormExport','C76'));
}
