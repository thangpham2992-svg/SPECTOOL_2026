const fs=require('fs'),vm=require('vm');let ctx={window:{},console};vm.createContext(ctx);
for(let f of ['engine-database.js','ceiling-finishes.js','cwt-dimensions.js','source-engine.js','legacy-step1.js','conditional-check.js','source-check.js'])vm.runInContext(fs.readFileSync(__dirname+'/'+f,'utf8'),ctx,{filename:f});
const DB=ctx.window.SPEC_ENGINE_DATABASE,Engine=ctx.window.SpecExcelSourceEngine.SourceEngine,check=ctx.window.SPEC_SOURCE_CHECK;
let pass=0,fail=0;function t(name,ok,detail=''){(ok?pass++:fail++);console.log(ok?'PASS':'FAIL',name,detail)}
const look=(changes)=>new Engine(DB.MRL,changes);const CWT=ctx.window.SPEC_CWT_ENGINE;
const base={D13:'P08-CO (630 KG)',D17:60,D31:800,D16:'1D-1G',D84:'Not supplied',D23:'Bên hông phòng thang (Side drop) - BWH-620D',D33:1800,D34:1720};
let result=CWT.lookupEngine(look(base));t('P08 620D base',result.ok&&result.AH===1800&&result.BH===1720,JSON.stringify(result).slice(0,130));
let alt={...base,D23:'Bên hông phòng thang (Side drop) - BWH-620E'};result=CWT.lookupEngine(look(alt));t('P08 620E min 1855x1720',result.ok&&result.AH===1855&&result.BH===1720);
let e=look(alt);t('F33/F34 dynamic E',e.safe('SpecForm','F33').value===1855&&e.safe('SpecForm','F34').value===1720);
let res=check('MRL',alt,DB.MRL);t('620E undersize reports AH',res.errors.some(x=>x.cell==='D33'&&/1855/.test(x.message)),res.errors.filter(x=>x.cell==='D33').map(x=>x.message).join(' | '));
res=check('MRL',{...alt,D33:1855},DB.MRL);t('620E no false undersize after AH correction',!res.errors.some(x=>x.id.startsWith('CWT!')),res.errors.filter(x=>x.id.startsWith('CWT!')));
const invalid={...base,D13:'P06-2S (450 KG)',D23:'Bên hông phòng thang (Side drop) - BWH-620E'};res=check('MRL',invalid,DB.MRL);t('450 620E unsupported => CWT D23',res.errors.some(x=>x.id==='CWT!TYPE'&&x.cell==='D23'));
let item=ctx.window.SPEC_CWT_DATA.find(x=>x.row===10);t('Corrected base DataTable AQ10',DB.MRL.sheets.DataTable.AQ10===1820&&item.baseAH===1820);
let changed=ctx.window.SPEC_CWT_DATA.filter(x=>x.altCwt!==x.baseCwt);t('154 alternate CWT rows',changed.length===154);
for(const [m,c,finish,expected] of [['MR','L210S','SUS-HL',true],['MR','L210S','Y033',false],['MRL','DL6-C','SUS-HL',true],['MRL','DL6-C','Y135D',false]]){
 let modifications=m==='MR'?{D45:c,D46:finish}:{D41:c,D42:finish};let sh=m==='MR'?'D509':'D508';let val=new Engine(DB[m],modifications).safe('DataTable',sh);t(`${m} ${c} / ${finish}`,val.ok&&val.value===expected,JSON.stringify(val));
}
t('MR L210S dropdown SUS-HL',JSON.stringify(ctx.window.SPEC_CEILING_FINISH.options('MR','L210S'))==='["SUS-HL"]');
t('MRL DL6 dropdown excludes painted',!ctx.window.SPEC_CEILING_FINISH.options('MRL','DL6-C').some(x=>x.startsWith('Y')));
res=check('MR',{D61:999},DB.MR);t('MR decor overweight attributed D61',res.errors.some(x=>x.cell==='D61'&&/Khối lượng|Decoration/.test(x.message)),res.errors.slice(0,7));
res=check('MR',{D45:'L210S',D46:'Y033'},DB.MR);t('MR ceiling wrong finish attributed D46',res.errors.some(x=>x.cell==='D46'),res.errors.slice(0,7));
res=check('MR',{D16:23,D65:'760'},DB.MR);t('MR COP/stops rule attributed D65',res.errors.some(x=>x.cell==='D65'),res.errors.slice(0,9));
console.log('TOTAL',pass,fail);if(fail)process.exit(2);
