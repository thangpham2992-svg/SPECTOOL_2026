'use strict';
const $=id=>document.getElementById(id);
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NSREL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const X=(txt)=>new DOMParser().parseFromString(String(txt).replace(/^\uFEFF/, '').replace(/^\s+(?=<\?xml\b)/, ''),'application/xml');
const each=(node,tag)=>Array.from(node.getElementsByTagNameNS(NS,tag));
const first=(node,tag)=>node.getElementsByTagNameNS(NS,tag)[0];
const get=(node,tag)=>first(node,tag)?.textContent||'';
const modelFiles={MR:'SpecToolGXLM2P.xlsx',MRL:'SpecToolGQXL2P.xlsx'};
let models={MR:structuredClone(window.SPEC_EMBEDDED.MR),MRL:structuredClone(window.SPEC_EMBEDDED.MRL)};
let layouts=structuredClone(window.SPEC_LAYOUT);
for(const k of ['MR','MRL']) models[k].fields.push(...structuredClone(window.SPEC_EXTRAS[k]));
let sourceFiles={},current=null,changes={};
function status(message,cls=''){$('status').className=cls;$('status').textContent=window.SPEC_PROOFREAD?.(message)??message}
function bytesFromB64(b64){const raw=atob(b64),bytes=new Uint8Array(raw.length);for(let i=0;i<raw.length;i++)bytes[i]=raw.charCodeAt(i);return bytes}
const colnum=col=>Array.from(col).reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
function colname(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
function split(addr){let m=/\$?([A-Z]+)\$?(\d+)/.exec(addr||'');return m?[colnum(m[1]),Number(m[2])]:[0,0]}
function normPath(path){let parts=[];for(const part of path.split('/')){if(part==='..')parts.pop();else if(part && part!=='.')parts.push(part)}return parts.join('/')}
function getSheetPath(target){return target.startsWith('/xl/')?target.slice(1):normPath('xl/'+target)}
function cellValue(c,shared){const typ=c.getAttribute('t');if(typ==='inlineStr')return each(c,'t').map(x=>x.textContent).join('');let raw=get(c,'v');if(typ==='s')return shared[Number(raw)]||'';if(['str','e'].includes(typ))return raw;if(typ==='b')return raw==='1';if(raw===''||raw==null)return '';const n=Number(raw);return Number.isFinite(n)?n:raw}
async function parseXlsx(data){
 const zip=await JSZip.loadAsync(data),read=async p=>zip.file(p)?await zip.file(p).async('string'):null;
 const sharedXml=await read('xl/sharedStrings.xml');const shared=sharedXml?each(X(sharedXml),'si').map(si=>each(si,'t').map(n=>n.textContent).join('')):[];
 const workbook=X(await read('xl/workbook.xml'));
 const relxml=X(await read('xl/_rels/workbook.xml.rels'));
 const rels=new Map(Array.from(relxml.documentElement.children).map(el=>[el.getAttribute('Id'),el.getAttribute('Target')]));
 const paths={};for(const sheet of each(workbook,'sheet')){let t=rels.get(sheet.getAttributeNS(NSREL,'id'));if(t)paths[sheet.getAttribute('name')]=getSheetPath(t)}
 function shiftSharedFormula(formula,from,to){
 const re=/^\$?([A-Z]{1,3})\$?(\d+)$/i,a=re.exec(from),b=re.exec(to);
 if(!a||!b)throw Error('Công thức Excel không hợp lệ.');
 const dr=Number(b[2])-Number(a[2]),dc=colnum(b[1])-colnum(a[1]);
 const parts=String(formula).split(/("(?:[^"]|"")*")/g);
 return parts.map((chunk,i)=>i%2?chunk:chunk.replace(/(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_.])/g,(whole,ca,col,ra,row)=>{
  const c=colnum(col)+(ca?0:dc),r=Number(row)+(ra?0:dr);
  return c<1||r<1?'#REF!':ca+colname(c)+ra+r;
 })).join('');
}
const roots={},cells={},rawCells={};for(const [sheet,path] of Object.entries(paths)){let xml=await read(path);if(!xml)continue;roots[sheet]=X(xml);cells[sheet]={};rawCells[sheet]={};for(const c of each(roots[sheet],'c')){const addr=c.getAttribute('r');if(!addr)continue;const value=cellValue(c,shared);cells[sheet][addr]=value;const f=first(c,'f');rawCells[sheet][addr]=f?{f:f.textContent||'',v:value}:value;}
  // Excel stores shared formula members as <f t="shared" si="n"/> with no text.
  // Rebuild their formulas from the master before running source checks.
  const masters=new Map();const pending=[];
  for(const c of each(roots[sheet],'c')){
   const f=first(c,'f');if(!f||f.getAttribute('t')!=='shared')continue;
   const si=f.getAttribute('si'),addr=c.getAttribute('r');
   if(f.textContent)masters.set(si,{addr,formula:f.textContent});
   else pending.push({addr,si});
  }
  for(const member of pending){
   const m=masters.get(member.si);
   if(!m)throw Error('Excel chứa công thức dùng chung thiếu nguồn: '+sheet+'!'+member.addr);
   rawCells[sheet][member.addr].f=shiftSharedFormula(m.formula,m.addr,member.addr);
  }
 }
 if(!cells.SpecForm||!cells.DataTable||!cells.SpecFormExport)throw Error('Workbook thiếu một trong ba sheet bắt buộc: SpecForm / DataTable / SpecFormExport.');
 const names={},namesScoped={global:{},local:{}};const sheetOrder=each(workbook,'sheet').map(s=>s.getAttribute('name'));for(let dn of each(workbook,'definedName')){const name=dn.getAttribute('name'),value=dn.textContent||'',local=dn.getAttribute('localSheetId');(names[name]??=[]).push(value);if(local!==null&&local!==undefined){const sheetName=sheetOrder[Number(local)];(namesScoped.local[sheetName]??={})[name]=value;}else namesScoped.global[name]=value;}
 const tables={};for(const [path,file] of Object.entries(zip.files)){
  if(!/^xl\/tables\/[^/]+\.xml$/.test(path)||file.dir)continue;
  const root=X(await read(path)).documentElement,name=root.getAttribute('name'),ref=root.getAttribute('ref'),columns=each(root,'tableColumn').map(x=>x.getAttribute('name'));
  const tableFilename=path.split('/').at(-1);let sheetName=null;
  for(const [sheet,sp] of Object.entries(paths)){
   const relpath=sp.replace(/([^/]+)$/,'_rels/$1.rels'),relsText=await read(relpath);if(!relsText)continue;
   const d=X(relsText);if(Array.from(d.documentElement.children).some(r=>r.getAttribute('Target')?.split('/').at(-1)===tableFilename)){sheetName=sheet;break}
  }
  if(sheetName&&name&&ref)tables[name]={sheet:sheetName,ref,columns};
 }
 const tableRows=(name)=>{let t=tables[name];if(!t)return [];let [a,b]=t.ref.split(':'),[c0,r0]=split(a),[c1,r1]=split(b);let out=[];for(let r=r0+1;r<=r1;r++){const obj={};t.columns.forEach((col,j)=>obj[col]=cells[t.sheet][colname(c0+j)+r]??'');out.push(obj)}return out};
 const options=(formula,depth=0)=>{
  if(depth>3)return [];
  let f=(formula||'').replace(/^=/,'').trim();if(/^".*"$/.test(f))return f.slice(1,-1).split(',');
  // Range references inside INDIRECT: e.g. INDIRECT("DataTable!$B$930:$B$939")
  const range=/['"]?((?:DataTable|SpecForm))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/i.exec(f);
  if(range){const sh=range[1],a=colnum(range[2]),b=+range[3],c=colnum(range[4]),d=+range[5],values=[];
   if((c-a+1)*(d-b+1)<3000)for(let r=b;r<=d;r++)for(let col=a;col<=c;col++){let v=cells[sh]?.[colname(col)+r];if(v!==undefined && v!==null && String(v)!=='')values.push(String(v))}
   return [...new Set(values)];
  }
  let m=/(Tbl[A-Za-z0-9_]+)\[([^\]]+)\]/.exec(f);if(m)return [...new Set(tableRows(m[1]).map(row=>String(row[m[2]]??'')).filter(Boolean))];
  if(/^[a-z_][\w.]*$/i.test(f)){
   let r=names[f]||[];for(let v of r){if(v.includes('[')){let o=options(v,depth+1);if(o.length)return o}
    let z=/'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/.exec(v);
    if(z){let sh=z[1],c1=colnum(z[2]),r1=+z[3],c2=colnum(z[4]),r2=+z[5],found=[];for(let ri=r1;ri<=r2;ri++)for(let ci=c1;ci<=c2;ci++){let v=cells[sh]?.[colname(ci)+ri];if(v!==''&&v!==undefined)found.push(String(v))}if(found.length)return found}
   }
  }
  return []
 };
 const validations=[];let validator={};for(let dv of each(roots.SpecForm,'dataValidation')){
  if(dv.getAttribute('type')!=='list')continue;let formula=get(dv,'formula1');let sqref=dv.getAttribute('sqref')||'';
  validations.push({ref:sqref,formula});for(let token of sqref.split(/\s+/)){
   const m=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(token);
   if(m){let c1=colnum(m[1]),c2=colnum(m[3]),r1=+m[2],r2=+m[4];if((c2-c1+1)*(r2-r1+1)<2000)for(let rr=r1;rr<=r2;rr++)for(let cc=c1;cc<=c2;cc++)validator[colname(cc)+rr]=formula}
   else validator[token]=formula;
  }
 }
 // Read workbook-native conditional formatting: DevExpress LightUp = prohibited input.
 // When a new Excel Database is loaded, these rules are rebuilt from that file.
 const conditional={lightUp:[],defaultLightUp:[],namedRanges:namesScoped.local.SpecForm||{}};
 const sx=await read('xl/styles.xml');
 if(sx){
  const styleDoc=X(sx),dxfNodes=each(styleDoc,'dxf');
  const dxfFill=dxfNodes.map(n=>first(n,'patternFill')?.getAttribute('patternType')||'');
  const fillsParent=first(styleDoc,'fills'),cellXfs=first(styleDoc,'cellXfs');
  const fills=fillsParent?each(fillsParent,'fill').map(n=>first(n,'patternFill')?.getAttribute('patternType')||''):[];
  const styles=cellXfs?each(cellXfs,'xf').map(n=>fills[Number(n.getAttribute('fillId')||0)]||''):[];
  for(const c of each(roots.SpecForm,'c')){if(styles[Number(c.getAttribute('s')||0)]==='lightUp')conditional.defaultLightUp.push(c.getAttribute('r'));}
  for(const cf of each(roots.SpecForm,'conditionalFormatting')){
   const ranges=[];
   for(const r of String(cf.getAttribute('sqref')||'').split(/\s+/)){
    if(!r)continue;const [left,right]=r.split(':');const [c1,r1]=split(left),[c2,r2]=split(right||left);
    if(c1&&r1&&c2&&r2)ranges.push([c1,r1,c2,r2]);
   }
   for(const cr of each(cf,'cfRule')){
    if(dxfFill[Number(cr.getAttribute('dxfId')||'-1')]!=='lightUp')continue;
    const formula=get(cr,'formula');if(!formula)continue;
    conditional.lightUp.push({ranges,formula,priority:Number(cr.getAttribute('priority')||99999),stop:cr.getAttribute('stopIfTrue')==='1'});
   }
  }
  conditional.lightUp.sort((a,b)=>a.priority-b.priority);
 }
 const fields=[],used=new Set();for(const [name,refs] of Object.entries(names)){
  if(!/^K\d+_/.test(name)||name==='K50_CheckSpec')continue;
  let addr=null;for(let ref of refs){let re=/^'?SpecForm'?!\$?([A-Z]+)\$?(\d+)/.exec(ref);if(re){addr=re[1]+re[2];break}}
  if(!addr||used.has(addr)||split(addr)[0]>8)continue;used.add(addr);
  let formula=validator[addr]||'',opts=formula?options(formula):[];
  if(name==='K5_Capacity')opts=tableRows('TblK5_Capacity').map(r=>String(r.EleCapacityName||'')).filter(Boolean);
  fields.push({name,cell:addr,label:name.replace(/^K\d+_/,'').replaceAll('_',' '),value:cells.SpecForm[addr]??'',options:opts,validation:formula,dynamic:!!(formula&&!opts.length),editable:true});
 }
 fields.sort((a,b)=>split(a.cell)[1]-split(b.cell)[1]);
 const layout=[];const maxRow=Object.keys(cells.SpecForm).some(k=>/^D84$/.test(k))?84:83;for(let r=7;r<=maxRow;r++){let info={row:r};for(let c of ['B','C','D'])info[c]=String(cells.SpecForm[c+r]??'');if(Object.values(info).some(x=>typeof x==='string'&&x.trim()))layout.push(info)}
 for(let v of validations)for(let addr of v.ref.split(/\s+/)){if(!/^D\d+$/.test(addr)||fields.some(f=>f.cell===addr))continue;let r=split(addr)[1];if(r<11||r>maxRow)continue;fields.push({name:'Excel_'+addr,cell:addr,label:String(cells.SpecForm['C'+r]??cells.SpecForm['B'+r]??addr),value:cells.SpecForm[addr]??'',options:options(v.formula),validation:v.formula,dynamic:true,editable:true})}
 fields.sort((a,b)=>split(a.cell)[1]-split(b.cell)[1]);
return {layout,model:'Imported Excel',fields,validations,capacity:tableRows('TblK5_Capacity'),table_count:Object.keys(tables).length,name_count:Object.keys(names).length,engineData:{sheets:rawCells,tables,names:namesScoped,order:sheetOrder,conditional}};
}

let featureTouched = new Set();
let activeFeatureTable='';
const SPEC_FEATURE_VISUALS={
 'EXCL-C':{src:'data:image/png;base64,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',alt:'EXCL-C',caption:'Hình minh họa EXCL-C'},
 EXCLC:{src:'data:image/png;base64,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',alt:'EXCL-C',caption:'Hình minh họa EXCL-C'},
 EXCL:{src:'data:image/png;base64,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',alt:'EXCL',caption:'Hình minh họa EXCL'},
 'DKO-TB':{src:'data:image/png;base64,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',alt:'DKO-TB',caption:'Hình minh họa DKO-TB'}
};
function xrules(){return window.SPEC_RULES?.[$('model').value]||{labels:{},features:[],sideDrop:[]}}
// Add missing MR 2C-2BC ITP variants to the in-use source database, including when
// customers import an older original SpecTool Excel. This leaves all existing rows intact.
function workbookEngineData(){
 const model=$('model').value;
 const db=sourceFiles[model]?.engineData||window.SPEC_ENGINE_DATABASE?.[model];
 if(model==='MRL')window.SPEC_MATERIAL_5825?.patchDatabase(db);
 if(model==='MR')window.SPEC_MR_BED_5871?.patchDatabase(db);
 if(model==='MR'&&db?.sheets?.DataTable){
  const cells=db.sheets.DataTable;
  const added=['24V – 5 Lines – 2 Master Stations','24V – 5 Lines – 3 Master Stations'];
  for(const name of added){
   let exists=false;
   for(let r=2;r<=24;r++)if(String(cells['DN'+r]||'').trim()==='2C-2BC'&&String(cells['DO'+r]||'').trim()===name){exists=true;break;}
   if(exists)continue;
   for(let r=2;r<=24;r++)if(!String(cells['DN'+r]||'').trim()){
    cells['DN'+r]='2C-2BC';cells['DO'+r]=name;cells['DP'+r]='N';break;
   }
  }
 }
 window.SPEC_V52_FEATURES?.fixDatabase(model,db);
 window.SPEC_MIRROR_RULES?.fixDatabase(model,db);
 return db;
}
function sourceFeatureRows(){
 const db=workbookEngineData();if(!db)return xrules().features||[];
 const fromSource=[];
 for(const [name,t] of Object.entries(db.tables||{})){
  if(!/^TblFeature_/.test(name)||t.sheet!=='SpecForm')continue;
  const m=/[A-Z]+(\d+):[A-Z]+(\d+)/.exec(t.ref);if(!m)continue;
  for(let r=Number(m[1])+1;r<=Number(m[2]);r++){
   const sh=db.sheets.SpecForm;
   const src=addr=>{const v=sh[addr];return v&&typeof v==='object'?(v.v??''):(v??'')};
   const code=String(src('B'+r)||'').trim();if(!code)continue;
   // End-of-production: remove MBSS from MR/MRL selection and clear stale saved checkmarks.
   if(code.toUpperCase()==='MBSS'){changes['A'+r]='0';continue;}
   // Some legacy workbooks store the complete Hall Call card-reader description
   // in the feature CODE column. That source line is the same NSCR-H option,
   // not another independently selectable feature. Use the single NSCR-H
   // supplementary row below; keep the workbook/source calculation untouched.
   if(/^Non-service\s+Temporary\s+Release\s+for\s+Hall\s+Call\s*[-–]\s*Card\s+Reader\s+Type/i.test(code))continue;
   fromSource.push({table:name,cell:'A'+r,code,en:String(src('C'+r)||''),vi:String(src('D'+r)||''),option:String(src('E'+r)||''),suggestion:String(src('F'+r)||''),apply:src('A'+r)});
  }
 }
 // NEXIEZ-MR is an independent product: show only Feature codes with ALPHA=Y
 // in the MR DataTable; do not borrow Version02 features or show MR Q-blank options.
 if($('model').value==='MR' && db.tables?.Tbl39_Feature && db.sheets?.DataTable){
  const dt=db.sheets.DataTable,ref=db.tables.Tbl39_Feature.ref.split(':');
  const begin=Number(ref[0].match(/\d+/)[0])+1,end=Number(ref[1].match(/\d+/)[0]);
  const allowed=new Set(),notAllowed=new Set();
  for(let r=begin;r<=end;r++){
   const code=String(dt['A'+r]||'').trim();if(!code)continue;
   if(String(dt['Q'+r]||'').trim().toUpperCase()==='Y')allowed.add(code);
   else notAllowed.add(code);
  }
  for(let i=fromSource.length-1;i>=0;i--){
    const f=fromSource[i];if(!allowed.has(f.code)){
      // Even if an older source checked the feature, never submit or export it.
      changes[f.cell]='0';fromSource.splice(i,1);
    }
  }
 }
 // The supplied MR source rule contains FE-C30, but older generated DB tables
 // may omit its row. Show this selectable variant without re-numbering source cells.
 if($('model').value==='MR' && fromSource.length && !fromSource.some(f=>f.code==='FE-C30')){
  const def=(xrules().features||[]).find(f=>f.code==='FE-C30');
  if(def){const idx=fromSource.findIndex(f=>f.code==='FE');
   fromSource.splice(idx>=0?idx+1:0,0,{...def,cell:'FE_C30',virtual:true,apply:0,table:'TblFeature_Emergency'});
  }
 }
 if($('model').value==='MRL') fromSource.push(...(window.SPEC_V52_FEATURES?.virtual('MRL',db)||[]));
 // Supplemental NSCR-H card-reader contact feature. Keep MR and MRL catalogues independent,
 // but present the requested option directly after NSCR-C in each Operation list.
 // Remove any duplicate supplemental records on repeated model loads.
 // NSCR-C is the CAR-call reader, NSCR-H is the HALL-call reader.
 const featureSeen=new Set();
 for(let i=fromSource.length-1;i>=0;i--){
  const key=String(fromSource[i].code||'').trim().toUpperCase();
  if(!['NSCR-C','NSCR-H'].includes(key))continue;
  if(featureSeen.has(key))fromSource.splice(i,1);else featureSeen.add(key);
 }
 if(!fromSource.some(f=>String(f.code).toUpperCase()==='NSCR-H')){
  const after=fromSource.findIndex(f=>String(f.code).toUpperCase()==='NSCR-C');
  const supplementary={table:'TblFeature_Operation',cell:'NSCR_H',code:'NSCR-H',virtual:true,apply:0,option:'A',
   en:'Non-service Temporary Release for Hall Call - Card Reader Type (Only dry contact, not supply Card reader and card)',
   vi:'Hạn chế phục vụ cuộc gọi ở tầng - Dạng thẻ từ (Chỉ cung cấp tiếp điểm khô, không bao gồm đầu đọc thẻ và thẻ từ)',
   suggestion:'Card reader and card by others / Chỉ cung cấp tiếp điểm khô'};
  fromSource.splice(after>=0?after+1:fromSource.length,0,supplementary);
 }

 // ACC is a supplemental traveling cable choice, distinct from the ITV cable.
 // Require the ITV feature in the respective model (never copy the other model's Feature table).
 if(fromSource.some(f=>(window.SPEC_V52_FEATURES?.canonical(f.code)||f.code)==='ITV')){
  fromSource.push({table:'TblFeature_Signal',cell:'AC_TRAVEL_CABLE',code:'ACC',virtual:true,apply:0,option:'A',
   vi:'Cáp động cấp nguồn cho điều hòa cabin, nằm bên trong giếng thang (hố thang máy).',
   en:'Power supply traveling cable for the elevator car air conditioner, inside the hoistway.',suggestion:'Requires ITV'});
 }
 return (fromSource.length?fromSource:xrules().features||[]).filter(f=>String(f.code||'').trim().toUpperCase()!=='MBSS');
}
const lockedNames=new Set(['K7_DoorType','K17_CarW','K17_CarD','K18_CarJJ']);
// MRL: K18_CarJJ is selectable; the capacity row gives a default, not a fixed value.
function capacityLockedField(name){return name!=='K18_CarJJ'||$('model').value==='MR'}
// Material-dependent inputs. Field names are stable workbook-defined names (MR/MRL may use different cells).
const finishDependents=[
 ['K23_WallFinish','K23_WallPaintedColor','K23_WallPattern'],
 ['K24_CarDoorFinish','K24_CarDoorPaintedColor','K24_CarDoorPattern'],
 ['K25_CarTransomFinish','K25_CarTransomPainted','K25_CarTransomPattern'],
 ['K26_CarFrontFinish','K26_CarFrontPainted',null],
 ['K36_JambMainFloorMaterial','K36_JambMainFloorPaint',null],
 ['K36_JambOtherFloorMaterial','K36_JambOtherFloorPaint',null],
 ['K36_DoorFinishMainFloorMaterial','K36_DoorFinishMainFloorPaint',null],
 ['K36_DoorFinishOtherFloorMaterial','K36_DoorFinishOtherFloorPaint',null]
];
const paintedMaterial=v=>/^Painted(?:\b|\s|\-)/i.test(String(v||'').trim());
const patternedMaterial=v=>/^(?:Pattern-printed steel sheet PPS|PPS|Laminated steel sheet)(?:\b|\s|\-)/i.test(String(v||'').trim());
function currentValue(addr){return changes[addr]??current.fields.find(f=>f.cell===addr)?.value??''}
function activeInput(addr){return document.querySelector('[data-cell="'+addr+'"]')}
// v5.8.68: Inline 11,000 mm warning for exactly two stops on BOTH models.
// Read live DOM values for immediate feedback; never change SpecForm inputs or constraints.
function refreshTwoStopTravelWarning(){
 const model=$('model')?.value;
 const travelCell=model==='MR'?'D39':'D35';
 const input=activeInput(travelCell);
 const hint=input?.closest('.field')?.querySelector('.travel-limit-hint');
 if(!input||!hint)return;
 const rule=window.SPEC_ALPHA_5865?.twoStopTravel(model,addr=>activeInput(addr)?.value??currentValue(addr));
 const invalid=Boolean(rule);
 hint.hidden=!invalid;
 hint.textContent=invalid
  ?`⚠ Hai điểm dừng: Travel ${rule.travel.toLocaleString('vi-VN')} mm vượt tối đa 11.000 mm. Vui lòng điều chỉnh Travel xuống ≤ 11.000 mm.`
  :'';
 input.classList.toggle('travel-over-limit',invalid);
 if(invalid)input.setAttribute('aria-invalid','true');
 else input.removeAttribute('aria-invalid');
}
function selectedCapacity(){return current.capacity.find(c=>String(c.EleCapacityName)===String(currentValue('D13')))||{}}
// FlooringWeight: evaluate the original SpecForm Excel expression in the browser.
// AA, BB and JJ are in mm. Thickness 0.02 m; density 2700 kg/m3.
// Hence AA * BB * 0.02 * 2.7 * 0.001 is weight in kg.
let flooringManual=false;
function flooringAddresses(model=$('model').value){
 return model==='MR'
  ?{type:'D59',weight:'D60',aa:'D29',bb:'D30',jj:'D32'}
  :{type:'D55',weight:'D56',aa:'D28',bb:'D29',jj:'D31'};
}
function computedFlooringWeight(){
 const a=flooringAddresses();
 if(String(currentValue(a.type)).trim()!=='Local granite (20 mm)')return 0;
 const aa=Number(currentValue(a.aa)),bb=Number(currentValue(a.bb)),jj=Number(currentValue(a.jj));
 if(![aa,bb,jj].every(Number.isFinite)||aa<=0||bb<=0||jj<=0)return null;
 // Match Excel ROUND(x,0): for nonnegative x, round halves away from zero.
 return (jj<1000?5:10)+Math.floor(aa*bb*0.02*2.7*0.001+0.5);
}
function updateFlooringWeight(force=false){
 if(flooringManual&&!force)return;
 const a=flooringAddresses(),v=computedFlooringWeight();
 if(v===null)return;
 flooringManual=false;
 const input=activeInput(a.weight);
 if(input){input.value=String(v);input.classList.add('dirty');}
 changes[a.weight]=String(v);
}
function flooringDependencyChanged(cell){
 const a=flooringAddresses();
 if([a.type,a.aa,a.bb,a.jj,'D13'].includes(cell)){
  flooringManual=false;
  updateFlooringWeight(true);
 }
}


const NY40_POPUP_DATA={
 'NY-40A':{vi:'Kính phía sau và hai bên',en:'Rear & both sides',img:'assets/NY-40A.png'},
 'NY-40B':{vi:'Kính phía sau',en:'Rear',img:'assets/NY-40B.png'},
 'NY-40C':{vi:'Kính hai bên',en:'Both sides',img:'assets/NY-40C.png'},
 'NY-40D':{vi:'Kính phía sau và một bên',en:'Rear & one side',img:'assets/NY-40D.png'},
 'NY-40E':{vi:'Kính phía sau và một bên',en:'Rear & one side',img:'assets/NY-40E.png'}
};
let ny40Active='NY-40A';
function syncObservationButton(input){
 const field=input?.closest('.field');const btn=field?.querySelector('.obs-preview-btn');
 if(!btn)return;const active=/^NY-40/i.test(String(input.value||''));btn.disabled=!active;btn.title=active?'Xem minh họa kiểu kính':'Chỉ dùng khi chọn NY-40';
}
function selectNy40Option(value){
 const data=NY40_POPUP_DATA[value]||NY40_POPUP_DATA['NY-40A'];ny40Active=data===NY40_POPUP_DATA[value]?value:'NY-40A';
 $('ny40-title').textContent=ny40Active+' — Quan sát';
 $('ny40-image').src=data.img;$('ny40-image').alt=ny40Active;
 $('ny40-vi').textContent=data.vi;$('ny40-en').textContent=data.en;
 for(const el of $('ny40-choices').children)el.classList.toggle('active',el.dataset.value===ny40Active);
}
function openNy40Popup(value){
 if(!/^NY-40/i.test(String(value||'')))return;
 const overlay=$('ny40-popup');if(!overlay)return;
 if(!$('ny40-choices').children.length){
  for(const key of Object.keys(NY40_POPUP_DATA)){
   const b=document.createElement('button');b.type='button';b.className='ny40-choice';b.dataset.value=key;
   b.innerHTML='<img src="'+NY40_POPUP_DATA[key].img+'" alt="'+key+'"><b>'+key+'</b>';
   b.onclick=()=>selectNy40Option(key);
   $('ny40-choices').append(b);
  }
 }
 // Only the upstream car/door/entrance conditions restrict NY-40 choices.
 // Ceiling and handrail choices are validated later by Check Spec, not by this popup.
 const allowed=new Set(allowedObservationCodes());
 for(const b of $('ny40-choices').children){
  b.disabled=!allowed.has(b.dataset.value);
  b.title=b.disabled?'Không phù hợp tải trọng hoặc loại cửa':'';
 }
 selectNy40Option(value);
 overlay.classList.add('show');overlay.setAttribute('aria-hidden','false');
}

function closeNy40Popup(){const overlay=$('ny40-popup');if(overlay){overlay.classList.remove('show');overlay.setAttribute('aria-hidden','true');}}
// Workbook-governed output fields: selected capacity fixes MR counterweight position;
// COP model is an Excel formula, not a user-entered choice.
function refreshCalculatedOutputs(){
 const model=$('model').value;
 if(model==='MR'){
  const cap=selectedCapacity(),pos=String(cap['CWT-Position']||'').trim();
  if(pos){
   const field=current.fields.find(f=>f.name==='K14_CWTLocation');
   if(field){
    const val=pos==='Rear'?'Phía sau phòng thang (Rear drop)':pos==='Side'?'Bên hông phòng thang (Side drop)':pos;
    const el=activeInput(field.cell);
    if(el){
     if(el.tagName==='SELECT'&&![...el.options].some(o=>o.value===val)){
      const op=document.createElement('option');op.value=val;op.textContent=val;el.append(op);
     }
     el.value=val;changes[field.cell]=val;
     el.disabled=true;el.closest('.field')?.classList.add('locked');
     el.title='Tự động theo tải trọng và loại cabin';
    }
   }
  }
 }
 const copField=current.fields.find(f=>f.name==='K31_COPModel');
 if(copField){
  const el=activeInput(copField.cell),db=workbookEngineData();
  if(el&&db&&window.SpecExcelSourceEngine){
   const calc=new window.SpecExcelSourceEngine.SourceEngine(db,changes);
   const result=calc.safe('SpecForm',copField.cell);
   if(result.ok){
    const val=String(result.value??'');
    if(el.tagName==='SELECT'&&![...el.options].some(o=>o.value===val)){
     const op=document.createElement('option');op.value=val;op.textContent=val;el.append(op);
    }
    el.value=val;
    // Keep the dependent result out of changes so the original workbook formula
    // remains authoritative in Source Check and Excel Export.
    delete changes[copField.cell];
   }
   el.disabled=true;el.closest('.field')?.classList.add('locked');
   el.title='Mẫu COP tự động theo nút COP và mã COP';
  }
 }
}
function enforceMRSingleEntrance(){
 // NEXIEZ-MR Door-Gate (SpecForm D20) is fixed to 1D-1G.
 // SpecForm D14 is the actual CO/2S Door type and must retain its Capacity value.
 if($('model').value!=='MR')return;
 const f=current.fields.find(x=>x.cell==='D20'); // Door-Gate type, NOT D14 (CO/2S).
 if(!f)return;
 const el=activeInput(f.cell);
 if(!el)return;
 const fixedValue='1D-1G';
 if(el.tagName==='SELECT'&&![...el.options].some(o=>o.value===fixedValue)){
  const op=document.createElement('option');op.value=fixedValue;op.textContent=fixedValue;el.append(op);
 }
 if(String(currentValue(f.cell))!==fixedValue){changes[f.cell]=fixedValue;el.classList.add('dirty');}
 // Repair the v5.8.59 regression only when D14 contains a Door-Gate token.
 // Never suppress a genuine JJ or Door type validation failure.
 const doorType=activeInput('D14');
 const capacityDoorType=String(selectedCapacity().DoorType||'').trim();
 if(doorType && /^(?:\d+D[-–]\d+G)$/i.test(String(currentValue('D14')).trim()) && /^(?:CO|2S)$/i.test(capacityDoorType)){
  changes.D14=capacityDoorType;doorType.value=capacityDoorType;doorType.classList.add('dirty');
 }
 el.value=fixedValue;
 el.disabled=true;
 el.closest('.field')?.classList.add('locked');
 el.title='NEXIEZ-MR: cố định 1D-1G / Fixed door-gate type';
}
function refreshLocks(){
 // Release only locks previously applied by the workbook's conditional formatting.
 for(const el of document.querySelectorAll('[data-cf-locked="1"]')){
  el.disabled=false;delete el.dataset.cfLocked;el.closest('.field')?.classList.remove('cf-locked','locked');
 }
 const fixed=Boolean(selectedCapacity().DoorType);
 for(const key of lockedNames){const f=current.fields.find(f=>f.name===key);if(!f)continue;let el=activeInput(f.cell);if(!el)continue;const isFixed=fixed&&capacityLockedField(key);el.disabled=isFixed;el.closest('.field')?.classList.toggle('locked',isFixed);el.title=isFixed?'Tự động theo tải trọng':''}
 for(const [parentName,colorName,patternName] of finishDependents){
  const parent=current.fields.find(f=>f.name===parentName);if(!parent)continue;
  const finish=String(currentValue(parent.cell));
  for(const [childName,allowed] of [[colorName,paintedMaterial(finish)],[patternName,patternedMaterial(finish)]]){
   if(!childName)continue;
   const f=current.fields.find(f=>f.name===childName);if(!f)continue;
   const el=activeInput(f.cell);if(!el)continue;
   el.disabled=!allowed;
   el.closest('.field')?.classList.toggle('locked',!allowed);
   el.title=allowed?'':('Không áp dụng với '+finish+' / Not applicable to '+finish);
   if(!allowed && String(currentValue(f.cell)).trim()){
    el.value='';changes[f.cell]='';el.classList.add('dirty');
   }
  }
 }
 // Additional LightUp rules from original Excel (COP, handrail, special material, etc.).
 const db=workbookEngineData(),Engine=window.SpecExcelSourceEngine?.SourceEngine;
 if(db&&Engine&&window.SPEC_CF_CHECK){
  const report=window.SPEC_CF_CHECK($('model').value,changes,db,new Engine(db,changes));
  for(const [addr,hatched] of Object.entries(report.locked)){
   if(!hatched)continue;
   const el=activeInput(addr);if(!el)continue;
   // If an invalid preloaded value exists, leave it editable so Sales can clear it.
   if(String(currentValue(addr)).trim()){
    el.title='Dữ liệu không được phép trong ô xám — hãy xóa giá trị / Clear prohibited data before Check Spec';
    el.closest('.field')?.classList.add('cf-locked');continue;
   }
   el.disabled=true;el.dataset.cfLocked='1';el.closest('.field')?.classList.add('cf-locked','locked');
   el.title='Ô bị khóa theo Conditional Formatting trong Excel / Locked by workbook LightUp rule';
  }
 }

 // MR v5.8.1: N300S/L210S/L200S have one permitted ceiling material,
 // SUS-HL. A stale LightUp source must not grey out this essential selection.
 if($('model').value==='MR'&&['L200S','L210S','N300S'].includes(String(currentValue('D45')||'').trim().toUpperCase())){
  const f=activeInput('D46');
  if(f){
   if(f.tagName==='SELECT'&&!Array.from(f.options).some(o=>o.value==='SUS-HL')){
    const option=document.createElement('option');option.value='SUS-HL';option.textContent='SUS-HL';f.replaceChildren(option);
   }
   f.value='SUS-HL';changes['D46']='SUS-HL';f.disabled=false;
   delete f.dataset.cfLocked;
   f.closest('.field')?.classList.remove('locked','cf-locked');
   f.title='Vật liệu trần cố định: SUS-HL / Fixed ceiling finish: SUS-HL';
  }
 }
 // Fixed factory configuration: counterweight safety gear is not supplied
 // on both MR and MRL Version02. Keep rail fields unrestricted.
 const cwtSafety=current.fields.find(f=>f.name==='K15_CWTSafety');
 if(cwtSafety){
  const el=activeInput(cwtSafety.cell);
  if(el){
   const fixedValue='Không cung cấp (Not provided)';
   if(el.tagName==='SELECT'&&!Array.from(el.options).some(o=>o.value===fixedValue)){
    const option=document.createElement('option');option.value=fixedValue;option.textContent=fixedValue;el.append(option);
   }
   el.value=fixedValue;
   if(String(cwtSafety.value)!==fixedValue)changes[cwtSafety.cell]=fixedValue;
   else delete changes[cwtSafety.cell];
   el.disabled=true;
   el.closest('.field')?.classList.add('locked');
   el.title='Cố định theo tiêu chuẩn / Fixed factory standard';
  }
 }
 enforceMRSingleEntrance();
 refreshCarDesignStates();
 refreshCalculatedOutputs();
 refreshPanelImagePreviews();
}

// Controlled constraints. Workbook is the source of allowed raw variants.
function capacityKg(){
 const cap=String(currentValue('D13'));
 const m=cap.match(/\((\d{3,4})\s*KG\)/i);
 return m?Number(m[1]):null;
}
function eligibleJJ(){
 if($('model').value!=='MRL')return [];
 const cap=selectedCapacity(),db=workbookEngineData();if(!cap.EleCapacityName)return [];
 const base=Number(cap.JJ),out=[];
 if(Number.isFinite(base)&&base>0)out.push(String(base));
 // User's Version02 car-size table: additional JJ values in DataTable CA:CC.
 // Do not put '(Option)' in the Excel input: the cell must remain numeric.
 const sh=db?.sheets?.DataTable||{};
 for(let r=1081;r<=1100;r++){
  if(String(sh['CA'+r]??'')!==String(cap.EleCapacityName))continue;
  const value=Number(sh['CB'+r]);
  if(Number.isFinite(value)&&value>0&&!out.includes(String(value)))out.push(String(value));
 }
 return out;
}
function jjOptionText(value){
 if($('model').value!=='MRL')return String(value);
 const cap=selectedCapacity(),db=workbookEngineData(),sh=db?.sheets?.DataTable||{};
 const mm=String(value);
 if(mm===String(cap.JJ))return mm;
 for(let r=1081;r<=1100;r++){
  if(String(sh['CA'+r]??'')===String(cap.EleCapacityName)&&String(sh['CB'+r])===mm)
   return mm+(String(sh['CC'+r]||'').trim().toLowerCase()==='option'?' (Option)':(String(sh['CC'+r]||'').includes('Note 1')?' (Option - Note 1)':''));
 }
 return mm;
}
function eligibleSideDrops(){
 const db=workbookEngineData();const source=db?.sheets?.DataTable;let vals=($('model').value==='MRL'&&source)?[921,922,923,924].map(r=>source['C'+r]).filter(v=>typeof v==='string'&&v.trim()):[];if(!vals.length)vals=(xrules().sideDrop||[]).filter(Boolean);
 if($('model').value!=='MRL')return vals;
 const kg=capacityKg();
 if(kg===null)return vals;
 if(kg<=1000)return vals.filter(v=>/BWH-620[DE]\b/i.test(v));
 if(kg>=1275)return vals.filter(v=>/BWH-720[FG]\b/i.test(v));
 return [];
}
function refreshSideDrop(){
 if($('model').value!=='MRL')return;
 const f=current.fields.find(f=>f.cell==='D23');const el=activeInput('D23');
 if(!f||!el)return;
 const opts=eligibleSideDrops();
 const previous=String(currentValue('D23'));
 if(el.tagName!=='SELECT')return;
 el.replaceChildren();
 for(const v of opts){const option=document.createElement('option');option.value=v;option.textContent=v;el.append(option)}
 // D23 is a constrained selection: never retain a wrong group from previous capacity.
 const next=opts.includes(previous)?previous:(opts[0]||'');
 el.value=next;
 if(next!==previous)changes.D23=next;
 f.options=[...opts];
 const fixed=opts.length===1;
 el.disabled=fixed||opts.length===0;
 el.closest('.field')?.classList.toggle('locked',el.disabled);
 el.title=opts.length?'Giới hạn theo tải trọng / Capacity-based option':'Không có Side drop hợp lệ / No valid side drop';
}
function featureIsSelected(f){const v=changes[f.cell]??f.apply;return v===true||v===1||String(v).toLowerCase()==='true'||String(v)==='1'||String(v)==='✓'||String(v).toLowerCase()==='v'}
function validateSpecPartial(){
 const errors=[],warnings=[],verified=[];
 const err=(cell,msg)=>errors.push(cell+': '+msg);
 // Check every changed input, and every declared list for which all candidates are known.
 for(const f of current.fields){
  const el=activeInput(f.cell);if(!el)continue;
  const value=String(currentValue(f.cell)).trim();
  if(el.disabled)continue;
  const choices=f.cell==='D23'&&$('model').value==='MRL'?eligibleSideDrops():technicalOptions(f);
  if(choices.length && value && !choices.map(String).includes(value)){
   // An unchanged Excel value can be a valid result of a conditional formula not yet ported.
   if(Object.hasOwn(changes,f.cell)||el.tagName==='SELECT')err(f.cell,'Giá trị không thuộc lựa chọn hiện hành / Value outside allowed list: '+value);
  }
 }
 verified.push('Danh sách lựa chọn đang hiển thị / Active selections');
 if($('model').value==='MRL'){
  const kg=capacityKg(),side=String(currentValue('D23'));
  const opts=eligibleSideDrops();
  if(kg===null)err('D13','Không xác định được tải trọng / Unknown rated load');
  else if(!opts.includes(side))err('D23','Side drop không đúng nhóm tải '+kg+' kg / Invalid side drop for capacity');
  verified.push('Side drop theo tải trọng (≤1000 kg: 620D/E; ≥1275 kg: 720F/G)');
  for(const cell of ['D14','D16']){
   const f=current.fields.find(x=>x.cell===cell);if(!f)continue;
   const allowed=technicalOptions(f).map(String),v=String(currentValue(cell));
   if(allowed.length&&!allowed.includes(v))err(cell,'Không đúng điều kiện Capacity / Entrance / Observation');
  }
  verified.push('Observation & Entrance (quy tắc đã triển khai)');
 }
 for(const [parentName,colorName,patternName] of finishDependents){
  const parent=current.fields.find(f=>f.name===parentName);if(!parent)continue;
  const finish=String(currentValue(parent.cell));
  for(const [childName,allowed] of [[colorName,paintedMaterial(finish)],[patternName,patternedMaterial(finish)]]){
   if(!childName)continue;
   const child=current.fields.find(f=>f.name===childName);if(!child)continue;
   if(!allowed&&String(currentValue(child.cell)).trim())err(child.cell,'Màu/hoa văn không áp dụng với '+finish+' / Inapplicable finish selection');
  }
 }
 verified.push('Vật liệu và màu sơn/hoa văn / Material and color/pattern');
 for(const f of xrules().features||[]){
  if(String(f.option).trim().toUpperCase()==='X'&&featureIsSelected(f))err(f.cell,'Feature '+f.code+' không áp dụng (X) / Not applicable');
  if(String(f.option).trim().toUpperCase()==='D'&&featureIsSelected(f))warnings.push(f.cell+': '+f.code+' cần Technical Inquiry (D) / Requires TI');
 }
 verified.push('Feature X; Feature D cảnh báo TI / Prohibited and TI-required features');
 const cap=selectedCapacity();if(cap.DoorType){
  const mapping={K7_DoorType:'DoorType',K17_CarW:'AA',K17_CarD:'BB',K18_CarJJ:'JJ'};
  for(const [name,column] of Object.entries(mapping)){
   const f=current.fields.find(x=>x.name===name);if(!f)continue;
   const ref=cap[column];if(ref===null||ref===undefined||String(ref)==='')continue;
   if(String(currentValue(f.cell))!==String(ref))err(f.cell,'Khác dữ liệu Capacity / Expected '+ref);
  }
 }
 verified.push('Giá trị kích thước cố định theo Capacity (các cột đã ánh xạ)');
 warnings.push('Chưa xác minh đầy đủ công thức động, Step 1 và logic Export của Tool.exe / Full formula evaluation, C# Step 1 and Export not verified');
 return {errors,warnings,verified};
}
function applyValue(addr,v){const el=activeInput(addr);if(!el)return;v=String(v??'');if(el.tagName==='SELECT'&&![...el.options].some(o=>o.value===v)){let o=document.createElement('option');o.value=v;o.textContent=v;el.append(o)}el.value=v;changes[addr]=v;el.classList.add('dirty')}
function setupSpecialOptions(){
 const side=activeInput('D23');if(side&&$('model').value==='MRL'){
  const f=current.fields.find(x=>x.cell==='D23'),old=String(currentValue('D23'));
  const sel=document.createElement('select');sel.dataset.cell='D23';
  side.replaceWith(sel);if(f)f.options=eligibleSideDrops();
  refreshSideDrop();
  sel.addEventListener('change',()=>{changes.D23=sel.value;sel.classList.add('dirty');refreshMRLShaftMinimum();invalidateCheck();status('Đã cập nhật loại đối trọng; AH/BH tối thiểu được tra lại theo CWT.')});
 }
 const obsField=current.fields.find(f=>/Observation/i.test(f.name));const obs=obsField?activeInput(obsField.cell):null;if(obs && obs.tagName!=='SELECT' && $('model').value==='MR'){
  const f=obsField;const values=['No','NY-40A','NY-40B','NY-40C','NY-40D','NY-40E'];let s=document.createElement('select');s.dataset.cell=obsField.cell;for(let v of values){let o=document.createElement('option');o.value=v;o.textContent=v;s.append(o)}s.value=String(obs.value||'No');s.addEventListener('change',()=>{changes[obsField.cell]=s.value;invalidateCheck()});obs.replaceWith(s);if(f)f.options=values;
 }
}
// ITP lives in the original Signal Feature list; there is no pinned panel.
// Still choose the first valid ITP type on load, even when the Signal tab is closed.
function ensureITPDefault(features){
 const feature=features.find(f=>f.code==='ITP');if(!feature)return;
 const cell=itpCell();
 const enabled=['1','TRUE','V','✓'].includes(String(changes[feature.cell]??feature.apply).toUpperCase());
 const db=workbookEngineData();
 const eng=db&&window.SpecExcelSourceEngine?new window.SpecExcelSourceEngine.SourceEngine(db,changes):null;
 const allowed=eng?.safe('SpecForm','E'+feature.cell.slice(1));
 const disabled=allowed?.ok&&String(allowed.value).trim().toUpperCase()==='X';
 if(!enabled||disabled){delete changes[cell];return;}
 const options=itpOptions();
 if(!options.includes(String(changes[cell]??''))){
  if(options.length)changes[cell]=options[0];else delete changes[cell];
 }
}
function ensureFreeFeatureDefaults(features){
 const model=$('model').value,ctl=String(currentValue(model==='MR'?'D21':'D22')||'');
 const v52=window.SPEC_V52_FEATURES;if(!v52)return;
 const db=workbookEngineData();
 for(const f of features){
  const code=v52.canonical(f.code),status=v52.status(model,code,ctl,db);
  if(v52.complimentary.has(code)){
   if(status==='X'){changes[f.cell]='0';featureTouched.delete(f.cell)}
   else if(status&&status!=='X'&&!featureTouched.has(f.cell))changes[f.cell]='1';
  }else if(f.virtual&&status==='X'){changes[f.cell]='0';featureTouched.delete(f.cell)}
 }
}
// v5.8.108 — Natural-light advisory: selected MBS + glass car door or NY-40 observation.
function shouldShowMBSNaturalLightNote(){
 const doorField=current?.fields?.find(f=>f.name==='K24_CarDoorFinish');
 const door=String(doorField?currentValue(doorField.cell):'');
 const observation=String(currentValue('D14')||'').trim();
 const glassDoor=/(?:Large\s*window|DSJ1|\bDS1\b)/i.test(door);
 return glassDoor || ($('model').value==='MRL' && /^NY-?40[A-E]?\b/i.test(observation));
}
function featureVisualInfo(code){
 const key=String(code||'').trim();
 return SPEC_FEATURE_VISUALS[key]||SPEC_FEATURE_VISUALS[(window.SPEC_V52_FEATURES?.canonical(key)||key)]||null;
}
function appendFeatureVisual(desc,code){
 const info=featureVisualInfo(code);
 if(!info||!info.src)return;
 const wrap=document.createElement('span');wrap.className='feature-visual-inline';
 const img=document.createElement('img');img.className='feature-visual-thumb';img.loading='lazy';img.decoding='async';img.src=info.src;img.alt=info.alt||code;img.title=info.caption||('Hình minh họa '+code);
 wrap.append(img);
 if(info.caption){const cap=document.createElement('small');cap.className='feature-visual-caption';cap.textContent=info.caption;wrap.append(cap);}
 desc.append(wrap);
}

function renderFeatures(){
 const features=sourceFeatureRows();ensureFreeFeatureDefaults(features);const tables=[...new Set(features.map(f=>f.table))];const tabs=$('feature-tabs'),main=$('features');if(!tabs||!main)return;tabs.replaceChildren();main.replaceChildren();ensureITPDefault(features);
 if(!tables.length){main.textContent='Chưa trích xuất được bảng Features / Features unavailable';return}
 if(!tables.includes(activeFeatureTable))activeFeatureTable=tables[0];
 tables.forEach(t=>{let b=document.createElement('button');const selected=features.some(f=>f.table===t&&featureIsSelected(f));b.className='btn feature-tab'+(selected?' feature-tab-selected':'')+(t===activeFeatureTable?' feature-tab-current':'');b.textContent=({TblFeature_Emergency:'Khẩn cấp / Emergency',TblFeature_Operation:'Vận hành / Operation',TblFeature_Door:'Cửa / Door',TblFeature_Signal:'Tín hiệu / Signal',TblFeature_Traffic:'Giao thông / Traffic'})[t]||t;b.onclick=()=>{activeFeatureTable=t;renderFeatures()};tabs.append(b)});
 const db=workbookEngineData(),featureEngine=db&&window.SpecExcelSourceEngine?new window.SpecExcelSourceEngine.SourceEngine(db,changes):null;
 for(const f of features.filter(f=>f.table===activeFeatureTable)){
  const line=document.createElement('div');line.className='feature-line';let addr=document.createElement('span');addr.textContent=f.cell;addr.className='addr';let code=document.createElement('span');code.textContent=f.code;code.className='fcode';let desc=document.createElement('div');desc.className='fdesc';desc.textContent=window.SPEC_PROOFREAD?.(f.vi||f.code)??(f.vi||f.code);let en=document.createElement('small');en.textContent=window.SPEC_PROOFREAD?.(f.en||'')??(f.en||'');desc.append(en);appendFeatureVisual(desc,window.SPEC_V52_FEATURES?.canonical(f.code)||f.code);
  let chk=document.createElement('input');chk.type='checkbox';chk.dataset.cell=f.cell;chk.checked=String(changes[f.cell]??f.apply)==='1'||String(changes[f.cell]??f.apply).toLowerCase()==='true';const optResult=f.virtual?null:featureEngine?.safe('SpecForm','E'+f.cell.slice(1));const manualOption=window.SPEC_V52_FEATURES?.status($('model').value,f.code,currentValue($('model').value==='MR'?'D21':'D22'),db);const dynamicOption=manualOption|| (optResult?.ok?optResult.value:f.option);let disabled=String(dynamicOption).trim().toUpperCase()==='X';
  const fCode=window.SPEC_V52_FEATURES?.canonical(f.code)||f.code;
  // NSCR-C and NSCR-H are independent options; the card-reader variant
  // must never lock or overwrite the card-reader-free NSCR-C selection.
  if(fCode==='NSCR-C'){
    const model=$('model').value;
    const ctl=currentValue(model==='MR'?'D21':'D22');
    const sourceStatus=window.SPEC_V52_FEATURES?.status(model,'NSCR-C',ctl,db);
    // Keep the native NSCR-C option available independently from virtual NSCR-H.
    // Technical requirements (e.g. EVRC-C/S for COP C760) remain Check Spec rules.
    if(sourceStatus==='A'||sourceStatus==='S')disabled=false;
    line.title='NSCR-C: hạn chế cuộc gọi trong cabin (CAR CALL). Khác NSCR-H: cuộc gọi tại tầng (HALL CALL).';
  }
  if(fCode==='HE-B'){
   const selected=name=>features.some(z=>(window.SPEC_V52_FEATURES?.canonical(z.code)||z.code)===name&&featureIsSelected(z));
   const missing=['DKO-TB','EXCL'].filter(name=>!selected(name));
   line.title='HE-B chỉ hợp lệ khi chọn cả DKO-TB và EXCL / Requires both DKO-TB and EXCL.';
   if(missing.length){const requirement=document.createElement('small');requirement.className='feature-dependency-note';
    requirement.textContent='Yêu cầu / Requires: DKO-TB + EXCL'+(chk.checked?' (còn thiếu / missing: '+missing.join(', ')+')':'');
    desc.append(requirement);}
  }
  if(fCode==='NSCR-H'){
    line.title='NSCR-H: hạn chế cuộc gọi tại tầng (HALL CALL). Không thay thế NSCR-C (CAR CALL).';
  }
  if(fCode==='ACC'){
   const itv=features.find(z=>(window.SPEC_V52_FEATURES?.canonical(z.code)||z.code)==='ITV');
   const itvEnabled=!!itv&&featureIsSelected(itv);
   disabled=!itvEnabled;
   if(!itvEnabled){changes.AC_TRAVEL_CABLE='0';chk.checked=false;}
   line.title=itvEnabled?'':'Yêu cầu bật tính năng ITV / ITV must be selected first';
  }
  if($('model').value==='MRL'&&fCode==='AIL'&&!window.SPEC_HALL_LANTERN?.any(changes))disabled=true;
  if($('model').value==='MRL'&&fCode==='SETS-L'&&disabled&&chk.checked){
    // Source Q472 disallows SETS-L at 60 m/min. Clear a stale selection
    // inherited from an earlier speed so D520 cannot remain FALSE with a
    // disabled checkbox that the user cannot untick.
    changes[f.cell]='0';featureTouched.delete(f.cell);chk.checked=false;
  }
  chk.disabled=disabled;if(chk.checked&&!disabled)line.classList.add('feature-line-selected');if(disabled)line.classList.add('x');chk.onchange=()=>{
   const code=window.SPEC_V52_FEATURES?.canonical(f.code)||f.code;
   if($('model').value==='MRL'&&code==='FHL'&&!chk.checked&&window.SPEC_HALL_LANTERN?.any(changes)){
     chk.checked=true;changes[f.cell]='1';status('FHL phải đi kèm Hall Lantern đã chọn / FHL is required while lanterns are selected.');return;
   }
   featureTouched.add(f.cell);changes[f.cell]=chk.checked?'1':'0';line.classList.toggle('feature-line-selected',chk.checked);
   if(code==='HE-B'&&chk.checked){
     const active=c=>sourceFeatureRows().some(f=> (window.SPEC_V52_FEATURES?.canonical(f.code)||f.code)===c&&featureIsSelected(f));
     const missing=['DKO-TB','EXCL'].filter(c=>!active(c));
     if(missing.length)status('HE-B yêu cầu chọn đồng thời DKO-TB và EXCL. Còn thiếu: '+missing.join(', '));
   }tabs.querySelectorAll('.feature-tab').forEach(b=>{const group=tables.find(t=>b.textContent===(({TblFeature_Emergency:'Khẩn cấp / Emergency',TblFeature_Operation:'Vận hành / Operation',TblFeature_Door:'Cửa / Door',TblFeature_Signal:'Tín hiệu / Signal',TblFeature_Traffic:'Giao thông / Traffic'})[t]||t));if(group)b.classList.toggle('feature-tab-selected',features.some(x=>x.table===group&&featureIsSelected(x)))});
   // v5.8.41: EVRC-C (not EVRC-S) is mandatory with NSCR-C/AS/SCS-B/NS-CB
   // for MR C760/C770/C780 and MRL V02 M*/P* COP codes.
   if(chk.checked&&['NSCR-C','AS','SCS-B','NS-CB'].includes(code)){
     const model=$('model').value;
     const copCell=model==='MR'?'D65':'D61';
     const selected=name=>features.some(x=>
       (window.SPEC_V52_FEATURES?.canonical(x.code)||x.code)===name&&featureIsSelected(x));
     const rule=window.SPEC_EVRC_C_RULE?.check(model,{ok:true,value:currentValue(copCell)},selected);
     if(rule?.required&&!rule.evrcC)status('COP '+rule.cop+' với '+rule.triggered.join(', ')+': bắt buộc chọn EVRC-C trong Vận hành / Operation để Check Spec đạt.');
   }
   // NSCR-C changes only its own source cell. NSCR-H remains fully independent.
   if(code==='ITV'&&!chk.checked){changes.AC_TRAVEL_CABLE='0';featureTouched.delete('AC_TRAVEL_CABLE');}

   if($('model').value==='MRL'&&chk.checked&&['AECC','AECH','AIL'].includes(code)){
     // Apply the exclusivity notes from the Version02 Sales Manual.
     const conflicts=code==='AECC'?['AECH']:code==='AECH'?['AECC','AIL']:['AECH'];
     for(const other of sourceFeatureRows())if(conflicts.includes(window.SPEC_V52_FEATURES?.canonical(other.code)||other.code)){
       changes[other.cell]='0';featureTouched.add(other.cell);
     }
   }
   if(f.code==='FE-C30'&&chk.checked)changes.A93='0';
   if(f.code==='FE'&&chk.checked)changes.FE_C30='0';
   if($('model').value==='MRL'&&f.cell==='A157')refreshMRLOverheadMinimum(true);
   if(code==='MBS')renderFeatures();
  if(['FE','FE-C30','ITP','ITV','AECC','AECH','AIL'].includes(code)){
     if(code==='ITP'&&!chk.checked)delete changes[itpCell()];renderFeatures();
   }
   invalidateCheck();status('Đã cập nhật Feature / Feature updated. Vui lòng chạy lại Check Spec.');
  };
  let op=document.createElement('span');op.className='option';op.textContent=dynamicOption;let suggest=document.createElement('span');suggest.className='suggest';suggest.textContent=f.suggestion;line.append(addr,code,desc,chk,op,suggest);if(f.code==='ITP'&&chk.checked&&!chk.disabled)renderITPOptions(line);main.append(line);
  if(fCode==='MBS' && chk.checked && !chk.disabled && shouldShowMBSNaturalLightNote()){
   const advisory=document.createElement('div');advisory.className='mbs-natural-light-note';advisory.setAttribute('role','note');
   advisory.textContent='Lưu ý: Chùm quang phổ ánh sáng tự nhiên có thể ảnh hưởng đến hiệu quả của mành tia MBS.';
   main.append(advisory);
  }
 }
}


// ITP choices always come from DataTable!DN:DP of the current workbook.
function itpCell(model=$('model').value){return model==='MR'?'H175':'H177'}
function itpFeatureCell(model=$('model').value){return model==='MR'?'A175':'A177'}
function itpOptions(model=$('model').value){
 const db=workbookEngineData(),group=String(currentValue(model==='MR'?'D21':'D22')).trim();
 const fe=sourceFeatureRows().find(x=>x.code==='FE');
 const feOn=!!fe && ['1','TRUE','V','✓'].includes(String(changes[fe.cell]??fe.apply).toUpperCase());
 const rows=[];
 for(let r=2;r<=24;r++){
  const c=String(db?.sheets?.DataTable?.['DN'+r]??'').trim();
  if(!c)break;
  if(c!==group)continue;
  const label=String(db.sheets.DataTable['DO'+r]??'').trim();
  if(!label)continue;
  const compulsory=String(db.sheets.DataTable['DP'+r]??'').trim().toUpperCase()==='Y';
  if(feOn&&group==='1C-2BC'&&!compulsory)continue;
  rows.push(label);
 }
 return rows;
}
function renderITPOptions(line){
 const cell=itpCell();const options=itpOptions();
 const area=document.createElement('div');area.className='feature-itp-config';
 const label=document.createElement('label');label.htmlFor='itp-type';label.textContent='Loại liên lạc nội bộ / Interphone type';
 const select=document.createElement('select');select.id='itp-type';select.dataset.cell=cell;
 // Default to the first valid row of the current model/operation/FE list.
 // Keep explicit user choices when they remain available after a re-render.
 // Save the default into changes so Check Spec and both exports use the visible choice.
 const previous=Object.hasOwn(changes,cell)?String(changes[cell]):'';
 if(options.length){
  for(const value of options){const o=document.createElement('option');o.value=value;o.textContent=value;select.appendChild(o)}
  const selected=options.includes(previous)?previous:options[0];
  select.value=selected;
  changes[cell]=selected;
 }else{
  const placeholder=document.createElement('option');placeholder.value='';placeholder.textContent='— Không có loại ITP phù hợp —';select.appendChild(placeholder);
  select.value='';
  delete changes[cell];
  select.disabled=true;
 }
 select.onchange=()=>{changes[cell]=select.value;invalidateCheck();status('Đã cập nhật loại ITP. Vui lòng chạy lại Check Spec.')};
 area.append(label,select);
 if(options.length===1){const hint=document.createElement('small');hint.textContent='FE / FE Hong Kong: bắt buộc chọn loại 3 Master Stations.';area.append(hint)}
 line.append(area);
}

// Dropdowns are resolved from the *active* workbook (including conditional Excel formulas).
// Never infer a selectable field only from its current displayed text.
const supplementalListSources={
 K14_CWTLocation:['TblK14_CWTLocation','CWTLocation'],
 K36_JambMainFloor:['TblK36_Jamb','Jamb'],
 K36_JambOtherFloor:['TblK36_Jamb','Jamb'],
 K36_JambMainFloorMaterial:['TblK36_EntranceDoorFinish','Finish'],
 K36_JambOtherFloorMaterial:['TblK36_EntranceDoorFinish','Finish'],
 K36_DoorFinishMainFloorMaterial:['TblK36_EntranceDoorFinish','Finish'],
 K36_DoorFinishOtherFloorMaterial:['TblK36_EntranceDoorFinish','Finish']
};
// Landing-door finish options are NOT the two-row entrance-finish table.
// Both models use an explicit workbook Data Validation source for these cells:
// MR: DataTable!AJ617:AJ621; MRL V02: DataTable!B930:B939.
// Read that source, preserving DSJ1/DS1/Large Window and any future changes
// made in the user's imported source workbook.
function landingDoorChoices(field,workbook){
 if(!['K36_DoorFinishMainFloorMaterial','K36_DoorFinishOtherFloorMaterial'].includes(field.name))return null;
 const model=$('model').value;
 const range=model==='MR'?['AJ',617,621]:['B',930,939];
 const sheet=workbook?.sheets?.DataTable||{};
 const result=[];
 for(let r=range[1];r<=range[2];r++){
  const v=sheet[range[0]+r];
  const raw=v&&typeof v==='object'&&'v' in v?v.v:v;
  const value=String(raw??'').trim();
  if(value&&!result.includes(value))result.push(value);
 }
 return result.length?result:null;
}
function workbookChoices(field,engine=null){
 // Source XLSX already validates this Table 4.3 exception in R14, but its
 // legacy data-validation ranges for D14/D16 don't expose the valid choice.
 // These two fields must use the manual-specific UI choices, independent of the
 // old INDIRECT dropdown; Check Spec still evaluates R13/R14/D525.
 if($('model').value==='MRL'&&['D14','D16'].includes(field.cell))return technicalOptions(field);
 // NEXIEZ-MR hospital bed handrail Rev.AM: YH-56S, not YH-55S.
 if($('model').value==='MRL'&&field.name==='K32_Handrail'){
  return [...new Set([...(field.options||[]),'YH-100SL (SUS304)'])];
 }
 if($('model').value==='MR'&&field.cell==='D67'&&/^B(?:750|1000)-2S\b/.test(String(currentValue('D13')))){
  const model=workbookEngineData(),dbTable=model?.tables?.TblK32_Handrail;
  const choices=(field.options||[]).slice();
  const allowed='YH-56S (SUS304)';if(!choices.includes(allowed))choices.push(allowed);
  return choices;
 }
 // NEXIEZ-MRL Version02 — G11L1P Sales Manual Rev.M, Table 2.1-C.
 // Speed 60 m/min: 22 stops; 96/105/120/150 m/min: 30 stops.
 // Preserve an old incompatible selection as invalid so Check Spec reports it.
 if($('model').value==='MRL'&&field.name==='K9_Stops'){
  const speed=Number(currentValue('D17'));
  const maxStops=speed===60?22:30;
  return Array.from({length:maxStops-1},(_,i)=>String(i+2));
 }
 if(field.name==='K35_Ventilation'&&$('model').value==='MRL')return ['No','DC axial flow fan','DC axial flow fan + Plasma Quad'];
 if(field.name==='K22_CeilingColor'&&window.SPEC_CEILING_FINISH){
  const model=$('model').value,ceiling=String(currentValue(model==='MR'?'D45':'D41'));
  return window.SPEC_CEILING_FINISH.options(model,ceiling);
 }
 const workbook=workbookEngineData();
 let list=[];
 if(field.validation&&workbook&&window.SpecExcelSourceEngine){
  try{
   const calc=engine||new window.SpecExcelSourceEngine.SourceEngine(workbook,changes);
   let v=calc.eval(calc.ast(field.validation),{sheet:'SpecForm',addr:field.cell});
   list=v?.isGrid?v.values():Array.isArray(v)?v.flat(Infinity):[v];
   // Literal Data Validation list (e.g. "1,2,3,4") is not an Excel range.
   if(list.length===1&&typeof list[0]==='string'&&/^".*"$/.test(String(field.validation).trim()))list=String(list[0]).split(',');
   list=list.map(v=>String(v??'').trim()).filter(Boolean);
  }catch(err){
   // INDIRECT("") intentionally means "not applicable". Retain the original
   // list only when the evaluator cannot resolve a supported static source.
   if(!/INDIRECT\(""\)/i.test(field.validation))list=field.options||[];
  }
 }
 if(!field.validation&&supplementalListSources[field.name]&&workbook&&window.SpecExcelSourceEngine){
  try{
   const [table,col]=supplementalListSources[field.name];
   const calc=engine||new window.SpecExcelSourceEngine.SourceEngine(workbook,changes);
   const vals=calc.eval(calc.ast(`${table}[${col}]`),{sheet:'SpecForm',addr:field.cell});
   list=vals?.isGrid?vals.values():[];
  }catch(err){list=[]}
 }
 if(!field.validation&& !supplementalListSources[field.name])list=technicalOptions(field);
 if(field.cell==='D23'&&$('model').value==='MRL')list=eligibleSideDrops();
 if(field.cell==='D65'&&$('model').value==='MRL'&&ny40PositionChoices())list=ny40PositionChoices();
 // v5.8.65: MRL Version02 only. Limit visible choices but NEVER overwrite an
 // existing selection; Check Spec flags incompatible saved/imported values.
 if($('model').value==='MRL'&&field.cell==='D62'&&String(currentValue('D60')).trim()==='Side')list=['No'];
 // v5.8.76: Default Handrail Type (D64) never presets or narrows
 // Handrail Quantity / Location (D65). Use its independent workbook choices.
 // Observation-car location choices remain governed by NY-40A..E above;
 // invalid combinations are reported by Check Spec, not silently replaced.
 if(field.name==='K18_CarJJ'&&$('model').value==='MRL')list=eligibleJJ();
 if(field.name==='K33_Mirror'&&window.SPEC_MIRROR_RULES){
   const model=$('model').value;
   list=window.SPEC_MIRROR_RULES.allOptions(model,currentValue('D13'),model==='MR'?currentValue('D29'):currentValue('D28'),model==='MR'?currentValue('D31'):undefined);
 }
 // Use the actual landing-door validation range rather than the two-item
 // TblK36_EntranceDoorFinish list used by door jamb and basic material fields.
 const landingOptions=landingDoorChoices(field,workbook);
 if(landingOptions)list=landingOptions;
 list=window.SPEC_MATERIAL_5825?.choices($('model').value,field,list)||list;
 // NY-40 design guide: restrict finishes for front return, transom and kickplate.
 // Existing incompatible values remain visibly invalid (syncDropdownOptions) and
 // are reported by SPEC_NY40_RULES.validate, rather than silently overwritten.
 if($('model').value==='MRL'&&window.SPEC_NY40_RULES?.allowedLocations(String(currentValue('D14')))
    &&['K26_CarFrontFinish','K25_CarTransomFinish','K27_Kickplate'].includes(field.name))list=['SUS-HL'];
 // Car wall is the master material for the transom panel in both MR and MRL.
 // Preserve factory restrictions (e.g. NY-40 SUS-HL): if the wall material is
 // not in the allowed Transom list, keep the actual allowed list and let
 // Check Spec flag the incompatible wall/transom combination.
 if(field.name==='K25_CarTransomFinish'){
  const wall=String(currentValue($('model').value==='MR'?'D47':'D43')||'').trim();
  if(wall&&list.some(x=>String(x).trim()===wall))list=[wall];
 }
 // v5.8.111: MR-specific UI catalog. The original GXLM2P workbook contains
 // legacy PIV "H" entries, but the MR order-entry dropdown excludes that suffix.
 // MRL Version02 is NOT affected; keep the actual workbook unchanged.
 if($('model').value==='MR'&&/^K36_PIV(?:MainFloor|OtherFloor)$/.test(field.name)){
   list=list.filter(v=>!/^PIV\d+-[A-Z]\d{3,4}H$/i.test(String(v).trim()));
 }
 return [...new Set(list.map(v=>String(v??'').trim()).filter(Boolean))];
}
function hasDropdown(field){return Boolean(field.validation||supplementalListSources[field.name]||(field.options||[]).length||field.cell==='D23'&&$('model').value==='MRL'||field.name==='K18_CarJJ'&&$('model').value==='MRL')}
function syncDropdownOptions(){
 const workbook=workbookEngineData();if(!workbook||!window.SpecExcelSourceEngine)return;
 const engine=new window.SpecExcelSourceEngine.SourceEngine(workbook,changes);
 for(const f of current.fields){
  const el=activeInput(f.cell);if(!el||el.tagName!=='SELECT')continue;
  // v5.8.157: the group selector owns the filtered MR D13 options.
  if($('model').value==='MR'&&f.name==='K5_Capacity'&&document.getElementById('mr-capacity-group'))continue;
  // Side-drop filtering is handled separately and must retain its capacity limits.
  if(f.cell==='D23'&&$('model').value==='MRL')continue;
  let opts=workbookChoices(f,engine),old=String(currentValue(f.cell)??'');
  if((f.cell==='D14'||f.cell==='D16')&&$('model').value==='MRL'){
   // The dynamic formulas in SpecForm are authoritative; if unavailable,
   // keep the last known Workbook-compatible options rather than a free text box.
   if(!opts.length)opts=technicalOptions(f);
  }
  // MR suffix-S ceilings have one fixed SUS-HL finish. Never retain a stale
  // Y033/Y055/Y073 color as an invalid extra dropdown choice after model change.
  if($('model').value==='MR'&&f.name==='K22_CeilingColor'&&opts.length===1&&opts[0]==='SUS-HL'){
   if(old!=='SUS-HL'){
    changes[f.cell]='SUS-HL';
    old='SUS-HL';
    el.classList.add('dirty');
   }
  }
  const before=[...el.options].map(o=>o.value);
  const isJJ=$('model').value==='MRL'&&f.name==='K18_CarJJ';
  if(before.length===opts.length&&before.every((v,i)=>v===opts[i]) &&
     (!isJJ||[...el.options].every((op,i)=>op.textContent===jjOptionText(opts[i]))))continue;
  el.replaceChildren();
  for(const v of opts){const op=document.createElement('option');op.value=v;op.textContent=isJJ?jjOptionText(v):v;el.append(op)}
  if(!opts.length){
   const op=document.createElement('option');op.value='';op.textContent='Không áp dụng';el.append(op);
   if(old&&Object.prototype.hasOwnProperty.call(changes,f.cell))changes[f.cell]='';
   continue;
  }
  // Do not silently replace an invalid technical specification with a convenient
  // default: keep the original as an invalid visible choice so Check Spec can flag it.
  if(old&&!opts.includes(old)){
   const invalid=document.createElement('option');invalid.value=old;invalid.textContent=old+' (không hợp lệ)';
   invalid.dataset.invalid='true';el.prepend(invalid);
  }
  el.value=old&&[...el.options].some(x=>x.value===old)?old:(opts[0]||'');
 }
}

// Mitsubishi design guide — visual reference beside the existing source-compatible dropdowns.
// The picture catalogue is deliberately separate from workbook authorization: a reference
// illustration must never make an unsupported Excel option pass Check Spec.
function designKindFor(field){
 if(field.name==='K22_CeilingModel')return 'ceiling';
 // Keep the original ceiling-color dropdown, but do not show a duplicated image/description panel here.
 if(field.name==='K32_Handrail')return 'handrail';
 if(field.name==='K33_Mirror')return 'mirror';
 return '';
}
function guideEntries(kind){
 const model=$('model').value;
 const items=window.SPEC_DESIGN_GUIDE?.[model]?.[kind==='ceilingColor'?'ceiling':kind]||[];
 if(model!=='MR'||!['ceiling','ceilingColor'].includes(kind))return items;
 const capacity=String(currentValue('D13')||'').trim();
 const useTwoLightS00=/^(?:B750|B1000|P20C4)(?:-|\s|\(|$)/i.test(capacity);
 return items.map(item=>{
  if(item.code!=='S00')return item;
  if(useTwoLightS00)return {...item,subtitle:'Trần S00 hai đèn – B750 / B1000 / P20C4',
    desc:'Mẫu trần S00 hai đèn dành riêng cho NEXIEZ-MR B750, B1000 và P20C4.'};
  return {...item,images:['assets/design/mr_ceiling_s00.jpg'],
    subtitle:'Trần S00 tiêu chuẩn',
    desc:'Mẫu trần S00 tiêu chuẩn theo catalogue cho các cấu hình NEXIEZ-MR còn lại.'};
 });
}
function selectedGuideItem(kind){
 const items=guideEntries(kind);let raw='';
 if(kind==='ceiling'||kind==='ceilingColor')raw=String(currentValue($('model').value==='MR'?'D45':'D41'));
 else raw=String(currentValue($('model').value==='MR'?(kind==='handrail'?'D67':'D69'):(kind==='handrail'?'D64':'D66')));
 return [...items].sort((a,b)=>b.code.length-a.code.length).find(x=>raw.toUpperCase().startsWith(x.code.toUpperCase()))||null;
}
function designSummaryFor(kind){
 const item=selectedGuideItem(kind);
 if(kind==='ceilingColor'){
  if(!item)return 'Chọn mẫu trần để xem màu sắc và vật liệu phù hợp.';
  return item.code+' · '+($('model').value==='MR'?'Màu trần theo SpecForm · xem vật liệu tham khảo':item.code==='DL6'?'Chỉ Stainless-steel / Laminated steel sheet (không Painted). Excel gốc chưa đủ mã vật liệu.':'Xem ảnh và các vật liệu / See material options');
 }
 if(!item)return 'Xem bảng mẫu tham khảo & mô tả thiết kế';
 return item.code+' · '+item.subtitle;
}
function refreshDesignHints(){
 for(const wrap of document.querySelectorAll('.design-inline[data-design-kind]')){
  const kind=wrap.dataset.designKind, item=selectedGuideItem(kind),img=wrap.querySelector('img'),hint=wrap.querySelector('.design-hint');
  if(img){
   if(item?.images?.[0]){img.src=item.images[0];img.alt=item.title;img.hidden=false;}
   else {img.removeAttribute('src');img.hidden=true;}
  }
  if(hint){hint.textContent=kind==='mirror'?'':designSummaryFor(kind);hint.hidden=kind==='mirror';hint.classList.toggle('warn',kind==='ceilingColor'&&item?.code==='DL6');}
 }
}
function renderDesignInline(box,kind){
 if(!guideEntries(kind).length)return;
 const holder=document.createElement('div');holder.className='design-inline';holder.dataset.designKind=kind;
 const img=document.createElement('img');img.className='design-thumb';img.alt='Ảnh mẫu';
 const note=document.createElement('span');note.className='design-hint';
 const button=document.createElement('button');button.className='btn design-inline-btn';button.type='button';button.textContent='Xem hình & mô tả';
 button.onclick=()=>openDesignGuide(kind);
 // v5.8.149: the thumbnail itself opens the catalogue; keep one compact icon next to the select.
 button.textContent='';button.classList.add('design-icon-button');
 button.setAttribute('aria-label','Xem hình và mô tả / View design catalogue');
 button.title='Xem hình và mô tả';button.append(img);
 holder.append(button,note);box.append(holder);
 const item=selectedGuideItem(kind);if(item?.images?.[0]){img.src=item.images[0];img.alt=item.title}else img.hidden=true;
 note.textContent=kind==='mirror'?'':designSummaryFor(kind);note.hidden=kind==='mirror';
}
function guideSelectableValue(kind,code){
 if(kind==='ceilingColor')return '';
 const cell=kind==='ceiling'?($('model').value==='MR'?'D45':'D41'):$('model').value==='MR'?(kind==='handrail'?'D67':'D69'):(kind==='handrail'?'D64':'D66');
 const input=activeInput(cell);if(!input||input.tagName!=='SELECT')return '';
 const values=[...input.options].filter(op=>!op.dataset.invalid).map(op=>op.value);
 if(kind==='ceiling'){
  if($('model').value==='MR')return values.find(v=>String(v).toUpperCase()===code.toUpperCase())||'';
  const raw=String(currentValue('D41'));
  const currentLight=raw.endsWith('-W')?'-W':'-C';
  return [code+currentLight,code+'-C',code+'-W',code].find(v=>values.includes(v))||'';
 }
 return values.find(v=>v.toUpperCase().startsWith(code.toUpperCase()))||'';
}
function openDesignGuide(kind){
 const modal=$('design-popup'),items=guideEntries(kind);if(!modal||!items.length)return;
 const label={ceiling:'Trần phòng thang / Car ceiling',ceilingColor:'Màu sắc & vật liệu trần / Ceiling finishes',handrail:'Tay vịn / Handrail',mirror:'Gương / Mirror'};
 $('design-title').textContent=label[kind]+' · '+($('model').value==='MR'?'NEXIEZ-MR':'NEXIEZ-MRL Version2');
 $('design-subtitle').textContent=kind==='mirror'?'':kind==='ceilingColor'?'Vật liệu theo hình catalogue. Các màu bổ sung chỉ áp dụng khi có dữ liệu SpecTool tương ứng.':'Chọn mẫu bằng ảnh. Các mẫu chưa có trong Excel source được hiển thị để tham khảo.';
 const generalFootnote=modal.querySelector('.design-footnote');if(generalFootnote)generalFootnote.hidden=kind==='mirror';
 const root=$('design-cards');root.replaceChildren();
 const selected=selectedGuideItem(kind)?.code;
 for(const item of items){
  const card=document.createElement('article');card.className='design-card';if(item.code===selected)card.classList.add('is-current');
  const images=document.createElement('div');images.className='design-card-images'+(item.images.length>1?' double':'');
  for(let i=0;i<item.images.length;i++){
   const pic=document.createElement('img');pic.loading='lazy';pic.src=item.images[i];pic.alt=item.title+(item.images.length>1?(i===0?' — cool white':' — warm white'):'');images.append(pic);
  }
  const body=document.createElement('div');body.className='design-card-text';
  const name=document.createElement('div');name.className='design-card-name';const strong=document.createElement('strong');strong.textContent=item.title;name.append(strong);
  if(item.code===selected){const tag=document.createElement('span');tag.className='design-current';tag.textContent='Đang chọn';name.append(tag)}
  const sub=document.createElement('div');sub.className='design-card-sub';sub.textContent=item.subtitle;
  const desc=document.createElement('div');desc.className='design-card-desc';desc.textContent=item.desc;
  body.append(name,sub,desc);
  if(item.warning){const warning=document.createElement('div');warning.className='design-card-warning';warning.textContent=item.warning;body.append(warning)}
  const foot=document.createElement('div');foot.className='design-card-footer';
  const allowed=guideSelectableValue(kind,item.code);
  if(allowed){const button=document.createElement('button');button.type='button';button.className='btn btn-small btn-secondary design-apply';button.textContent='Áp dụng vào Spec';
   button.onclick=()=>{
    const cell=kind==='ceiling'?($('model').value==='MR'?'D45':'D41'):$('model').value==='MR'?(kind==='handrail'?'D67':'D69'):(kind==='handrail'?'D64':'D66');
    const input=activeInput(cell);if(!input)return;input.value=allowed;input.dispatchEvent(new Event('change',{bubbles:true}));closeDesignGuide();
   };foot.append(button);
  }else if(kind==='mirror'){
   // No informational note for mirror selection; the technical validation is authoritative.
  }else if(kind!=='ceilingColor'){
   const note=document.createElement('span');note.className='design-unavailable';note.textContent='Mẫu tham khảo · chưa có trong dropdown Excel hiện tại';foot.append(note);
  }else{
   const note=document.createElement('span');note.className='design-unavailable';note.textContent='Vật liệu tham khảo từ hình catalogue';foot.append(note);
  }
  card.append(images,body,foot);root.append(card);
 }
 modal.classList.add('show');modal.setAttribute('aria-hidden','false');$('design-close')?.focus();
}
function closeDesignGuide(){const modal=$('design-popup');if(!modal)return;modal.classList.remove('show');modal.setAttribute('aria-hidden','true')}
$('design-close')?.addEventListener('click',closeDesignGuide);
$('design-popup')?.addEventListener('click',event=>{if(event.target===$('design-popup'))closeDesignGuide()});
document.addEventListener('keydown',event=>{if(event.key==='Escape'&&$('design-popup')?.classList.contains('show'))closeDesignGuide()});

// v5.8.21 – User-supplied display/indicator illustrations. No changes to source values.
function panelIllustrationKey(kind,value){
 const s=String(value??'').trim().toUpperCase();
 if(!s||/^(NO|NONE|NOT SUPPLIED|N\/A|STANDARD|0)$/.test(s))return '';
 if(kind==='COP'){
  const m=s.match(/(?:^|[-\s])([CMPN]\d{3,4}(?:ETL|E)?)\b/);
  return m?m[1]:s;
 }
 if(kind==='CBV'){
  const m=s.match(/CBV\d+/);
  return m?m[0]:s;
 }
 if(kind==='PIV'){
  const m=s.match(/PIV\d*[- ]?([A-Z]\d{3,4})/);
  return m?'PIV*-'+m[1]+'*':'';
 }
 if(kind==='PIH_PID')return s;
 return '';
}
function renderPanelImagePreview(fieldBox,kind,value){
 const source=window.SPEC_PANEL_IMAGES;
 if(!source)return;
 let panel=fieldBox.querySelector('.panel-image-preview');
 const key=panelIllustrationKey(kind,value);
 // MRL Version02: N-series COP panel models share the corresponding M-series
 // illustration when no N-specific artwork is available. Display only.
 const catalog=source.catalog?.[kind]||{};
 let filenames=catalog[key]||[];
 if(kind==='COP'&&$('model')?.value==='MRL'&&/^N\d{3,4}(?:ETL|E)?$/.test(key)&&!filenames.length){
  const equivalent='M'+key.slice(1);
  filenames=catalog[equivalent]||(!/E(?:TL)?$/.test(equivalent)?catalog[equivalent+'E']:null)||[];
 }
 if(!filenames.length){
  panel?.remove();fieldBox.classList.remove('panel-preview-inline');
  return;
 }
 fieldBox.classList.add('panel-preview-inline');
 if(!panel){panel=document.createElement('div');panel.className='panel-image-preview';fieldBox.append(panel)}
 if(panel.dataset.key===key&&panel.dataset.kind===kind)return;
 panel.dataset.key=key;panel.dataset.kind=kind;
 panel.replaceChildren();
 const gallery=document.createElement('div');gallery.className='panel-preview-gallery';
 for(const filename of filenames){
  const url=source.media?.[filename];if(!url)continue;
  const img=document.createElement('img');img.src=url;img.alt=key+' – '+filename;
  img.title=kind==='CBV'?'Nút bấm cabin / Car operating button':'Ảnh chỉ thể hiện loại màn hình hiển thị';
  img.loading='lazy';img.decoding='async';gallery.append(img);
 }
 if(!gallery.children.length){panel.remove();fieldBox.classList.remove('panel-preview-inline');return}
 panel.append(gallery);
 if(kind!=='CBV'){
  const note=document.createElement('small');note.className='panel-preview-note';
  note.textContent='Ảnh chỉ thể hiện loại màn hình hiển thị';
  panel.append(note);
 }
}
function refreshPanelImagePreviews(){
 if(!window.SPEC_PANEL_IMAGES)return;
 for(const input of document.querySelectorAll('#fields [data-cell]')){
  const box=input.closest('.field');if(!box)continue;
  const field=current.fields.find(x=>x.cell===input.dataset.cell);if(!field)continue;
  const name=String(field.name||'');
  // No images for HBV or Hall Lantern – those are handled separately.
  let kind='';
  if(/COPButton$/i.test(name))kind='CBV';
  else if(/COPModel$/i.test(name))kind='COP';
  else if(/PIV/i.test(name)||(/K36_/i.test(name)&&/Hall.*(?:Indicator|Button)/i.test(name)))kind='PIV';
  else if(/PIH_PID|PIH|PID/i.test(name))kind='PIH_PID';
  if(!kind)continue;
  renderPanelImagePreview(box,kind,input.value);
 }
}
function updateFinishPreviews(){
 const rules=window.SPEC_MATERIAL_5825,model=$('model').value;
 if(!rules)return;
 for(const field of current.fields){
  const el=activeInput(field.cell);if(!el)continue;
  const box=el.closest('.field');if(!box)continue;
  const v=rules.preview(model,field,String(el.value||''));
  let panel=box.querySelector('.finish-preview-box');
  if(!v || !v.src){panel?.remove();box.classList.remove('finish-preview-inline');continue;}
  if(!panel){
   panel=document.createElement('div');panel.className='finish-preview-box';
   const img=document.createElement('img');img.className='finish-preview-thumb';
   const caption=document.createElement('small');caption.className='finish-preview-caption';
   panel.append(img,caption);box.append(panel);
  }
  box.classList.add('finish-preview-inline');
  panel.querySelector('img').src=v.src;
  panel.querySelector('img').alt=String(el.value);
  panel.querySelector('small').textContent=v.caption+(v.approx?' · Reference color':'');
 }
}

// v5.8.147: default landing-door finish follows the car-door finish.
// This is a one-way convenience default, never a lock or a validation rule.
function syncLandingFinishFromCar(){
 const car=current?.fields?.find(f=>f.name==='K24_CarDoorFinish');
 if(!car)return;
 const v=String(currentValue(car.cell)||'').trim();if(!v)return;
 const family=/^Painted\b/i.test(v)?'Painted':/^SUS-HL\b/i.test(v)?'SUS-HL':'';
 if(!family)return;
 const windowSize=/W300xH1300|300x1300/i.test(v)?'300':/W200xH1300|200x1300/i.test(v)?'200':'';
 const glass=/DSJ1/i.test(v)?'DSJ1':/\bDS1\b/i.test(v)?'DS1':/Large\s*window/i.test(v)?'Large':'';
 for(const key of ['K36_DoorFinishMainFloorMaterial','K36_DoorFinishOtherFloorMaterial']){
  const f=current.fields.find(x=>x.name===key);if(!f)continue;
  const values=workbookChoices(f).map(String);
  let candidates=values.filter(x=>x.startsWith(family));
  if(glass==='DSJ1')candidates=candidates.filter(x=>/DSJ1/i.test(x));
  else if(glass==='DS1')candidates=candidates.filter(x=>/\bDS1\b/i.test(x)&&!/DSJ1/i.test(x));
  else if(glass==='Large'){
   candidates=candidates.filter(x=>/Large\s*window/i.test(x));
   if(windowSize)candidates=candidates.filter(x=>new RegExp(windowSize+'x1300|'+windowSize+'xH1300','i').test(x));
  }else candidates=candidates.filter(x=>!/DSJ1|\bDS1\b|Large\s*window/i.test(x));
  if(candidates.length && String(currentValue(f.cell))!==candidates[0])applyValue(f.cell,candidates[0]);
 }
}

function row(field,displayName='',compact=false){
 const box=document.createElement('div');box.className='field';
 box.dataset.sourceCell=field.cell;
 // Source identifiers remain as data-cell attributes for Check Spec, focus and Export.
 const name=document.createElement('label');name.className='nm';name.textContent=window.SPEC_PROOFREAD?.(displayName||xrules().labels[field.cell]||field.label||field.name.replace(/^K\d+_/,'').replaceAll('_',' '))??(displayName||xrules().labels[field.cell]||field.label||field.name.replace(/^K\d+_/,'').replaceAll('_',' '));
 const choices=workbookChoices(field);const input=hasDropdown(field)?document.createElement('select'):document.createElement('input');input.dataset.cell=field.cell;input.id='spec-'+field.cell;
 name.htmlFor=input.id;
 let value=String(field.value??'');
 if(input.tagName==='SELECT'){
  const options=[...choices.map(String)];if(value&&!choices.includes(value)&&!field.validation&&!supplementalListSources[field.name])options.unshift(value);
  if(!options.length)options.push('');
  for(const option of [...new Set(options)]){const opt=document.createElement('option');opt.value=option;opt.textContent=window.SPEC_PROOFREAD?.(option||'— Chọn / Select —')??(option||'— Chọn / Select —');input.append(opt)}
 }else input.placeholder='';
 input.value=value;
 input.addEventListener('change',()=>{
   changes[field.cell]=input.value;input.classList.add('dirty');
   if(['MR','MRL'].includes($('model').value))refreshTwoStopTravelWarning();
   if(field.name==='K5_Capacity')capacityUpdate(input.value);
   if($('model').value==='MRL'&&['D13','D16'].includes(field.cell)){
    const opts=allowedObservationCodes();
    if(!opts.includes(String(currentValue('D14')))&&String(currentValue('D14'))!=='No')applyValue('D14','No');
   }
   // SETS-L status uses a dynamic source formula (DataTable!Q472 / K472).
   // Rebuild the feature controls when rated speed changes; prior UI kept
   // a stale disabled X state from the previous 60 m/min selection.
   if($('model').value==='MRL'&&field.name==='K8_Speed')renderFeatures();
   if(field.name==='K28_FlooringWeight')flooringManual=true;
   else flooringDependencyChanged(field.cell);
   if(field.name==='K5_Capacity' || field.name==='K17_CarW')refreshMirrorWeightHint();
   if($('model').value==='MR'&&['D13','D34','D35'].includes(field.cell))refreshMRBedShaftNotes();
   // A new ceiling model immediately changes its finish choices, instead of
   // falsely flagging the old paint color when switching to SUS-HL.
   if(field.name==='K22_CeilingModel'&&window.SPEC_CEILING_FINISH){
    const model=$('model').value,fin=current.fields.find(x=>x.name==='K22_CeilingColor');
    if(fin){const chosen=window.SPEC_CEILING_FINISH.next(model,input.value,currentValue(fin.cell));
     if(String(currentValue(fin.cell))!==chosen)applyValue(fin.cell,chosen);}
   }
   // Material-only sync: changing Car wall also sets Transom material, if that
   // value is available under the current factory / observation restrictions.
   // Paint color and PPS pattern remain independent selections.
   if(field.name==='K23_WallFinish'){
    const transom=current.fields.find(x=>x.name==='K25_CarTransomFinish');
    if(transom){
     const allowed=workbookChoices(transom);
     if(allowed.includes(String(input.value).trim())&&String(currentValue(transom.cell))!==String(input.value).trim())
      applyValue(transom.cell,String(input.value).trim());
    }
   }
   if($('model').value==='MRL'&&field.cell==='D14'&&window.SPEC_NY40_RULES){
     const allowed=window.SPEC_NY40_RULES.allowedLocations(input.value);
     if(allowed&&!allowed.includes(String(currentValue('D65'))))applyValue('D65',allowed[0]);
   }
   syncDropdownOptions();
   if(field.name==='K24_CarDoorFinish')syncLandingFinishFromCar();
   if(/^K36_(?:PIV|HBV|PIH_PID)(?:TopFloor|MainFloor|OtherFloor)$/.test(field.name)){
    hallManualSelections.add(field.cell);hallAutoSelections.delete(field.cell);
    autoFillHallSiblingFields(field);
   }
   if(field.name==='K0_Qty')refreshHallQuantities();
   refreshNY40CopChoices();
   refreshSideDrop();
   refreshLocks();
   if(field.cell==='D14'){refreshCarDesignStates();renderFeatures();}
   if(field.name==='K24_CarDoorFinish')renderFeatures();
   refreshJambDepthFields();
   refreshDesignHints();
   updateFinishPreviews();
   refreshMirrorWeightHint();
   if($('model').value==='MRL'){
    refreshMRLOverheadMinimum(['K5_Capacity','K8_Speed','K21_CarTR','K16_Regulation'].includes(field.name));
    if(['D13','D17','D31','D16','D84','D23'].includes(field.cell))refreshMRLShaftMinimum();
   }
   if(['K13_Operation','K5_Capacity','K8_Speed'].includes(field.name))renderFeatures();
   if(field.name==='K13_Operation'){refreshGroupControlPreview();refreshHallQuantities();}
   if($('model').value==='MRL'&&field.cell==='D14'){syncObservationButton(input);if(/^NY-40/i.test(String(input.value||'')))openNy40Popup(input.value);}
   if($('model').value==='MRL'&&['D36','D37'].includes(field.cell))refreshMRLOverheadMinimum(false);
   invalidateCheck();status('Đã thay đổi thông số. Cần chạy lại Check Spec.');
 });
 if($('model').value==='MRL'&&field.cell==='D14'){
  const wrap=document.createElement('div');wrap.className='obs-inline';
  const btn=document.createElement('button');btn.type='button';btn.className='btn btn-small obs-preview-btn';btn.textContent='Xem mô tả';
  btn.onclick=()=>openNy40Popup(input.value);
  wrap.append(input,btn);box.append(name,wrap);syncObservationButton(input);
 }else box.append(name,input);
 if($('model').value==='MR'&&['D34','D35'].includes(field.cell)){
  const note=document.createElement('small');note.className='mr-bed-shaft-note';note.hidden=true;
  note.id='mr-bed-shaft-note-'+field.cell;
  note.setAttribute('role','alert');note.setAttribute('aria-live','polite');
  input.setAttribute('aria-describedby',note.id);
  box.append(note);
  input.addEventListener('input',refreshMRBedShaftNotes);
 }
 if(($('model').value==='MR'&&field.cell==='D39')||($('model').value==='MRL'&&field.cell==='D35')){
  const hint=document.createElement('small');
  hint.className='travel-limit-hint';hint.hidden=true;
  hint.id='travel-limit-warning-'+$('model').value;
  hint.setAttribute('role','alert');hint.setAttribute('aria-live','polite');
  input.setAttribute('aria-describedby',hint.id);
  box.append(hint);
  // Travel is a free-text input: warn while typing, not only when leaving the cell.
  input.addEventListener('input',refreshTwoStopTravelWarning);
 }
 if(field.name==='K28_FlooringWeight'){
  const btn=document.createElement('button');btn.type='button';btn.className='btn btn-small';
  btn.textContent='Tính lại / Recalculate';
  btn.title='Tính lại từ AA × BB, JJ và loại sàn theo công thức Excel';
  btn.addEventListener('click',()=>{updateFlooringWeight(true);invalidateCheck();status('Đã tính lại khối lượng sàn đá theo công thức nguồn / Flooring weight recalculated.');});
  box.append(btn);
 }
 if($('model').value==='MRL'&&['K22_SUSType','K24_CarDoorFinish','K36_DoorFinishMainFloorMaterial','K36_DoorFinishOtherFloorMaterial'].includes(field.name)){
   const hint=document.createElement('small');hint.className='technical-material-hint';
   hint.textContent='SUS Standard vẫn hợp lệ với NY-40. Chỉ cấu kiện cửa có kính DSJ1/DS1/Large Window dùng SUS304; khung cửa, transom và front return giữ vật liệu tiêu chuẩn.';
   hint.style.cssText='display:block;margin-top:5px;color:#52677b;font-size:11px;';box.append(hint);
 }
 if($('model').value==='MRL'&&field.name==='K29_DecorWeight'){
  box.classList.add('mrl-decoration-field');
  const hint=document.createElement('span');hint.className='mrl-decoration-maximum unavailable';
  hint.id='mrl-decor-maximum';hint.textContent='Maximum: —';box.append(hint);
 }
 if(field.name==='K31_BrailleLetter'){
  box.classList.add('braille-preview-inline');
  const image=document.createElement('img');image.className='braille-preview-thumb';
  image.alt='Minh họa nút bấm có ký tự Braille không vít';image.title='Braille without screw';
  image.src='data:image/jpeg;base64,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';image.loading='lazy';
  box.append(image);
 }
 const guideKind=designKindFor(field);if(guideKind)renderDesignInline(box,guideKind);
 if($('model').value==='MRL' && field.name==='K35_Ventilation'){
  const hint=document.createElement('small');hint.className='fan-guidance';
  hint.textContent='DC axial flow fan: Quạt hướng trục DC; DC axial flow fan + Plasma Quad: Quạt hướng trục DC kết hợp với diệt khuẩn Plasma.';
  box.append(hint);
 }
 return box;
}
// BED750/BED1000 only: offer a structural narrowing note alongside AH/BH
// when an entered dimension exceeds the published maximum. This is an
// engineering suggestion, NOT approval of an out-of-range factory configuration.
function refreshMRBedShaftNotes(){
 const bed=$('model').value==='MR'&&window.SPEC_MR_BED_5871?.cases?.find(
   x=>String(currentValue('D13')).startsWith(x.code+'-2S'));
 for(const [cell,key,maxKey,label] of [['D34','ah','ahmax','AH'],['D35','bh','bhmax','BH']]){
  const input=activeInput(cell),note=document.getElementById('mr-bed-shaft-note-'+cell);
  if(!note||!input)continue;
  const val=Number(input.value),max=bed?.[maxKey];
  const excessive=!!bed&&input.value.trim()!==''&&Number.isFinite(val)&&val>max;
  note.hidden=!excessive;
  note.textContent=excessive?`${label} vượt Maximum ${max.toLocaleString('vi-VN')} mm. Vui lòng bổ sung dầm thu hố (dầm thép hình, dầm bê tông...) để giảm kích thước hữu dụng về giới hạn cho phép; cần xác nhận phương án kết cấu và kích thước với bộ phận kỹ thuật.`:'';
  input.classList.toggle('mr-bed-over-max',excessive);
  input.setAttribute('aria-invalid',excessive?'true':'false');
 }
}
// Additional MRL shaft minima from the user's 156-row, four-CWT V02 source.
// These are displayed without overwriting a manually entered shaft size.
function refreshMRLShaftMinimum(){
 if($('model').value!=='MRL'||!window.SPEC_CWT_ENGINE||!window.SpecExcelSourceEngine)return;
 const db=workbookEngineData();if(!db)return;
 const calc=new window.SpecExcelSourceEngine.SourceEngine(db,changes);
 const res=window.SPEC_CWT_ENGINE.lookupEngine(calc);
 for(const [key,cell] of [['AH','D33'],['BH','D34']]){
  const input=activeInput(cell);if(!input)continue;
  const field=input.closest('.field');if(!field)continue;
  let hint=field.querySelector('.cwt-minimum-hint');
  if(!hint){hint=document.createElement('small');hint.className='cwt-minimum-hint';field.append(hint);}
  if(res.ok){
   hint.textContent=key+' tối thiểu: '+Number(res[key]).toLocaleString('vi-VN')+' mm · '+res.cwt;
   const val=Number(currentValue(cell));const wrong=!Number.isFinite(val)||val<Number(res[key]);
   hint.classList.toggle('cwt-minimum-error',wrong);input.classList.toggle('cwt-below-minimum',wrong);
  }else{
   hint.textContent=res.reason==='cwt'?'Đối trọng chưa có dữ liệu AH/BH cho tổ hợp này.':'CHECK LAYOUT: chưa có tổ hợp AH/BH phù hợp.';
   hint.classList.add('cwt-minimum-error');input.classList.remove('cwt-below-minimum');
  }
 }
}
// MRL OH/PD source-of-truth: normalized DataTable!DA5:DL164.
// Minima come from the same SpecForm F36/F37 formulas as exported workbook.
// Auto-raise on source configuration change only; never overwrite a larger manual value.
function refreshMRLOverheadMinimum(adjustIfBelow=false){
 if($('model').value!=='MRL')return;
 const db=workbookEngineData();const Engine=window.SpecExcelSourceEngine?.SourceEngine;
 if(!db||!Engine)return;
 const calc=new Engine(db,changes);
 const setsAlternative=window.SPEC_MRL_SETSL_SUGGESTION?.compare(db,changes);
 for(const [cell,source,label] of [['D36','F36','PD'],['D37','F37','OH']]){
  const el=activeInput(cell);if(!el)continue;
  const r=calc.safe('SpecForm',source),numeric=r.ok?Number(r.value):NaN;
  const minimum=typeof r.value==='number'&&Number.isFinite(numeric)&&numeric>0?numeric:null;
  let hint=el.closest('.field')?.querySelector('.oh-minimum-hint');
  if(!hint){hint=document.createElement('div');hint.className='oh-minimum-hint';hint.style.cssText='font-size:11px;color:#526a82;line-height:1.4;margin-top:2px';el.closest('.field')?.append(hint);}
  if(minimum===null){hint.textContent=`${label}: CHECK LAYOUT — không có giá trị cho tổ hợp Capacity / Speed / TR hiện tại.`;hint.style.color='#a32832';continue;}
  const actual=Number(currentValue(cell));
  const advice=window.SPEC_MRL_SETSL_SUGGESTION?.advise(label,actual,setsAlternative)||'';
  hint.textContent=`${label} tối thiểu theo DataTable: ${minimum} mm (có thể nhập lớn hơn)`+(advice?' · '+advice:'');
  hint.style.color=advice?'#8b521e':'#526a82';
  if(adjustIfBelow){const actual=Number(currentValue(cell));if(!Number.isFinite(actual)||actual<minimum){el.value=String(minimum);changes[cell]=String(minimum);el.classList.add('dirty');}}
 }
}
// v5.8.75: factory-requested defaults for NEXIEZ-MR hospital-bed models.
const BED_DEFAULT_FEATURES=new Set(['DKO-TB','HE-B','EXCL','EXCLC']);
function isMRBedCapacity(value=String(currentValue('D13'))){
 return $('model').value==='MR'&&/^B(?:750|1000)-2S(?:\s|\(|$)/i.test(String(value));
}
function applyMRBedDefaults(capacity){
 if(!isMRBedCapacity(capacity))return;
 // These are defaults on switching capacity; subsequent user edits are respected.
 for(const [cell,value] of [['D64','Side'],['D65','710'],['D67','YH-56S (SUS304)'],['D68','3 pcs on 3 walls']])applyValue(cell,value);
 const features=sourceFeatureRows();
 for(const feat of features){
  const code=window.SPEC_V52_FEATURES?.canonical(feat.code)||feat.code;
  if(BED_DEFAULT_FEATURES.has(code)){
   changes[feat.cell]='1';featureTouched.delete(feat.cell);
  }
 }
}
// v5.8.95: Sales consultation defaults for NEXIEZ-MR only.
// Apply once on MR load. Subsequent user selections and FER toggles remain editable.
// Store overrides in `changes` so source Check Spec and Excel/PDF Export agree with the UI.
function setMRDefaultGroupAndFER(){
 if($('model').value!=='MR')return;
 const groupField=current.fields.find(f=>f.name==='K13_Operation');
 if(groupField)applyValue(groupField.cell,'1C-2BC');
 const fer=sourceFeatureRows().find(f=>(window.SPEC_V52_FEATURES?.canonical(f.code)||f.code)==='FER');
 if(fer){
  const option=window.SPEC_V52_FEATURES?.status('MR','FER','1C-2BC',workbookEngineData());
  if(String(option||fer.option).trim().toUpperCase()!=='X'){
   changes[fer.cell]='1';
   featureTouched.delete(fer.cell);
  }
 }
}
function setMRLHandrailDefault(){
 if($('model').value!=='MRL')return;
 const field=current.fields.find(x=>x.name==='K32_Handrail');
 if(!field)return;
 // Set only Handrail Type (D64): D65 quantity/location remains untouched.
 // The source SpecForm may initially show "1 pc on rear wall"; that is its own
 // source value, not a quantity default tied to YH-100SL.
 applyValue(field.cell,'YH-100SL (SUS304)');
}
function capacityUpdate(capacity){
 let row=current.capacity.find(c=>String(c.EleCapacityName)===capacity);if(!row)return;
 // Mirrors mapped parameters from the base prototype, not yet a complete C# rule port.
 const fireField=current.fields.find(f=>f.name==='K36_FireproofDoor');
 const fireValue=String(fireField?currentValue(fireField.cell):'Not supplied');
 const fireProvided=fireValue!=='Not supplied'&&fireValue!=='GHCL E30'&&fireValue!=='';
 const map={K7_DoorType:'DoorType',K17_CarW:'AA',K17_CarD:'BB',K18_CarJJ:'JJ',K19_CarAH:fireProvided?'AH-CWT1-FIRE':'AH-CWT1',K19_CarBH:fireProvided?'BH-CWT1-FIRE':'BH-CWT1',K20_CarAM:'AM',K20_CarBM:'BM',K21_CarOH:'OH',K21_CarPD:'PD'};
 for(let [fieldName,tableCol] of Object.entries(map)){
  const f=current.fields.find(f=>f.name===fieldName);if(!f||row[tableCol]===''||row[tableCol]===undefined)continue;
  const el=document.querySelector('[data-cell="'+f.cell+'"]');if(!el)continue;
  const v=String(row[tableCol]);if(el.tagName==='SELECT'&&![...el.options].some(o=>o.value===v)){let op=document.createElement('option');op.value=v;op.textContent=v;el.append(op)}
  el.value=v;changes[f.cell]=v;el.classList.add('dirty');
 }
  // v5.8.157: capacity and counterweight are one source-backed configuration.
  if($('model').value==='MR'){
    const cwtField=current.fields.find(f=>f.name==='K14_CWTLocation');
    const pos=String(row['CWT-Position']||'');
    if(cwtField && ['Rear','Side'].includes(pos)){
      const opts=workbookChoices(cwtField).map(String);
      const v=opts.find(x=>pos==='Rear'?/Rear drop|Phía sau/i.test(x):/Side drop|Bên hông/i.test(x));
      if(v)applyValue(cwtField.cell,v);
    }
  }
 applyMRBedDefaults(capacity);
}
function allowedObservationCodes(){
 if($('model').value!=='MRL')return [];
 const cap=String(currentValue('D13')),entrance=String(currentValue('D16'));
 // Table 4.3 Version02: 1D-2G is exclusively NY-40C on P11-CO/P13W-CO.
 // Do not hide NY-40C before the salesman switches COP to Front: that
 // creates a circular selection blocker. R13/R14 remain checked by Check Spec.
 if(entrance==='1D-2G')return /^(?:P11-CO|P13W-CO)(?:\s|\(|$)/i.test(cap)?['NY-40C']:[];
 const db=workbookEngineData();
 if(db&&window.SpecExcelSourceEngine?.SourceEngine){
  const e=new window.SpecExcelSourceEngine.SourceEngine(db,changes);
  const gate=e.safe('SpecForm','R13');
  if(gate.ok&&gate.value===false)return [];
 }
 if(entrance==='1D-1G')return Object.keys(NY40_POPUP_DATA);
 return [];
}
function ny40PositionChoices(){
 if($('model').value!=='MRL')return null;
 return window.SPEC_NY40_RULES?.allowedLocations(String(currentValue('D14')))||null;
}
function technicalOptions(field){
 if($('model').value!=='MRL')return field.options||[];
 const cap=String(currentValue('D13')),entrance=String(currentValue('D16'));
 if(field.cell==='D16')return /^(P08|P11|P13|P17D|P21D)/.test(cap)?['1D-1G','1D-2G']:['1D-1G'];
 if(field.cell==='D14')return ['No',...allowedObservationCodes()];
 if(field.cell==='D65'&&ny40PositionChoices())return ny40PositionChoices();
 return field.options||[];
}
function updateConditional(){
 for(let cell of ['D16','D14','D65']){
  const f=current.fields.find(x=>x.cell===cell),el=activeInput(cell);if(!f||!el)continue;
  const opts=technicalOptions(f);if(el.tagName!=='SELECT')continue;
  const old=String(currentValue(cell));el.replaceChildren();for(let value of opts){let op=document.createElement('option');op.value=value;op.textContent=value;el.append(op)}
  const next=opts.includes(old)?old:opts[0]||'';el.value=next;if(next!==old)changes[cell]=next;
 }
}
// Presentation-only labels. Technical source values, sheet positions and Excel logic remain unchanged.
const INPUT_DETAIL_LABELS={
  Material:'Vật liệu / Material',Paint:'Màu sơn / Paint color',
  COPModel:'Mẫu hộp COP / COP model',COPButton:'Nút bấm cabin / Car buttons',
  JambMainFloor:'Mẫu khung cửa / Jamb model',JambOtherFloor:'Mẫu khung cửa / Jamb model'
};
function subLabel(raw){
 const s=String(raw||'').trim();if(!s)return '';
 const replacements={
   'Color':'Màu sắc / Color','Model':'Mẫu / Model',
   'Finish / Material':'Vật liệu / Finish',
   'Color (Only Painted)':'Màu sơn / Paint color',
   'Pattern (Only PPS)':'Hoa văn / Pattern',
   'Location':'Vị trí / Location','Code':'Mã / Code',
   'Type':'Loại / Type','Car button':'Nút bấm / Car button',
   'Weight (Only Local granite)':'Khối lượng đá địa phương / Local granite weight',
   'Số lượng / Vị trí':'Số lượng / Vị trí (Quantity / Location)',
   'AA (Rộng-W)':'AA – Rộng / Width', 'BB (Sâu-D)':'BB – Sâu / Depth',
   'HL (Cao-H)':'HL – Cao / Height', 'JJ (Rộng-W)':'JJ – Rộng / Width',
   'HH (Cao-H)':'HH – Cao / Height', 'AH (Rộng-W)':'AH – Rộng / Width',
   'BH (Sâu-D)':'BH – Sâu / Depth', 'AM (Rộng-W)':'AM – Rộng / Width',
   'BM (Sâu-D)':'BM – Sâu / Depth', 'MH (Cao-H)':'MH – Cao / Height'
 };
 return replacements[s]||s;
}
function fieldRole(field){
 const n=field.name||'';
 for(const [suffix,label] of Object.entries(INPUT_DETAIL_LABELS))if(n.endsWith(suffix))return label;
 if(/Material$|Finish$/.test(n))return 'Vật liệu / Material';
 if(/Painted(Color)?$/.test(n))return 'Màu sơn / Paint color';
 if(/Pattern$/.test(n))return 'Hoa văn / Pattern';
 if(/COPButton$/.test(n))return 'Nút bấm / Car buttons';
 return xrules().labels[field.cell]||field.label||n.replace(/^K\d+_/,'').replaceAll('_',' ');
}
function headingText(label){return String(label||'').replace(/\s+/g,' ').trim().replace(/\s*:\s*$/,'');}
function refreshMirrorWeightHint(){ /* Hidden: mirror mass corrections run only in DataTable formulas. */ }
// Visual Hall Lantern catalog: only names and extracted manual images.
function openHallLanternGallery(fieldKey,select){
 const source=window.SPEC_HALL_LANTERN;if(!source||$('model').value!=='MRL')return;
 document.querySelector('.hall-lantern-overlay')?.remove();
 const overlay=document.createElement('div');overlay.className='hall-lantern-overlay';overlay.setAttribute('role','dialog');overlay.setAttribute('aria-modal','true');overlay.setAttribute('aria-label','Hall Lanterns gallery');
 const modal=document.createElement('div');modal.className='hall-lantern-modal';
 const bar=document.createElement('div');bar.className='hall-lantern-header';
 const title=document.createElement('h3');title.textContent='Đèn báo tầng / Hall Lanterns — NEXIEZ-MRL Version2';
 const close=document.createElement('button');close.className='btn';close.type='button';close.textContent='Đóng / Close';close.onclick=()=>overlay.remove();bar.append(title,close);
 const grid=document.createElement('div');grid.className='hall-lantern-gallery';
 for(const rec of source.catalog){
  const button=document.createElement('button');button.type='button';button.className='hall-lantern-card';
  if(select.value===rec.code)button.classList.add('is-current');
  const img=document.createElement('img');img.src=rec.image;img.alt=rec.code;img.loading='lazy';
  const name=document.createElement('strong');name.textContent=rec.code;button.append(img,name);
  button.onclick=()=>{select.value=rec.code;select.dispatchEvent(new Event('change',{bubbles:true}));overlay.remove()};grid.append(button);
 }
 modal.append(bar,grid);overlay.append(modal);overlay.addEventListener('click',e=>{if(e.target===overlay)overlay.remove()});
 overlay.addEventListener('keydown',e=>{if(e.key==='Escape')overlay.remove()});document.body.append(overlay);close.focus();
}
// v5.8.26: one device per line, moving original SpecForm nodes (no cloned input).
function buildCarDesignLines(){
 const section=[...document.querySelectorAll('#fields .spec-section')].find(x=>x.dataset.section==='III');
 if(!section)return;
 const grid=section.querySelector('.section-fields');if(!grid)return;
 const model=$('model').value;
 const specs=[
  ['K22_SUSType','Inox sọc nhuyễn / Full SUS-HL','single'],
  ['K22_CeilingModel','Trần phòng thang / Car ceiling','K22_CeilingColor'],
  ['K23_WallFinish','Vách phòng thang / Car wall','K23_WallPaintedColor','K23_WallPattern'],
  ['K24_CarDoorFinish','Cửa phòng thang / Car door','K24_CarDoorPaintedColor','K24_CarDoorPattern'],
  ['K25_CarTransomFinish','Đố cửa trên / Transom panel','K25_CarTransomPainted','K25_CarTransomPattern'],
  ['K26_CarFrontFinish','Vách trước / Front return panel','K26_CarFrontPainted']
 ];
 const wrapper=document.createElement('div');wrapper.className='car-design-lines';wrapper.dataset.model=model;
 for(const [parentName,title,paintName,patternName] of specs){
  const parent=current.fields.find(f=>f.name===parentName);if(!parent)continue;
  const node=activeInput(parent.cell)?.closest('.field');if(!node)continue;
  const line=document.createElement('div');line.className='car-design-line';line.dataset.device=parentName;
  const label=document.createElement('div');label.className='car-design-label';label.textContent=title;
  const material=document.createElement('div');material.className='car-design-material';material.append(node);
  const variants=document.createElement('div');variants.className='car-design-variants';
  if(paintName!=='single')for(const [name,kind] of [[paintName,'paint'],[patternName,'pattern']]){
   if(!name)continue;
   const f=current.fields.find(x=>x.name===name);if(!f)continue;
   const fieldNode=activeInput(f.cell)?.closest('.field');if(!fieldNode)continue;
   fieldNode.dataset.finishKind=kind;
   const nm=fieldNode.querySelector('.nm');if(nm)nm.textContent=kind==='paint'?'Màu sơn / Paint color':'Hoa văn / Pattern';
   variants.append(fieldNode);
  }
  // v5.8.107: advisory only for MRL Version02 NY-40 observation cars.
  if(parentName==='K22_SUSType'){
   const note=document.createElement('small');
   note.className='ny40-sus304-note';
   note.id='ny40-sus304-advisory';
   note.setAttribute('role','note');
   note.hidden=true;
   note.innerHTML='Thang máy quan sát, khuyến nghị nên chọn <strong>SUS304</strong> do yếu tố ánh sáng, môi trường, tăng khả năng chống ăn mòn.';
   // Keep advisory in the material cell: variants is hidden for single-value Full SUS-HL rows.
   material.append(note);
  }
  line.append(label,material,variants);wrapper.append(line);
 }
 // Move just the car-design block, retaining all SpecForm cell IDs for Export & Check.
 grid.prepend(wrapper);grid.classList.add('car-design-section-fields');
 // Empty source wrappers may be deleted only after all input nodes have been moved.
 for(const oldRow of [...grid.querySelectorAll(':scope > .spec-row')]){
  if(!oldRow.querySelector('[data-cell]') && !oldRow.querySelector('.readonly-value'))oldRow.remove();
 }
 refreshCarDesignStates();
}
function refreshCarDesignStates(){
 const panel=document.querySelector('.car-design-lines');if(!panel)return;
 const model=$('model').value;
 // Do not enforce SUS304: the designer retains full control of the material selection.
 const sus304Note=panel.querySelector('#ny40-sus304-advisory');
 if(sus304Note)sus304Note.hidden=!(model==='MRL'&&/^NY-40(?:[A-E])?\b/i.test(String(currentValue('D14')||'').trim()));
 for(const line of panel.querySelectorAll('.car-design-line')){
  const parent=current.fields.find(f=>f.name===line.dataset.device);if(!parent)continue;
  const value=String(currentValue(parent.cell)||'');
  const usePaint=paintedMaterial(value),usePattern=patternedMaterial(value);
  const variants=line.querySelector('.car-design-variants');
  const hasVariant=!!variants?.querySelector('.field');
  if(variants)variants.classList.toggle('car-design-no-variant',!hasVariant||(!usePaint&&!usePattern && parent.name!=='K22_CeilingModel'));
  for(const box of variants?.querySelectorAll('.field')||[]){
   const input=box.querySelector('[data-cell]');if(!input)continue;
   const isCeiling=parent.name==='K22_CeilingModel';
   const active=isCeiling || box.dataset.finishKind===(usePaint?'paint':usePattern?'pattern':'');
   box.classList.toggle('car-finish-inactive',!active);
   if(!active)continue;
   // Conditional formatting from the old PPS sheet must not hide a valid
   // Painted/Laminated selection. Preserve workbook options and checks.
   input.disabled=false;delete input.dataset.cfLocked;
   box.classList.remove('cf-locked','locked');input.title='';
   if(input.tagName==='SELECT'){
    const choices=[...input.options].filter(o=>!o.dataset.invalid && o.value).map(o=>o.value);
    const currentVal=String(currentValue(input.dataset.cell)||'');
    if(choices.length && !choices.includes(currentVal)){
     const next=choices[0];input.value=next;changes[input.dataset.cell]=next;
    }
   }
  }
 }
}
// v5.8.92 — conditional JD (Jamb depth) for G11L1P, independent for main/other floors.
function mountJambDepthFields(){
 if(!['MR','MRL'].includes($('model').value)||!window.SPEC_JAMB_DEPTH)return;
 for(const config of window.SPEC_JAMB_DEPTH.floorsForModel($('model').value)){
  const jamb=activeInput(config.cell),host=jamb?.closest('.field');
  if(!host||host.querySelector('.jamb-jd-field'))continue;
  const block=document.createElement('div');block.className='jamb-jd-field';block.dataset.jdKey=config.key;
  const label=document.createElement('label');label.className='nm jamb-jd-label';label.textContent='JD (mm) – Chiều rộng khung bao bản rộng / Jamb depth';
  const select=document.createElement('select');select.id='spec-'+config.key;
  select.dataset.cell=config.key;select.className='jamb-jd-input';label.htmlFor=select.id;
  select.addEventListener('change',()=>{
   changes[config.key]=select.value;select.classList.add('dirty');
   invalidateCheck();status('Đã thay đổi chiều rộng khung bao (JD). Cần kiểm tra lại Check Spec.');
  });
  const help=document.createElement('small');help.className='jamb-jd-range';
  block.append(label,select,help);host.append(block);
 }
 refreshJambDepthFields();
}
function refreshJambDepthFields(){
 if(!['MR','MRL'].includes($('model').value)||!window.SPEC_JAMB_DEPTH)return;
 for(const config of window.SPEC_JAMB_DEPTH.floorsForModel($('model').value)){
  const block=document.querySelector('.jamb-jd-field[data-jd-key="'+config.key+'"]');
  if(!block)continue;
  const selectedType=window.SPEC_JAMB_DEPTH.jambType(currentValue(config.cell));
  const range=window.SPEC_JAMB_DEPTH.limits(selectedType);
  block.hidden=!range;
  if(!range)continue;
  const select=block.querySelector('select');
  const prev=String(changes[config.key]??'');
  // The customer-approved JD selector uses exactly 5-mm increments within each published range.
  if(select.dataset.min!==String(range.min)){
   select.replaceChildren();
   const blank=document.createElement('option');blank.value='';blank.textContent='— Chọn JD (mm) / Select —';select.append(blank);
   for(let value=range.min;value<=range.max;value+=range.step){
    const option=document.createElement('option');option.value=String(value);option.textContent=value+' mm';select.append(option);
   }
   select.dataset.min=String(range.min);
  }
  select.value=(/^\d+$/.test(prev)&&Number(prev)>=range.min&&Number(prev)<=range.max&&(Number(prev)-range.min)%range.step===0)?prev:'';
  block.querySelector('.jamb-jd-range').textContent=selectedType+': '+range.min+' ≤ JD ≤ '+range.max+' mm · Bước / Step '+range.step+' mm';
 }
}

// Preview is presentation-only and is anchored to the K13_Operation field for
// EACH model (MR=D21, MRL=D22). It disappears five seconds after a selection.
let groupControlPreviewTimer=null;
function clearGroupControlPreview(){
 if(groupControlPreviewTimer!==null){clearTimeout(groupControlPreviewTimer);groupControlPreviewTimer=null;}
 document.querySelectorAll('.group-control-preview').forEach(panel=>panel.remove());
 document.querySelectorAll('.group-preview-inline').forEach(host=>host.classList.remove('group-preview-inline'));
}
function refreshGroupControlPreview(){
 // Cancel the previous model/selection preview even when a new one is hidden.
 clearGroupControlPreview();
 const model=$('model')?.value;
 if(!['MR','MRL'].includes(model))return;
 const field=current?.fields?.find(f=>f.name==='K13_Operation');
 if(!field)return;
 const input=activeInput(field.cell),host=input?.closest('.field');
 if(!host)return;
 const value=String(currentValue(field.cell)||'').trim();
 if(!value||value==='1C-2BC')return;
 const panel=document.createElement('div');panel.className='group-control-preview';
 const img=document.createElement('img');img.className='group-control-preview-thumb';
 img.loading='eager';img.decoding='async';img.alt='At bottom floor — group control arrangement';
 img.src='assets/group-control-non-1c-2bc.png';
 const note=document.createElement('small');note.className='group-control-preview-note';
 note.textContent='Hình minh họa bố trí thang nhóm / Group control arrangement preview';
 panel.append(img,note);host.append(panel);host.classList.add('group-preview-inline');
 // Auto-dismiss without changing the selected group control or Check Spec state.
 groupControlPreviewTimer=setTimeout(()=>{
  panel.remove();host.classList.remove('group-preview-inline');
  groupControlPreviewTimer=null;
 },5000);
}
function refreshNY40CopChoices(){
 if($('model').value!=='MRL')return;
 const obs=String(currentValue('D14')||'').trim();
 const isNY40=/^NY-40[A-E]$/.test(obs);
 for(const [cell,disallow] of [
  ['D61',v=>/^[MP]\d/i.test(v)],
  ['D62',v=>v!==''&&!/^(?:NO|NONE|NOT SUPPLIED|KHÔNG CUNG CẤP)$/i.test(v)]
 ]){
   const input=activeInput(cell);if(!input||input.tagName!=='SELECT')continue;
   for(const option of input.options){
     option.disabled=isNY40&&disallow(String(option.value).trim());
   }
   let help=input.closest('.field')?.querySelector('.ny40-cop-hint');
   if(isNY40){
     if(!help){help=document.createElement('small');help.className='ny40-cop-hint';input.closest('.field')?.append(help);}
     if(help){help.textContent=cell==='D61'?'NY-40: không chọn COP M/P / M and P COP not permitted.':'NY-40: không sử dụng Auxiliary/Wheelchair COP / Not permitted.';help.hidden=false;}
   }else if(help)help.hidden=true;
 }
}
// v5.8.110 - Hall panel rapid entry. Suggestions never override a manually selected position.
const HALL_QTY_KEY = cell=>'HALL_QTY_'+cell;
let hallManualSelections=new Set(),hallAutoSelections=new Set(),hallManualQuantities=new Set();
function hallIndicatorCount(code,kind){
 if(kind==='HBV')return 0;
 if(kind==='PIH_PID')return 1;
 if(kind!=='PIV')return null;
 const match=String(code||'').toUpperCase().match(/PIV\d*[- ]?([A-Z])(\d{3,4})/);
 if(!match)return null;
 const family=match[1]+match[2];
 const singles=new Set(['A1010','A1030','A1080','C710','C730','C766','C780']);
 const doubles=new Set(['A1020','A1040','A1090','C720','C740','C776','C790']);
 return singles.has(family)?1:doubles.has(family)?2:null;
}
function hallGroupCars(){
 const model=$('model').value;
 const raw=String(currentValue(model==='MR'?'D21':'D22')||'').toUpperCase().trim();
 const hit=raw.match(/(?:^|\D)([1-4])\s*C(?:\D|$)/);
 return hit?Number(hit[1]):1;
}
function hallProjectCars(){
 const n=Number(currentValue('D11'));
 return Number.isInteger(n)&&n>=1?Math.min(4,n):hallGroupCars();
}
function hallQuantityMax(kind,code){
 const cars=hallGroupCars(),displays=hallIndicatorCount(code,kind);
 if(displays===2&&cars>=2)return cars-1;
 return Math.max(1,Math.min(4,cars));
}
function hallQuantityDefault(kind,code,max){
 // Only one-indicator call panels take the initial Quantity of lifts (D11).
 // A double-indicator panel remains 1 by default; it may be increased manually.
 return kind==='PIV'&&hallIndicatorCount(code,kind)===1 ? Math.min(max,hallProjectCars()) : 1;
}
function refreshHallQuantities(){
 document.querySelectorAll('.hall-qty-pair').forEach(pair=>{
  const input=pair.querySelector('select[data-cell]'),qty=pair.querySelector('select.hall-qty');
  if(!input||!qty)return;
  const kind=pair.dataset.kind||'PIV';
  const val=String(input.value||''),absent=!val||/^(?:Not supplied|Không cung cấp)/i.test(val);
  const num=hallIndicatorCount(val,kind),max=hallQuantityMax(kind,val);
  const key=HALL_QTY_KEY(input.dataset.cell);
  const fallback=hallQuantityDefault(kind,val,max);
  const saved=Number(changes[key]);
  const desired=hallManualQuantities.has(key)&&Number.isInteger(saved) ? Math.max(1,Math.min(max,saved)) : fallback;
  qty.replaceChildren();for(let n=1;n<=max;n++){const opt=document.createElement('option');opt.value=String(n);opt.textContent=String(n);qty.append(opt)}
  qty.value=String(desired);qty.disabled=absent;
  // Keep the export quantity in sync even if the user has not manually changed the selector.
  changes[key]=String(desired);
  const info=pair.closest('.hall-choice-item')?.querySelector('.hall-display-info');
  if(info){
   info.hidden=absent;
   info.textContent=kind==='HBV'?'Hộp nút gọi (không có màn hình hiển thị).':kind==='PIH_PID'?'Bảng hiển thị vị trí tầng.':num===1?'Loại 1 hiển thị / Single indicator':num===2?'Loại 2 hiển thị / Double indicator':'Chưa xác định số màn hình từ database; cần xác minh mã.';
  }
 });
}
// Match only the model/family body, e.g. PIV1-A1010H2 -> PIV1-A1010N.
// Only values listed in the TARGET dropdown may be chosen.
function hallCodeBody(code){
 const match=String(code||'').trim().toUpperCase().match(/^(PIV\d*|HBV\d*|PIH|PID)[- ]?([A-Z]\d{3,4})([A-Z0-9]*)$/);
 return match?{prefix:match[1],body:match[2],suffix:match[3]}:null;
}
function hallSiblingSuggestion(source,select,floor){
 const parsed=hallCodeBody(source);
 if(!parsed||!select||select.tagName!=='SELECT')return '';
 const candidates=[...select.options].filter(o=>!o.disabled&&!o.dataset.invalid).map(o=>({value:o.value,code:hallCodeBody(o.value)}))
  .filter(o=>o.code&&o.code.body===parsed.body&&o.code.prefix.replace(/\d+/g,'')===parsed.prefix.replace(/\d+/g,''));
 if(!candidates.length)return '';
 const preferred=floor==='TopFloor'?['H2','H','HA','N','NA']:['N','NA','B','BA','H2','H'];
 candidates.sort((a,b)=>{
  const rank=o=>{
   const prefixPenalty=o.code.prefix===parsed.prefix?0:30;
   const variant=preferred.indexOf(o.code.suffix);
   return prefixPenalty+(variant<0?20:variant);
  };
  return rank(a)-rank(b);
 });
 return candidates[0].value;
}
function autoFillHallSiblingFields(sourceField){
 const sourceName=sourceField?.name||'';
 const match=sourceName.match(/^K36_(PIV|HBV|PIH_PID)(TopFloor|MainFloor)$/);
 if(!match)return;
 const kind=match[1],from=match[2],model=$('model').value;
 const sourceValue=String(currentValue(sourceField.cell)||'');
 if(!hallCodeBody(sourceValue))return;
 // NEXIEZ-MR: Main Floor (1) -> Other Floors (2), PIV/PIH only.
 // HBV (call-button box) and LCD PIV families C766/C776 are manually selected.
 // NEXIEZ-MRL Version02 retains its existing independent autofill behavior.
 if(model==='MR'){
  if(from!=='MainFloor'||kind==='HBV')return;
  if(kind==='PIV'&&/C(?:766|776)(?!\d)/i.test(sourceValue))return;
 }
 const destinations=model==='MRL' ? (from==='TopFloor'?['MainFloor','OtherFloor']:['TopFloor','OtherFloor']) : ['OtherFloor'];
 for(const floor of destinations){
  const target=current.fields.find(f=>f.name==='K36_'+kind+floor);
  if(!target||target.cell===sourceField.cell||hallManualSelections.has(target.cell))continue;
  const el=activeInput(target.cell);
  const next=hallSiblingSuggestion(sourceValue,el,floor);
  if(!next||next===String(currentValue(target.cell)))continue;
  if(model==='MR'&&kind==='PIV'&&/C(?:766|776)(?!\d)/i.test(next))continue;
  applyValue(target.cell,next);
  hallAutoSelections.add(target.cell);
  // A newly suggested single panel gets its project-quantity default.
  hallManualQuantities.delete(HALL_QTY_KEY(target.cell));
 }
 refreshHallQuantities();
}
function mountHallQuantity(item,input,kind){
 if(kind==='PIH_PID')return;
 const wrapper=input.closest('.field');if(!wrapper)return;
 const pair=document.createElement('div');pair.className='hall-qty-pair';pair.dataset.kind=kind;
 const qtyWrap=document.createElement('div');qtyWrap.className='hall-qty-wrap';
 const qtyLabel=document.createElement('label');qtyLabel.className='hall-qty-label';qtyLabel.textContent='Số hộp / Qty';
 const qty=document.createElement('select');qty.className='hall-qty';qty.setAttribute('aria-label','Số lượng hộp gọi');
 qty.onchange=()=>{const key=HALL_QTY_KEY(input.dataset.cell);hallManualQuantities.add(key);changes[key]=qty.value;invalidateCheck();};
 qtyWrap.append(qtyLabel,qty);wrapper.parentNode.insertBefore(pair,wrapper);pair.append(wrapper,qtyWrap);
 const info=document.createElement('small');info.className='hall-display-info';item.append(info);
 input.addEventListener('change',()=>{hallManualQuantities.delete(HALL_QTY_KEY(input.dataset.cell));refreshHallQuantities();});
 refreshHallQuantities();
}
 // v5.8.157 – Group the MR source capacity rows without rewriting SpecForm or Excel rules.
 // D13, D14 and CWT retain their ORIGINAL nodes/cell addresses for Check Spec and Export.
 // v5.8.158: one-way controlled selection: P > Door > CWT > AA/BB > JJ > exact source row.
 // Original workbook-backed cells remain present (hidden only where a separate filter controls them).
 function mountMRCapacitySelector(root){
  if($('model').value!=='MR')return;
  const capacity=current.fields.find(f=>f.name==='K5_Capacity');
  if(!capacity)return;
  const native=activeInput(capacity.cell),parent=native?.closest('.field');
  if(!native||!parent)return;
  const source=current.capacity.filter(r=>String(r.EleCapacityName||'').trim());
  const groupName=r=>{const m=String(r.EleCapacityName||'').match(/^P0?(6|8|9|10|11|13|15|17|20)(?=[A-Z0-9-]|\b)/i);return m?'P'+m[1]:''};
  const groups=[...new Set(source.map(groupName).filter(Boolean))];if(!groups.length)return;
  const field=(name)=>current.fields.find(f=>f.name===name);
  const doorNative=activeInput(field('K7_DoorType')?.cell),cwtNative=activeInput(field('K14_CWTLocation')?.cell);
  const aaNative=activeInput(field('K17_CarW')?.cell),bbNative=activeInput(field('K17_CarD')?.cell),jjNative=activeInput(field('K18_CarJJ')?.cell);
  if(!doorNative||!cwtNative||!aaNative||!bbNative||!jjNative)return;
  const makeSelect=(id,host,label)=>{
   const box=document.createElement('div');box.className='mr-selector-control';
   // The existing row label already identifies the three Section I inputs.
   // Do not repeat a second label above their dropdowns (mobile compact layout).
   const compact=['mr-capacity-group','mr-door-choice','mr-cwt-choice'].includes(id);
   if(compact)box.classList.add('mr-selector-compact');
   if(!compact){const title=document.createElement('label');title.className='mr-selector-label';title.htmlFor=id;title.textContent=label;box.append(title)}
   const el=document.createElement('select');el.id=id;el.setAttribute('aria-label',label);box.append(el);host.append(box);return el;
  };
  const group=makeSelect('mr-capacity-group',parent,'Nhóm tải trọng P / Capacity group');
  // Keep Section I fields and original workbook cells; only hide duplicate internal controls.
  const door=makeSelect('mr-door-choice',doorNative.closest('.field'),'Door Type / Loại cửa');
  const cwt=makeSelect('mr-cwt-choice',cwtNative.closest('.field'),'CWT Position / Vị trí đối trọng');
  const aa=makeSelect('mr-aa-choice',aaNative.closest('.field'),'AA – Width (mm)');
  const bb=makeSelect('mr-bb-choice',bbNative.closest('.field'),'BB – Depth (mm)');
  const jj=makeSelect('mr-jj-choice',jjNative.closest('.field'),'JJ – Door width (mm)');
  for(const el of [native,doorNative,cwtNative,aaNative,bbNative,jjNative]){
   el.classList.add('mr-original-internal');el.setAttribute('aria-hidden','true');el.tabIndex=-1;
   const oldLabel=el.closest('.field')?.querySelector(':scope > label.nm');if(oldLabel)oldLabel.classList.add('mr-original-internal');
  }
  // Capacity source code is only requested for ambiguous, physically identical configurations.
  const exact=document.createElement('div');exact.className='mr-exact-selector';exact.hidden=true;
  const exactLabel=document.createElement('label');exactLabel.textContent='Mã P trùng cấu hình / Source code';
  const exactChoice=document.createElement('select');exactChoice.id='mr-source-choice';
  exact.append(exactLabel,exactChoice);parent.append(exact);
  const warning=document.createElement('small');warning.className='capacity-duplicate-warning';parent.append(warning);
  const getDoor=r=>String(r.DoorType||'').trim();const getCwt=r=>String(r['CWT-Position']||'').trim();
  const vals=(rows,fn)=>[...new Set(rows.map(fn).filter(v=>v!==''&&v!=null).map(String))];
  const filter=(rows,fn,v)=>rows.filter(r=>String(fn(r))===String(v));
  const fill=(el,options,preferred,labels={})=>{
   const opts=options.map(String);const prior=String(el.value||'');
   el.replaceChildren();for(const v of opts){const op=document.createElement('option');op.value=v;op.textContent=labels[v]||v;el.append(op)}
   const chosen=opts.includes(String(preferred))?String(preferred):opts.includes(prior)?prior:opts[0]||'';
   el.value=chosen;
   // Only Door Type uses grey/disabled styling for a single permitted option.
   // CWT, AA, BB and JJ remain visually active (blue) even when the
   // currently filtered configuration produces one candidate value.
   const singleDoor=(el===door && opts.length===1);
   el.disabled=singleDoor || opts.length===0;
   el.closest('.mr-selector-control')?.classList.toggle('single-option',singleDoor);
   el.closest('.mr-selector-control')?.classList.toggle('is-selectable',!el.disabled);
   return chosen;
  };
  for(const g of groups){const r=source.find(x=>groupName(x)===g);const op=document.createElement('option');op.value=g;op.textContent=g+' — '+r.Capacity+' kg';group.append(op)}
  let busy=false;
  const sourceVal=()=>String(currentValue(capacity.cell)||native.value);
  function run(level=0,prefs={}){
   if(busy)return;busy=true;
   try{
    const original=source.find(r=>String(r.EleCapacityName)===sourceVal());
    const base=original||source[0];
    const g=groups.includes(prefs.group)?prefs.group:(level===0?groupName(base):group.value)||groups[0];
    group.value=g;
    let rows=source.filter(r=>groupName(r)===g);
    const doorVal=fill(door,vals(rows,getDoor),level===0?getDoor(base):prefs.door??door.value);rows=filter(rows,getDoor,doorVal);
    const cwtVal=fill(cwt,vals(rows,getCwt),level===0?getCwt(base):prefs.cwt??cwt.value);rows=filter(rows,getCwt,cwtVal);
    const aaVal=fill(aa,vals(rows,r=>r.AA),level===0?base.AA:prefs.aa??aa.value);rows=filter(rows,r=>r.AA,aaVal);
    const bbVal=fill(bb,vals(rows,r=>r.BB),level===0?base.BB:prefs.bb??bb.value);rows=filter(rows,r=>r.BB,bbVal);
    const jjVal=fill(jj,vals(rows,r=>r.JJ),level===0?base.JJ:prefs.jj??jj.value);rows=filter(rows,r=>r.JJ,jjVal);
    if(!rows.length){warning.textContent='Không tìm thấy mã P ứng với cấu hình.';return}
    exact.hidden=rows.length===1;
    const last=sourceVal();const preferred=rows.find(r=>String(r.EleCapacityName)===String(prefs.code))||rows.find(r=>String(r.EleCapacityName)===last)||rows[0];
    exactChoice.replaceChildren();for(const r of rows){const op=document.createElement('option');op.value=String(r.EleCapacityName);op.textContent=String(r.EleCapacityName).split(' (')[0];exactChoice.append(op)}
    exactChoice.value=String(preferred.EleCapacityName);
    warning.textContent=rows.length>1?'⚠ Có '+rows.length+' mã P cùng cấu hình. Chọn mã gốc để đảm bảo Check Spec chính xác.':'';
    if(String(preferred.EleCapacityName)!==last){
      // Fire the ORIGINAL change listener: this updates all formulas, dependent specs and export changes.
      native.value=String(preferred.EleCapacityName);
      // If the capacity option is not present, add the selected exact code.
      if(native.value!==String(preferred.EleCapacityName)){
       const op=document.createElement('option');op.value=String(preferred.EleCapacityName);op.textContent=String(preferred.EleCapacityName);native.append(op);native.value=op.value;
      }
      native.dispatchEvent(new Event('change',{bubbles:true}));
    }
    // Explicitly synchronize row lookup when the exact code did not change.
    door.value=getDoor(preferred);cwt.value=getCwt(preferred);aa.value=String(preferred.AA);bb.value=String(preferred.BB);jj.value=String(preferred.JJ);
   }finally{busy=false}
  }
  group.addEventListener('change',()=>run(1,{group:group.value,door:door.value,cwt:cwt.value}));
  door.addEventListener('change',()=>run(1,{door:door.value,cwt:cwt.value,aa:aa.value,bb:bb.value,jj:jj.value}));
  cwt.addEventListener('change',()=>run(1,{door:door.value,cwt:cwt.value,aa:aa.value,bb:bb.value,jj:jj.value}));
  aa.addEventListener('change',()=>run(1,{aa:aa.value,bb:bb.value,jj:jj.value}));
  bb.addEventListener('change',()=>run(1,{aa:aa.value,bb:bb.value,jj:jj.value}));
  jj.addEventListener('change',()=>run(1,{jj:jj.value}));
  exactChoice.addEventListener('change',()=>run(1,{code:exactChoice.value}));
  native.addEventListener('change',()=>{if(!busy)queueMicrotask(()=>run(0))});
  run(0);
 }
function renderAllSpecRows(){
 const rows=current.layout||layouts[$('model').value]||[];
 const byRow=new Map();for(const f of current.fields){const n=split(f.cell)[1];if(!byRow.has(n))byRow.set(n,[]);byRow.get(n).push(f)}
 const root=$('fields'),project=$('project-fields');root.replaceChildren();project.replaceChildren();
 let rendered=0,section=null,sectionRows=null,lastMajor='';
 for(const info of rows){
  const fields=byRow.get(info.row)||[];
  if(info.row===7||info.row===8){
   for(const f of fields){const label=info.row===7?'Công trình / Project':'Hạng mục / Item';project.append(row(f,label));rendered++}
   continue;
  }
  const major=String(info.B||'').trim();const sub=String(info.C||'').trim();
  if(/^([IVX]+)\.\s*/.test(major)){
   section=document.createElement('section');section.className='spec-section';section.dataset.section=major.match(/^([IVX]+)/)[1];
   const heading=document.createElement('h2');heading.className='section-heading';heading.textContent=window.SPEC_PROOFREAD?.(headingText(major))??headingText(major);
   sectionRows=document.createElement('div');sectionRows.className='section-fields';
   section.append(heading,sectionRows);root.append(section);lastMajor='';continue;
  }
  if(!sectionRows)continue;
  if(major)lastMajor=major;
  const isInfoNote=/^(?:Ex:\s*|Không tính cửa thoát hiểm)/i.test(sub);
  const currentLabel=major||(sub?subLabel(sub):'');
  const label=major&&sub&&!isInfoNote?major+' · '+subLabel(sub):currentLabel||lastMajor;
  if(!fields.length && !String(info.D||'').trim())continue;
  const sectionRow=document.createElement('div');sectionRow.className='spec-row';sectionRow.dataset.row=info.row;
  const text=document.createElement('div');text.className='spec-row-title';text.textContent=window.SPEC_PROOFREAD?.(headingText(label||lastMajor))??headingText(label||lastMajor);
  const area=document.createElement('div');area.className='spec-row-inputs';
  if(fields.length>1){
   sectionRow.classList.add('span-full','multi');
   for(const f of fields){const title=fieldRole(f);area.append(row(f,title,true));rendered++}
  }else if(fields.length===1){
   const f=fields[0];area.append(row(f,'',false));rendered++;
  }else{
   const value=document.createElement('div');value.className='readonly-value';value.textContent=String(info.D||'');area.append(value);
  }
  sectionRow.append(text,area);sectionRows.append(sectionRow);
 }
  // Preserve original field node and bindings; move the CWT row just below Door Type.
  {
    const model=$('model').value;
    const sectionI=root.querySelector('.spec-section[data-section="I"] .section-fields');
    const doorRow=sectionI?.querySelector('.spec-row[data-row="'+(model==='MR'?'14':'15')+'"]');
    const cwtRow=sectionI?.querySelector('.spec-row[data-row="'+(model==='MR'?'22':'23')+'"]');
    if(doorRow&&cwtRow)doorRow.after(cwtRow);
  }
  mountMRCapacitySelector(root);
  buildCarDesignLines();
 // v5.8.16: Arrange both models in wide, independently editable dimension groups.
 // Move the existing input nodes (never clone them) so Check Spec / Export bindings stay unchanged.
 {
  const model=$('model').value;
  const dimensionSection=[...root.querySelectorAll('.spec-section')].find(x=>x.dataset.section==='II');
  const dimensionGrid=dimensionSection?.querySelector('.section-fields');
  if(dimensionGrid){
   const groups=model==='MRL' ? [
    {title:'Kích thước thang máy',en:'Elevator dimensions',rows:[28,29,30,31,32],labels:['AA – Rộng / Width','BB – Sâu / Depth','HL – Cao / Height','JJ – Rộng cửa / Door width','HH – Cao cửa / Door height']},
    {title:'Kích thước giếng thang máy',en:'Hoistway dimensions',rows:[33,34,35,37,36],labels:['AH – Rộng / Width','BH – Sâu / Depth','Travel – Hành trình','OH – Đỉnh hố thang','PD – Chiều sâu Pit']}
   ] : [
    {title:'Kích thước thang máy',en:'Elevator dimensions',rows:[29,30,31,32,33],labels:['AA – Rộng / Width','BB – Sâu / Depth','HL – Cao / Height','JJ – Rộng cửa / Door width','HH – Cao cửa / Door height']},
    {title:'Kích thước giếng thang máy',en:'Hoistway dimensions',rows:[34,35,36,37,38,39,41,40],labels:['AH – Rộng / Width','BH – Sâu / Depth','AM – Rộng phòng máy / MR width','BM – Sâu phòng máy / MR depth','MH – Cao phòng máy / MR height','TR – Hành trình / Travel','OH – Đỉnh hố thang','PD – Chiều sâu Pit']}
   ];
   const grouped=document.createElement('div');grouped.className='dimension-groups';
   grouped.dataset.model=model;
   for(let groupIndex=0;groupIndex<groups.length;groupIndex++){
    const group=groups[groupIndex];
    const card=document.createElement('section');card.className='dimension-group';
    card.classList.add(groupIndex===0?'dimension-car':'dimension-hoistway');
    const heading=document.createElement('h3');heading.className='dimension-group-title';
    heading.textContent=group.title+' / '+group.en;card.append(heading);
    group.rows.forEach((sourceRow,i)=>{
     const original=dimensionGrid.querySelector('.spec-row[data-row="'+sourceRow+'"]');
     if(!original)return;
     original.classList.add('dimension-entry');
     const label=original.querySelector('.spec-row-title');
     if(label)label.textContent=group.labels[i];
     // MR only: fixed maintenance headroom note beside MH (source row 38).
     // Informational display; do not bind it to SpecForm, Check Spec, or Export.
     if(model==='MR'&&sourceRow===38){
      const note=document.createElement('small');
      note.className='mr-mh-maintenance-note';
      note.textContent='Chiều cao bảo trì tối thiểu 1770 mm – Sử dụng thép tấm thay thế Underbeam H125';
      original.append(note);
     }
     card.append(original);
    });
    grouped.append(card);
   }
   dimensionGrid.classList.add('dimension-section-fields');
   dimensionGrid.prepend(grouped);
  }
 }
 // Consolidate hall call / indicator fields by floor, rather than by Excel source row.
 // The original input nodes and their data-cell bindings are MOVED, not cloned.
 const hallFields = $('model').value==='MRL' ? [
  {key:'TopFloor',vi:'Tầng trên cùng',en:'Top floor',cells:['E75','E76','E77']},
  {key:'OtherFloor',vi:'Tầng khác',en:'Other floors',cells:['E81','E82','E83']},
  {key:'MainFloor',vi:'Tầng chính',en:'Main floor',cells:['E78','E79','E80']}
 ] : [
  {key:'MainFloor',vi:'Tầng chính',en:'Main floor',cells:['E78','E79','E80']},
  {key:'OtherFloor',vi:'Tầng khác',en:'Other floors',cells:['E81','E82']}
 ];
 const entranceSection=[...root.querySelectorAll('.spec-section')].find(x=>x.dataset.section==='IV');
 const landingRows=entranceSection?.querySelector('.section-fields');
 if(landingRows){
  const panel=document.createElement('div');panel.className='hall-choice-panel';
  const title=document.createElement('h3');title.className='hall-choice-title';title.textContent='Bảng gọi tầng và hiển thị / Hall indicators & buttons';panel.append(title);
  const grid=document.createElement('div');grid.className='hall-floor-grid';
  for(const group of hallFields){
   const card=document.createElement('section');card.className='hall-floor-card';card.dataset.floor=group.key;
   const head=document.createElement('h4');head.className='hall-floor-title';
   const index=document.createElement('span');index.className='hall-floor-marker';index.textContent=String(grid.children.length+1).padStart(2,'0');
   head.append(index,document.createTextNode(group.vi+' / '+group.en));card.append(head);
   const content=document.createElement('div');content.className='hall-floor-options';
   const labels=['Hiển thị & nút bấm (PIV) / Hall indicators & buttons', 'Nút gọi tầng (HBV) / Hall buttons','Hiển thị tầng (PIH/PID) / Hall position indicators'];
   for(let i=0;i<group.cells.length;i++){
    const cell=group.cells[i],input=landingRows.querySelector('[data-cell="'+cell+'"]');
    if(!input)continue;
    const wrapper=input.closest('.field'),oldRow=input.closest('.spec-row');
    if(!wrapper||!oldRow)continue;
    const item=document.createElement('div');item.className='hall-choice-item';item.dataset.sourceRow=oldRow.dataset.row;
    const label=document.createElement('label');label.className='hall-choice-label';label.textContent=labels[i];
    item.append(label,wrapper);content.append(item);mountHallQuantity(item,input,['PIV','HBV','PIH_PID'][i]);oldRow.remove();
   }
   if($('model').value==='MRL'&&window.SPEC_HALL_LANTERN){
    const opts=window.SPEC_HALL_LANTERN;
    const key=opts.keys[group.key];
    const item=document.createElement('div');item.className='hall-choice-item hall-lantern-choice';
    const label=document.createElement('label');label.className='hall-choice-label';
    label.textContent='Đèn báo tầng (Hall Lanterns)';
    const select=document.createElement('select');select.className='hall-lantern-select';select.dataset.cell=key;
    const all=[opts.none,...opts.catalog.map(x=>x.code)];
    for(const code of all){const el=document.createElement('option');el.value=code;el.textContent=code;select.append(el)}
    let value=opts.selection(changes,key);select.value=value;
    const thumb=document.createElement('img');thumb.className='hall-lantern-thumb';thumb.alt='Hall Lantern';
    const view=document.createElement('button');view.type='button';view.className='btn hall-lantern-view';view.textContent='Xem bảng mẫu / Gallery';
    function preview(){const val=opts.byCode.get(select.value);if(val){thumb.src=val.image;thumb.alt=val.code;thumb.hidden=false}else{thumb.hidden=true;thumb.removeAttribute('src')}}
    select.onchange=()=>{
      changes[key]=select.value;preview();
      const fhl=sourceFeatureRows().find(f=>f.code==='FHL');
      if(fhl){if(opts.any(changes)){changes[fhl.cell]='1';featureTouched.delete(fhl.cell)}else{changes[fhl.cell]='0';featureTouched.delete(fhl.cell)}}
      renderFeatures();invalidateCheck();status('Đã cập nhật đèn báo tầng và tính năng FHL / Hall Lantern & FHL updated. Cần Check Spec.');
    };
    view.onclick=()=>openHallLanternGallery(key,select);
    const control=document.createElement('div');control.className='hall-lantern-controls';control.append(thumb,view);
    item.append(label,select,control);content.append(item);preview();
   }
   if(content.children.length){card.append(content);grid.append(card)}
  }
  panel.append(grid);
  if(grid.children.length){
   const firstHallRow=landingRows.querySelector('[data-row="'+($('model').value==='MRL'?75:78)+'"]');
   // Insert before fireproof-door selector to keep entrance design in workbook order.
   const fireRow=landingRows.querySelector('[data-row="'+($('model').value==='MRL'?84:83)+'"]');
   landingRows.insertBefore(panel,fireRow||null);
  }
 }
 mountJambDepthFields();
 // Build a compact section navigator without changing any workbook inputs or source-cell bindings.
 const nav=$('spec-nav');if(nav){nav.replaceChildren();
  const headings=[...root.querySelectorAll('.spec-section')];
  headings.forEach((sec,i)=>{const id='spec-section-'+(i+1);sec.id=id;
   const text=sec.querySelector('.section-heading')?.textContent||'';
   const link=document.createElement('a');link.href='#'+id;
   const numeral=document.createElement('b');numeral.textContent=(text.match(/^([IVX]+)\./)||[])[1]||String(i+1);
   const label=document.createElement('span');label.textContent=(text.replace(/^[IVX]+\.\s*/,'').split('(')[0].trim()||'Specifications');
   link.append(numeral,label);nav.append(link);
  });
  const features=document.querySelector('.features-panel');if(features){features.id='spec-section-features';const link=document.createElement('a');link.href='#spec-section-features';link.innerHTML='<b>V</b> <span>Tính năng / Features</span>';nav.append(link)}
 }
 $('numfields').textContent=String(rendered);
 const missing=rows.filter(x=>x.row>=10 && /^([IVX]+)\.\s*/.test(String(x.B||''))===false && (byRow.get(x.row)||[]).length && !document.querySelector('[data-row="'+x.row+'"]') && !document.querySelector('.hall-choice-item[data-source-row="'+x.row+'"]'));
 if(missing.length)status('Thiếu trường trong giao diện: '+missing.map(x=>x.row).join(', '),'error');
}
function invalidateCheck(){
 if($('model').value==='MRL')refreshMRLDecorationMaximum();
 window.__lastCheck=null;
 const top=$('check-top');if(top){top.className='check-top partial';top.textContent='Đã thay đổi — cần kiểm tra lại';}
 const issues=$('check-issues');if(issues)issues.replaceChildren();
}
function restoreHallLanternFromSource(){
 if($('model').value!=='MRL'||!window.SPEC_HALL_LANTERN)return;
 const rules=window.SPEC_HALL_LANTERN,db=workbookEngineData();
 const sh=db?.sheets?.HallLantern_V02;
 if(!sh)return;
 for(const [key,addr] of [['HL_TOP','B21'],['HL_OTHER','B22'],['HL_MAIN','B23']]){
  const entry=sh[addr],v=typeof entry==='object'?(entry.v??''):entry;
  if(v&&rules.allowed(String(v))&&String(v)!==rules.none)changes[key]=String(v);
 }
 if(rules.any(changes)){
  const fhl=sourceFeatureRows().find(f=>f.code==='FHL');if(fhl)changes[fhl.cell]='1';
 }
}

/* v5.8.47: ETD lead time is display-only, deliberately NOT a SpecForm field,
   not included in Excel/PDF or Check Spec payloads. MRL rules require its
   separate factory table; MR rules follow the supplied (revised) table. */
function calculateMRLeadTime(){
 const val=(cell)=>String(currentValue(cell)||'').trim();
 const reasons50=[],reasons90=[];
 const ceiling=val('D45').toUpperCase();
 const mirror=val('D69').toUpperCase();
 const handrail=val('D67').toUpperCase();
 const carDoor=val('D50').toUpperCase();
 const landingDoor=[val('E76'),val('E77')].map(x=>x.toUpperCase());
 const features=sourceFeatureRows().filter(x=>featureIsSelected(x)).map(x=>String(x.code||'').toUpperCase());
 const obs=[val('D14'),val('D50'),...landingDoor].some(x=>/OBSERVATION|NY-40|SEE[ -]?THROUGH/i.test(x));
 if(obs||features.some(x=>/^(?:OBS|STH)$/.test(x)))reasons50.push('Observation elevator / Thang quan sát');
 const allDoorValues=[carDoor,...landingDoor];
 const glassDoor=allDoorValues.filter(x=>/GLASS WINDOW|GLASS-WINDOW|LARGE WINDOW|200\s*[X×]\s*1300/i.test(x));
 if(glassDoor.some(x=>!/(?:200\s*[X×]\s*1300)/i.test(x)))reasons50.push('Glass window (khác 200×1300 mm)');
 const seeThrough=allDoorValues.some(x=>/SEE[ -]?THROUGH|DSJ/i.test(x));
 if(seeThrough){
  const dsj1=allDoorValues.some(x=>/DSJ1/i.test(x));
  const exception=dsj1&&Number(val('D33'))===2100&&[800,900].includes(Number(val('D32')));
  if(!exception)reasons50.push('See-through door (không thuộc ngoại lệ DSJ1, HH 2100, JJ 800/900)');
 }
 if(['L200','L200S','N130','L400'].some(x=>ceiling===x||ceiling.startsWith(x+'-')))reasons50.push('Ceiling '+val('D45'));
 if(/YZ[ -]?5(?:2A|3A|5SN)/i.test(mirror))reasons50.push('Mirror '+val('D69'));
 if(/YH[ -]?59M/i.test(handrail))reasons50.push('Handrail '+val('D67'));
 const reasons120=[];
 if(/(?:^|[^A-Z0-9])(?:FEVN10|FE10)(?=$|[^A-Z0-9])|VIETNAM/i.test(ceiling))reasons120.push('Ceiling FEVN10 (FE10 / Vietnam type): Extra 120 days');
 const extra=reasons120.length?120:(reasons90.length?90:(reasons50.length?50:0)); // Factory tier, 0 means standard 30 days.
 return {extra,days:(extra || 30)+5,reasons:[...reasons120,...reasons90,...reasons50]};
}
/* v5.8.47: factory table for NEXIEZ-MRL Version02; ETD equals
   the applicable production lead-time tier (30/50/90/120) + 5 days.
   Only assess factory conditions with an actual corresponding input in
   the current SpecForm; never assume country/regulatory/COP material. */
function calculateMRLLeadTime(){
 const val=(cell)=>String(currentValue(cell)||'').trim();
 const norm=(v)=>String(v||'').trim().toUpperCase().replace(/\s+/g,' ');
 const ceiling=norm(val('D41'));
 const cap=norm(val('D13')).split(/\s*\(/)[0].trim();
 const aa=Number(val('D28')),jj=Number(val('D31'));
 const mirror=norm(val('D66'));
 const featureCodes=new Set(sourceFeatureRows().filter(f=>featureIsSelected(f)).map(f=>norm(f.code)));
 const triggered=[];
 const add=(extra,why)=>triggered.push({extra,why});
 // 50 days: eligible window dimensions, DS1/DSJ1 except JJ 800/900,
 // selected mirrors (with AA restriction) and NY-40 observation car.
 const doorValues=['D46','E73','E74','D14'].map(val).map(norm);
 if(doorValues.some(v=>/(?:200\s*[X×]\s*800|300\s*[X×]\s*1300)/.test(v)))
   add(50,'Glass window 200×800 / 300×1300 mm');
 if(![800,900].includes(jj)&&doorValues.some(v=>/\bDSJ?1\b|SEE[ -]?THROUGH/.test(v)))
   add(50,'See-through door DS1/DSJ1 (JJ khác 800/900 mm)');
 if((/YZ[- ]?5(?:2A|3A)/.test(mirror)&&![1000,1100].includes(aa))||/YZ[- ]?55SN/.test(mirror))
   add(50,'Mirror '+val('D66')+( /YZ[- ]?55SN/.test(mirror)?'':' (AA không phải 1000/1100 mm)'));
 if(/^NY-40/.test(norm(val('D14'))))add(50,'Observation NY-40');
 // 50-day CL1 exception is for named regulation packages, not every CL1.
 // 90 days: CL2/DL5/DL6/DL7 only when SUS-M finish is positively selected.
 const finishInfo=[val('D40'),val('D42'),ceiling].map(norm).join(' ');
 if(/^(?:CL2|DL5|DL6|DL7)(?:[- ]|$)/.test(ceiling)&&/SUS[- ]?M\b/.test(finishInfo))
   add(90,'Ceiling '+val('D41')+' with SUS-M finish');
 // 120 days: CL1 for EN81-20/50 Malaysia and insulation require
 // an explicit regulatory option, which current SpecForm does not expose.
 const specialCar=new Set(['P06-2S','P13D-CO','P13D-2S']);
 if(/^DL7-[CW](?:$|[- ])/i.test(ceiling)&&specialCar.has(cap))
   add(120,'Ceiling '+val('D41')+' with car '+cap);
 if(/^FL1(?:$|[- ])/i.test(ceiling)&&!new Set(['P13D-CO','P13D-2S']).has(cap))
   add(120,'Ceiling FL1 (car khác P13D-CO/2S)');
 const pivCells=['E75','E78','E81'];
 const pivTriggered=pivCells.map(val).some(v=>/PIV\w*[- ]?(?:1210|1230|1280|1220|1240|1290)BA\b/i.test(v));
 if(featureCodes.has('HOS')&&pivTriggered)add(120,'HOS switch với PIVx-1210BA / 1230BA / 1280BA / 1220BA / 1240BA / 1290BA');
 if(featureCodes.has('SNCD'))add(120,'SNCD feature');
 const extra=triggered.length?Math.max(...triggered.map(x=>x.extra)):0; // Factory tier, 0 means standard 30 days.
 return {extra,days:(extra || 30)+5,reasons:triggered.sort((a,b)=>b.extra-a.extra).map(x=>x.why)};
}
function refreshLeadTime(){
 const days=document.getElementById('leadtime-days'),reason=document.getElementById('leadtime-reasons');if(!days||!reason||!current)return;
 const r=$('model').value==='MRL'?calculateMRLLeadTime():calculateMRLeadTime();days.textContent=r.days+' days';days.classList.toggle('leadtime-standard',Number(r.days)<=35);days.classList.toggle('leadtime-extended',Number(r.days)>35);reason.replaceChildren();
 const lines=r.reasons.length?r.reasons:[];
 for(const line of lines){const div=document.createElement('div');div.textContent=line;reason.append(div)}
}
// Re-evaluate after the existing event handler updates `changes`, including
// delegated controls and model/database reloads. No export pipeline changes.
document.addEventListener('change',()=>queueMicrotask(refreshLeadTime));

function load(){clearGroupControlPreview();hallManualSelections=new Set();hallAutoSelections=new Set();hallManualQuantities=new Set();featureTouched=new Set();window.__lastCheck=null;if($('check-top')){$('check-top').textContent='Chưa kiểm tra';$('check-top').className='check-top';}if($('check-issues'))$('check-issues').replaceChildren();changes={};let model=$('model').value;current=models[model];restoreHallLanternFromSource();renderAllSpecRows();setMRDefaultGroupAndFER();setMRLHandrailDefault();if(isMRBedCapacity())applyMRBedDefaults(currentValue('D13'));refreshTwoStopTravelWarning();refreshMRBedShaftNotes();setupSpecialOptions();flooringManual=false;updateFlooringWeight(true);syncDropdownOptions();syncLandingFinishFromCar();refreshNY40CopChoices();refreshSideDrop();refreshLocks();refreshDesignHints();updateFinishPreviews();refreshMirrorWeightHint();refreshMRLOverheadMinimum(false);refreshMRLShaftMinimum();renderFeatures();refreshGroupControlPreview();refreshHallQuantities();refreshMRLDecorationMaximum();$('numvalid').textContent=current.validations.length;$('numtables').textContent=current.table_count;$('filename').textContent=sourceFiles[model]?.name||modelFiles[model]+' (embedded copy)';status('Đã hiển thị '+(current.layout||layouts[model]).length+' dòng có nội dung trong SpecForm của '+$('model').selectedOptions[0].text+'.\nCheck Spec: tính lại công thức DataTable. Export Excel/A4 theo dữ liệu SpecFormExport, chỉ khi Check Spec PASS.','');refreshLeadTime();}

function issueDetails(raw){
 const m=/^([A-Z]{1,3}\d+):\s*(.*)$/.exec(raw);
 const cell=m?.[1]||'',reason=m?.[2]||raw;
 const f=current.fields.find(x=>x.cell===cell);
 const feature=sourceFeatureRows().find(x=>x.cell===cell);
 const label=(xrules().labels||{})[cell]||f?.label||feature?.vi||feature?.code||'Thông số kỹ thuật / Specification';
 const value=feature?(featureIsSelected(feature)?'✓':'✕'):cell?String(currentValue(cell)??''):'';
 let action='Kiểm tra và chọn lại giá trị hợp lệ trong Excel Database / Review allowed value in Excel Database';
 if(/Side drop/i.test(reason))action='Chọn BWH-620D/E nếu tải ≤1000 kg, BWH-720F/G nếu tải ≥1275 kg / Select correct side drop for capacity';
 else if(/Màu|hoa văn|finish/i.test(reason))action='Xóa màu sơn/hoa văn không áp dụng; SUS-HL không có màu sơn / Clear inapplicable paint or pattern';
 else if(/Feature.*X|không áp dụng/i.test(reason))action='Bỏ chọn Feature bị cấm (Option X) / Deselect prohibited feature';
 else if(/Capacity|Expected|kích thước/i.test(reason))action='Đồng bộ thông số theo Capacity trong DataTable / Match the capacity-dependent specification';
 else if(/Observation|Entrance/i.test(reason))action='Chọn lại loại cửa/quan sát hợp lệ / Select valid entrance or observation';
 return {cell,reason,label,value,action,feature};
}
function focusIssue(cell){
 // Dimension rows use display:contents for .field and .spec-row-inputs;
 // scrolling those elements has no physical box. Scroll a visible row instead.
 const feature=sourceFeatureRows().find(x=>x.cell===cell||(x.code==='ITP'&&cell===('H'+x.cell.slice(1))));
 if(feature&&activeFeatureTable!==feature.table){activeFeatureTable=feature.table;renderFeatures()}
 const input=Array.from(document.querySelectorAll('[data-cell]')).find(x=>x.dataset.cell===cell);
 if(!input)return;
 const section=input.closest('.spec-section,.features-panel');
 // Reveal the destination if a section has been collapsed elsewhere.
 if(section?.hidden)section.hidden=false;
 const row=input.closest('.dimension-entry,.hall-choice-item,.spec-row,.feature-line,.field')||input;
 const visible=node=>node&&node.getClientRects().length&&getComputedStyle(node).display!=='contents';
 const target=visible(row)?row:input;
 // Use a viewport coordinate for a stable offset below the top header.
 const topOffset=Math.max(14,Math.min(96,(document.querySelector('.site-header')?.getBoundingClientRect().bottom||0)+12));
 const y=window.scrollY+target.getBoundingClientRect().top-topOffset-18;
 window.scrollTo({top:Math.max(0,y),behavior:'smooth'});
 row.classList.add('issue-focus');input.classList.add('issue-input-focus');
 setTimeout(()=>{row.classList.remove('issue-focus');input.classList.remove('issue-input-focus')},3000);
 // Do not let focus trigger a second scrolling action on mobile browsers.
 if(!input.disabled)try{input.focus({preventScroll:true})}catch(e){input.focus()}
}

// v5.8.105: Floor Name errors must start from COP No. 1, then hall panels.
// Ranking is independent of worksheet row numbers (MRL report otherwise sorts bottom-up).
function floorDisplayErrorRank(err){
 const id=String(err?.id||''),cell=String(err?.cell||'').toUpperCase();
 if(!id.startsWith('FloorName!Display!'))return null;
 const model=$('model').value;
 const cop=model==='MR'?'D65':'D61';
 if(cell===cop)return 0; // COP No. 1 (main COP)
 const order=model==='MRL'
  ? ['E75','E78','E81','E76','E79','E82','E77','E80','E83']
  : ['E78','E81','E79','E82','E80','E83'];
 const ix=order.indexOf(cell);
 return ix>=0?10+ix:90;
}
function orderedCheckErrors(errors){
 return [...errors].sort((a,b)=>{
  const aRank=floorDisplayErrorRank(a),bRank=floorDisplayErrorRank(b);
  if(aRank!==null||bRank!==null){
   if(aRank===null)return 1;
   if(bRank===null)return -1;
   return aRank-bRank;
  }
  const num=e=>{const m=/\d+/.exec(String(e.cell||''));return m?Number(m[0]):9999;};
  return $('model').value==='MRL'?num(b)-num(a):num(a)-num(b);
 });
}
function renderCheckIssues(result){
 const wrap=$('check-issues');if(!wrap)return;wrap.replaceChildren();
 const all=[...orderedCheckErrors(result.errors).map(x=>({raw:x,type:'error'})),...result.warnings.filter(x=>/^[A-Z]{1,3}\d+:/.test(x)).map(x=>({raw:x,type:'warning'}))];
 const head=document.createElement('div');head.className='check-summary';
 head.textContent=result.errors.length?`FAIL — ${result.errors.length} lỗi cần sửa / ${result.errors.length} errors to fix`:'PASS PARTIAL — Đạt quy tắc đã triển khai; vẫn còn điều kiện chưa xác minh / Implemented rules passed; some conditions unverified';
 head.classList.add(result.errors.length?'check-fail':'check-pass');wrap.append(head);
 if(!all.length){const p=document.createElement('p');p.className='muted';p.textContent='Chưa phát hiện lỗi Partial. Đây không phải phê duyệt Check Spec Full / Partial pass is not full approval.';wrap.append(p);return}
 for(const item of all){const info=issueDetails(item.raw);const card=document.createElement('div');card.className='issue-card '+(item.type==='error'?'issue-error':'issue-warning');
  const title=document.createElement('strong');title.textContent=`${info.cell||'RULE'} — ${info.label}`;
  const detail=document.createElement('div');detail.textContent=`Giá trị hiện tại / Current value: ${info.value||'(trống / blank)'}`;
  const why=document.createElement('div');why.textContent=`Nguyên nhân / Reason: ${info.reason}`;
  const fix=document.createElement('div');fix.textContent=`Cách xử lý / Action: ${info.action}`;
  card.append(title,detail,why,fix);
  if(info.cell){const btn=document.createElement('button');btn.type='button';btn.className='btn issue-jump';btn.textContent='Đến ô lỗi / Go to field';btn.onclick=()=>focusIssue(info.cell);card.append(btn)}wrap.append(card)
 }
}
function renderFormulaAudit(audit){
 const wrap=$('formula-audit');if(!wrap)return;wrap.replaceChildren();
 const h=document.createElement('p');h.textContent='Danh mục công thức từ workbook. Check Spec tính lại các điều kiện nguồn và ràng buộc C# đã chuyển đổi; công thức không tính được được ghi là INCOMPLETE. / Source formulas are recalculated during Check Spec; uncomputable conditions are never treated as PASS.';wrap.append(h);
 for(const sheet of ['SpecForm','DataTable']){
  const heading=document.createElement('h4');heading.textContent=sheet;wrap.append(heading);
  for(const [addr,formula] of Object.entries(audit.form[sheet]||{})){
   const line=document.createElement('div');line.style.cssText='padding:5px;border-bottom:1px solid #ddd;overflow-wrap:anywhere';
   const k=sheet+'!'+addr;
   const status=Object.hasOwn(audit.checked,addr)?(audit.checked[addr]?'PASS – partial expression':'FAIL'):'NOT VERIFIED';
   line.textContent=k+'  |  '+status+'  |  ='+formula;wrap.append(line);
  }
 }
}
function runCheckSpec(){
 const top=$('check-top');
 if(top){top.textContent='Đang kiểm tra…';top.className='check-top running';}
 try{
  const errors=check();
  if(top){
   if(!window.__lastCheck){top.textContent='CHECK NOT AVAILABLE';top.className='check-top fail';}
   top.scrollIntoView({behavior:'smooth',block:'start'});
  }
  return errors;
 }catch(err){
  const message='CHECK SPEC ERROR — Không thể hoàn tất kiểm tra / Validation crashed: '+(err?.message||String(err));
  if(top){top.textContent=message;top.className='check-top fail';top.scrollIntoView({behavior:'smooth',block:'start'});}
  const box=$('check-issues');if(box){box.replaceChildren();const h=document.createElement('div');h.className='check-summary check-fail';h.textContent=message;box.append(h);}
  status(message,'error');console.error('Check Spec failed',err);
  return [message];
 }
}
function mrlDecorationFeatureSelected(code){
 if($('model').value!=='MRL')return false;
 const canon=x=>String(x||'').toUpperCase().replace(/[\s_-]+/g,'').trim();
 const want=canon(code);
 return sourceFeatureRows().some(f=>{
  if(canon(f.code)!==want)return false;
  const state=changes[f.cell]??f.apply;
  return ['1','TRUE','YES','V','✓'].includes(String(state).toUpperCase());
 });
}
function evaluateMRLDecoration(){
 if($('model').value!=='MRL'||!window.SPEC_MRL_V02_DECOR)return null;
 return window.SPEC_MRL_V02_DECOR.validate(currentValue,mrlDecorationFeatureSelected);
}
function refreshMRLDecorationMaximum(){
 if($('model').value!=='MRL')return;
 const hint=$('mrl-decor-maximum');if(!hint)return;
 let r;
 try{r=evaluateMRLDecoration()}
 catch(e){r={state:'unavailable',reason:String(e?.message||e)}}
 hint.textContent=r?.state==='ready'?'Maximum: '+r.max.toLocaleString('en-US')+' kg':'Maximum: —';
 if(r?.overweightFlooring){
  const note=document.createElement('span');note.className='mrl-flooring-overlimit-note';
  note.textContent=' — Flooring vượt Maximum. Vui lòng giảm trọng lượng sàn đá.';
  hint.append(note);
 }
 hint.classList.toggle('unavailable',r?.state!=='ready');
 hint.classList.toggle('overlimit',!!r?.overweightFlooring||!!r?.errors?.some(x=>x.id==='MRL_V02_DECOR!Maximum'));
 hint.title=r?.reason||'Chưa xác định được kết quả theo Technical Study G11L1';
}
function check(){
 const model=$('model').value;
 const engineData=workbookEngineData();
 const result=window.SPEC_SOURCE_CHECK?.(model,changes,engineData);
 if(!result)throw Error('Source Check Engine chưa được tải / Calculation engine not loaded');
 if(model==='MRL'){
  const decor=evaluateMRLDecoration();
  if(decor){
   result.errors.push(...decor.errors);
   if(decor.state==='unavailable')result.unverified.push('Decoration Weight Maximum NEXIEZ-MRL Version2: '+decor.reason);
   if(decor.errors.length)result.status='FAIL';
   else if(decor.state==='unavailable'&&result.status==='PASS')result.status='INCOMPLETE';
   if(decor.state==='ready')result.verified.push('Technical Study G11L1 v2.5.0.0: '+decor.max+' kg excluding flooring');
  }
 }
 if(model==='MRL'&&window.SPEC_MATERIAL_5825){
  const r=window.SPEC_MATERIAL_5825;
  for(const f of current.fields){if(!/^(K22_CeilingColor|K23_WallFinish|K24_CarDoorFinish|K25_CarTransomFinish|K23_WallPaintedColor|K24_CarDoorPaintedColor|K25_CarTransomPainted|K26_CarFrontPainted|K23_WallPattern|K24_CarDoorPattern|K25_CarTransomPattern)$/.test(f.name))continue;
   const v=String(currentValue(f.cell)||'').trim();if(v&&!r.check(model,f,v))result.errors.push({id:'Material!'+f.name,cell:f.cell,message:'Mã vật liệu / màu không hợp lệ với MRL V02: '+v});
  }
  if(result.errors.length)result.status='FAIL';
 }
 // v5.8.146: Landing-door glass-window finishes cannot be used with
 // EN81-58(II). Check both main-floor and other-floor landing door fields;
 // never use car-door finish for this landing-door restriction.
 if(['MR','MRL'].includes(model)){
  const byName=name=>current.fields.find(f=>f.name===name);
  const fireField=byName('K36_FireproofDoor');
  const fire=fireField?String(currentValue(fireField.cell)||''):'';
  if(/EN81\s*[-–]?\s*58\s*\(II\)/i.test(fire)){
   const landingDoorFields=[
    ['K36_DoorFinishMainFloorMaterial','cửa tầng chính'],
    ['K36_DoorFinishOtherFloorMaterial','cửa tầng khác']
   ];
   for(const [name,location] of landingDoorFields){
    const f=byName(name);
    if(!f)continue;
    const finish=String(currentValue(f.cell)||'');
    if(/DSJ1|LARGE\s*[- ]?\s*WINDOW/i.test(finish)){
     result.errors.push({
      id:'LandingDoor!WindowFireproof!'+name,
      cell:f.cell,
      message:'Không được kết hợp '+finish+' ở '+location+' với cửa chống cháy EN81-58(II). Hãy bỏ lựa chọn DSJ1/Large Window hoặc đổi điều kiện chống cháy. / DSJ1/Large Window landing door is incompatible with EN81-58(II).'
     });
     result.status='FAIL';
    }
   }
  }
 }
 if(['MR','MRL'].includes(model)&&window.SPEC_JAMB_DEPTH){
  const jdErrors=window.SPEC_JAMB_DEPTH.validate(model,currentValue);
  if(jdErrors.length){result.errors.push(...jdErrors);result.status='FAIL';}
 }
 if(window.SPEC_MIRROR_RULES && model==='MR'){
  const mr=true,val=currentValue('D69'),check=window.SPEC_MIRROR_RULES.validate(model,currentValue('D13'),currentValue('D29'),val,currentValue('D31'));
  if(!check.ok){result.errors.push({id:'Mirror!Eligibility',cell:mr?'D69':'D66',message:check.message});result.status='FAIL';}
 }
 renderSourceCheck(result);
 window.__lastCheck=result;
 status(result.status==='PASS'?'PASS — Đạt':result.status==='FAIL'?`FAIL — ${result.errors.length} lỗi`:'CHƯA XÁC MINH',result.status==='FAIL'?'error':'');
 return result.status==='PASS'?[]:result.errors.length?result.errors.map(x=>x.id+': '+x.message):['INCOMPLETE: '+result.unverified.join('; ')];
}
function vietnameseFieldName(label){
 return String(label||'').split(/\s+\/\s+/)[0]
   .replace(/\s*\([^)]*[a-zA-Z][^)]*\)/g,'')
   .replace(/\s*·\s*/g,' — ').replace(/\s{2,}/g,' ').trim();
}
function renderSourceCheck(result){
 const top=$('check-top'),wrap=$('check-issues');
 const failed=result.status==='FAIL',incomplete=result.status==='INCOMPLETE';
 if(top){top.textContent=failed?'Không đạt':incomplete?'Chưa xác minh':'Đạt';
  top.className='check-top '+(failed?'fail':incomplete?'partial':'source-pass')}
 if(!wrap)return;
 wrap.replaceChildren();

 if(!failed){
  // When the Version02 mass table is incomplete, show a direct link at D57
  // without pretending this is an incorrect user input.
  if(incomplete && $('model').value==='MRL' && (result.unverified||[]).some(x=>/Decoration Weight Maximum NEXIEZ-MRL Version2/.test(String(x)))){
   const notice=document.createElement('button');notice.type='button';notice.className='check-error-row';
   notice.textContent='Chưa xác minh Maximum Decoration Weight (D57) — '+(result.unverified||[]).filter(x=>/Decoration Weight Maximum NEXIEZ-MRL Version2/.test(String(x))).map(x=>String(x).replace(/^Decoration Weight Maximum NEXIEZ-MRL Version2:\s*/, '')).join('; ');
   // Observation wall MethodDef 69 can be unsupported for the selected COP.
   // Link to the actual incompatible COP selection instead of only to D57.
   const relatedToCop=(result.unverified||[]).some(x=>/Decoration Weight Maximum NEXIEZ-MRL Version2:.*không có nhánh tính khối lượng vách.*COP/i.test(String(x)));
   notice.onclick=()=>focusIssue(relatedToCop?'D61':'D57');wrap.append(notice);
  }
  return;
 }
 const aggregates=/^(?:DataTable!D512|DataTable!D527|SpecForm!B210|SpecForm!B206|SpecForm!B208)$/;
 const primaries=result.errors.filter(x=>!aggregates.test(x.id));
 const detailedObservation=primaries.some(x=>String(x.id).startsWith('DataTable!D525 /'));
 const dimensionDetails=primaries.some(x=>/^DataTable!D52[23]\//.test(String(x.id)));
 const filtered=primaries.filter(x=>{
  if(detailedObservation&&x.id==='SpecForm!B209'&&/Observation/i.test(x.message))return false;
  if(dimensionDetails&&x.id==='SpecForm!B209'&&/OH\/PD|DataTable!D522|AH\/BH/i.test(x.message))return false;
  return true;
 });
 // Display errors of MRL Version02 from the bottom of SpecForm upwards.
 // This is UI ordering only: no input, shaft-size or validation formula changes.
 const issues=orderedCheckErrors(filtered.length?filtered:result.errors);
 const fields=new Map(),features=sourceFeatureRows();
 const sizeLabel={D33:'Chiều rộng hố thang (AH)',D34:'Chiều sâu hố thang (BH)',D36:'Chiều sâu hố PIT (PD)',D37:'Đỉnh hố thang (OH)',D40:'Chiều sâu hố PIT (PD)',D41:'Đỉnh hố thang (OH)',D35:'Chiều sâu hố thang (BH)'};
 const isDimension=cell=>$('model').value==='MRL'?['D33','D34','D36','D37'].includes(cell):['D34','D35','D40','D41'].includes(cell);
 const dimensionPriority=err=>/^(?:DataTable!D52[23]\/|DataTable!AP2[45]$)/.test(String(err.id))?3:err.id==='Tool.exe CheckSpecStep1'?2:1;
 for(const err of issues){
  const cell=String(err.cell||'').trim();if(!cell)continue;
  const f=current.fields.find(x=>x.cell===cell),feature=features.find(x=>x.cell===cell);
  let name=vietnameseFieldName(xrules().labels?.[cell]||f?.label||feature?.vi||feature?.code||'Thông số kỹ thuật');
  if(isDimension(cell))name=($('model').value==='MR'&&cell==='D34'?'Chiều rộng hố thang (AH)':sizeLabel[cell])||name;
  if(String(err.id).startsWith('JD!'))name='Chiều rộng khung bao bản rộng (JD) — '+(cell==='JD_MAIN'?'Tầng chính':'Tầng khác');
  if(String(err.id)==='ITP!Required')name='Vui lòng chọn loại ITP';
  if(String(err.id).startsWith('ITP!')&&String(err.id)!=='ITP!Required')name=String(err.message||'Loại ITP không hợp lệ');
  if(!name||name===cell)name='Thông số kỹ thuật';
  // NSCR-C/C760: explain required EVRC choice, instead of merely showing an
  // unlabeled failed NSCR-C row. Clicking the error opens Operation at EVRC-C.
  const message=(isDimension(cell)||String(err.id)==='ALPHA!TwoStopsTravel'||String(err.id).startsWith('MRL_V02_DECOR!')||String(err.id).startsWith('NY40!')||String(err.id)==='FloorName!AAN-G'||String(err.id).startsWith('FloorName!Display!')||String(err.id)==='Feature!HE-B-Dependencies'||(String(err.id)==='MR!SideDropBHMax'||String(err.id)==='MRL!Standard1G_BHMax'||String(err.id)==='MRL!2G_BH_Exact')||String(err.id)==='Feature!MR_COP_OLHL_EXCLC'||String(err.id).startsWith('LandingDoor!WindowFireproof!')||String(err.id).startsWith('JD!')||($('model').value==='MR'&&String(err.id)==='DataTable!D494'))?String(err.message||'').trim():'';
  const priority=(String(err.id)==='MR!SideDropBHMax'||String(err.id)==='MRL!Standard1G_BHMax'||String(err.id)==='MRL!2G_BH_Exact')?10:(isDimension(cell)?dimensionPriority(err):0);
  const old=fields.get(cell);
  if(!old||priority>old.priority)fields.set(cell,{id:String(err.id),name,message,priority});
 }
 if(top&&failed)top.textContent=`Không đạt — ${fields.size||1} lỗi`;
 if(!fields.size){const row=document.createElement('div');row.className='check-error-row';row.textContent='Có lỗi kỹ thuật cần kiểm tra';wrap.append(row);return}
 let index=0;
 for(const [cell,item] of fields){
  const link=document.createElement('button');link.type='button';link.className='check-error-row';
  const heading=document.createElement('strong');heading.textContent=`${++index}. ${item.name}`;link.append(heading);
  if(item.message){const explanation=document.createElement('span');explanation.className='check-error-explanation';
   let detail=item.message;
   // Avoid repeating the same field name in title and explanation, without hiding the limit.
   detail=detail.replace(/^(?:Đỉnh hố thang \(OH\)|Chiều sâu hố PIT \(PD\)|Chiều rộng hố thang \(AH\)|Chiều sâu hố thang \(BH\))\s*/,'');
   explanation.textContent=detail;link.append(explanation);
  }
  // A limited allowable decoration weight can show as D57 (Decoration Weight)
  // or D56 (Flooring Weight exceeding the same Technical Study allowance).
  // Show the same counterweight suggestion for either case in MRL Version2.
  if($('model').value==='MRL' && (
      item.id==='MRL_V02_DECOR!Maximum' ||
      item.id==='MRL_V02_DECOR!FlooringOverMaximum' ||
      (cell==='D57' && /[Dd]ecoration [Ww]eight|khối lượng trang trí/i.test(item.message)) ||
      (cell==='D56' && /Flooring Weight.*(?:vượt Maximum|exceeds)/i.test(item.message))
     )){
   const note=document.createElement('span');note.className='check-error-explanation';
   note.textContent='Bạn có thể chuyển qua đối trọng 620E hoặc 720G để tăng khối lượng trang trí (lưu ý: AH có thể sẽ tăng).';
   link.append(note);
  }
  if((($('model').value==='MR'&&cell==='D35')||($('model').value==='MRL'&&cell==='D34'))&&item.priority===10&&item.id!=='MRL!2G_BH_Exact'){
   const img=document.createElement('img');img.src='assets/mr_side_drop_bh_safety_5474.png';
   img.alt='Minh họa lưới/vách bảo vệ phía sau, lan can tối thiểu 700 mm và gờ chắn chân tối thiểu 100 mm';
   img.className='bh-safety-illustration';img.loading='lazy';link.append(img);
  }
  link.setAttribute('aria-label','Đến thông số lỗi: '+item.name+(item.message?' — '+item.message:''));
  link.onclick=()=>focusIssue(cell);
  wrap.append(link);
 }
}

function download(bytes,name,type){
 // Keep the exact XLSX bytes produced by the original v5.8.1 export engine.
 // iOS has different download/Share behavior; never convert XLSX to a data URL.
 const blob=new Blob([bytes],{type});
 const url=URL.createObjectURL(blob);
 const isIOS=/iPad|iPhone|iPod/.test(navigator.userAgent)||
   (navigator.platform==='MacIntel'&&navigator.maxTouchPoints>1);
 if(isIOS){
   // A fresh tap is required for Web Share on iOS after an async XLSX build.
   document.getElementById('ios-xlsx-save-dialog')?.remove();
   const panel=document.createElement('div');panel.id='ios-xlsx-save-dialog';
   panel.setAttribute('role','dialog');panel.setAttribute('aria-label','Lưu Excel trên iOS');
   panel.style.cssText='position:fixed;inset:0;z-index:2147483647;background:#0008;display:flex;align-items:center;justify-content:center;padding:20px;font-family:Arial,sans-serif';
   const card=document.createElement('div');card.style.cssText='background:#fff;color:#172b4d;border-radius:12px;max-width:400px;width:100%;padding:22px;box-shadow:0 12px 48px #0005';
   const msg=document.createElement('p');msg.textContent='File Excel đã tạo xong. Chọn Chia sẻ / Lưu vào Tệp để lưu đúng định dạng .xlsx.';
   msg.style.cssText='line-height:1.5;margin:0 0 16px';
   const save=document.createElement('button');save.type='button';save.textContent='Chia sẻ / Lưu vào Tệp';
   save.style.cssText='width:100%;border:0;background:#1764a5;color:white;padding:12px;border-radius:8px;font-size:16px';
   const fallback=document.createElement('a');fallback.href=url;fallback.download=name;
   fallback.textContent='Tải Excel theo cách thông thường';fallback.style.cssText='display:block;text-align:center;margin-top:15px;color:#1764a5';
   const close=document.createElement('button');close.type='button';close.textContent='Đóng';
   close.style.cssText='width:100%;margin-top:12px;border:0;background:#e5e7eb;padding:10px;border-radius:8px';
   const dispose=()=>{panel.remove();setTimeout(()=>URL.revokeObjectURL(url),120000);};
   close.onclick=dispose;
   save.onclick=async()=>{
     try{
       const file=new File([blob],name,{type,lastModified:Date.now()});
       if(!navigator.share||!navigator.canShare?.({files:[file]}))throw Error('Thiết bị không hỗ trợ chia sẻ file');
       await navigator.share({files:[file],title:name});dispose();
     }catch(err){if(err?.name!=='AbortError')alert('Không mở được Share Sheet. Hãy chọn Tải Excel theo cách thông thường.');}
   };
   card.append(msg,save,fallback,close);panel.append(card);document.body.append(panel);
   return;
 }
 const a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();
 setTimeout(()=>URL.revokeObjectURL(url),120000);
}
async function buildFinalExport(){
  const model=$('model').value,db=workbookEngineData();
  // Source validation MUST pass before any export side effect.
  const evaluation=window.SPEC_EXPORT.exportModel(model,changes,db);
  // Optional presentation-only detail rows; defaults OFF for both models.
  evaluation.includeTechDetails=Boolean($('export-tech-details')?.checked);
  window.SPEC_EXPORT.attachMRMaterialThickness?.(model,changes,db,evaluation);
  const bytes=sourceFiles[model]?.bytes||bytesFromB64(window.SPEC_WORKBOOKS[model]);
  const zip=await JSZip.loadAsync(bytes);
  const packageResult=await window.SPEC_EXPORT.xlsxFromModel(zip,model,db,evaluation);
  const data=await zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
  const project=String(evaluation.cells.B8||'Project').replace(/[^A-Za-z0-9_-]+/g,'_').slice(0,35);
  return {model,evaluation,data,filename:'Specification_'+model+'_'+project};
}
function canExportBriefly(){
 try{
  const model=$('model').value;
  // Only MRL Version02 requires the decoration limit to be checked in addition
  // to the existing source check. The same validation is used on the Check Spec
  // button and before Excel/PDF export, so an over-limit input cannot bypass it.
  if(model==='MRL'){
   check();
   if(window.__lastCheck?.status==='PASS')return true;
   alert('Spec chưa đạt hoặc Maximum Decoration Weight chưa được xác minh. Vui lòng xem Check Spec.');
   return false;
  }
  const result=window.SPEC_SOURCE_CHECK?.(model,changes,workbookEngineData());
  if(result?.status==='PASS')return true;
  renderSourceCheck(result||{status:'INCOMPLETE',errors:[]});
  alert('Spec có lỗi. Vui lòng sửa lại spec.');
  return false;
 }catch(error){console.error('Kiểm tra trước Export:',error);alert('Spec có lỗi. Vui lòng sửa lại spec.');return false;}
}
async function exportFinalExcel(){
 if(!canExportBriefly())return;
 const button=$('final-xlsx');button.disabled=true;
 try{status('Đang tạo SpecFormExport từ logic Tool.exe / Building final export...');
  const result=await buildFinalExport();
  download(result.data,result.filename+'.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  status('Đã xuất Excel SpecFormExport theo source C# / Excel ready. '+result.evaluation.includedFeatures.length+' features applied; '+result.evaluation.removedFeatures.length+' removed; '+result.evaluation.hiddenHall+' hall rows hidden; '+result.evaluation.totalFormulaUnresolved+' formula(s) excluded in unused feature rows.','ok');
 }catch(error){console.error('Excel Export:',error);alert('Không thể xuất Excel. Chi tiết: '+(error?.message||String(error)))}
 finally{button.disabled=false;}
}
async function exportFinalPdf(){
 if(!canExportBriefly())return;
 const button=$('final-pdf');button.disabled=true;
 // Create window in direct click handler to avoid browser popup blockers during async processing.
 const popup=window.open('','_blank');
 if(popup){popup.document.write('<!doctype html><html><meta charset="utf-8"><title>Creating A4 Preview</title><body>Đang tạo A4 PDF / Preparing A4 layout...</body></html>');popup.document.close();}
 try{
  const result=await buildFinalExport();
  const html=await window.SPEC_PRINT_LAYOUT.fromXlsx(result.data,result.model);
  if(!popup)throw Error('Trình duyệt đã chặn tab mới. Hãy cho phép mở cửa sổ / Enable popups.');
  popup.document.open();popup.document.write(html);popup.document.close();
  popup.focus();
  status('Đã mở bản in A4 dọc theo bố cục SpecFormExport. Chọn In / Save as PDF để lưu tài liệu.','ok');
 }catch(error){
   console.error('PDF Export:',error);
   const detail=String(error?.message||error||'Unknown error');
   if(popup&&!popup.closed){
     try{
       popup.document.open();popup.document.write('<!doctype html><html lang="vi"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><body style="font:16px Arial;padding:22px;line-height:1.5"><h3>Không tạo được bản PDF / PDF export failed</h3><pre style="white-space:pre-wrap;overflow-wrap:anywhere"></pre><button onclick="window.close()">Đóng / Close</button></body></html>');popup.document.close();
       popup.document.querySelector('pre').textContent=detail;
     }catch(_e){}
   }
   alert('Không thể xuất PDF. Chi tiết: '+detail);
 }
 finally{button.disabled=false;}
}
async function exportWorking(){
 if(runCheckSpec().length)return;
 try{
  const model=$('model').value,bytes=sourceFiles[model]?.bytes||bytesFromB64(window.SPEC_WORKBOOKS[model]);
  const zip=await JSZip.loadAsync(bytes);
  let workbook=X(await zip.file('xl/workbook.xml').async('string'));
  let rels=X(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
  let rid=each(workbook,'sheet').find(s=>s.getAttribute('name')==='SpecForm')?.getAttributeNS(NSREL,'id');
  let target=Array.from(rels.documentElement.children).find(x=>x.getAttribute('Id')===rid)?.getAttribute('Target');
  if(!target)throw Error('Không tìm thấy SpecForm in workbook relationships.');
  const path=getSheetPath(target);let xml=X(await zip.file(path).async('string'));
  let data=first(xml,'sheetData');let cells=new Map(each(data,'c').map(c=>[c.getAttribute('r'),c]));
  for(let [addr,value] of Object.entries(changes)){
   let c=cells.get(addr);if(!c)continue;
   for(let n of Array.from(c.childNodes)){if(['v','f','is'].includes(n.localName))c.removeChild(n)}
   if(value.trim()!==''&&!isNaN(Number(value))){c.removeAttribute('t');let v=xml.createElementNS(NS,'v');v.textContent=value;c.append(v)}
   else {c.setAttribute('t','inlineStr');let is=xml.createElementNS(NS,'is'),t=xml.createElementNS(NS,'t');t.textContent=value;is.append(t);c.append(is)}
  }
  zip.file(path,new XMLSerializer().serializeToString(xml));
  // Preserve the additional MRL Hall Lantern selections when exporting an editable source workbook.
  if(model==='MRL'&&window.SPEC_HALL_LANTERN){
   const hallSheet=each(workbook,'sheet').find(x=>x.getAttribute('name')==='HallLantern_V02');
   const hrid=hallSheet?.getAttributeNS(NSREL,'id');
   const target2=Array.from(rels.documentElement.children).find(x=>x.getAttribute('Id')===hrid)?.getAttribute('Target');
   const hallPath=target2?getSheetPath(target2):null;
   if(hallPath&&zip.file(hallPath)){
    const hxml=X(await zip.file(hallPath).async('string'));
    const cells2=new Map(each(hxml,'c').map(c=>[c.getAttribute('r'),c]));
    for(const [key,addr] of [['HL_TOP','B21'],['HL_OTHER','B22'],['HL_MAIN','B23']]){
     let c=cells2.get(addr);if(!c)continue;
     for(const child of Array.from(c.childNodes))if(['v','f','is'].includes(child.localName))c.removeChild(child);
     c.setAttribute('t','inlineStr');
     const is=hxml.createElementNS(NS,'is'),t=hxml.createElementNS(NS,'t');
     t.textContent=window.SPEC_HALL_LANTERN.selection(changes,key);is.append(t);c.append(is);
    }
    zip.file(hallPath,new XMLSerializer().serializeToString(hxml));
   }
  }
  const output=await zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
  download(output,'SpecForm_Working_'+model+'.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  status('Đã tạo bản Working Excel: các ô đã chỉnh sửa được ghi vào SpecForm, các sheet khác giữ nguyên.\nĐây KHÔNG phải kết quả Export Spec chính thức của Tool.exe.','ok');
 }catch(err){status('Lỗi xuất working Excel: '+err.message,'error')}
}
async function applyWorkbook(model,name,bytes,quiet=false){
 const parsed=await parseXlsx(bytes);
 const title=String(parsed.engineData?.sheets?.SpecForm?.B5||'').toUpperCase();
 const validModel=model==='MRL'?title.includes('NEXIEZ-MRL'):title.includes('NEXIEZ-MR')&&!title.includes('NEXIEZ-MRL');
 if(!validModel)throw Error('Workbook không đúng model '+model+' / Wrong model database: '+title.slice(0,90));
 if(!parsed.engineData?.tables?.TblCondition1?.columns?.includes('Condition'))throw Error('Thiếu bảng TblCondition1[Condition] / Missing Check Spec source table.');
 for(const x of (window.SPEC_EXTRAS[model]||[])){
  if(parsed.fields.some(f=>f.cell===x.cell))continue;
  const raw=parsed.engineData?.sheets?.SpecForm?.[x.cell];
  parsed.fields.push({...x,value:raw&&typeof raw==='object'?(raw.v??x.value):(raw??x.value)});
 }
 const fire=parsed.fields.find(f=>f.name==='K36_FireproofDoor');
 if(fire){fire.options=['Not supplied','GHCL E30','EN81-58(II)'];fire.validation='"Not supplied,GHCL E30,EN81-58(II)"';}
 window.SPEC_MATERIAL_5825?.patchFields(model,parsed.fields);
 window.SPEC_MATERIAL_5825?.patchDatabase(model==='MRL'?parsed.engineData:null);
 parsed.fields.sort((a,b)=>split(a.cell)[1]-split(b.cell)[1]);
 // Apply the versioned CWT source table and ceiling-finish checks to imported old databases.
 window.SPEC_CWT_ENGINE?.applyDataset(parsed.engineData,model);
 if(model==='MR')window.SPEC_MR_BED_5871?.patchEmbedded(parsed);
 models[model]=parsed;sourceFiles[model]={name,bytes,engineData:parsed.engineData};
 if($('model').value===model){load();if(!quiet)status('Đã nạp Excel Database / Loaded: '+name+'\n'+parsed.fields.length+' trường SpecForm; '+parsed.validations.length+' Data Validations; '+parsed.table_count+' DataTable tables.','ok')}
 return parsed;
}
const remoteLoaded=new Set();
async function loadRepositoryDatabase(model,quiet=false){
 const protocol=window.location?.protocol||'';
 if(protocol!=='https:'&&protocol!=='http:'){
  if(!quiet)$('file').click();
  return false;
 }
 const path='database/'+modelFiles[model];
 const response=await fetch(path+'?fresh='+Date.now(),{cache:'no-store'});
 if(!response.ok)throw Error('Không tải được '+path+' / HTTP '+response.status);
 const bytes=new Uint8Array(await response.arrayBuffer());
 await applyWorkbook(model,path,bytes,quiet);remoteLoaded.add(model);
 return true;
}
$('model').onchange=()=>{load();/* v5.8.77: same embedded source for offline and GitHub Pages; workbook loading is manual only. */};
$('reload-db').onclick=()=>{const model=$('model').value;if(Object.keys(changes).length&&!confirm('Nạp Database sẽ thay các giá trị đang chỉnh sửa. Tiếp tục? / Reload discards unsaved edits. Continue?'))return;
 loadRepositoryDatabase(model,false).catch(e=>status('Lỗi Reload Database / Reload failed: '+e.message,'error'));
};

$('check').addEventListener('click',runCheckSpec);
$('working').onclick=exportWorking;
$('final-xlsx').onclick=exportFinalExcel;
$('final-pdf').onclick=exportFinalPdf;
$('reset').onclick=()=>{models[$('model').value]=structuredClone(window.SPEC_EMBEDDED[$('model').value]);models[$('model').value].fields.push(...structuredClone(window.SPEC_EXTRAS[$('model').value]));delete sourceFiles[$('model').value];$('file').value='';load()};
$('file').addEventListener('change',async(e)=>{const file=e.target.files[0];if(!file)return;
 try{status('Đang đọc Excel / Reading workbook...');const model=$('model').value;
 const bytes=new Uint8Array(await file.arrayBuffer());await applyWorkbook(model,file.name,bytes);
 }catch(err){status('Không nạp được Excel / Load failed: '+err.message,'error')}
});
$('ny40-close')?.addEventListener('click',closeNy40Popup);
$('ny40-ok')?.addEventListener('click',closeNy40Popup);
$('ny40-apply')?.addEventListener('click',()=>{
 const input=activeInput('D14');
 if($('model').value!=='MRL'||!input||!allowedObservationCodes().includes(ny40Active))return;
 input.value=ny40Active;
 input.dispatchEvent(new Event('change',{bubbles:true}));
 closeNy40Popup();
});
$('ny40-popup')?.addEventListener('click',e=>{if(e.target.id==='ny40-popup')closeNy40Popup()});
window.addEventListener('keydown',e=>{if(e.key==='Escape')closeNy40Popup()});
load();
// v5.8.77: Avoid silent database replacement on GitHub Pages.
// Use the embedded workbook in both file:// and https:// contexts.
// Reload Database remains available only through an explicit user action.
