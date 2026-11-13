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
function status(message,cls=''){$('status').className=cls;$('status').textContent=message}
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
function xrules(){return window.SPEC_RULES?.[$('model').value]||{labels:{},features:[],sideDrop:[]}}
// Add missing MR 2C-2BC ITP variants to the in-use source database, including when
// customers import an older original SpecTool Excel. This leaves all existing rows intact.
function workbookEngineData(){
 const model=$('model').value;
 const db=sourceFiles[model]?.engineData||window.SPEC_ENGINE_DATABASE?.[model];
 if(model==='MRL')window.SPEC_MATERIAL_5825?.patchDatabase(db);
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
function renderFeatures(){
 const features=sourceFeatureRows();ensureFreeFeatureDefaults(features);const tables=[...new Set(features.map(f=>f.table))];const tabs=$('feature-tabs'),main=$('features');if(!tabs||!main)return;tabs.replaceChildren();main.replaceChildren();ensureITPDefault(features);
 if(!tables.length){main.textContent='Chưa trích xuất được bảng Features / Features unavailable';return}
 if(!tables.includes(activeFeatureTable))activeFeatureTable=tables[0];
 tables.forEach(t=>{let b=document.createElement('button');b.className='btn'+(t===activeFeatureTable?' red':'');b.textContent=({TblFeature_Emergency:'Khẩn cấp / Emergency',TblFeature_Operation:'Vận hành / Operation',TblFeature_Door:'Cửa / Door',TblFeature_Signal:'Tín hiệu / Signal',TblFeature_Traffic:'Giao thông / Traffic'})[t]||t;b.onclick=()=>{activeFeatureTable=t;renderFeatures()};tabs.append(b)});
 const db=workbookEngineData(),featureEngine=db&&window.SpecExcelSourceEngine?new window.SpecExcelSourceEngine.SourceEngine(db,changes):null;
 for(const f of features.filter(f=>f.table===activeFeatureTable)){
  const line=document.createElement('div');line.className='feature-line';let addr=document.createElement('span');addr.textContent=f.cell;addr.className='addr';let code=document.createElement('span');code.textContent=f.code;code.className='fcode';let desc=document.createElement('span');desc.className='fdesc';desc.textContent=f.vi||f.code;let en=document.createElement('small');en.textContent=f.en||'';desc.append(en);
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
  chk.disabled=disabled;if(disabled)line.classList.add('x');chk.onchange=()=>{
   const code=window.SPEC_V52_FEATURES?.canonical(f.code)||f.code;
   if($('model').value==='MRL'&&code==='FHL'&&!chk.checked&&window.SPEC_HALL_LANTERN?.any(changes)){
     chk.checked=true;changes[f.cell]='1';status('FHL phải đi kèm Hall Lantern đã chọn / FHL is required while lanterns are selected.');return;
   }
   featureTouched.add(f.cell);changes[f.cell]=chk.checked?'1':'0';
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
   if(['FE','FE-C30','ITP','ITV','AECC','AECH','AIL'].includes(code)){
     if(code==='ITP'&&!chk.checked)delete changes[itpCell()];renderFeatures();
   }
   invalidateCheck();status('Đã cập nhật Feature / Feature updated. Vui lòng chạy lại Check Spec.');
  };
  let op=document.createElement('span');op.className='option';op.textContent=dynamicOption;let suggest=document.createElement('span');suggest.className='suggest';suggest.textContent=f.suggestion;line.append(addr,code,desc,chk,op,suggest);if(f.code==='ITP'&&chk.checked&&!chk.disabled)renderITPOptions(line);main.append(line)
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
 return [...new Set(list.map(v=>String(v??'').trim()).filter(Boolean))];
}
function hasDropdown(field){return Boolean(field.validation||supplementalListSources[field.name]||(field.options||[]).length||field.cell==='D23'&&$('model').value==='MRL'||field.name==='K18_CarJJ'&&$('model').value==='MRL')}
function syncDropdownOptions(){
 const workbook=workbookEngineData();if(!workbook||!window.SpecExcelSourceEngine)return;
 const engine=new window.SpecExcelSourceEngine.SourceEngine(workbook,changes);
 for(const f of current.fields){
  const el=activeInput(f.cell);if(!el||el.tagName!=='SELECT')continue;
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
function guideEntries(kind){return window.SPEC_DESIGN_GUIDE?.[$('model').value]?.[kind==='ceilingColor'?'ceiling':kind]||[]}
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
 holder.append(img,note,button);box.append(holder);
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
 $('design-title').textContent=label[kind]+' · '+($('model').value==='MR'?'NEXIEZ-MR':'NEXIEZ-MRL Version02');
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

function row(field,displayName='',compact=false){
 const box=document.createElement('div');box.className='field';
 box.dataset.sourceCell=field.cell;
 // Source identifiers remain as data-cell attributes for Check Spec, focus and Export.
 const name=document.createElement('label');name.className='nm';name.textContent=displayName||xrules().labels[field.cell]||field.label||field.name.replace(/^K\d+_/,'').replaceAll('_',' ');
 const choices=workbookChoices(field);const input=hasDropdown(field)?document.createElement('select'):document.createElement('input');input.dataset.cell=field.cell;input.id='spec-'+field.cell;
 name.htmlFor=input.id;
 let value=String(field.value??'');
 if(input.tagName==='SELECT'){
  const options=[...choices.map(String)];if(value&&!choices.includes(value)&&!field.validation&&!supplementalListSources[field.name])options.unshift(value);
  if(!options.length)options.push('');
  for(const option of [...new Set(options)]){const opt=document.createElement('option');opt.value=option;opt.textContent=option||'— Chọn / Select —';input.append(opt)}
 }else input.placeholder='';
 input.value=value;
 input.addEventListener('change',()=>{
   changes[field.cell]=input.value;input.classList.add('dirty');
   if(field.name==='K5_Capacity')capacityUpdate(input.value);
   if(field.name==='K28_FlooringWeight')flooringManual=true;
   else flooringDependencyChanged(field.cell);
   if(field.name==='K5_Capacity' || field.name==='K17_CarW')refreshMirrorWeightHint();
   // A new ceiling model immediately changes its finish choices, instead of
   // falsely flagging the old paint color when switching to SUS-HL.
   if(field.name==='K22_CeilingModel'&&window.SPEC_CEILING_FINISH){
    const model=$('model').value,fin=current.fields.find(x=>x.name==='K22_CeilingColor');
    if(fin){const chosen=window.SPEC_CEILING_FINISH.next(model,input.value,currentValue(fin.cell));
     if(String(currentValue(fin.cell))!==chosen)applyValue(fin.cell,chosen);}
   }
   if($('model').value==='MRL'&&field.cell==='D14'&&window.SPEC_NY40_RULES){
     const allowed=window.SPEC_NY40_RULES.allowedLocations(input.value);
     if(allowed&&!allowed.includes(String(currentValue('D65'))))applyValue('D65',allowed[0]);
   }
   syncDropdownOptions();
   refreshSideDrop();
   refreshLocks();
   refreshDesignHints();
   updateFinishPreviews();
   refreshMirrorWeightHint();
   if($('model').value==='MRL'){
    refreshMRLOverheadMinimum(['K5_Capacity','K8_Speed','K21_CarTR','K16_Regulation'].includes(field.name));
    if(['D13','D17','D31','D16','D84','D23'].includes(field.cell))refreshMRLShaftMinimum();
   }
   if(['K13_Operation','K5_Capacity','K8_Speed'].includes(field.name))renderFeatures();
   if($('model').value==='MRL'&&field.cell==='D14'){syncObservationButton(input);if(/^NY-40/i.test(String(input.value||'')))openNy40Popup(input.value);}
   invalidateCheck();status('Đã thay đổi thông số. Cần chạy lại Check Spec.');
 });
 if($('model').value==='MRL'&&field.cell==='D14'){
  const wrap=document.createElement('div');wrap.className='obs-inline';
  const btn=document.createElement('button');btn.type='button';btn.className='btn btn-small obs-preview-btn';btn.textContent='Xem mô tả';
  btn.onclick=()=>openNy40Popup(input.value);
  wrap.append(input,btn);box.append(name,wrap);syncObservationButton(input);
 }else box.append(name,input);
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
 const guideKind=designKindFor(field);if(guideKind)renderDesignInline(box,guideKind);
 if($('model').value==='MRL' && field.name==='K35_Ventilation'){
  const hint=document.createElement('small');hint.className='fan-guidance';
  hint.textContent='DC axial flow fan: Quạt hướng trục DC; DC axial flow fan + Plasma Quad: Quạt hướng trục DC kết hợp với diệt khuẩn Plasma.';
  box.append(hint);
 }
 return box;
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
 for(const [cell,source,label] of [['D36','F36','PD'],['D37','F37','OH']]){
  const el=activeInput(cell);if(!el)continue;
  const r=calc.safe('SpecForm',source),numeric=r.ok?Number(r.value):NaN;
  const minimum=typeof r.value==='number'&&Number.isFinite(numeric)&&numeric>0?numeric:null;
  let hint=el.closest('.field')?.querySelector('.oh-minimum-hint');
  if(!hint){hint=document.createElement('div');hint.className='oh-minimum-hint';hint.style.cssText='font-size:11px;color:#526a82;line-height:1.4;margin-top:2px';el.closest('.field')?.append(hint);}
  if(minimum===null){hint.textContent=`${label}: CHECK LAYOUT — không có giá trị cho tổ hợp Capacity / Speed / TR hiện tại.`;hint.style.color='#a32832';continue;}
  hint.textContent=`${label} tối thiểu theo DataTable: ${minimum} mm (có thể nhập lớn hơn)`;hint.style.color='#526a82';
  if(adjustIfBelow){const actual=Number(currentValue(cell));if(!Number.isFinite(actual)||actual<minimum){el.value=String(minimum);changes[cell]=String(minimum);el.classList.add('dirty');}}
 }
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
}
function allowedObservationCodes(){
 if($('model').value!=='MRL')return [];
 const db=workbookEngineData();
 if(db&&window.SpecExcelSourceEngine?.SourceEngine){
  const e=new window.SpecExcelSourceEngine.SourceEngine(db,changes);
  const gate=e.safe('SpecForm','R13');
  if(!gate.ok||gate.value!==true)return [];
 }
 const cap=String(currentValue('D13')),entrance=String(currentValue('D16'));
 if(entrance==='1D-1G')return Object.keys(NY40_POPUP_DATA);
 if(entrance==='1D-2G' && (/^P11-CO/.test(cap)||/^P13W-CO/.test(cap)) && String(currentValue('D60'))==='Front')return ['NY-40C'];
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
 const title=document.createElement('h3');title.textContent='Đèn báo tầng / Hall Lanterns — NEXIEZ-MRL Version02';
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
   const heading=document.createElement('h2');heading.className='section-heading';heading.textContent=headingText(major);
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
  const text=document.createElement('div');text.className='spec-row-title';text.textContent=headingText(label||lastMajor);
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
  {key:'OtherFloor',vi:'Tầng khác',en:'Other floors',cells:['E81','E82']},
  {key:'MainFloor',vi:'Tầng chính',en:'Main floor',cells:['E78','E79','E80']}
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
    item.append(label,wrapper);content.append(item);oldRow.remove();
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
 const r=$('model').value==='MRL'?calculateMRLLeadTime():calculateMRLeadTime();days.textContent=r.days+' days';reason.replaceChildren();
 const lines=r.reasons.length?r.reasons:[];
 for(const line of lines){const div=document.createElement('div');div.textContent=line;reason.append(div)}
}
// Re-evaluate after the existing event handler updates `changes`, including
// delegated controls and model/database reloads. No export pipeline changes.
document.addEventListener('change',()=>queueMicrotask(refreshLeadTime));

function load(){featureTouched=new Set();window.__lastCheck=null;if($('check-top')){$('check-top').textContent='Chưa kiểm tra';$('check-top').className='check-top';}if($('check-issues'))$('check-issues').replaceChildren();changes={};let model=$('model').value;current=models[model];restoreHallLanternFromSource();renderAllSpecRows();setupSpecialOptions();flooringManual=false;updateFlooringWeight(true);syncDropdownOptions();refreshSideDrop();refreshLocks();refreshDesignHints();updateFinishPreviews();refreshMirrorWeightHint();refreshMRLOverheadMinimum(false);refreshMRLShaftMinimum();renderFeatures();$('numvalid').textContent=current.validations.length;$('numtables').textContent=current.table_count;$('filename').textContent=sourceFiles[model]?.name||modelFiles[model]+' (embedded copy)';status('Đã hiển thị '+(current.layout||layouts[model]).length+' dòng có nội dung trong SpecForm của '+$('model').selectedOptions[0].text+'.\nCheck Spec: tính lại công thức DataTable. Export Excel/A4 theo dữ liệu SpecFormExport, chỉ khi Check Spec PASS.','');refreshLeadTime();}

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

function renderCheckIssues(result){
 const wrap=$('check-issues');if(!wrap)return;wrap.replaceChildren();
 const all=[...result.errors.map(x=>({raw:x,type:'error'})),...result.warnings.filter(x=>/^[A-Z]{1,3}\d+:/.test(x)).map(x=>({raw:x,type:'warning'}))];
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
function check(){
 const model=$('model').value;
 const engineData=workbookEngineData();
 const result=window.SPEC_SOURCE_CHECK?.(model,changes,engineData);
 if(!result)throw Error('Source Check Engine chưa được tải / Calculation engine not loaded');
 if(model==='MRL'&&window.SPEC_MATERIAL_5825){
  const r=window.SPEC_MATERIAL_5825;
  for(const f of current.fields){if(!/^(K22_CeilingColor|K23_WallFinish|K24_CarDoorFinish|K25_CarTransomFinish|K23_WallPaintedColor|K24_CarDoorPaintedColor|K25_CarTransomPainted|K26_CarFrontPainted|K23_WallPattern|K24_CarDoorPattern|K25_CarTransomPattern)$/.test(f.name))continue;
   const v=String(currentValue(f.cell)||'').trim();if(v&&!r.check(model,f,v))result.errors.push({id:'Material!'+f.name,cell:f.cell,message:'Mã vật liệu / màu không hợp lệ với MRL V02: '+v});
  }
  if(result.errors.length)result.status='FAIL';
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
 const addFloorLedWarning=()=>{
  if(!Array.isArray(result.dotLedFloorNames)||result.dotLedFloorNames.length===0)return;
  const notice=document.createElement('div');notice.className='check-led-warning';notice.setAttribute('role','note');
  notice.textContent='Tồn tại tên tầng cần phải sử dụng bảng gọi Dotled';
  notice.title='Tên tầng theo DataTable: '+result.dotLedFloorNames.join(', ');
  wrap.append(notice);
 };
 if(!failed){addFloorLedWarning();return;}
 const aggregates=/^(?:DataTable!D512|DataTable!D527|SpecForm!B210|SpecForm!B206|SpecForm!B208)$/;
 const primaries=result.errors.filter(x=>!aggregates.test(x.id));
 const detailedObservation=primaries.some(x=>String(x.id).startsWith('DataTable!D525 /'));
 const dimensionDetails=primaries.some(x=>/^DataTable!D52[23]\//.test(String(x.id)));
 const filtered=primaries.filter(x=>{
  if(detailedObservation&&x.id==='SpecForm!B209'&&/Observation/i.test(x.message))return false;
  if(dimensionDetails&&x.id==='SpecForm!B209'&&/OH\/PD|DataTable!D522|AH\/BH/i.test(x.message))return false;
  return true;
 });
 const issues=[...(filtered.length?filtered:result.errors)].sort((a,b)=>{
  const num=error=>{const cell=String(error.cell||'');const m=/\d+/.exec(cell);return m?Number(m[0]):9999;};
  return num(a)-num(b);
 });
 const fields=new Map(),features=sourceFeatureRows();
 const sizeLabel={D33:'Chiều rộng hố thang (AH)',D34:'Chiều sâu hố thang (BH)',D36:'Chiều sâu hố PIT (PD)',D37:'Đỉnh hố thang (OH)',D40:'Chiều sâu hố PIT (PD)',D41:'Đỉnh hố thang (OH)',D35:'Chiều sâu hố thang (BH)'};
 const isDimension=cell=>$('model').value==='MRL'?['D33','D34','D36','D37'].includes(cell):['D34','D35','D40','D41'].includes(cell);
 const dimensionPriority=err=>/^(?:DataTable!D52[23]\/|DataTable!AP2[45]$)/.test(String(err.id))?3:err.id==='Tool.exe CheckSpecStep1'?2:1;
 for(const err of issues){
  const cell=String(err.cell||'').trim();if(!cell)continue;
  const f=current.fields.find(x=>x.cell===cell),feature=features.find(x=>x.cell===cell);
  let name=vietnameseFieldName(xrules().labels?.[cell]||f?.label||feature?.vi||feature?.code||'Thông số kỹ thuật');
  if(isDimension(cell))name=sizeLabel[cell]||name;
  if(String(err.id)==='ITP!Required')name='Vui lòng chọn loại ITP';
  if(String(err.id).startsWith('ITP!')&&String(err.id)!=='ITP!Required')name=String(err.message||'Loại ITP không hợp lệ');
  if(!name||name===cell)name='Thông số kỹ thuật';
  // NSCR-C/C760: explain required EVRC choice, instead of merely showing an
  // unlabeled failed NSCR-C row. Clicking the error opens Operation at EVRC-C.
  const message=(isDimension(cell)||($('model').value==='MR'&&String(err.id)==='DataTable!D494'))?String(err.message||'').trim():'';
  const priority=isDimension(cell)?dimensionPriority(err):0;
  const old=fields.get(cell);
  if(!old||priority>old.priority)fields.set(cell,{name,message,priority});
 }
 if(top&&failed)top.textContent=`Không đạt — ${fields.size||1} lỗi`;
 if(!fields.size){const row=document.createElement('div');row.className='check-error-row';row.textContent='Có lỗi kỹ thuật cần kiểm tra';wrap.append(row);addFloorLedWarning();return}
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
  link.setAttribute('aria-label','Đến thông số lỗi: '+item.name+(item.message?' — '+item.message:''));
  link.onclick=()=>focusIssue(cell);
  wrap.append(link);
 }
 addFloorLedWarning();
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
$('model').onchange=()=>{load();if(!remoteLoaded.has($('model').value)&&/^https?:$/.test(window.location.protocol))loadRepositoryDatabase($('model').value,true).catch(e=>status('Không thể nạp từ GitHub; đang dùng bản dữ liệu tích hợp: '+e.message,'error'))};
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
if(/^https?:$/.test(window.location.protocol))loadRepositoryDatabase($('model').value,true).catch(e=>status('Đã dùng Database tích hợp / Embedded fallback: '+e.message));
