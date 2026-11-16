/* v5.8.71 – AES13-006 Rev.AM (2026-02-13), Tables 5,7,8,11.
 * Hospital bed models are independent of passenger models with same kg.
 * NOTE: 220kg B1000 decoration requires Table 11-A prerequisites; use the conservative
 * 120kg baseline until all prerequisites have been verified from the source workbook.
 */
(function(){'use strict';
 const cases=[
 {code:'B750',kg:750,persons:11,aa:1300,bb:2300,jj:1100,ah:2070,bh:2730,ahmax:2762,bhmax:3357,decor60:210,decor90:190,maxTravel:80},
 {code:'B1000',kg:1000,persons:15,aa:1500,bb:2500,jj:1200,ah:2270,bh:2930,ahmax:2962,bhmax:3557,decor60:120,decor90:120,maxTravel:60}
 ];
 const asName=x=>`${x.code}-2S (${x.kg} KG)-CWT Side`;
 const rows=cases.map(x=>({
  Capacity:x.kg,EleCapacityName:asName(x),EleCapacityNameEN:`Hospital bed ${x.kg} KG / ${x.persons} persons`,DoorType:'2S',AA:x.aa,BB:x.bb,JJ:x.jj,'CWT-Position':'Side',
  'AH-CWT1':x.ah,'BH-CWT1':x.bh,'AH-CWT1-FIRE':x.ah,'BH-CWT1-FIRE':x.bh,'AH-Max':x.ahmax,'BH-Max':x.bhmax,'AH-Max-FIRE':x.ahmax,'BH-Max-FIRE':x.bhmax,
  'AH-Valid':'N-Min','AH-Valid-Value':x.ah,'BH-Valid':'N-Min','BH-Valid-Value':x.bh,
  'DecorWeight-CWT1-S00-60MPM':x.decor60,'DecorWeight-CWT1-S00-90MPM':x.decor90,
  'DecorWeight-CWT1-L210-210S-60MPM':null,'DecorWeight-CWT1-L210-210S-90MPM':null,
  'DecorWeight-CWT1-S300-N300-N300S-L310-N120-N130-N140-60MPM':null,
  'DecorWeight-CWT1-S300-N300-N300S-L310-N120-N130-N140-90MPM':null,
  'CheckValid-DecorWeight':x.decor90,'IncWeight-Wall-SUS':0,'IncWeight-Wall-PPS':0,'IncWeight-YZ-52A':0,'IncWeight-NoHandrail':0,
  AM:x.ah,BM:x.bh
 }));
 const toCol=n=>{let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s;};
 function patchEmbedded(model){
  if(!model||!model.capacity)return;
  for(const row of rows){if(!model.capacity.some(x=>x.EleCapacityName===row.EleCapacityName))model.capacity.push({...row});}
  const field=model.fields?.find(f=>f.name==='K5_Capacity');
  if(field){field.options=field.options||[];for(const r of rows)if(!field.options.includes(r.EleCapacityName))field.options.push(r.EleCapacityName);}
 }
 function patchDatabase(db){
  if(!db?.tables?.TblK5_Capacity||!db.sheets?.DataTable)return;
  const t=db.tables.TblK5_Capacity,sh=db.sheets.DataTable;
  const m=String(t.ref).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);if(!m)return;
  const originalEnd=Number(m[4]);let end=originalEnd;
  for(const row of rows){
   let existing=false;
   for(let r=Number(m[2])+1;r<=end;r++)if(String(sh['B'+r]||'')===row.EleCapacityName){existing=true;break;}
   if(existing)continue;
   end++;for(let i=0;i<t.columns.length;i++)sh[toCol(i+1)+end]=row[t.columns[i]]??'';
  }
  if(end>originalEnd){
   t.ref=`${m[1]}${m[2]}:${m[3]}${end}`;
   // Existing native workbook lookups may use a fixed row-56 range.
   // Extend only explicit DataTable capacity lookup blocks, not other tables.
   for(const cells of Object.values(db.sheets))for(const entry of Object.values(cells)){
    if(!entry||typeof entry!=='object'||typeof entry.f!=='string')continue;
    entry.f=entry.f.replace(/(DataTable!?\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?)56\b/g,`$1${end}`);
   }
  }
 }
 function validate(get){
  const c=String(get('D13')??'');const bed=cases.find(x=>c.includes(x.code+'-2S'));if(!bed)return [];
  const issues=[],add=(id,cell,message)=>issues.push({id:'MR_BED_'+id,cell,message});
  const speed=Number(get('D15')); // MR workbook: speed cell D15
  if(Number.isFinite(speed)&&speed!==0&&![60,90,105].includes(speed))add('SPEED','D15','Bed elevator: tốc độ phải là 60/90/105 m/min theo cấu hình được hỗ trợ.');
  const travel=Number(get('D39'));
  const limit=bed.maxTravel*1000;
  if(Number.isFinite(travel)&&travel>limit)add('TRAVEL','D39',`${bed.code}-2S: Travel không vượt ${bed.maxTravel} m theo AES13-006 Rev.AM, Table 8.`);
  const ceiling=String(get('D45')??'').trim();
  if(ceiling&&ceiling!=='S00')add('CEILING','D45',`${bed.code}-2S chỉ áp dụng ceiling S00 theo Table 11 Note *4.`);
  const door=String(get('D14')??'').trim();
  if(door&&door!=='2S')add('DOOR','D14',`${bed.code} bắt buộc dùng cửa 2S.`);
  return issues;
 }
 window.SPEC_MR_BED_5871={cases,rows,patchEmbedded,patchDatabase,validate};
 patchEmbedded(window.SPEC_EMBEDDED?.MR);
 patchDatabase(window.SPEC_ENGINE_DATABASE?.MR);
})();
