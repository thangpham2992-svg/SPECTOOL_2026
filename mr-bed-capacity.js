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
  'CheckValid-DecorWeight':x.decor90,'IncWeight-Wall-SUS':x.code==='B750'?24:27,'IncWeight-Wall-PPS':x.code==='B750'?39:31,'IncWeight-YZ-52A':0,'IncWeight-NoHandrail':x.code==='B750'?39:44,
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
    // Legacy native AP10/AP15 use unqualified fixed $B$10:$B$56
    // (and matching output columns). Widen only capacity-row references;
    // otherwise the newly appended hospital rows remain invisible to Excel.
    if(/\$B\$10:\$B\$56/.test(entry.f))
      entry.f=entry.f.replace(/\$([A-Z]+)\$10:\$\1\$56\b/g,(_,c)=>`$${c}$10:$${c}$${end}`);
   }
  }
  // B1000 Table 11-A: AP15 must allow 220 kg only if *all* conditions
  // are satisfied. Keep the legacy source formula as the otherwise-branch.
  const ap15=sh.AP15;
  if(ap15&&typeof ap15==='object'&&typeof ap15.f==='string'&&!ap15.mrBedTable11A){
   const eligible='AND(LEFT(SpecForm!D13,6)="B1000-",SpecForm!D45="S00",SpecForm!D33=2100,SpecForm!D31=2200,SpecForm!D47="SUS-HL",SpecForm!D50="SUS-HL",SpecForm!D53="SUS-HL",SpecForm!D56="SUS-HL",SpecForm!D69="Not supplied",ISNUMBER(SEARCH("YH-56S",SpecForm!D67)),OR(SpecForm!D58="Aluminum",SpecForm!D58="SUS-HL"),SpecForm!D39>0,SpecForm!D39<=60000)';
   ap15.f='IF('+eligible+',220,'+ap15.f+')';
   ap15.mrBedTable11A=true;
  }
 }
 function validate(get){
  const c=String(get('D13')??'');const bed=cases.find(x=>c.startsWith(x.code+'-2S'));
  if(!bed)return [];
  const issues=[],add=(id,cell,message)=>issues.push({id:'MR_BED_'+id,cell,message});
  const speed=Number(get('D15')),travel=Number(get('D39'));
  if(speed&&![60,90,105].includes(speed))add('SPEED','D15','Hospital bed: chỉ cho phép tốc độ 60/90/105 m/min (Rev.AM Table 8).');
  const maxTravel=(speed===60?60:bed.maxTravel)*1000;
  if(travel>maxTravel)add('TRAVEL','D39',`${bed.code}: Travel tối đa ${maxTravel/1000} m ở tốc độ ${speed} m/min (Table 8 Note *6).`);
  const minTravel={60:3400,90:5700,105:7300}[speed];
  if(minTravel&&travel>0&&travel<minTravel)add('MIN_TRAVEL','D39',`${bed.code}: Travel nhỏ hơn ${minTravel} mm tối thiểu theo tốc độ (Table 8).`);
  const floors=Number(get('D16'))+Number(get('D18'));
  const stopsMax=speed===60?18:24;
  if(Number.isFinite(floors)&&floors>stopsMax)add('STOPS','D16',`${bed.code}: tổng số tầng, bao gồm tầng không phục vụ, tối đa ${stopsMax} ở tốc độ ${speed} m/min (Table 8).`);
  const handrail=String(get('D67')??'').trim();
  if(handrail && !['YH-56S (SUS304)','Not supplied'].includes(handrail))add('HANDRAIL','D67',`${bed.code}: Tay vịn YH-56S (SUS304) hoặc không cung cấp, theo Rev.AM Table 11-C.`);
  const ceiling=String(get('D45')??'').trim();
  if(ceiling&&ceiling!=='S00')add('CEILING','D45',`${bed.code}: chỉ áp dụng ceiling S00 (Table 11 Note *4).`);
  const door=String(get('D14')??'').trim();
  if(door&&door!=='2S')add('DOOR','D14',`${bed.code}: cửa cabin bắt buộc là 2S.`);
  const geometry=[['D29',bed.aa,'AA'],['D30',bed.bb,'BB'],['D32',bed.jj,'JJ']];
  for(const [cell,expected,label] of geometry){const actual=Number(get(cell));if(actual>0&&actual!==expected)add(label,cell,`${bed.code}: ${label} phải là ${expected} mm theo Table 7.`)}
  const ah=Number(get('D34')),bh=Number(get('D35'));
  if(ah>0&&(ah<bed.ah||ah>bed.ahmax))add('AH','D34',`${bed.code}: AH nằm trong ${bed.ah}–${bed.ahmax} mm (Table 7).`);
  if(bh>0&&(bh<bed.bh||bh>bed.bhmax))add('BH','D35',`${bed.code}: BH nằm trong ${bed.bh}–${bed.bhmax} mm (Table 7).`);
  const flooring=Number(get('D60'))||0,decor=Number(get('D61'))||0;
  // Table 11 Note 1: total allowable decoration includes flooring weight.
  // B1000, 220 kg only under the complete Table 11-A exact specification.
  let allowance=speed===60?bed.decor60:bed.decor90;
  const wall=String(get('D47')||'');
  if(wall==='SUS-HL')allowance+=bed.code==='B750'?24:27;
  else if(/^(?:Pattern-printed steel sheet PPS|PPS)$/.test(wall))allowance+=bed.code==='B750'?39:31;
  const railSelected=String(get('D67')||'');
  if(/^Not supplied$/i.test(railSelected))allowance+=bed.code==='B750'?39:44;
  if(bed.code==='B1000'){
   const specifications=[['D45','S00'],['D33','2100'],['D31','2200'],
      ['D47','SUS-HL'],['D50','SUS-HL'],['D53','SUS-HL'],['D56','SUS-HL'],
      ['D58','Aluminum']];
   const mirror=String(get('D69')||'');
   const rail=String(get('D67')||'');
   const heel=String(get('D58')||'');
   const eligible=travel>0&&travel<=60000&&specifications.slice(0,-1).every(([cell,v])=>String(get(cell)||'').trim()===v)&&
     (!mirror||/^(No|Not supplied|Without)$/i.test(mirror))&&rail.includes('YH-56S')&&(/Aluminum|SUS-HL/i.test(heel));
   if(eligible)allowance=220;
   else if(travel>=60000)add('DECOR_60M','D39','B1000: mức trang trí 120 kg chỉ áp dụng TR < 60 m. Tại TR = 60 m cần đáp ứng Table 11-A để dùng mức 220 kg.')
  }
  if(decor+flooring>allowance)add('DECOR','D61',`${bed.code}: Decoration + Flooring (${decor+flooring} kg) vượt ${allowance} kg. B1000 chỉ được 220 kg khi đáp ứng toàn bộ Table 11-A.`);
  return issues;
 }
 window.SPEC_MR_BED_5871={cases,rows,patchEmbedded,patchDatabase,validate};
 patchEmbedded(window.SPEC_EMBEDDED?.MR);
 patchDatabase(window.SPEC_ENGINE_DATABASE?.MR);
})();
