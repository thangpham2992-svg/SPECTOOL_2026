/* SpecWebTool v5.3: MR mirror allowance adjustments, AES13-006 Rev AM Table 11-C, physical PDF page 26.
  These are kg ADJUSTMENTS to permissible decoration weight, never the rated capacity. */
(function(){'use strict';
const source='AES13-006 Rev AM, Table 11-C (14/N), 2026-02-13';
const weights={"P6":[-26.0,-41.0,-3.3],"P6C3":[-23.0,-35.0,-3.3],"P8":[-26.0,-41.0,-3.3],"P9":[-26.0,-41.0,-3.3],"P10":[-26.0,-41.0,-3.3],"P11":[-26.0,-41.0,-3.3],"P11C1":[-27.0,-41.0,-3.3],"P11C3":[-26.0,-39.0,-3.3],"P13":[-31.0,-47.0,-2.1],"P15":[-31.0,-47.0,-2.1],"P15W":[-30.0,-46.0,-7.4],"P15D":[-20.0,-32.0,-1.0],"P15C1":[-21.0,-32.0,-1.0],"P15C2":[-31.0,-47.0,-2.1],"P15C8":[-21.0,-32.0,-3.9],"P15C10":[-23.0,-35.0,-3.3],"P15C11":[-23.0,-35.0,-3.3],"P17":[-30.0,-46.0,-7.4],"P17W":[-34.0,-52.0,-7.4],"P17C4":[-23.0,-35.0,-3.3],"P17C5":[-23.0,-35.0,-3.3],"P20":[-30.0,-46.0,-7.4],"P20W":[-33.0,-52.0,-7.4],"P20C4":[-25.0,-38.0,-3.3],"P20C5":[-25.0,-38.0,-3.3],"P20C6":[-27.0,-41.0,-3.3],"P20C7":[-27.0,-41.0,-3.3],"P20C8":[-27.0,-41.0,-3.3],"P20C9":[-27.0,-41.0,-3.3]};
const models=['YZ-52A','YZ-53A','YZ-55SN'];
const norm=code=>String(code||'').toUpperCase().replace(/^P0+(\d+)/,'P$1').replace(/[-_\s].*$/,'');
const codeOf=name=>norm(String(name||'').split('-')[0]);
const selected=model=>String(model||'').trim().toUpperCase();
const allOptions=(model,capacity,aa)=>{
 const base=['Not supplied'];
 if(model==='MR'){
   const code=codeOf(capacity);
   if(!/^B(?:750|1000)$/.test(code))base.push(...models,'By Local');
 } else {
   const width=Number(aa);
   if(width>=1000&&width<2201)base.push('YZ-52A','YZ-53A');
   if(width>=1100&&width<2101)base.push('YZ-55SN');
   base.push('By Local');
 }
 return base;
};
function adjustment(capacity,mirror){
 const code=codeOf(capacity),option=selected(mirror),vals=weights[code];
 if(!vals)return null;
 if(option==='NOT SUPPLIED')return 0;
 if(option==='BY LOCAL')return null; // Mass of local-supply mirror must be provided in decoration input.
 const i=models.indexOf(option);
 return i<0?null:vals[i];
}
function fixDatabase(model,db){
 const dt=db?.sheets?.DataTable;if(!dt)return db;
 if(model==='MR'){
  for(const [row,code] of Object.entries(weights)){} // no-op: populated from keyed source below
  for(let r=10;r<=56;r++){
   const base=String(dt['B'+r]||'');const k=codeOf(base),a=weights[k];
   if(!a)continue;
   const orig=dt['AA'+r];if(!orig||typeof orig!=='object'||!orig.f)continue;
   // Excel source had hardcoded YZ-52A adjustment for ANY mirror; swap only that component.
   const target=/IF\(SpecForm!\$D\$69="Not supplied",0,INDEX\(\$AD\$10:\$AD\$56,MATCH\(SpecForm!\$D\$13,\$B\$10:\$B\$56,0\)\)\)/;
   const rule=`IF(SpecForm!$D$69="Not supplied",0,IF(OR(SpecForm!$D$69="By Local",SpecForm!$D$69="By local"),0,IF(SpecForm!$D$69="YZ-52A",${a[0]},IF(SpecForm!$D$69="YZ-53A",${a[1]},IF(SpecForm!$D$69="YZ-55SN",${a[2]},0)))))`;
   if(target.test(orig.f))orig.f=orig.f.replace(target,rule);
  }
  // MR only: link the live Technical Check to the same Table 11-C mirror data
  // that is stored in the MR source workbook (never use the MRL Version02 table).
  const mrMass=db.sheets.Mirror_Mass_Table11C||(db.sheets.Mirror_Mass_Table11C={});
  const massHeaders=['Cabin code','Not supplied (kg)','YZ-52A (kg)','YZ-53A (kg)','YZ-55SN (kg)'];
  for(let i=0;i<massHeaders.length;i++){const col='ABCDE'[i];if(!Object.hasOwn(mrMass,col+'1'))mrMass[col+'1']=massHeaders[i];}
  Object.entries(weights).forEach(([code,nums],i)=>{
   const row=i+2;mrMass['A'+row]=code;mrMass['B'+row]=0;
   mrMass['C'+row]=nums[0];mrMass['D'+row]=nums[1];mrMass['E'+row]=nums[2];
  });
  const original='IF(SpecForm!D69="Not supplied",0,INDEX($AD$10:$AD$56,MATCH(SpecForm!D13,$B$10:$B$56,0)))';
  const carCode='LEFT(SpecForm!D13,FIND("-",SpecForm!D13)-1)';
  const massRow=`IFERROR(MATCH(${carCode},Mirror_Mass_Table11C!$A$2:$A$30,0),MATCH(SUBSTITUTE(${carCode},"P0","P"),Mirror_Mass_Table11C!$A$2:$A$30,0))`;
  const massColumn='MATCH(SpecForm!D69&" (kg)",Mirror_Mass_Table11C!$C$1:$E$1,0)';
  const formula=`IF(OR(SpecForm!D69="Not supplied",SpecForm!D69="By Local"),0,IFERROR(INDEX(Mirror_Mass_Table11C!$C$2:$E$30,${massRow},${massColumn}),0))`;
  if(dt.AP15&&typeof dt.AP15==='object'&&dt.AP15.f&&dt.AP15.f.includes(original))dt.AP15.f=dt.AP15.f.replace(original,formula);
  dt.A193='YZ-53A';dt.B193='YZ-53A - Bộ 2 gương vách sau';dt.C193='Model YZ-53A - Full-size mirror set with aluminum frame';
  dt.A194='YZ-55SN';dt.B194='YZ-55SN - Gương inox toàn chiều cao';dt.C194='Model YZ-55SN - Full-height mirror-finish SUS304 wall';
  if(db.tables?.TblK33_Mirror)db.tables.TblK33_Mirror.ref='A189:C194';
 }else if(model==='MRL'){
  dt.A201='YZ-53A';dt.B201='YZ-53A - Gương toàn thân với khung nhôm';dt.C201='Model YZ-53A - Full-size glass mirror with aluminum frame';
  dt.A202='YZ-55SN';dt.B202='YZ-55SN - Vách inox gương toàn chiều cao';dt.C202='Model YZ-55SN - Full-height SUS-M(SUS304) rear wall';
  if(db.tables?.TblK33_Mirror)db.tables.TblK33_Mirror.ref='A197:C202';
  // Version02 only: supplement the in-memory workbook with the actual source
  // Mirror_AA_V02 lookup table and wire the test into the DataTable technical flow.
  // This is NOT a shared MR feature rule.
  const mirrorTable=db.sheets.Mirror_AA_V02||(db.sheets.Mirror_AA_V02={});
  const rows=[['Not supplied',null,null],['YZ-52A',1000,2200],['YZ-53A',1000,2200],['YZ-55SN',1100,2100],['By Local',null,null]];
  for(let i=0;i<rows.length;i++){
   const row=i+2;
   if(!Object.hasOwn(mirrorTable,'A'+row))mirrorTable['A'+row]=rows[i][0];
   if(!Object.hasOwn(mirrorTable,'B'+row))mirrorTable['B'+row]=rows[i][1];
   if(!Object.hasOwn(mirrorTable,'C'+row))mirrorTable['C'+row]=rows[i][2];
  }
  dt.A528='G11 Mirror';dt.B528='Mirror model and cabin internal width AA';
  dt.C528='Mirror YZ-52A/YZ-53A: AA 1000-2200 mm; YZ-55SN: AA 1100-2100 mm';
  if(!dt.D528||typeof dt.D528!=='object'||!dt.D528.f)
   dt.D528={f:'IF(OR(SpecForm!$D$66="Not supplied",SpecForm!$D$66="By Local"),TRUE,IFERROR(AND(ISNUMBER(SpecForm!$D$28),SpecForm!$D$28>=INDEX(Mirror_AA_V02!$B$2:$B$6,MATCH(SpecForm!$D$66,Mirror_AA_V02!$A$2:$A$6,0)),SpecForm!$D$28<=INDEX(Mirror_AA_V02!$C$2:$C$6,MATCH(SpecForm!$D$66,Mirror_AA_V02!$A$2:$A$6,0))),FALSE))'};
  const aggregate=dt.D512;
  if(aggregate&&typeof aggregate==='object'&&aggregate.f&&!/\bD528\b/.test(aggregate.f))
   aggregate.f=aggregate.f.replace(/\)$/,',D528)');
  const spec=db.sheets.SpecForm||{};
  const count=spec.B210;
  if(count&&typeof count==='object'&&count.f&&!count.f.includes('$D$528'))
   count.f+='+IF(DataTable!$D$528,0,1)';
  const detail=spec.B209;
  if(detail&&typeof detail==='object'&&detail.f&&!detail.f.includes('DataTable!$D$528'))
   detail.f='IF(NOT(DataTable!$D$528),"Err - DataTable!D528: Mirror model not applicable to car internal width AA (Gương không phù hợp chiều rộng AA)",'+detail.f.replace(/^=/,'')+')';
 }
 return db;
}
function validate(model,capacity,aa,mirror){
 const opt=selected(mirror);const allowed=allOptions(model,capacity,aa);
 if(!allowed.some(x=>selected(x)===opt))return {ok:false,message:model==='MR'?'Gương không áp dụng cho cabin B750/B1000 / Mirrors are not applicable to bed lifts':'Loại gương không phù hợp chiều rộng AA / Mirror option not applicable to car width AA'};
 return {ok:true};
}
window.SPEC_MIRROR_RULES={weights,source,models,allOptions,adjustment,fixDatabase,validate,codeOf};
})();