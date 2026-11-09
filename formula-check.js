// Formula audit v1.2: original formula catalog is extracted verbatim from XLSX.
// Only explicitly verified formula-equivalent checks may report TRUE/FALSE.
// Other original formulas are UNKNOWN, never treated as TRUE from cached Excel results.
'use strict';
window.SPEC_FORMULA_CHECK=function(model,getValue,capacityKg,selectedFeature){
 const errors=[],warnings=[],verified=[],checked={},unverified=[];
 const form=window.SPEC_FORMULAS[model];
 const v=(addr)=>String(getValue(addr)??'').trim();
 const yes=(addr,ok,reason)=>{checked[addr]=ok; if(!ok)errors.push(addr+': '+reason);};
 const num=(addr)=>Number(v(addr));
 const contains=(s,p)=>String(s).includes(p);
 if(model==='MRL'){
  // DataTable!D519: capacity table DoorType must match SpecForm!K7_DoorType.
  const f=window.__specCurrent?.fields?.find(x=>x.name==='K7_DoorType');
  const cap=window.__specCapacity?.();
  if(f && cap && cap.DoorType){yes('D519',v(f.cell)===String(cap.DoorType),'DataTable D519: Loại cửa không trùng DoorType của Capacity / Door type differs from the Capacity row');verified.push('DataTable!D519');}
  else unverified.push('DataTable!D519');
  // DataTable!D521: 1D-2G permitted only on specific Capacity families.
  const validEntrance=v('D16')==='1D-1G'||(v('D16')==='1D-2G'&&/^(P08|P11|P13|P17D|P21D)/.test(v('D13')));
  yes('D521',validEntrance,'DataTable D521: 1D-2G chỉ áp dụng các Capacity P08/P11/P13/P17D/P21D / Invalid entrance type');verified.push('DataTable!D521');
  // SpecForm!R23 and DataTable!D524 includes 3 other R cells, not all derived.
  const kg=capacityKg(),side=v('D23');
  let r23=null;
  if(kg!==null && Number.isFinite(kg)) r23=kg<=1000?/BWH-620[DE]$/.test(side):kg>=1275?/BWH-720[FG]$/.test(side):false;
  if(r23!==null){yes('R23',r23,'SpecForm R23 / DataTable D524: Side drop không phù hợp tải trọng / Wrong side-drop capacity family');verified.push('SpecForm!R23');}else unverified.push('SpecForm!R23');
  // First branch of B209: SUS304 required on SUS-HL DSJ1 / Large Window.
  const special=[v('D46'),v('E73'),v('E74')].some(s=>s.startsWith('SUS-HL')&&(contains(s,'DSJ1')||contains(s,'Large window')));
  yes('B209',!special||v('D40')==='SUS304','SpecForm B209: Cửa SUS-HL có DSJ1/Large Window phải chọn SUS304 ở D40 / SUS304 required');verified.push('SpecForm!B209 – SUS304 branch');
  // Formula R14 / D525 also depends on R13 and selected handrail and ceiling.
  if(v('D14')==='No'){checked.D525=true;verified.push('DataTable!D525 (Observation = No)')}
  else {
   const ceiling=window.__specCurrent?.fields?.find(x=>x.name==='K22_CeilingModel');
   const rail=window.__specCurrent?.fields?.find(x=>x.name==='K32_Handrail');
   const r13ok=v('D15')==='CO' && /^(P08|P11|P13|P17D|P21D)/.test(v('D13')); // insufficient to reproduce exact COUNTIFS
   if(ceiling&&v(ceiling.cell)==='FL1')yes('D525',false,'DataTable D525: Observation không áp dụng với trần FL1 / FL1 ceiling inapplicable');
   else if(rail && v(rail.cell)!=='YH-100SL (SUS304)')yes('D525',false,'DataTable D525: Observation yêu cầu tay vịn YH-100SL (SUS304) / Required handrail');
   else {unverified.push('DataTable!D525 (R14/R13 capacity lookup)')}
  }
  // D524 depends on R23, R46, R73, R74; other members not evaluated.
  unverified.push('DataTable!D524 (R46/R73/R74)', 'DataTable!D512 (entire condition range)', 'DataTable!D527 (TblCondition1)', 'SpecForm!B210 (aggregate FALSE count)', 'SpecForm!B206 (final decision)');
 } else {
  unverified.push('SpecForm!B207 (SDE condition)', 'SpecForm!B208 (SUBTOTAL TblSpecFailure)', 'DataTable!AP2:AP10 (feature and condition checks)');
 }
 // The catalog is the authority for traceability, not an executable formula engine.
 for(const [sheet,entries] of Object.entries(form))for(const addr of Object.keys(entries)){
  const id=sheet+'!'+addr;
  if(!(id in checked) && !verified.includes(id) && !unverified.includes(id))unverified.push(id);
 }
 return {errors,warnings,verified,unverified:[...new Set(unverified)],checked,form};
};
