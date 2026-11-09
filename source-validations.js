// Checked against the original workbook DataTable, not Excel cached cells.
'use strict';
window.SPEC_SOURCE_VALIDATION=function(model,context){
 const errors=[],verified=[],unverified=[],checked={};
 const raw=a=>String(context.value(a)??'').trim();
 const mark=(ref,cell,valid,reason)=>{checked[ref]=valid;if(!valid)errors.push(`${cell}: [${ref}] ${reason}`);verified.push(ref)};
 const feats=context.features||[];
 const feature=code=>feats.find(f=>String(f.code).trim().toLowerCase()===code.toLowerCase());
 const selected=code=>{const f=feature(code);return f?!!context.featureSelected(f):false};
 const fcell=code=>feature(code)?.cell||'A116';
 const one=selected('SDE - One Side'),both=selected('SDE - Both Sides');
 const conflictRef=model==='MR'?'DataTable!D489':'DataTable!D498';
 mark(conflictRef,fcell('SDE - Both Sides'),!(one&&both),'SDE - One Side và SDE - Both Sides không được chọn đồng thời / Mutually exclusive SDE features. Bỏ một trong hai lựa chọn.');
 if(model==='MR'){
  // DataTable!AP10 = NOT(AND(SpecForm!D14="CO", OR(SEARCH("Large window",D50),SEARCH("DSJ1",D50)), SpecForm!A116=1))
  const special=/large\s*window|dsj1/i.test(raw('D50'));
  mark('DataTable!AP10',fcell('SDE - One Side'),!(one&&raw('D14')==='CO'&&special),'Cửa CO + Large Window/DSJ1 không được chọn SDE - One Side / CO with special door requires No SDE or Both Sides.');
  // Minimum OH is a physical lookup in the workbook; do not use cached values or a constant.
  const rows=window.SPEC_ENGINEERING_SOURCE?.MR?.OH||[];
  const speed=Number(raw('D15')),hl=Number(raw('D31')),actual=Number(raw('D41'));
  const ceiling=raw('D45').toUpperCase().replace(/\s+/g,'');
  let group=null;
  if(['S00','L210','L210S'].includes(ceiling))group='S00-L210-L210S';
  if(['S300','N300','N300S','L310','N120','N130','N140','FEVN10'].includes(ceiling))group='S300-N300-N300S-L310-N120-N130-N140';
  // FEVN10: per project specification, uses S300 as source; same group.
  const row=rows.find(r=>r.speed===speed&&r.hl===hl);
  if(row&&group&&Number.isFinite(actual)&&Number.isFinite(row.groups[group])){
   const min=Number(row.groups[group]);
   mark('DataTable!TblK21_CarOH / '+row.source,'D41',actual>=min,`OH = ${actual} mm, nhỏ hơn OH tối thiểu ${min} mm theo Speed ${speed}, HL ${hl}, Ceiling ${ceiling} / OH below workbook minimum.`);
  } else {
   unverified.push(`DataTable!TblK21_CarOH (OH unresolved: Speed ${raw('D15')}, HL ${raw('D31')}, Ceiling ${ceiling||'(blank)'}, OH ${raw('D41')})`);
  }
  // Direct translations of selected TblCondition1 feature constraints.
  const pair=(ref,a,b)=>mark(ref,fcell(b),!(selected(a)&&selected(b)),`${a} và ${b} không được chọn cùng lúc / Mutually exclusive features.`);
  pair('DataTable!D495','EVRC-C','EVRC-S');
  pair('DataTable!D500','AECH','AECC');
  pair('DataTable!D501','AAN-G','AAN-B');
  if(selected('SDE - Both Sides'))mark('DataTable!D487',fcell('SDE - Both Sides'),raw('D14')==='CO','SDE - Both Sides chỉ dùng cho cửa CO / Both-side SDE requires CO door.');
  const required=[['D482','OEPS-A',['WP','WP-W']],['D483','OEPS-M',['WP']],['D499','HE-B',['DKO-TB']],['D502','BPL',['ABP']]];
  for(const [ref,code,need] of required){if(selected(code))mark('DataTable!'+ref,fcell(code),need.some(selected),`${code} yêu cầu ${need.join(' hoặc ')} / Required feature missing.`)}
  // Table 485 restricts SR-1, SR-2, multi-beam and MBS3 co-selection.
  const sensors=['SR-1 Beam','SR-2 Beams','Multi - Beam Door Sensor','MBS3'];
  if(sensors.filter(selected).length>1)mark('DataTable!D485',fcell(sensors.find(x=>selected(x))),'false'==='true','Các cảm biến SR-1/SR-2/Multi-Beam/MBS3 không được chọn đồng thời / Door sensors conflict.');
  const stopsField=context.fields.find(x=>x.name==='K9_Stops');const stops=Number(stopsField?raw(stopsField.cell):raw('D16'));
  for(const [code,minimum,ref] of [['ABP',3,'D490'],['FCC-P',3,'D491'],['FCC-A',6,'D492']])if(selected(code)){
   if(Number.isFinite(stops)&&stops>0)mark('DataTable!'+ref,fcell(code),stops>=minimum,`${code} yêu cầu ít nhất ${minimum} điểm dừng / Insufficient stops (${stops}).`);
   else unverified.push('DataTable!'+ref+' (Stops not resolved)');
  }
 }else{
  // MRL workbook DataTable rules 495,496,497,498.
  if(selected('MBSS')&&(one||both))mark('DataTable!D495',fcell('MBSS'),false,'MBSS không được dùng với SDE / MBSS and SDE conflict.');
  if(both)mark('DataTable!D496',fcell('SDE - Both Sides'),raw('D15')==='CO'&&!/DS1(?!J)/i.test(raw('D46')),'SDE Both Sides không áp dụng cho Door Type hiện tại / Wrong door type for Both Sides.');
  if(one){const finish=raw('D46');const valid=!/DS1(?!J)/i.test(finish)&&(!/DSJ1/i.test(finish)||raw('D15')==='2S');mark('DataTable!D497',fcell('SDE - One Side'),valid,'SDE One Side không phù hợp Door Type/DSJ1/DS1 / Invalid one-side SDE combination.');}
 }
 return {errors,verified,unverified,checked};
};
