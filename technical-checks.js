// Exact direct translations of selected MRL DataTable rules. No use of cached Excel TRUE/FALSE.
'use strict';
window.SPEC_TECHNICAL_CHECK=function(model,context){
 const errors=[],verified=[],checked={},unverified=[];
 const fail=(ref,cell,description)=>{checked[ref]=false;errors.push(cell+': ['+ref+'] '+description)};
 const pass=(ref,ok,cell,description)=>{checked[ref]=!!ok;verified.push('DataTable!'+ref);if(!ok)fail(ref,cell,description)};
 const v=addr=>String(context.value(addr)??'').trim();
 const named=name=>{const f=context.fields.find(f=>f.name===name);return f?v(f.cell):''};
 const n=x=>x===''||x===null||x===undefined?NaN:Number(x);
 const kg=n(context.capacity?.Capacity??context.kg());
 const cap=v('D13'),speed=n(named('K8_Speed')||v('D17'));
 if(model!=='MRL')return {errors,verified,checked,unverified};
 const speedOK=Number.isFinite(kg)&&(([60,96,105].includes(speed)&&kg>0)||([120,150].includes(speed)&&kg>=825&&kg<=1600));
 pass('D513',speedOK,'D17',`Tốc độ ${speed} m/min không phù hợp tải trọng ${kg} kg theo điều kiện DataTable D513 / Invalid Speed–Capacity combination`);
 const travel=n(named('K21_CarTR')); // Travel unit in workbook is millimeters
 if(Number.isFinite(travel)&&Number.isFinite(kg)&&Number.isFinite(speed)){
  const limit=speed===60?60000:kg===1800?60000:kg<=630?80000:kg<=1000?(speed>=120?105000:80000):105000;
  pass('D514',kg>0&&travel<=limit,'D35',`Hành trình nâng ${travel} mm vượt giới hạn ${limit} mm theo tải trọng/tốc độ / Travel exceeds capacity-speed limit`);
 }else unverified.push('DataTable!D514 (Travel missing or not numeric)');
 const stops=n(named('K9_Stops')||v('D18'));
 if(Number.isFinite(stops)&&Number.isFinite(speed))pass('D515',stops>0&&stops<=(speed===60?22:30),'D18',`Số điểm dừng ${stops} vượt giới hạn ${speed===60?22:30} của tốc độ đã chọn / Stop count outside allowed limit`);
 else unverified.push('DataTable!D515 (Stops unavailable)');
 const ceiling=named('K22_CeilingModel')||v('D41'),h=n(named('K17_CarH')||v('D30'));
 if(ceiling&&Number.isFinite(h)){
  const validHeight=ceiling.startsWith('DL5')?[2150,2250].includes(h):ceiling.startsWith('DL6')?[2160,2260].includes(h):ceiling==='BN1'?true:[2200,2300].includes(h);
  const color=v('D42'),colors=context.rules?.ceilingColors||[];
  const validColor=ceiling==='BN1'?(!color||color==='No color (BN1)'):colors.length?colors.includes(color):null;
  if(validColor!==null)pass('D516',validHeight&&validColor,'D41',`Trần ${ceiling} không phù hợp chiều cao cabin ${h} mm hoặc màu ${color} / Ceiling height or color mismatch`);
  else {if(!validHeight)fail('D516','D41',`Trần ${ceiling} không phù hợp chiều cao ${h} mm / Invalid ceiling-height combination`);else unverified.push('DataTable!D516 (CeilingColor table not resolved)')}
 }else unverified.push('DataTable!D516 (Ceiling/CarH unavailable)');
 const restricted=['P17','P18','P21','P24'].some(x=>cap.startsWith(x));
 pass('D517',!(['CL1','DL6','DL7'].some(x=>ceiling.startsWith(x))&&restricted),'D41','Mẫu trần CL1/DL6/DL7 không áp dụng cho nhóm Capacity P17/P18/P21/P24 / Ceiling restricted for capacity');
 pass('D518',!(ceiling==='FL1'&&(['P06','P11','P13W'].some(x=>cap.startsWith(x)))),'D41','Trần FL1 không áp dụng cho Capacity P06/P11/P13W / FL1 ceiling incompatible');
 if(context.capacity?.DoorType)pass('D519',v('D15')===String(context.capacity.DoorType),'D15','Loại cửa không trùng dữ liệu DoorType của Capacity trong DataTable / Wrong door type');
 else unverified.push('DataTable!D519 (Capacity lookup unresolved)');
 pass('D521',v('D16')==='1D-1G'||(v('D16')==='1D-2G'&&['P08','P11','P13','P17D','P21D'].some(x=>cap.startsWith(x))),'D16','Loại mở cửa 1D-2G không áp dụng cho Capacity hiện tại / Entrance type restricted');
 const obs=v('D14');
 if(obs==='No') {checked.D525=true;verified.push('DataTable!D525 (Observation No branch)')}
 else {
  if(ceiling==='FL1')fail('D525','D41','Observation không được sử dụng với trần FL1 / Observation incompatible with ceiling FL1');
  const handrail=named('K32_Handrail');
  if(handrail&&handrail!=='YH-100SL (SUS304)')fail('D525',context.fields.find(x=>x.name==='K32_Handrail')?.cell||'D14','Observation yêu cầu tay vịn YH-100SL (SUS304) / Handrail required');
  unverified.push('SpecForm!R13/R14 (Observation capacity table and conditions)');
  unverified.push('DataTable!D525 (complete condition)');
 }
 // Source checks are not yet sufficiently supported for an official Excel-equivalent final decision.
 unverified.push('DataTable!D490:D511 (TblCondition1 feature combinations)','DataTable!D520 (SETS-L and speed/travel case)','DataTable!D522:D523 (minimum dimensions)','DataTable!D524 (R46, R73, R74 door detail)','DataTable!D526 (COP disabled)','DataTable!D527 (TblCondition1 summary)','DataTable!D512 (aggregate)','SpecForm!B210 (FALSE count)','SpecForm!B206 (MISTAKE/NO MISTAKE)');
 return {errors,verified,checked,unverified};
};
