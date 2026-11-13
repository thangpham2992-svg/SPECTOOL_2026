'use strict';
/* NEXIEZ-MRL Version02: Technical Study G11L1 2.5.0.0 ONLY.
 * This module reads the existing SpecForm selections and does NOT change any
 * input/dropdown/dimension/other model. Excel Rev.C is intentionally not used.
 */
(()=>{
 const CODES=new Set(['P06-2S','P08-CO','P08-2S','P11-CO','P11-2S','P13W-CO','P13W-2S','P13D-CO','P13D-2S','P17W-CO','P17D-CO','P17D-2S','P18-CO','P21W-CO','P21D-CO','P21D-2S','P24-CO']);
 const KNOWN_OBS=new Set(['No','NY-40A','NY-40B','NY-40C','NY-40D','NY-40E']);
 const txt=v=>String(v??'').trim();
 const num=v=>{if(v==null||txt(v)==='')return null;const n=+v;return Number.isFinite(n)?n:null;};
 const normalizeObs=v=>{const s=txt(v).toUpperCase().replace(/^NY40([A-E])$/,'NY-40$1');return ['NO','NONE','NOT SUPPLIED','KHÔNG'].includes(s)?'No':s;};
 const codeOf=v=>txt(v).match(/^P(?:06|08|11|13W|13D|17W|17D|18|21W|21D|24)-(?:CO|2S)/i)?.[0]?.toUpperCase()||'';
 const capOf=v=>num(txt(v).match(/\((\d{3,4})\s*KG\)/i)?.[1]);
 const material=v=>{
  const s=txt(v);
  if(/^SUS-HL$/i.test(s))return 'SUS-HL';
  if(/^PAINTED$/i.test(s)||/Painted steel sheet/i.test(s))return 'Painted';
  if(/^Pattern-printed/i.test(s)||/\bPPS\b/i.test(s))return 'Pattern-printed';
  if(/^Laminated steel sheet$/i.test(s)||/Laminated.*(?:CP|LA|ML)/i.test(s))return /ML/i.test(s)?'Laminated(MLxx)':'Laminated(CPxx/LAxx)';
  return '';
 };
 const doorFinish=v=>{
  const s=txt(v),base=/^Painted/i.test(s)?'Painted':/^SUS-HL/i.test(s)?'SUS-HL':'';
  if(!base)return material(s);
  if(/Large window W200xH1300/i.test(s))return base+' with window (1300X200)';
  if(/Large window W300xH1300/i.test(s))return base+' with window (1300X300)';
  if(/DSJ1/i.test(s))return base+' with See through (DSJ1)';
  if(/DS1/i.test(s))return base+' with See through (DS1)';
  return material(s);
 };
 const nameCwt=v=>txt(v).match(/BWH[- ]*(620D|620E|720F|720G)\b/i)?.[1]?.toUpperCase()||'';
 const handrail=v=>{
  if(/^YH-100SL/i.test(txt(v)))return 'YH-100SL';
  if(/^YH-57S/i.test(txt(v)))return 'YH-57S';
  if(/not supplied|^no$|không/i.test(txt(v)))return 'No';
  return '';
 };
 const railPos=v=>({
  '1 pc on rear wall':'Rear wall','1 pc on right side wall':'Right side wall',
  '1 pc on left side wall':'Left side wall','2 pcs on both side walls':'Both side walls',
  '2 pcs on rear & right side wall':'Right and rear walls',
  '2 pcs on rear & left side wall':'Left and rear walls',
  '3 pcs on 3 walls':'Three side walls'
 })[txt(v)]||'';
 const cop=v=>{
  const s=txt(v).toUpperCase();
  if(/^D7\d\dE\d?$/.test(s))return 'CB**-D7**E (Swing type)';
  const f=s.match(/^([CMNPD])10\d\dE(?:TL)?$/)?.[1];
  if(!f)return '';
  return {C:'CB**-C10**E (F.R.panel)',M:'CB**-M10**E (Wall hanging)',N:'CB**-N10**E (Full height)',P:'CB**-P10**E (Wall hanging)',D:'CB**-D10**E (Swing type)'}[f]||'';
 };
 const floorType=v=>/vinyl|PVC/i.test(txt(v))?'PVC tile 2t':'By Local';
 const emExit=v=>/provided|cung cấp|standard/i.test(txt(v))?'Standard size':'No';
 const fan=v=>/Plasma Quad/i.test(txt(v))?'Plasma Quad':/axial flow fan/i.test(txt(v))?'DC axial flow fan':'No';
 function input(get,has){
  const v=k=>get(k),cap=codeOf(v('D13')),capacity=capOf(v('D13'));
  const mreg=/EN81[-\s]*20/i.test(txt(v('D26')))?'EN81-20':/MITSUBISHI|MELCO/i.test(txt(v('D26')))?'MELCO std.':'';
  const obs=normalizeObs(v('D14'));
  const rawMirror=txt(v('D66'));
  const mirror=/not supplied|^no$|^none$|không/i.test(rawMirror)?'No':/^YZ-52A\b/i.test(rawMirror)?'YZ-52A full set':/^YZ-53A\b/i.test(rawMirror)?'YZ-53A full set':/^YZ-55SN\b/i.test(rawMirror)?'YZ-55SN mirror wall':'';
  const mHandrail=handrail(v('D64'));
  const rawCop=txt(v('D61'));
  const rawAux=txt(v('D62'));
  // Spectool D62 lists F650E/F660E/F670E: wheelchair COP (PCNV),
  // not the technical-study auxiliary COP table ACO1. PCNV is read only
  // by rail-clearance/drawing methods, not CageMass/Maximum methods.
  const pcnv=/^F6(?:50|60|70)E$/i.test(rawAux)?'Yes':'No';
  const aux=/^(?:no|not supplied|none)$/i.test(rawAux)||pcnv==='Yes'?'No':cop(rawAux);
  return {code:cap,capacity,doorType:txt(v('D15')).toUpperCase(),entrance:txt(v('D16')).toUpperCase(),
   speed:num(v('D17')),stops:num(v('D18')),
   cwt:nameCwt(v('D23')),regulation:mreg,
   aa:num(v('D28')),bb:num(v('D29')),hl:num(v('D30')),jj:num(v('D31')),hh:num(v('D32')),
   travel:num(v('D35')),pd:num(v('D36')),
   observation:obs,ceiling:txt(v('D41')).match(/^(CL1|CL2|DL5|DL6|DL7|BN1|FL1)/i)?.[1]?.toUpperCase()||'',
   rearWall:material(v('D43')),sideWall:material(v('D43')),
   carDoor:doorFinish(v('D46')),transom:material(v('D49')),frontPanel:material(v('D52')),
   flooring:floorType(v('D55')),flooringThickness:num(txt(v('D55')).match(/(\d+(?:\.\d+)?)\s*mm/i)?.[1])||2,
   flooringMass:num(v('D56')),
   mainCop:cop(rawCop),auxCop:aux,pcnv,rawAux,rawCop,
   handrails:mHandrail,handrailsLocation:mHandrail==='No'?'No':railPos(v('D65')),
   mirror,rawMirror,
   ventilation:fan(v('D68')),emergencyExit:emExit(v('D67')),
   setsL:has('SETS-L')?'Yes':'No',fireproofDoor:/EN81-58/i.test(txt(v('D84')))?'EN81-58(II) E120':'No',
   observationRaw:txt(v('D14')),codeRaw:txt(v('D13'))
  };
 }
 function evaluate(get,has){
  const i=input(get,has),errors=[];
  const err=(cell,message)=>errors.push({id:'MRL_V02_DECOR!Input',cell,message});
  if(!CODES.has(i.code))err('D13','Car size không thuộc danh mục Alpha Version02.');
  if(!['MELCO std.','EN81-20'].includes(i.regulation))err('D26','Regulation chỉ áp dụng Mitsubishi Standard hoặc EN81-20.');
  if(!['620D','620E','720F','720G'].includes(i.cwt))err('D23','CWT chỉ áp dụng 620D/E hoặc 720F/G, không có Safety Gear.');
  if(!['CO','2S'].includes(i.doorType))err('D15','Door type không được hỗ trợ.');
  if(!['1D-1G','1D-2G'].includes(i.entrance))err('D16','Entrance không được hỗ trợ.');
  if(!KNOWN_OBS.has(i.observation))err('D14','Observation chỉ áp dụng No hoặc NY-40A/B/C/D/E.');
  if(!i.ceiling)err('D41','Không xác định được mã trần Alpha.');
  for(let [cell,k] of [['D43','rearWall'],['D46','carDoor'],['D49','transom'],['D52','frontPanel']])if(!i[k])err(cell,'Vật liệu không có ánh xạ khối lượng trong Technical Study.');
  if(!i.handrails)err('D64','Chưa xác định khối lượng tay vịn do local cung cấp.');
  if(!i.handrailsLocation)err('D65','Chưa xác định vị trí tay vịn trong Technical Study.');
  if(!i.mirror)err('D66','Chưa có tương đương chắc chắn cho loại gương này.');
  if(!i.mainCop)err('D61','Mã COP chưa có trong Technical Study.');
  if(!i.auxCop)err('D62','Auxiliary / wheelchair COP chưa được ánh xạ chắc chắn.');
  if(i.flooringMass===null||i.flooringMass<0)err('D56','Flooring Weight phải là số không âm.');
  for(const [cell,k] of [['D28','aa'],['D29','bb'],['D30','hl'],['D31','jj'],['D32','hh'],['D35','travel'],['D36','pd'],['D17','speed']])if(i[k]===null)err(cell,'Thiếu số liệu đầu vào từ Spectool.');
  if(errors.length)return {state:'invalid',max:null,errors,reason:'Input ngoài phạm vi Technical Study Alpha',info:{source:'Technical Study G11L1 2.5.0.0',inputs:i}};
  const calc=window.G11L1_SOURCE_DECOR?.evaluate(i)||{state:'unverified',reason:'Technical Study engine not loaded'};
  if(calc.state==='ready')return {state:'ready',max:calc.max,rawMaximum:calc.rawMax,errors:[],reason:'Technical Study G11L1 v2.5.0.0 (including flooring - Flooring Weight once)',info:{source:'Technical Study G11L1 2.5.0.0',inputs:i,detail:calc}};
  return {state:'unavailable',max:null,errors:[],reason:calc.reason||'Cannot verify Technical Study source logic',info:{source:'Technical Study G11L1 2.5.0.0',inputs:i,detail:calc}};
 }
 function validate(get,has){
  const result=evaluate(get,has),errors=[...result.errors];const v=get('D57');
  const entered=txt(v),n=num(v);
  if(entered!==''&&(n===null||n<0))errors.push({id:'MRL_V02_DECOR!Input',cell:'D57',message:'Decoration Weight phải là số không âm (kg).'});
  if(result.state==='ready'&&n!==null&&n>result.max){errors.push({id:'MRL_V02_DECOR!Maximum',cell:'D57',message:`Decoration Weight ${n.toLocaleString('en-US')} kg vượt Maximum: ${result.max.toLocaleString('en-US')} kg (Excluding Flooring).`});}
  return {...result,errors};
 }
 window.SPEC_MRL_V02_DECOR={evaluate,validate,ALL_V02:[...CODES]};
})();
