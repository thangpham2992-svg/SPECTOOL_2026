'use strict';
/* TECHNICAL STUDY G11L1 v2.5.0.0 - source-only independent maximum engine.
 * This module DOES NOT read any Excel Rev.C worksheets or touch the Spectool UI/state.
 * Imported source: cage IL MethodDef 67-79, mechanical IL MethodDef 43-62,
 * selectors and source TSV tables. Any undefined or source sentinel (9999) fails closed.
 * Supported scope: NEXIEZ-MRL Version02 Alpha standard sizes, EN81-20/MELCO,
 * CWT safety gear absent; four CWT sizes. All dimensions internal to the solver
 * are temporary values in a freshly allocated object; no shaft values are written.
 */
(function(root){
const CE=root.G11L1_CAGE_SOURCE,ME=root.G11L1_MACHINE_SOURCE,TRA=root.G11L1_TRACTION_SOURCE,DUT=root.G11L1_DUTY_SOURCE;
const allowedCwt=new Set(['620D','620E','720F','720G']);
const capacityRange=cap=>cap<=450?"A0'":cap<=630?"A1'":cap<=1000?"A2'":cap<=1600?"B1'":cap<=1800?"B2'":null;
const number=v=>{if(v===null||v===undefined||String(v).trim()==='')return null;const n=+v;return Number.isFinite(n)?n:null;};
const assert=(test,message)=>{if(!test)throw Error(message);};
// 9999 is the Technical Study sentinel for an unsupported table lookup.
// Real mechanical limits can be 10,000 kg or more (two 5,000 kg buffers).
const within=v=>typeof v==='number'&&Number.isFinite(v)&&v>=0&&v!==9999;
const range=['CageMass','CarDoorMass','MachineMass','MaxCarMass','MaxCwtMass','MaxCarBufferLoad','MaxCwtBufferLoad','MaxCarSafetyLoad','CableMass','CompMass'];
function single(i,side,opening){
 const cap=number(i.capacity),speed=number(i.speed),travel=number(i.travel),floor=number(i.flooringMass);
 const aa=number(i.aa),bb=number(i.bb),jj=number(i.jj),hh=number(i.hh),hl=number(i.hl),pd=number(i.pd);
 assert(capacityRange(cap),'Technical Study: invalid capacity range');
 assert([speed,travel,floor,aa,bb,jj,hh,hl,pd].every(v=>v!==null),'Technical Study: missing numeric SpecForm input');
 assert(travel>=0&&floor>=0&&pd>=0&&aa>0&&bb>0&&jj>0,'Technical Study: invalid input dimension/weight');
 assert(allowedCwt.has(i.cwt),'Technical Study: unsupported CWT');
 assert(['MELCO std.','EN81-20'].includes(i.regulation),'Technical Study: unsupported regulation');
 const cwtName={'620D':'650X240','620E':'650X300','720F':'750X345','720G':'750X375'}[i.cwt];
 const d={
  ModelGroup:'G11L1',Regulation:i.regulation,RegPtn:i.regulation==='EN81-20'?'EN':'MELCO',
  FireproofDoor:i.fireproofDoor||'No',PowerSupplyVoltage:400,
  CodeNumber:i.code,Capacity:cap,Speed:speed,Stops:number(i.stops),
  SetsL:i.setsL||'No',DoorType:opening,Entrance:i.entrance,
  AAxx:aa,BBxx:bb,JJxx:jj,HHxx:hh,HLxx:hl,
  Observation:i.observation,Ceiling:i.ceiling,
  EmergencyExit:i.emergencyExit||'No',Ventilation:i.ventilation||'No',
  RearWall:i.rearWall,SideWall:i.sideWall,CarTransom:i.transom,
  FrontReturnPanel:i.frontPanel,CarDoor:i.carDoor,
  Flooring:i.flooring||'By Local',TTxx:number(i.flooringThickness)||2,
  FlooringMass:floor,DecorationMass:0,
  CarSill:'Aluminum',Handrails:i.handrails,HandrailsLocation:i.handrailsLocation,
  Mirror:i.mirror,MainCop:i.mainCop,AuxCop:i.auxCop||'No',Pcnv:i.pcnv||'No',
  BalustradeHeight:i.regulation==='EN81-20'?'Fixed type 1100mm':'Fixed type 700mm',
  CwtLocation:side,CwtSafetyReqd:'No',CwtType:`BWH-${i.cwt} (${cwtName})`,
  TRxm:travel,PDxx:pd,ShortOHMode:'No',BracketSpan:0,
 };
 assert(Number.isFinite(d.Stops)&&d.Stops>0,'Technical Study: missing stops');
 const stage={};let phase='dimensions';
 try{
  d.HCxx=ME.run(518,[d.CapacityRange=capacityRange(cap),d.Ceiling,d.HLxx]);
  assert(within(d.HCxx),'Technical Study: invalid canopy height');
  ME.run(523,[d]);
  d.EExx=ME.run(520,[d.CapacityRange,d.Entrance,d.DoorType,d.BBxx]);
  assert([d.KKxx,d.KUPx,d.ASxx,d.BSxx,d.BGxx,d.HBxx,d.EExx].every(within),'Technical Study: unsupported cabin geometry');
  phase='cage mass';
  const cage=CE.calculate({aa,bb,jj,hh,hl,kk:d.KKxx,kup:d.KUPx,capacity:cap,speed,ceiling:i.ceiling,
    rearWall:i.rearWall,sideWall:i.sideWall,frontPanel:i.frontPanel,transom:i.transom,
    carDoor:i.carDoor,mainCop:i.mainCop,auxCop:i.auxCop||'No',stops:d.Stops,
    handrails:i.handrails,handrailsLocation:i.handrailsLocation,mirror:i.mirror,
    doorType:opening,entrance:i.entrance,observation:i.observation,
    regulation:i.regulation,emergencyExit:d.EmergencyExit,ventilation:d.Ventilation,ee:d.EExx});
  d.CageMass=cage.cage;d.CarDoorMass=cage.door;
  assert(within(d.CageMass)&&within(d.CarDoorMass),'Technical Study: cage/door mass unresolved');
  stage.cage=cage;
  phase='traction select';
  TRA.maxMin(d);DUT.balance(d);
  assert(within(d.MaxCarMass)&&within(d.MinCarMass),'Technical Study: traction car limit unresolved');
  d.CarMass=d.MaxCarMass;
  TRA.select(d,d.CarMass);
  assert(d.TractionMachine&&d.TractionMachine!=='-'&&d.HoistingRope&&d.HoistingRope!=='-', 'Technical Study: traction choice unresolved');
  phase='rope & compensation';
  ME.run(44,[d]);ME.run(45,[d]);
  phase='car safety/buffers';
  DUT.duty(d);
  phase='machine mass';
  ME.run(46,[d]);ME.run(45,[d]);
  phase='maximum';
  const active=[...range];if(d.CapacityRange.includes('A'))active.push('MaxSafetyFrameMoment','BGxx');
  const missing=active.filter(k=>!within(d[k]));
  if(missing.length)throw Error('Technical Study: unresolved mechanical limits '+missing.join(', '));
  assert(d.Counterweight==='BWH-'+i.cwt,'Technical Study: selected CWT mismatch');
  ME.run(43,[d]);
  assert(Number.isFinite(d.MaxDecorationMass)&&Math.abs(d.MaxDecorationMass)<9999,'Technical Study: invalid MaxDecorationMass');
  const maxInclFloor=d.MaxDecorationMass,maxExclFloor=maxInclFloor-floor;
  if(maxExclFloor<0)throw Error('Technical Study: flooring already exceeds allowance');
  const cwtConstraint=ME.round(d.MaxCwtMass-.5*cap-.5*d.CableMass-4);
  const controlling={car:ME.round(d.MaxCarMass),cwt:cwtConstraint,carBuffer:ME.round(d.MaxCarBufferLoad-cap),cwtBuffer:ME.round(d.MaxCwtBufferLoad-.5*cap-.5*d.CableMass),carSafety:ME.round(d.MaxCarSafetyLoad-cap-d.CompMass-d.CableMass)};
  if(d.CapacityRange.includes('A'))controlling.frame=ME.round(d.MaxSafetyFrameMoment/d.BGxx*1e6-cap);
  return {state:'ready',max:Math.floor(maxExclFloor+1e-8),rawMax:maxExclFloor,includingFloor:maxInclFloor,
   controls:controlling,limiting:Object.keys(controlling).sort((a,b)=>controlling[a]-controlling[b])[0],
   masses:Object.fromEntries([...range,'BGxx','MaxSafetyFrameMoment'].map(k=>[k,d[k]])),
   selected:{machine:d.TractionMachine,rope:d.HoistingRope,carSafety:d.CarSafety,cwt:d.Counterweight,buffer:d.CarBuffer},
   cage:stage.cage,phase:'complete'};
 }catch(e){return {state:'unverified',max:null,reason:phase+': '+(e?.message||String(e)),phase};}
}
function evaluate(i){
 try{
  assert(CE&&ME&&TRA&&DUT,'Technical Study source scripts not loaded');
  assert(i&&typeof i==='object','Technical Study missing input');
  const variants=[];const doors=i.doorType==='2S'?['2S-L','2S-R']:i.doorType==='CO'?['CO']:[];
  assert(doors.length,'Unsupported door type');
  for(const orientation of ['LB','RB'])for(const door of doors){
   const x=single(i,orientation,door);
   if(x.state!=='ready')return {...x,orientation,door};
   variants.push(x);
  }
  const ref=variants[0];
  // Spectool does not expose LB/RB, nor 2S left/right. Only emit a single
  // Maximum where all plausible Technical Study variants produce the same kg.
  if(!variants.every(x=>x.max===ref.max))return {state:'unverified',max:null,reason:'LB/RB or 2S left/right changes Maximum but Spectool does not distinguish these choices',variants:variants.map(x=>({max:x.max,limiting:x.limiting}))};
  return {...ref,variantsChecked:variants.length};
 }catch(e){return {state:'unverified',max:null,reason:String(e?.message||e)};}
}
root.G11L1_SOURCE_DECOR={evaluate,groups:['CageMass','CarDoorMass','MachineMass','MaxCarMass','MaxCwtMass','MaxCarBufferLoad','MaxCwtBufferLoad','MaxCarSafetyLoad','MaxCwtSafetyLoad (not applicable)','MaxSafetyFrameMoment/BGxx','CableMass','CompMass']};
})(typeof window!=='undefined'?window:globalThis);
