'use strict';
/* Independent port of original G11L1 v2.5.0.0 CageMassAmec MethodDef 65.
   Uses pure IL function evaluator for methods 67..79. Check fail-closed. */
(function(root){
 const T=root.G11L1_TS_ALPHA_TABLES||{};
 const IL=root.G11L1_IL_PURE;
 const nd=(v)=>Number.isFinite(+v)&&v!=null&&String(v)!==''?+v:null;
 const good=(n)=>Number.isFinite(n)&&n>=0&&n<9999;
 const round=(x,d=0)=>{const p=10**d,y=x*p,q=Math.floor(y),f=y-q;return (f===0.5?(q%2===0?q:q+1):Math.round(y))/p;};
 function calcOne(id,args){ const v=IL.run(id,args);if(!good(v))throw Error(`IL ${id} returned technical sentinel: ${v}`);return round(v,1); }
 function ceiling(i){
  const aa=+i.aa,bb=+i.bb,cap=+i.capacity,typ=i.ceiling,ent=i.entrance,dt=i.doorType;
  let mass=null;
  const exact=(T.CageMassCeiling||[]).find(r=>+r.AA===aa&&+r.BB===bb);
  if(cap<=1000&&exact&&good(+exact[typ]))mass=+exact[typ];
  if(mass==null&&cap<=1000){
   const slo=(T[`CageMass_SLOPE_${typ}`]||[]).find(r=>aa>=+r['AA min.']&&aa<+r['AA Max.']&&bb>=+r['BB min.']&&bb<+r['BB Max.']);
   const int=(T[`CageMass_INTERCEPT_${typ}`]||[]).find(r=>aa>=+r['AA min.']&&aa<+r['AA Max.']&&bb>=+r['BB min.']&&bb<+r['BB Max.']);
   if(slo&&int)mass=aa*bb*Number(slo[typ])+Number(int[typ]);
  }
  if(mass==null&&cap>1000){
   let dt2=dt==='CO'?'CO':(dt==='2S'||dt==='2S-L'||dt==='2S-R')?'2S':null;
   let gate=ent==='1D-1G'?'1G':ent==='1D-2G'?'2G':null;
   const rows=T[`Extended_${typ}_${dt2}_${gate}`]||[];
   const line=rows.find(r=>+r.AA===aa&&+r.BB===bb);
   if(line&&good(+line.Mass))mass=+line.Mass;
  }
  if(!good(mass))throw Error(`Ceiling mass missing: ${typ} ${aa}x${bb} ${ent}`);
  const extraEmergency=i.emergencyExit==='No'?0:3;
  let extraFan=0;
  if(['DC axial flow fan','Line flow fan'].includes(i.ventilation))extraFan=cap<1050?1.5:3;
  else if(i.ventilation==='Plasma Quad')extraFan=6.5;
  return mass+extraEmergency+extraFan;
 }
 function carDoor(i){
  const dt=i.doorType,reg=i.regulation,finish=i.carDoor,HH=+i.hh,KUP=+i.kup,jj=+i.jj;
  let f=finish;
  if(['Colored SUS-HL','Colored SUS-HE','SUS-Shiny vibration'].includes(f))f='SUS-HE';
  if(['Pattern-printed','Laminated(CPxx/LAxx)','Laminated(MLxx)'].includes(f))f='Painted';
  const heavy=['EN81-20','Hong Kong','China','Singapore','Malaysia','IU','UK','Vietnam'].includes(reg);
  if(f==='SUS-HL')f=heavy||HH+KUP>=2401?'SUS-HL t1.5':'SUS-HL t1.2';
  else if(f==='Painted')f=heavy||HH+KUP>=2401?'Painted t1.6':'Painted t1.2';
  else if(f==='SUS-HE')f=heavy?'SUS-HE t1.5':'SUS-HE t1.2';
  const rows=T[dt==='CO'?'CageMassCarDoorCO':'CageMassCarDoor2S']||[];
  const sel=rows.find(r=>(r.finish||r.Finish)===f&&+r.HH===HH&&+r.KUP===KUP);
  if(!sel||!good(+sel[String(jj)]))throw Error('Car door table missing '+JSON.stringify({dt,f,HH,KUP,jj}));
  return +sel[String(jj)] + (dt==='CO'?3.5:1.6);
 }
 function calculate(i){
   if(!IL)throw Error('missing original IL');
   const aa=+i.aa,bb=+i.bb,jj=+i.jj,hh=+i.hh,hl=+i.hl,kk=+i.kk,
      ent=i.entrance,dt=i.doorType,cap=+i.capacity,cop=i.mainCop,aux=i.auxCop,stop=+i.stops;
   const al=dt==='CO'?(aa-jj)/2:aa-jj-10;
   let bonus=i.ceiling==='DL5'?100:i.ceiling==='DL6'?90:50;
   const wallHeight=hl-75+bonus-0.5,transomHeight=hl+bonus-0.5-hh;
   let twoG=ent==='1D-1G'?1:2;
   const comp={};
   comp.ceiling=round(ceiling(i),1);
   if(i.observation==='No'){
    comp.wall=calcOne(ent==='1D-1G'?67:68,ent==='1D-1G'
     ?[aa,bb,wallHeight,i.rearWall,i.sideWall,cop,aux,stop,dt,i.mirror,i.regulation]
     :[aa,bb,wallHeight,i.rearWall,i.sideWall,cop,aux,stop,dt,i.regulation]);
    comp.handrail=i.handrails==='No'?0:calcOne(ent==='1D-1G'?77:78,ent==='1D-1G'
      ?[aa,bb,i.handrails,i.handrailsLocation,cop,aux,stop]
      :[aa,bb,i.handrails,i.handrailsLocation,cop,aux,stop,dt]);
   } else {
    comp.wall=calcOne(69,[aa,bb,+i.ee,wallHeight,i.rearWall,i.sideWall,i.observation,cop,aux,stop,dt,ent,i.regulation]);
    comp.handrail=i.handrails==='No'?0:calcOne(79,[aa,bb,i.handrails,i.handrailsLocation,i.observation,ent,cop,aux,stop]);
   }
   comp.front=round(calcOne(70,[aa,hh,jj,al,wallHeight,i.frontPanel,dt,cop,aux,kk])*twoG,1);
   comp.transom=round(calcOne(71,[aa,hh,jj,transomHeight,i.transom,dt,cop,aux,kk])*twoG,1);
   comp.column=dt==='CO'?0:round(calcOne(72,[wallHeight,i.frontPanel,dt,kk])*twoG,1);
   comp.kickplate=calcOne(73,[aa,bb,ent,cap]);
   comp.mainCop=calcOne(74,[al,cop,aux,stop,wallHeight,hh,ent,dt,bb,kk]);
   comp.auxCop=calcOne(75,[al,cop,aux,stop,wallHeight,ent]);
   comp.mirror=calcOne(76,[aa,wallHeight,i.mirror,i.rearWall,i.sideWall]);
   const parts=['ceiling','wall','handrail','front','transom','column','kickplate','mainCop','auxCop','mirror'];
   const sum=parts.reduce((s,k)=>s+comp[k],0);
   const cage=round(sum*1.1);
   const door=round(round(carDoor(i),1)*twoG*1.1);
   return {cage,door,components:comp,sum,wallHeight,al};
 }
 root.G11L1_CAGE_SOURCE={calculate,ceiling,carDoor};
})(typeof window!=='undefined'?window:globalThis);
