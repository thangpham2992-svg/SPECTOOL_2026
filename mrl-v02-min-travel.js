"use strict";
/* G11L1 Version 2 panel-layout minimum Travel from GQ Manual S-92-GQ-D0300-E,
 * Application of Panels and Min. Travel 1/2, 2025-07-28.
 * Not the maximum travel or two-stop limit. Autotransformer and option transformer
 * boxes are EXCLUDED by project scope; never silently substitute their rows.
 */
window.SPEC_MRL_MIN_TRAVEL=(function(){
 const criticalFeatures=['OEPS','EL-SC','EL-SCA','NSCR-C','CSB','WP','FNS'];
 const present=(selected,term)=>selected.some(raw=>{
   const f=String(raw).trim().toUpperCase();
   return f===term||f.startsWith(term+'-');
 });
 const fail=reason=>({state:'unavailable',reason});
 function calculate(arg){
   const speed=Number(arg.speed),cap=Number(arg.capacity),travel=Number(arg.travel);
   if(!Number.isFinite(cap)||cap<=0||!Number.isFinite(speed)||!Number.isFinite(travel))return fail('Thiếu Capacity / Speed / Travel');
   if(![60,96,105,120,150].includes(speed))return fail('Speed không thuộc bảng Min. Travel (60/96/105/120/150 m/min)');
   const group=String(arg.operation||'').trim().toUpperCase();
   let panel;
   if(/^1C(?:-|$)/.test(group))panel='none';
   else if(/^2C(?:-|$)/.test(group))panel='2c';
   else if(/^[34]C(?:-|$)/.test(group))panel='sigma';
   else return fail('Operation System '+group+' chưa ánh xạ vào 1C, 2C, 3C, 4C');
   const selected=Array.isArray(arg.features)?arg.features:[];
   const io=panel==='none'&&criticalFeatures.some(f=>present(selected,f));
   // MELD Panel is EXCLUDED from Minimum Travel for all Feature inputs.
   // Always use the MELD Panel = X column; MELD and MELD-H have no effect.
   // Reactor Box exists only with PCNV; never use a separate/manual default.
   const reactor=selected.some(raw=>String(raw).trim().toUpperCase()==='PCNV');
   const en8172=arg.en8172===true;
   const entrance=String(arg.entrance||'').trim().toUpperCase();
   const twoGate=/^(?:1D-2G|2D-2G|2D-1G)/.test(entrance);
   let minimum,table,starred=false;
   if(cap<=1000&&speed<=105){
      table='Page 1';
      // No auto transformer, no option transformer: only X/X rows 9,11,13,15.
      // Always MELD Panel = X: [Aux 2C, Aux ΣAI, I/O except EN81-72,
      // I/O EN81-72, without Aux/I/O except EN81-72, with EN81-72].
      const row=reactor
       ? [8150,8400,5710,5910,4860,4860]
       : [5650,5900,4900,4900,3000,3000];
      minimum=panel==='2c'?row[0]:panel==='sigma'?row[1]:io?row[en8172?3:2]:row[en8172?5:4];
      // *1 minimum 7300 mm for speeds 1.6/1.75 m/s. Where a cell has no
      // *1 annotation, taking max is numerically neutral because it is >=7300.
      if([96,105].includes(speed)){minimum=Math.max(minimum,7300);starred=true;}
   }else if((cap<=1000&&[120,150].includes(speed)) ||
        (cap>1000&&cap<=1600&&[60,96,105,120,150].includes(speed)) ||
        (cap>1600&&cap<=1800&&[60,96,105].includes(speed))){
      table='Page 2';
      // X/X auto transformer, X option transformer: the last table row.
      minimum=panel==='2c'?5750:panel==='sigma'?6000:io?5000:3000;
      if(twoGate)minimum+=800; // Page 2 Note *1
   }else return fail('Capacity/Speed ngoài phạm vi hai trang G11L1 được cung cấp');
   return {state:'ready',minimum,table,panel,io,reactor,en8172,twoGate,starred,travel,
    activeFeatures:criticalFeatures.filter(f=>present(selected,f))};
 }
 return Object.freeze({calculate});
})();
