'use strict';
(function(root){
 const t=root.G11L1_TS_ALPHA_TABLES||{};
 const num=x=>{const a=Number(x);return Number.isFinite(a)?a:NaN};
 const match=(x,check)=>x==='-'||check(x);
 function candidates(d){
  const reg=d.RegPtn||({'EN81-20':'EN','MELCO std.':'MELCO','Mitsubishi Standard':'MELCO'}[d.Regulation]);
  return(t.Traction1||[]).filter(r=>
   match(r.RegPtn,v=>v.includes(reg))&&
   match(r['Lower CAP [kg]'],v=>Math.ceil(num(v))<=d.Capacity)&&
   match(r['Upper CAP [kg]'],v=>d.Capacity<=Math.ceil(num(v)))&&
   match(r['Max speed [m/s]'],v=>v.includes(String(d.Speed)))&&
   match(r['Lower TR [m]'],v=>Math.ceil(num(v))<=d.TRxm)&&
   match(r['Upper TR [m]'],v=>d.TRxm<=Math.ceil(num(v)))&&
   match(r['Chain reqd.'],v=>v.includes('Yes'))&&
   match(r['SETS-L'],v=>v.includes(d.SetsL)));
 }
 function select(d,carMass){const a=candidates(d);if(!a.length)throw Error('No traction candidate for '+JSON.stringify({RegPtn:d.RegPtn,Capacity:d.Capacity,Speed:d.Speed,TRxm:d.TRxm}));
  let row=a.find(r=> Math.ceil(num(r['Upper car mass [kg]']))>=carMass)||a[a.length-1];
  const mapping={'LowerCarMass':'Lower car mass [kg]','HoistingRope':'Hoisting rope','RopeUnitMass':'Rope unit mass [kg/m]','CompChain':'Comp. chain','CompUnitMass':'Comp unit mass [kg/m]','CableUnitMass':'Cable unit mass [kg/m]'};
  for(let [key,column] of Object.entries(mapping)){const val=row[column];d[key]=/Mass|Weight/.test(key)?+val:val;}
  const V=+d.PowerSupplyVoltage;
  d.TractionMachine=(V>=200&&V<=240)?row['T/M type (200V)']:(V>=380&&V<=480)?row['T/M type (400V)']:null;
  return {candidates:a.length,row};
 }
 function maxMin(d){const a=candidates(d);if(!a.length)throw Error('No maxMin candidates');d.MinCarMass=+a[0]['Lower car mass [kg]'];d.MaxCarMass=+a[a.length-1]['Upper car mass [kg]'];return a;}
 root.G11L1_TRACTION_SOURCE={maxMin,select,candidates};
})(typeof window!=='undefined'?window:globalThis);
