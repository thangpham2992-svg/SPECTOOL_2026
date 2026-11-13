'use strict';global.window=global;require('./source-engine.js');require('./g11l1-data.js');require('./g11l1-engine.js');require('./mrl-decoration-v02.js');
const V=window.SPEC_MRL_V02_DECOR;
const capkg={'P06-2S':450,'P08-CO':630,'P08-2S':630,'P11-CO':825,'P11-2S':825,'P13W-CO':1000,'P13W-2S':1000,'P13D-CO':1000,'P13D-2S':1000,'P17W-CO':1275,'P17D-CO':1275,'P17D-2S':1275,'P18-CO':1350,'P21W-CO':1600,'P21D-CO':1600,'P21D-2S':1600,'P24-CO':1800};
let stats={tested:0,ready:0,unavailable:0,invalid:0,eligErrors:0,ny40Valid:0,dualGateValid:0};
for(const [code,kg] of Object.entries(capkg)){
 for(const obs of ['No','NY-40A','NY-40B','NY-40C','NY-40D','NY-40E']){
  for(const ent of ['1D-1G','1D-2G']){
   const d={D13:`${code} (${kg} KG)`,D14:obs,D16:ent,D17:60,D18:6,D23:kg<=1000?'BWH-620E':'BWH-720F',D26:'Mitsubishi Standard',D28:1100,D29:1400,D30:2200,D31:900,D32:2100,D35:22000,D36:1330,D41:'CL2-C',D43:'SUS-HL',D46:'SUS-HL',D49:'SUS-HL',D52:'SUS-HL',D54:'SUS-HL',D55:'Local granite (20 mm)',D56:88,D57:25,D60:ent==='1D-2G'?'Front':'Side',D61:'N1010E',D62:'No',D64:'YH-100SL',D65:'3 pcs on 3 walls',D66:'Not supplied',D67:'Provided on ceiling',D68:'DC axial flow fan',D84:'Not supplied'};
   let r=V.validate(k=>d[k]??'',()=>false);stats.tested++;stats[r.state]++;
   if(r.errors.length)stats.eligErrors++;
   if(obs!=='No'&&ent==='1D-1G'&&V.OBS_SIZES.includes(code)&&r.state!=='invalid')stats.ny40Valid++;
   if(obs==='NY-40C'&&ent==='1D-2G'&&['P11-CO','P13W-CO'].includes(code)&&r.state!=='invalid')stats.dualGateValid++;
   if(obs==='NY-40A'&&ent==='1D-2G'&&r.state!=='invalid')throw Error('NY40A 1D2G invalid not blocked: '+code);
   if(obs!=='No'&&!V.OBS_SIZES.includes(code)&&r.state!=='invalid')throw Error('Unsupported obs was accepted: '+code);
  }
 }
}
console.log(stats);
if(stats.tested!==204||stats.ny40Valid!==50||stats.dualGateValid!==2)throw Error('scenario coverage mismatch');
