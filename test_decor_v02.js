'use strict';
global.window=global;
require('./source-engine.js');require('./g11l1-data.js');require('./g11l1-engine.js');require('./mrl-decoration-v02.js');
const V=window.SPEC_MRL_V02_DECOR;
const base={D13:'P08-CO (630 KG)',D14:'No',D15:'CO',D16:'1D-1G',D17:60,D18:6,D23:'Bên hông phòng thang (Side drop) - BWH-620E',D26:'Mitsubishi Standard',D28:1100,D29:1400,D30:2200,D31:900,D32:2100,D35:22000,D36:1330,D41:'CL2-C',D43:'SUS-HL',D46:'SUS-HL',D49:'SUS-HL',D52:'SUS-HL',D54:'SUS-HL',D55:'Local granite (20 mm)',D56:88,D57:25,D60:'Side',D61:'N1010E',D64:'YH-57S (SUS304)',D65:'1 pc on rear wall',D66:'Not supplied',D67:'Trên trần phòng thang (Provided on the ceiling)',D68:'DC axial flow fan',D84:'Not supplied'};
function run(delta={},feat=[]){const values={...base,...delta};return V.validate(k=>values[k]??'',k=>feat.includes(k));}
function report(name,delta={},features=[]){let s=Date.now(),r=run(delta,features);console.log(name,JSON.stringify({state:r.state,max:r.max,errors:r.errors.map(x=>x.cell+' '+x.message),reason:r.state==='unavailable'?r.reason:undefined}),Date.now()-s+'ms');return r}
report('BASE');report('OVER',{D57:999});report('FLOOR CHANGE',{D56:128});report('CWT 620D',{D23:'Bên hông (Side drop) - BWH-620D'});
report('OBS NY40A 630',{D14:'NY-40A',D64:'YH-100SL (SUS304)'});
report('OBS NY40E 1600',{D13:'P21W-CO (1600 KG)',D14:'NY-40E',D64:'YH-100SL (SUS304)',D23:'Bên hông - BWH-720G'});
report('OBS NY40C 1D2G 825',{D13:'P11-CO (825 KG)',D14:'NY-40C',D16:'1D-2G',D60:'Front',D64:'YH-100SL (SUS304)'});
report('OBS NY40A 1D2G invalid',{D13:'P11-CO (825 KG)',D14:'NY-40A',D16:'1D-2G',D60:'Front',D64:'YH-100SL (SUS304)'});
report('OBS NY40C 1D2G 1000',{D13:'P13W-CO (1000 KG)',D14:'NY-40C',D16:'1D-2G',D60:'Front',D64:'YH-100SL (SUS304)'});
report('OBS NY40C 1D2G 1275 invalid',{D13:'P17W-CO (1275 KG)',D14:'NY-40C',D16:'1D-2G',D60:'Front',D64:'YH-100SL (SUS304)',D23:'BWH-720F'});
report('OBS NY40A P06 invalid',{D13:'P06-2S (450 KG)',D14:'NY-40A',D64:'YH-100SL (SUS304)'});
report('P17W nonOBS',{D13:'P17W-CO (1275 KG)',D23:'BWH-720F'});
report('P24 nonOBS',{D13:'P24-CO (1800 KG)',D23:'BWH-720G'});
report('1D2G No',{D13:'P11-CO (825 KG)',D16:'1D-2G'});
report('MBS feature',{},['MBS']);
