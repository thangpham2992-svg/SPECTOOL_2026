/* Hall Lanterns — NEXIEZ-MRL Alpha Package Version02 (G11L1P),
   Sales Manual S-1-U4010-E, Table 5.6.2, page 27.
   Independent from NEXIEZ-MR Feature and equipment catalogs. */
(function(){'use strict';
 const names=['HLV-A16S','HLV-E65','HLV-E66','HLH-A16S',
  'HLV1-A80S','HLV3-A80S','HLV1-A90S','HLV3-A90S','HLV1-E62','HLV3-E62',
  'HLH1-A80S','HLH3-A80S','HLH1-A90S','HLH3-A90S'];
 // Faceplate materials from Sales Manual S-1-U4010-E, Table 5.6.2, p.27.
 // A90S variants have a black plastic faceplate; all other listed variants use SUS-HL (SUS304).
 const blackPlastic=new Set(['HLV1-A90S','HLV3-A90S','HLH1-A90S','HLH3-A90S']);
 const catalog=names.map(code=>({code,image:'assets/hall-lanterns/'+code+'.jpeg',
  finish:blackPlastic.has(code)?'Black plastic':'SUS-HL (SUS304)'}));
 function exportDescription(code){
   const rec=catalog.find(item=>item.code===code);
   if(!rec)return null;
   if(rec.finish==='Black plastic')return {
     vi:`Đèn báo tầng model ${code}, mặt hộp nhựa đen.`,
     en:`(Hall lantern model ${code}, black plastic faceplate.)`
   };
   return {
     vi:`Đèn báo tầng model ${code}, mặt hộp inox sọc nhuyễn SUS-HL (SUS304).`,
     en:`(Hall lantern model ${code}, SUS-HL (SUS304) stainless steel hairline faceplate.)`
   };
 }
 const keys={TopFloor:'HL_TOP',OtherFloor:'HL_OTHER',MainFloor:'HL_MAIN'};
 const none='Not supplied';
 const byCode=new Map(catalog.map(a=>[a.code,a]));
 const selection=(changes,key)=>{
  const code=String(changes?.[key]||none).trim();return code===none||byCode.has(code)?code:none;
 };
 const selected=(changes)=>Object.values(keys).map(key=>selection(changes,key)).filter(x=>x!==none);
 const any=(changes)=>selected(changes).length>0;
 window.SPEC_HALL_LANTERN={catalog,keys,none,byCode,selection,selected,any,exportDescription,allowed:code=>code===none||byCode.has(code)};
})();
