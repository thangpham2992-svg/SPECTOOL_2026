
'use strict';
/* Reconstruct DevExpress LightUp hatched-cell rules from actual workbook
   conditional formatting (dxf fill pattern=lightUp). Never use cached colours. */
window.SPEC_CF_CHECK=function(model,changes,db,engine){
 const meta=db?.conditional||window.SPEC_CONDITIONAL?.[model];
 if(!meta?.lightUp||!meta?.namedRanges)return{errors:[],unknown:['Workbook conditional formatting (LightUp) unavailable'],verified:[],locked:{}};
 const rules=meta.lightUp, errors=[],unknown=[],verified=[],locked={};
 const a1=(str)=>{const m=/^([A-Z]+)(\d+)$/.exec(str||'');if(!m)return null;let c=0;for(const a of m[1])c=26*c+(a.charCodeAt(0)-64);return[c,Number(m[2])];};
 const contains=(ranges,addr)=>{const xy=a1(addr);return xy&&ranges.some(([c1,r1,c2,r2])=>xy[0]>=c1&&xy[0]<=c2&&xy[1]>=r1&&xy[1]<=r2)};
 const seen=new Set();
 for(const [name,ref] of Object.entries(meta.namedRanges)){
  const match=/^'?SpecForm'?!\$?([A-Z]+)\$?(\d+)$/i.exec(String(ref).trim());
  if(!match)continue;
  const cell=match[1].toUpperCase()+match[2];if(seen.has(cell))continue;seen.add(cell);
  let hatched=(meta.defaultLightUp||[]).includes(cell), valid=true;
  for(const rule of rules){
   if(!contains(rule.ranges,cell))continue;
   try{
    const v=engine.eval(engine.ast(rule.formula),{sheet:'SpecForm',addr:cell});
    if(typeof v!=='boolean'&&typeof v!=='number')throw Error('Condition did not return TRUE/FALSE');
    if(Boolean(v)){hatched=true;if(rule.stop)break;}
   }catch(err){valid=false;unknown.push('SpecForm!'+cell+' ('+name+'): Conditional Formatting ='+rule.formula+' → '+(err.message||String(err)));}
  }
  if(!valid)continue;
  locked[cell]=hatched;
  const result=engine.safe('SpecForm',cell);
  if(!result.ok){unknown.push('SpecForm!'+cell+' ('+name+'): '+result.error);continue;}
  const val=result.value;
  const empty=val===null||val===undefined||val==='';
  if(hatched&&!empty)errors.push({id:'SpecForm!'+cell+' / LightUp',cell,message:'Ô đang bị khóa theo Excel nhưng vẫn có dữ liệu / Value must be cleared from hatched field',detail:name+' — Conditional Formatting (LightUp)'});
  else if(!hatched&&empty)errors.push({id:'SpecForm!'+cell+' / Required',cell,message:'Ô bắt buộc chưa nhập dữ liệu / Required input is blank',detail:name+' — Named Range / Non-hatched input'});
  verified.push('SpecForm!'+cell+' / '+name+' (LightUp='+hatched+')');
 }
 return{errors,unknown,verified,locked};
};
