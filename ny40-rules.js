"use strict";
/* NEXIEZ-MRL Version02 observation glass – Note 5, exact source HandrailLocation values.
   Do not apply any rule to the separate NEXIEZ-MR database. */
window.SPEC_NY40_RULES=Object.freeze({
  positionChoices:Object.freeze({
    'NY-40A':['3 pcs on 3 walls'],
    'NY-40B':['1 pc on rear wall','3 pcs on 3 walls'],
    'NY-40C':['2 pcs on both side walls','3 pcs on 3 walls'],
    'NY-40D':['2 pcs on rear & left side wall','3 pcs on 3 walls'],
    'NY-40E':['2 pcs on rear & right side wall','3 pcs on 3 walls']
  }),
  isGlassDoor:function(s){return /(?:DSJ1|DS1|\bDS\b|Large\s*window)/i.test(String(s||''));},
  allowedLocations:function(obs){return this.positionChoices[String(obs||'').trim()]||null;},
  validate:function(get){
    const errors=[],obs=String(get('D14')||'').trim(),allowed=this.allowedLocations(obs);
    if(!allowed)return errors;
    const loc=String(get('D65')||'').trim(),rail=String(get('D64')||'').trim();
    if(!rail||/^not supplied$/i.test(rail))errors.push({cell:'D64',id:'NY40!Handrail',message:obs+' có vách kính nên phải cung cấp tay vịn / Handrail required on glass walls.'});
    else if(!allowed.includes(loc))errors.push({cell:'D65',id:'NY40!HandrailLocation',message:obs+': vị trí tay vịn phải là '+allowed.join(' hoặc ')+' / Wrong handrail position for observation glass.'});
    // Source R14 is the authority for special two-gate COP Front; we do not
    // invent a side restriction for one-gate or an unrecorded COP left/right side.
    const entrance=String(get('D16')||''),cop=String(get('D60')||'');
    if(entrance==='1D-2G'&&cop==='Side')errors.push({cell:'D60',id:'NY40!COPLocation',message:obs+' với 1D-2G yêu cầu kiểm tra COP Front; COP Side không thỏa điều kiện NY-40 / Review COP location.'});
    return errors;
  }
});
