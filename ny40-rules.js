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
    // Design guide: every NY-40 observation car requires SUS-HL for all
    // three front/trim finishes. Keep these checks independent from the
    // original workbook formulas and never silently normalize user selections.
    for(const item of [
      {cell:'D52',name:'Front return panel',id:'NY40!FrontReturnFinish'},
      {cell:'D49',name:'Transom panel',id:'NY40!TransomFinish'},
      {cell:'D54',name:'Kickplate',id:'NY40!KickplateFinish'}
    ]){
      const finish=String(get(item.cell)||'').trim();
      if(finish!=='SUS-HL')errors.push({cell:item.cell,id:item.id,
        message:obs+': '+item.name+' phải là Stainless-steel hairline (SUS-HL), hiện tại: '+(finish||'chưa chọn')+' / SUS-HL required.'});
    }
    // NY-40 (A/B/C/D/E): M/P main COP, auxiliary COP and wheelchair COP
    // are not offered. Keep incompatible previous selections visible and
    // report the precise field instead of silently replacing a customer's COP.
    const mainCop=String(get('D61')||'').trim().toUpperCase();
    const auxCop=String(get('D62')||'').trim();
    if(/^[MP]\d/i.test(mainCop))errors.push({cell:'D61',id:'NY40!MainCOP',
      message:obs+': không được sử dụng COP M hoặc COP P. Hãy chọn COP phù hợp (N/C hoặc loại được phép) / M/P COP is not permitted for NY-40.'});
    if(auxCop && !/^(?:NO|NONE|NOT SUPPLIED|KHÔNG CUNG CẤP)$/i.test(auxCop))
      errors.push({cell:'D62',id:'NY40!AuxiliaryWheelchairCOP',
        message:obs+': không được sử dụng Auxiliary COP hoặc Wheelchair COP. Chọn No / Not supplied. / Auxiliary and wheelchair COP are not permitted for NY-40.'});
    const loc=String(get('D65')||'').trim(),rail=String(get('D64')||'').trim();
    if(!rail||/^not supplied$/i.test(rail))errors.push({cell:'D64',id:'NY40!Handrail',message:obs+' có vách kính nên phải cung cấp tay vịn / Handrail required on glass walls.'});
    else if(!allowed.includes(loc))errors.push({cell:'D65',id:'NY40!HandrailLocation',message:obs+': vị trí tay vịn phải là '+allowed.join(' hoặc ')+' / Wrong handrail position for observation glass.'});
    // Source R14 is the authority for special two-gate COP Front; we do not
    // invent a side restriction for one-gate or an unrecorded COP left/right side.
    const entrance=String(get('D16')||''),cop=String(get('D60')||'');
    const cap=String(get('D13')||'');
    if(entrance==='1D-2G'){
      if(obs!=='NY-40C'||! /^(?:P11-CO|P13W-CO)(?:\s|\(|$)/i.test(cap))
        errors.push({cell:'D14',id:'NY40!TwoGateApplication',message:'1D-2G với kính NY-40 chỉ áp dụng NY-40C cho P11-CO/P13W-CO (Table 4.3).'});
      if(cop!=='Front')errors.push({cell:'D60',id:'NY40!COPLocation',message:'1D-2G + NY-40C yêu cầu COP đặt tại Front return panel (Table 4.3).'});
    }
    return errors;
  }
});
