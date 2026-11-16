'use strict';
/* v5.8.65: additive Check Spec conditions. No change to inputs or dimension formulas.
 * MR: only the two-stop travel limit is added.
 * MRL Version02: same travel limit, wheelchair COP and handrail location.
 * SETS-L remains governed by the native G11L1P DataTable!D520 and F36/F37.
 */
window.SPEC_ALPHA_5865 = Object.freeze({
  validate(model,get){
    const errors=[];
    const value=addr=>String(get(addr)??'').trim();
    const stopsCell=model==='MR'?'D16':'D18';
    const travelCell=model==='MR'?'D39':'D35';
    if(model==='MR'||model==='MRL'){
      const stops=Number(value(stopsCell));
      const travel=Number(value(travelCell));
      if(value(stopsCell)!=='' && stops===2 && value(travelCell)!=='' && Number.isFinite(travel) && travel>11000){
        errors.push({id:'ALPHA!TwoStopsTravel',cell:travelCell,
          message:`Thang chỉ có 2 điểm dừng: Travel ${travel.toLocaleString('vi-VN')} mm vượt giới hạn 11.000 mm / Maximum travel for exactly 2 stops is 11,000 mm.`});
      }
    }
    if(model==='MRL'){
      const copLocation=value('D60'),wheelchair=value('D62');
      if(wheelchair && wheelchair.toLowerCase()!=='no' && copLocation!=='Front'){
        errors.push({id:'MRL_V02!WheelchairCOPLocation',cell:'D62',
          message:'Wheelchair COP chỉ được chọn khi COP Location = Front. Với COP Location = Side, Wheelchair COP phải chọn No / Wheelchair COP requires Front location.'});
      }
      const handrail=value('D64'),quantity=value('D65');
      if(handrail && !/^not supplied$/i.test(handrail) && quantity!=='1 pc on rear wall'){
        errors.push({id:'MRL_V02!HandrailQuantity',cell:'D65',
          message:'Khi có tay vịn, Quantity / Location chỉ được chọn "1 pc on rear wall" / Only one handrail on the rear wall is permitted.'});
      }
    }
    return errors;
  }
});


/* v5.8.67: advisory ONLY. Recalculate OH/PD from the exact workbook formulas
 * with SETS-L toggled ON in a temporary scenario. Never mutate Spectool input,
 * actual OH/PD, dropdowns or existing validation thresholds. */
window.SPEC_MRL_SETSL_SUGGESTION=Object.freeze({
  compare(db,changes){
    const Engine=window.SpecExcelSourceEngine?.SourceEngine;
    if(!db||!Engine)return null;
    const readNumber=r=>r?.ok&&typeof r.value==='number'&&Number.isFinite(r.value)&&r.value>0?r.value:null;
    const baseline=new Engine(db,changes);
    const chosen=baseline.safe('SpecForm','A157');
    if(!chosen.ok||[1,true,'1','TRUE','V','✓'].includes(String(chosen.value).toUpperCase()))return null;
    const option=baseline.safe('SpecForm','E157');
    if(!option.ok||!['A','S'].includes(String(option.value).trim().toUpperCase()))return null;
    // Feature is valid only when the *whole* source SETS-L rule becomes TRUE.
    const alternative=new Engine(db,{...changes,A157:1});
    const allowed=alternative.safe('DataTable','D520');
    if(!allowed.ok||allowed.value!==true)return null;
    const dimensions={};let reducible=false;
    for(const [kind,cell] of [['PD','F36'],['OH','F37']]){
      const standard=readNumber(baseline.safe('SpecForm',cell));
      const withSETSL=readNumber(alternative.safe('SpecForm',cell));
      if(standard===null||withSETSL===null)return null;
      dimensions[kind]={standard,withSETSL,reduced:withSETSL<standard};
      if(withSETSL<standard)reducible=true;
    }
    return reducible?{feature:'SETS-L',featureCell:'A157',dimensions}:null;
  },
  advise(kind,actual,scenario){
    const data=scenario?.dimensions?.[kind];
    if(!data?.reduced||!(Number(actual)<data.standard))return '';
    const fmt=n=>Number(n).toLocaleString('vi-VN');
    const hasEnough=Number.isFinite(Number(actual))&&Number(actual)>=data.withSETSL;
    return `Gợi ý: Có thể bật SETS-L để giảm ${kind} tối thiểu từ ${fmt(data.standard)} mm xuống ${fmt(data.withSETSL)} mm.`+
      (hasEnough?` Kích thước đang nhập sẽ đạt giới hạn ${kind} này nếu các điều kiện kỹ thuật khác phù hợp.`:
       ` Tuy nhiên kích thước đang nhập vẫn thấp hơn ${fmt(data.withSETSL)} mm; cần điều chỉnh thêm.`);
  }
});
