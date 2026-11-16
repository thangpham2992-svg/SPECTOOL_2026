'use strict';
/* v5.8.68: common two-stop rule for inline Travel warning and Check Spec.
 * v5.8.65: additive Check Spec conditions. No change to inputs or dimension formulas.
 * MR: only the two-stop travel limit is added.
 * MRL Version02: same travel limit, wheelchair COP and handrail location.
 * SETS-L remains governed by the native G11L1P DataTable!D520 and F36/F37.
 */
window.SPEC_ALPHA_5865 = Object.freeze({
  // Pure, read-only predicate shared by the Travel UI and Check Spec.
  // When the number of stops is not exactly two, this extra 11 m rule does not apply.
  twoStopTravel(model,get){
    if(!['MR','MRL'].includes(model))return null;
    const stopsCell=model==='MR'?'D16':'D18';
    const travelCell=model==='MR'?'D39':'D35';
    const stopsRaw=String(get(stopsCell)??'').trim();
    const travelRaw=String(get(travelCell)??'').trim();
    if(!stopsRaw||!travelRaw)return null;
    const stops=Number(stopsRaw),travel=Number(travelRaw);
    if(stops!==2||!Number.isFinite(travel)||travel<=11000)return null;
    return {stopsCell,travelCell,travel,limit:11000};
  },
  // NEXIEZ-MRL Version02: Sales Manual G11L1P, Rev.M (2026-05-08),
  // Table 2.1-B/C. Values are maxima for this exact rated capacity and speed.
  // Speeds in m/min; 60=1.0, 96=1.6, 105=1.75, 120=2.0, 150=2.5 m/s.
  mrlV02Limits(capacityKg,speedMpm){
    const cap=Number(capacityKg),speed=Number(speedMpm);
    if(![450,630,825,1000,1275,1350,1600,1800].includes(cap)||
       ![60,96,105,120,150].includes(speed))return null;
    if(cap<=630&&speed>=120 || cap===1800&&speed>=120)return {applicable:false};
    let maxTravelM;
    if(speed===60 || cap===1800)maxTravelM=60;
    else if(cap<=1000)maxTravelM=speed>=120?105:80;
    else maxTravelM=105;
    return {applicable:true,maxTravelMm:maxTravelM*1000,maxStops:speed===60?22:30};
  },
  mrlV02CapacityKg(code){
    const x=String(code??'').trim();
    // Alpha standard Pxx code (including P13D/P17W/P21D variants).
    const prefix=(x.match(/^P(06|08|11|13|17|18|21|24)(?:[A-Z])?-(?:CO|2S)/i)||[])[1];
    if(prefix)return ({'06':450,'08':630,'11':825,'13':1000,'17':1275,'18':1350,'21':1600,'24':1800})[prefix];
    return null;
  },
  mrlV02Validate(get){
    const cap=this.mrlV02CapacityKg(get('D13'));
    const speed=Number(get('D17'));
    const limits=this.mrlV02Limits(cap,speed);
    if(!limits)return []; // Original DataTable will handle unresolved/non-standard configurations.
    const errors=[];
    if(!limits.applicable){
      errors.push({id:'MRL_V02!SpeedCapacity',cell:'D17',
        message:`Tốc độ ${speed} m/min không áp dụng cho tải trọng ${cap} kg theo Table 2.1-B/C Version02 / Speed not applicable.`});
      return errors;
    }
    const stops=Number(get('D18'));
    if(Number.isFinite(stops)&&stops>limits.maxStops)errors.push({id:'MRL_V02!MaximumStops',cell:'D18',
      message:`Số điểm dừng ${stops} vượt giới hạn ${limits.maxStops} (bao gồm tầng không phục vụ), theo Table 2.1-C Version02 / Maximum stops exceeded.`});
    const travel=Number(get('D35'));
    if(Number.isFinite(travel)&&travel>limits.maxTravelMm)errors.push({id:'MRL_V02!MaximumTravel',cell:'D35',
      message:`Travel ${travel.toLocaleString('vi-VN')} mm vượt giới hạn ${limits.maxTravelMm.toLocaleString('vi-VN')} mm cho ${cap} kg, ${speed} m/min theo Table 2.1-B Version02.`});
    return errors;
  },
  validate(model,get){
    const errors=[];
    const value=addr=>String(get(addr)??'').trim();
    const exceeded=this.twoStopTravel(model,get);
    if(exceeded){
      errors.push({id:'ALPHA!TwoStopsTravel',cell:exceeded.travelCell,
        message:`Thang chỉ có 2 điểm dừng: Travel ${exceeded.travel.toLocaleString('vi-VN')} mm vượt giới hạn 11.000 mm / Maximum travel for exactly 2 stops is 11,000 mm.`});
    }
    if(model==='MR') errors.push(...(window.SPEC_MR_BED_5871?.validate(get)||[]));
    if(model==='MRL'){
      errors.push(...this.mrlV02Validate(get));
      const copLocation=value('D60'),wheelchair=value('D62');
      if(wheelchair && wheelchair.toLowerCase()!=='no' && copLocation!=='Front'){
        errors.push({id:'MRL_V02!WheelchairCOPLocation',cell:'D62',
          message:'Wheelchair COP chỉ được chọn khi COP Location = Front. Với COP Location = Side, Wheelchair COP phải chọn No / Wheelchair COP requires Front location.'});
      }
      const handrail=value('D64'),quantity=value('D65');
      const obs=value('D14');
      if((!/^NY-40[A-E]$/i.test(obs)) && handrail && !/^not supplied$/i.test(handrail) && quantity!=='1 pc on rear wall'){
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
