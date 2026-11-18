/* G11L1P jamb-depth choices. Scope: NEXIEZ-MR and NEXIEZ-MRL Version2. */
'use strict';
(()=>{
  // The two workbooks use different source cell addresses for jamb models.
  const mrlFloors=Object.freeze([
    {cell:'D71',key:'JD_MAIN',floorVi:'Tầng chính',floorEn:'Main floor',exportRow:68},
    {cell:'D72',key:'JD_OTHER',floorVi:'Tầng khác',floorEn:'Other floors',exportRow:70}
  ]);
  const mrFloors=Object.freeze([
    {cell:'D74',key:'JD_MAIN',floorVi:'Tầng chính',floorEn:'Main floor',exportRow:68},
    {cell:'D75',key:'JD_OTHER',floorVi:'Tầng khác',floorEn:'Other floors',exportRow:70}
  ]);
  function floorsForModel(model){return model==='MR'?mrFloors:model==='MRL'?mrlFloors:[];}
  function jambType(value){
    const match=String(value??'').toUpperCase().match(/\bE-?(102|202|302)\b/);
    return match?'E-'+match[1]:'';
  }
  function limits(type){
    if(type==='E-202')return {min:130,max:600,step:5};
    if(type==='E-302')return {min:135,max:600,step:5};
    return null; // E-102 is fixed narrow jamb: no JD selector and no JD export.
  }
  function validate(model,get){
    if(!['MR','MRL'].includes(model))return [];
    const errors=[];
    for(const floor of floorsForModel(model)){
      const jamb=jambType(get(floor.cell)),range=limits(jamb);
      if(!range)continue;
      const raw=String(get(floor.key)??'').trim();
      const valid=/^\d+$/.test(raw)&&Number(raw)>=range.min&&Number(raw)<=range.max&&(Number(raw)-range.min)%range.step===0;
      if(!valid){
        errors.push({id:'JD!'+floor.key,cell:floor.key,floor:floor.floorVi,
          message:`${floor.floorVi}: ${jamb} yêu cầu JD từ ${range.min} đến ${range.max} mm, bước ${range.step} mm. Hãy chọn JD hợp lệ / Select JD ${range.min}–${range.max} mm (step ${range.step} mm).`});
      }
    }
    return errors;
  }
  function description(jd){
    const mm=String(Number(jd));
    return `Chiều rộng khung bao bản rộng JD=${mm} mm (Jamb depth JD=${mm} mm).`;
  }
  window.SPEC_JAMB_DEPTH=Object.freeze({floors:mrlFloors,floorsForModel,jambType,limits,validate,description});
})();
