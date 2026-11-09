"use strict";
// B209 is an ordered error selector, not an independent hardcoded pass flag.
// Original: SpecForm!B209 of GQXL2P workbook, linked to DataTable!D490:D527.
window.SPEC_MRL_B209=function(context, priorChecks={}){
 if(context.model!=="MRL")return {errors:[],verified:[],unverified:[],checked:{}};
 const errors=[],verified=[],unverified=[],checked={};
 const v=addr=>String(context.value(addr)??"").trim();
 const num=addr=>{const s=v(addr);return s===''?NaN:Number(s)};
 const feats=context.features||[];
 const feature=code=>feats.find(x=>String(x.code).trim().toLowerCase()===code.toLowerCase());
 const selected=code=>{const f=feature(code);return f?!!context.featureSelected(f):false};
 const record=(ref,success,cell,reason)=>{
  checked[ref]=!!success;verified.push(ref);
  if(!success)errors.push(cell+': ['+ref+'] '+reason);
 };
 const one=selected('SDE - One Side'),both=selected('SDE - Both Sides');
 const specialFinish=x=>/^SUS-HL/i.test(x)&&(/DSJ1/i.test(x)||/Large window/i.test(x));
 const susError=v('D40')!=='SUS304'&&['D46','E73','E74'].some(c=>specialFinish(v(c)));
 record('SpecForm!B209 / SUS304',!susError,'D40','DSJ1/Large Window với SUS-HL bắt buộc chọn SUS304 tại D40 / SUS304 is required for SUS-HL DSJ1/Large Window.');
 // Exact second branch of the supplied B209 formula.
 const invalidSde=(/DSJ1/i.test(v('D46'))&&((v('D15')==='CO'&&one)||(v('D15')==='2S'&&both)))||(/DS1/i.test(v('D46'))&&(one||both));
 record('SpecForm!B209 / SDE',!invalidSde,'A118/A119','SDE không hợp lệ: CO-DSJ1 chỉ None/Both; 2S-DSJ1 chỉ None/One; DS1 chỉ None / Invalid door-edge option.');
 // D520 direct translation: speed rule plus conditional requirement for SETS-L.
 const speed=num('D17'),travel=num('D35'),kg=Number(context.kg());
 const sets=selected('SETS-L');
 if(Number.isFinite(speed)&&Number.isFinite(travel)&&Number.isFinite(kg)){
  const special=speed===150&&travel>60000&&kg>=1275&&kg<=1600;
  const d520=(!sets||speed>=96)&&(!special||sets);
  record('DataTable!D520',d520,'A157',`SETS-L: tốc độ ${speed} m/min, tải trọng ${kg} kg, hành trình ${travel} mm; không thỏa điều kiện trong DataTable!D520 / SETS-L selection violates source formula.`);
 }else unverified.push('DataTable!D520 – cannot resolve Speed, Capacity or Travel');
 // D523: actual input D36/D37 and calculated minimum F36/F37, not a generic OH limit.
 const pit=num('D36'),oh=num('D37');
 const minPit=context.derived?.F36, minOh=context.derived?.F37;
 if(Number.isFinite(minPit)&&Number.isFinite(minOh)&&Number.isFinite(pit)&&Number.isFinite(oh)){
  record('DataTable!D523',pit>=minPit&&oh>=minOh,'D36/D37',`OH/PD dưới mức tối thiểu: Pit ${pit} / Min ${minPit} mm; OH ${oh} / Min ${minOh} mm / Below database minimum.`);
 }else{
  unverified.push('DataTable!D523 / SpecForm!F36:F37 – minimum OH/PD lookup unresolved; cannot report PASS');
 }
 // Special cases missing full evaluated R cells must not silently pass.
 const sourceMap={'DataTable!D524':['R23','R46','R73','R74'],'DataTable!D525':['R14']};
 for(const [ref,cells] of Object.entries(sourceMap)){
  if(cells.every(c=>typeof context.derived?.[c]==='boolean')){
   const good=cells.every(c=>context.derived[c]);
   record(ref,good,cells.join('/'),`Điều kiện ${cells.join(', ')} không đạt / Dependent workbook formula is FALSE.`);
  }else unverified.push(`${ref} – requires actual calculated ${cells.join(', ')}`);
 }
 // The priority selector can show its first certain failure; unresolved earlier branches mean
 // later errors cannot be asserted as the *first* one from B209, although still listed as failures.
 const priority=['SpecForm!B209 / SUS304','SpecForm!B209 / SDE','DataTable!D520','DataTable!D523','DataTable!D525','DataTable!D524'];
 const first=priority.find(p=>checked[p]===false);
 if(first)verified.push('SpecForm!B209 – detected '+first);
 else unverified.push('SpecForm!B209 – first-failing-condition selection incomplete until DataTable!D490:D527 is fully evaluated');
 // B210 always requires a complete count of FALSE results, not cached workbook values.
 unverified.push('SpecForm!B210 – complete COUNTIF of D490:D511,D513:D523,D525:D526 and D524');
 unverified.push('SpecForm!B206 – cannot assert NO MISTAKE without B210,D512,D527');
 return {errors,verified,unverified,checked};
};
