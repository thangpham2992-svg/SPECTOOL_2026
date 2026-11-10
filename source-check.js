'use strict';
/* Source workbook audit: every Condition row is recalculated, never passed from saved TRUE/FALSE. */
window.SPEC_SOURCE_CHECK=function(model,changes,db){
 const Engine=window.SpecExcelSourceEngine?.SourceEngine;
 if(!Engine||!db)return {errors:['SOURCE: Không tải được Excel calculation engine / Missing source engine'],warnings:[],verified:[],unverified:['Excel engine unavailable'],details:[],pass:false};
 const eng=new Engine(db,changes);
 const issues=[],unverified=[],verified=[],warnings=[],details=[],conditionCount={total:0,pass:0,fail:0,unknown:0};
 const addrPattern=/\$?([A-Z]{1,3})\$?(\d+)/g;
 const value=(sh,addr)=>eng.safe(sh,addr);
 const base=str=>typeof str==='string'?str:'';
 const featureTables=Object.fromEntries(Object.entries(db.tables||{}).filter(([k])=>/^TblFeature_/.test(k)));
 const featureByCode={};
 for(const [name,t] of Object.entries(featureTables)){
  const match=/([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(t.ref);if(!match)continue;
  for(let r=Number(match[2])+1;r<=Number(match[4]);r++){
   const key=String(db.sheets.SpecForm['B'+r]||'').toLowerCase().trim();if(key)featureByCode[key]='A'+r;
  }
 }
 const cellForCondition=(row)=>{
  if(model==='MRL'){
   if(row===522)return 'D33';if(row===523)return 'D37';if(row===524)return 'D23';
   if(row===525)return 'D14';if(row===526)return 'D62';
   if(row===520)return 'D17';if(row===519)return 'D15';if(row===521)return 'D16';if(row===513)return 'D17';
   if(row===514)return 'D35';if(row===515)return 'D18';if(row===516)return 'D41';if(row===517||row===518)return 'D13';
  }else{
   if(row===513)return 'D45';if(row===514||row===515||row===516)return 'D31';
   if(row===517)return 'D39';if(row===518)return 'D15';
  }
  for(const col of ['E','F','G','H','I']){
   const f=db.sheets.DataTable[col+row];if(!f||typeof f!=='object'||!f.f)continue;
   const m=Array.from(f.f.matchAll(/VLOOKUP\("([^"]+)"/g));
   for(const z of m){const code=String(z[1]).toLowerCase();if(featureByCode[code])return featureByCode[code];}
  }
  return 'D13';
 };
 const issue=(id,cell,message,detail='')=>{issues.push({id,cell,message:base(message),detail:base(detail)});};
 // All displayed dimension limits are calculated by the workbook's current formulas.
 // Never guess a minimum/maximum or silently treat an unresolved lookup as a valid value.
 const fmtMm=x=>Number(x).toLocaleString('vi-VN')+' mm';
 const dimensionName={OH:'Đỉnh hố thang (OH)',PD:'Chiều sâu hố PIT (PD)',AH:'Chiều rộng hố thang (AH)',BH:'Chiều sâu hố thang (BH)'};
 const dimensionMessage=(kind,actual,limit,relation='min')=>{
  const name=dimensionName[kind]||kind;
  const found=actual.ok?Number(actual.value):NaN;
  const validLimit=limit.ok?Number(limit.value):NaN;
  if(!Number.isFinite(validLimit)||validLimit<=0)return name+': chưa có dữ liệu giới hạn phù hợp Capacity / Speed / TR trong DataTable (CHECK LAYOUT).';
  if(!Number.isFinite(found))return name+': vui lòng nhập kích thước hợp lệ (tiêu chuẩn '+fmtMm(validLimit)+').';
  return relation==='max'
   ?name+' '+fmtMm(found)+' vượt quá giới hạn tối đa '+fmtMm(validLimit)+'.'
   :name+' '+fmtMm(found)+' nhỏ hơn tiêu chuẩn tối thiểu '+fmtMm(validLimit)+'.';
 };

 function checkRule(sheet,addr,kind='condition',tracked=true){
  const reference=sheet+'!'+addr;let x=value(sheet,addr);
  if(!x.ok){if(tracked)conditionCount.unknown++;unverified.push(reference+' — '+x.error);details.push({ref:reference,state:'UNKNOWN',message:x.error,cell:addr});return null;}
  if(kind==='boolean'){
   if(typeof x.value!=='boolean'){if(tracked)conditionCount.unknown++;unverified.push(reference+' — expected TRUE/FALSE; got '+JSON.stringify(x.value));details.push({ref:reference,state:'UNKNOWN',message:'Not Boolean',cell:addr});return null;}
   if(x.value){if(tracked)conditionCount.pass++;verified.push(reference);details.push({ref:reference,state:'PASS'});return true;}
   if(tracked)conditionCount.fail++;
   let row=Number(/\d+$/.exec(addr)[0]);let source=base(db.sheets.DataTable['C'+row]);let msg=source||`Rule ${reference} returned FALSE`;
   let field=sheet==='DataTable'&&addr==='AP5'?'D61':cellForCondition(row);
   if(model==='MRL'&&sheet==='DataTable'&&row===522){
    // D522 combines AH and BH. Show the exact failing field and dynamic F33/F34 minimum.
    let found=false;
    for(const [kind,addr,lookup] of [['AH','D33','F33'],['BH','D34','F34']]){
     const actual=value('SpecForm',addr),expected=value('SpecForm',lookup);
     const n=expected.ok?Number(expected.value):NaN,got=actual.ok?Number(actual.value):NaN;
     if(!Number.isFinite(n)||n<=0||!Number.isFinite(got)||got<n){
      issue(reference+'/'+kind,addr,dimensionMessage(kind,actual,expected));found=true;
     }
    }
    if(!found)issue(reference,'D33',msg);
   }else if(model==='MRL'&&sheet==='DataTable'&&row===523){
    // Report OH and PD independently; use minimum calculated by F36/F37.
    let found=false;
    for(const [kind,addr,lookup] of [['PD','D36','F36'],['OH','D37','F37']]){
     const actual=value('SpecForm',addr),expected=value('SpecForm',lookup);
     const n=expected.ok?Number(expected.value):NaN,got=actual.ok?Number(actual.value):NaN;
     if(!Number.isFinite(n)||n<=0||!Number.isFinite(got)||got<n){
      issue(reference+'/'+kind,addr,dimensionMessage(kind,actual,expected));found=true;
     }
    }
    if(!found)issue(reference,'D37',msg);
   }else if(model==='MRL'&&sheet==='DataTable'&&row===525){
    const obs=value('SpecForm','D14'),gate=value('SpecForm','R14'),ceiling=value('SpecForm','D41'),handrail=value('SpecForm','D64');
    let found=false;
    if(!gate.ok||gate.value!==true){issue(reference+' / upstream','D14','Mẫu cabin quan sát không phù hợp tải trọng hoặc loại cửa.');found=true;}
    if(ceiling.ok&&String(ceiling.value)==='FL1'){issue(reference+' / ceiling','D41','Trần FL1 không phù hợp với cabin quan sát.');found=true;}
    if(handrail.ok&&String(handrail.value)!=='YH-100SL (SUS304)'){issue(reference+' / handrail','D64','Cabin quan sát yêu cầu tay vịn YH-100SL (SUS304).');found=true;}
    if(!found)issue(reference,'D14',msg);
   }else issue(reference,field,msg,'Calculated FALSE from workbook formula ='+db.sheets[sheet][addr]?.f);
   details.push({ref:reference,state:'FAIL',message:msg,cell:field});return false;
  }
  verified.push(reference);details.push({ref:reference,state:'PASS',value:x.value});return x.value;
 }
 // Derive every Condition row dynamically from TblCondition1, not from version-specific row numbers.
 const conditionTable=db.tables?.TblCondition1;
 const tableRef=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(conditionTable?.ref||'');
 const columnNum=s=>[...s].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
 const columnText=n=>{let s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s;};
 const columnIndex=conditionTable?.columns?.findIndex(x=>String(x).toLowerCase()==='condition')??-1;
 if(!tableRef||columnIndex<0||conditionTable.sheet!=='DataTable'){
  conditionCount.unknown++;unverified.push('DataTable!TblCondition1: Cannot locate Condition column or table range');
 }else{
  const first=Number(tableRef[2])+1,last=Number(tableRef[4]),column=columnText(columnNum(tableRef[1])+columnIndex);
  for(let r=first;r<=last;r++){
   const a=column+r,cell=db.sheets.DataTable[a];conditionCount.total++;
   if(!cell||typeof cell!=='object'||!cell.f){conditionCount.unknown++;unverified.push(`DataTable!${a}: workbook Condition formula missing`);continue;}
   checkRule('DataTable',a,'boolean');
  }
 }
 if(model==='MR'){
  for(const n of [2,3,4,5,10]){
   const a='AP'+n,raw=db.sheets.DataTable[a];if(raw&&typeof raw==='object'&&raw.f){const r=checkRule('DataTable',a,'boolean',false);if(r===false){const msg=n===10?'SDE One Side không phù hợp cửa CO với DSJ1/Large Window / Invalid SDE for special CO door':n===5?'Khối lượng trang trí vượt giới hạn / Decoration weight exceeds limit':'Điều kiện cửa/vật liệu không đạt / Door finish or weight condition failed';issue('DataTable!'+a,n===10?featureByCode['sde - one side']||'D50':n===5?'D61':'D50',msg);}}}
  const result=value('DataTable','D512');
  const b207=value('SpecForm','B207');if(!b207.ok)unverified.push('SpecForm!B207: '+b207.error);
  else if(typeof b207.value==='string'&&/^Err\s*-/i.test(b207.value)){issue('SpecForm!B207','D50',b207.value);}
  // MR uses TblSpecFailure (B206:B208), not MRL's B209/B210/B206 flow.
  const b208=value('SpecForm','B208');
  if(!b208.ok||typeof b208.value!=='number')unverified.push('SpecForm!B208 — Failed to recalculate TblSpecFailure count: '+(b208.error||'non-numeric value'));
  else {verified.push('SpecForm!B208');details.push({ref:'SpecForm!B208',state:b208.value===0?'PASS':'FAIL',value:b208.value});
   if(b208.value>0)issue('SpecForm!B208','B207',`B208 = ${b208.value}: Có lỗi trong TblSpecFailure / Original workbook failure counter is not zero`);
  }
  // MR November workbook: OH/PD, special ceiling and SDE are evaluated in
  // DataTable!AP16:AP27, using B529/E521 (short/standard OH) and E524/E523
  // (short/standard PD). Older TblK21_CarOH lookup does not support L400.
  // Never hard-code a ceiling group: read the *current workbook's* formulas.
  const sourceChecks=[
    ['AP24','D41','Đỉnh hố thang (OH) không đạt yêu cầu tối thiểu'],
    ['AP25','D40','Chiều sâu hố PIT không đạt yêu cầu tối thiểu'],
    ['AP26','D45','Mẫu trần không phù hợp'],
    ['AP27','D31','Chiều cao cabin không phù hợp mẫu trần'],
    ['AP16','D50','Cửa cabin không phù hợp'],
    ['AP17','E76','Cửa tầng không phù hợp'],
    ['AP18','D50','Vật liệu cửa yêu cầu SUS304'],
    ['AP19',featureByCode['sde - one side']||'D50','SDE không phù hợp'],
    ['AP20','D61','Khối lượng trang trí vượt giới hạn'],
    ['AP21','D32','Kích thước mở cửa không phù hợp']
  ];
  for(const [ref,field,desc] of sourceChecks){
    const raw=db.sheets.DataTable[ref];
    if(!raw||typeof raw!=='object'||!raw.f){unverified.push('DataTable!'+ref+' — workbook formula missing');continue;}
    const out=value('DataTable',ref);
    if(!out.ok){unverified.push('DataTable!'+ref+' — '+out.error);continue;}
    if(typeof out.value!=='string'){unverified.push('DataTable!'+ref+' — expected an error message or empty string');continue;}
    verified.push('DataTable!'+ref);
    if(out.value.trim()){
     if(ref==='AP24'||ref==='AP25'){
      // Honor the actual MR priority: short OH B529 over E521, short PD E524 over E523.
      const isOH=ref==='AP24',kind=isOH?'OH':'PD',actual=value('SpecForm',field);
      const order=isOH?['B529','E521']:['E524','E523'];
      let lookup=null;
      for(const candidate of order){const result=value('DataTable',candidate);
       if(result.ok&&typeof result.value==='number'&&Number.isFinite(result.value)&&result.value>0){lookup=result;break;}}
      issue('DataTable!'+ref,field,lookup?dimensionMessage(kind,actual,lookup):desc+': '+out.value.trim());
     }else issue('DataTable!'+ref,field,desc+': '+out.value.trim());
    }
  }
  for(const ref of ['E521','B529','E523','E524']){
    const out=value('DataTable',ref);
    if(!out.ok){unverified.push('DataTable!'+ref+' — '+out.error);continue;}
    verified.push('DataTable!'+ref);
  }
  if(!result.ok||typeof result.value!=='boolean')unverified.push('DataTable!D512 aggregate not verifiable');
 }else{
  // Recalculate all required helper formulas and the workbook's source outcome.
  for(const addr of ['R13','R14','R23','R46','R73','R74','F33','F34','F36','F37']){
   const r=value('SpecForm',addr);if(!r.ok)unverified.push('SpecForm!'+addr+' — '+r.error);else verified.push('SpecForm!'+addr);
  }
  for(const a of ['D527','D512']){
   const r=value('DataTable',a);if(!r.ok||typeof r.value!=='boolean')unverified.push(`DataTable!${a} — `+(r.error||'not boolean'));
   else {verified.push('DataTable!'+a);details.push({ref:'DataTable!'+a,state:r.value?'PASS':'FAIL',message:'Workbook aggregate result'});} 
  }
  for(const a of ['B209','B210','B206']){
   const r=value('SpecForm',a);if(!r.ok)unverified.push(`SpecForm!${a} — `+r.error);
   else{verified.push('SpecForm!'+a);details.push({ref:'SpecForm!'+a,state:'CALCULATED',value:r.value});
    if(a==='B209'&&typeof r.value==='string'&&r.value.startsWith('Err -'))issue('SpecForm!B209',guessB209Cell(r.value),r.value,db.sheets.SpecForm.B209?.f||'');
    if(a==='B210'&&(typeof r.value!=='number'||r.value!==0))issue('SpecForm!B210','D13',`B210 = ${r.value}: Có điều kiện FALSE / One or more false conditions`);
    if(a==='B206'&&r.value!=='NO MISTAKE')issue('SpecForm!B206','D13',`Check Spec = ${r.value} (không đạt / not approved)`);
   }
  }
 }
 // Input Feature applies/options: use the Excel formula for Option (E) when available.
 for(const [code,addr] of Object.entries(featureByCode)){
  const r=Number(addr.slice(1));const apply=value('SpecForm',addr);const option=value('SpecForm','E'+r);
  if(!apply.ok||!option.ok){unverified.push('SpecForm!'+addr+'/E'+r+' — '+(apply.error||option.error));continue;}
  const isOn=apply.value===1||apply.value===true||String(apply.value).toUpperCase()==='V';
  if(isOn&&String(option.value).trim().toUpperCase()==='X')issue('SpecForm!E'+r,addr,`${db.sheets.SpecForm['B'+r]} không được áp dụng trong Group Control hiện tại / Feature unavailable (Option X)`);
  verified.push('SpecForm!E'+r);
 }
 // MRL JJ: the configured width is selectable, but only among supported values.
 // Other capacities use the default width declared in TblK5_Capacity.
 if(model==='MRL'){
  const capRes=value('SpecForm','D13'),jjRes=value('SpecForm','D31');
  if(!capRes.ok||!jjRes.ok)unverified.push('SpecForm!D31 — cannot evaluate JJ/Capacity');
  else {
   const cap=String(capRes.value),jj=Number(jjRes.value);
   const t=db.tables?.TblK5_Capacity,match=/^([A-Z]+)(\d+):[A-Z]+(\d+)$/.exec(t?.ref||'');
   if(!match)unverified.push('DataTable!TblK5_Capacity — missing JJ defaults');
   else{
    const low=Number(match[2])+1,high=Number(match[3]);let base=null;
    for(let r=low;r<=high;r++)if(String(db.sheets.DataTable['B'+r])===cap){base=Number(db.sheets.DataTable['G'+r]);break;}
    if(!Number.isFinite(base))unverified.push('DataTable!TblK5_Capacity — no JJ value for '+cap);
    else{
     const allowed=[base];for(let ri=1081;ri<=1100;ri++){if(String(db.sheets.DataTable['CA'+ri]??'')===cap){const v=Number(db.sheets.DataTable['CB'+ri]);if(Number.isFinite(v))allowed.push(v);}}
     if(!allowed.includes(jj))issue('SpecForm!D31','D31','Chiều rộng cửa JJ không phù hợp với tải trọng đã chọn.');
     else verified.push('SpecForm!D31 — JJ theo Capacity');
    }
   }
  }
 }
 // Additional non-formula CheckSpecStep1 checks recovered from the legacy C# source.
 // The original hatch-pattern UI state cannot be reconstructed from formula values alone,
 // so its coverage is explicitly distinguished from the 37 Excel Condition results.
 const legacy=window.SPEC_LEGACY_STEP1?.(model,changes,db);
 if(legacy){
  for(const problem of legacy.errors){issue(problem.id,problem.cell,problem.message,problem.detail);}
  for(const finding of legacy.verified)verified.push(finding);
  for(const advisory of legacy.warnings||[])warnings.push(advisory);
  for(const gap of legacy.unverified){
   if(!gap.includes('dynamic black-hatch')&&!gap.includes('LightUp'))unverified.push(gap);
  }
 }else unverified.push('Legacy Tool.exe CheckSpecStep1 is unavailable');
 // The final legacy Tool.exe step checks named-range black hatch / required inputs.
 // These rules are extracted from Excel conditionalFormatting rather than guessed from field labels.
 const hatch=window.SPEC_CF_CHECK?.(model,changes,db,eng);
 if(hatch){for(const x of hatch.errors)issues.push(x);for(const u of hatch.unknown)unverified.push(u);for(const v of hatch.verified)verified.push(v);}
 else unverified.push('Conditional Formatting / LightUp checker not loaded');
 // No silent full PASS: errors in unknown cells make verification INCOMPLETE / BLOCKED.
 const unique=[];const seen=new Set();for(const issueItem of issues){const key=issueItem.id+'|'+issueItem.message;if(!seen.has(key)){unique.push(issueItem);seen.add(key)}}
 const status=unique.length?'FAIL':unverified.length?'INCOMPLETE':'PASS';
 return {status,pass:status==='PASS',errors:unique,warnings,verified:[...new Set(verified)],unverified:[...new Set(unverified)],details,conditions:conditionCount,source:'Excel DataTable & SpecForm + legacy C# Step1 checks',excelConditionsPass:conditionCount.fail===0&&conditionCount.unknown===0,legacyVerified:!!legacy,legacyGaps:legacy?.unverified||[],calls:eng.callCount};
 function guessB209Cell(message){if(/SDE|door edge/i.test(message))return featureByCode['sde - one side']||'D46';if(/SUS304/.test(message))return 'D46';if(/OH\/PD|Pit /i.test(message))return 'D37';if(/SETS-L/.test(message))return featureByCode['sets-l']||'D17';if(/Observation/.test(message))return 'D14';if(/Additional inputs/.test(message))return 'D23';return 'D13';}
};
