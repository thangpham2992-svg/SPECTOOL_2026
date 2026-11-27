'use strict';
/* Port of the non-formula rules in FormSpecGXLM2P / FormSpecGQXL2P CheckSpecStep1_ItemClick.
   Excel formulas and DataTable conditions are evaluated separately in source-check.js.
   UI/conditional-fill-only checks are not treated as certified unless we can observe their state. */
// Shared PIV family classification: enclosure suffix does not change display count.
window.SPEC_PIV_RULES=(()=>{
 const families={A1010:[1,'SEGMENT'],A1020:[2,'SEGMENT'],A1030:[1,'DOT'],A1040:[2,'DOT'],A1080:[1,'DOT'],A1090:[2,'DOT'],C710:[1,'SEGMENT'],C720:[2,'SEGMENT'],C730:[1,'DOT'],C740:[2,'DOT'],C766:[1,'LCD'],C776:[2,'LCD'],C780:[1,'DOT'],C790:[2,'DOT']};
 const classify=code=>{const m=/^PIV[1-6][- ]([AC]\d{3,4})[A-Z0-9]*$/i.exec(String(code??'').trim());const info=m&&families[m[1].toUpperCase()];return info?{count:info[0],technology:info[1]}:null;};
 const requiresSingle=(model,group)=>model==='MRL'?['1C-2BC','3C-AI22'].includes(group):group==='1C-2BC';
 const allowed=(model,group,code)=>!requiresSingle(model,group)||classify(code)?.count===1;
 return {classify,requiresSingle,allowed};
})();
window.SPEC_LEGACY_STEP1=function(model,changes,db){
 const Engine=window.SpecExcelSourceEngine?.SourceEngine;
 if(!Engine||!db)return{errors:[],verified:[],unverified:['Legacy C# engine unavailable']};
 const calc=new Engine(db,changes),errors=[],verified=[],unverified=[],warnings=[];
 const nameCells=db.names?.local?.SpecForm||{},raw=(s,a)=>{const result=calc.safe(s,a);return result.ok?result.value:null};
 const namedAddr=n=>{const f=nameCells[n];const m=/SpecForm!\$?([A-Z]+)\$?(\d+)/.exec(f||'');return m?m[1]+m[2]:null;};
 const n=n=>{const addr=namedAddr(n);return addr===null?null:raw('SpecForm',addr)};
 const num=v=>{const x=Number(v);return Number.isFinite(x)?x:NaN};
 const str=v=>String(v??'').trim();
 const has=(a,b)=>str(a).includes(b);
 const mark=(name,description)=>{const addr=namedAddr(name)||'D13';errors.push({id:'Tool.exe CheckSpecStep1',cell:addr,message:description,detail:name+' (C# CheckSpecStep1_ItemClick)'});};
 const tbl=(table)=>{
  const t=db.tables?.[table];if(!t)return null;
  const m=/([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(t.ref);if(!m)return null;
  const cn=s=>Array.from(s).reduce((n,c)=>26*n+c.charCodeAt(0)-64,0), ct=x=>{let s='';while(x){x--;s=String.fromCharCode(65+x%26)+s;x=Math.floor(x/26)}return s;};
  const col=cn(m[1]);const out=[];
  for(let r=Number(m[2])+1;r<=Number(m[4]);r++){
   const row={};t.columns.forEach((k,j)=>row[k]=raw(t.sheet,ct(col+j)+r));out.push(row);
  }
  return out;
 };
 const f=name=>str(n(name));const capacity=f('K5_Capacity');const carHeight=num(n('K17_CarH'));
 const speed=num(n('K8_Speed'));
 const travel=num(n('K21_CarTR'));
 const stops=num(n('K9_Stops'));
 const group=f('K13_Operation');
 // Service floor is a free-text project description, intentionally independent of Stops.
 // Do not parse floor labels, enforce count, special-floor groups, or validate its syntax.
 // Preserve the exact entered string in SpecForm and Export.
 verified.push('Service floor: free-text input, not checked against Stops');
 // 2. Core dimension and travel checks. Workbook's newest formulas have priority for OH/PD in MRL.
 const capacities=tbl('TblK5_Capacity');
 const row=capacities?.find(x=>str(x.EleCapacityName)===capacity);
 if(!row)unverified.push('Tool.exe: no matching TblK5_Capacity row for '+capacity);
 else{
  const capKg=num(row.Capacity),family=['S300','N300','N300S','L310','N120','N130','N140','FEVN10'];
  const ceiling=f('K22_CeilingModel');
  if(carHeight===2500&&family.includes(ceiling))mark('K22_CeilingModel',`HL=2500 không dùng được với trần ${ceiling} / Ceiling not compatible with HL=2500`);
  const mm=x=>Number(x).toLocaleString('vi-VN')+' mm';
  const shaftName={AH:'Chiều rộng hố thang (AH)',BH:'Chiều sâu hố thang (BH)'};
  const isBedMR=model==='MR'&&/^(?:B750|B1000)-2S(?:\s|\()/i.test(capacity);
  for(const [code,dimension] of [['AH','K19_CarAH'],['BH','K19_CarBH']]){
   // BED rows were appended as static records; their N-Min values are not formula states.
   // BED envelopes are validated dynamically by mr-bed-capacity.js, preventing permanent false alarms.
   if(isBedMR)continue;
   const state=str(row[code+'-Valid']),limit=num(row[code+'-Valid-Value']),actual=num(n(dimension));
   // MRL V02 BH special limits are checked against the selected CWT DataTable below.
   // Avoid stale workbook N-Max when 1D-1G Mitsubishi Standard permits +440 mm,
   // or when 1D-2G requires equality with its CWT-specific source dimension.
   const ent=str(raw('SpecForm','D16')),reg=str(raw('SpecForm','D26'));
   if(model==='MRL'&&code==='BH'&&(ent==='1D-2G'||(ent==='1D-1G'&&/Mitsubishi\s*Standard|MELCO\s*std\.?/i.test(reg))))continue;
   // These state codes are workbook-calculated against model/door/fire/CWT-specific ranges.
   if(state==='N-Max'&&Number.isFinite(limit))mark(dimension,`${shaftName[code]} ${mm(actual)} vượt quá giới hạn tối đa ${mm(limit)}.`);
   if(state==='N-Min'&&Number.isFinite(limit))mark(dimension,`${shaftName[code]} ${mm(actual)} nhỏ hơn tiêu chuẩn tối thiểu ${mm(limit)}.`);
  }
  if(model==='MRL'){
   // V02 Mitsubishi Standard: 1D-1G BH max is source CWT-specific BH min + 440 mm.
   // 1D-2G is an exact-source-dimension layout, regardless of regulation.
   // The CWT dataset gives the correct BH for model/speed/JJ/entrance/fire/CWT.
   const entrance=f('K7_EntranceType')||str(raw('SpecForm','D16'));
   const regulation=f('K16_Regulation')||str(raw('SpecForm','D26'));
   const isMitsubishiStandard=/Mitsubishi\s*Standard|MELCO\s*std\.?/i.test(regulation);
   const strictTwoGate=entrance==='1D-2G';
   const standardSingleGate=entrance==='1D-1G'&&isMitsubishiStandard;
   if(strictTwoGate||standardSingleGate){
    const lookup=window.SPEC_CWT_ENGINE?.lookupEngine?.(calc);
    const minimumBH=lookup?.ok?Number(lookup.BH):NaN;
    const actualBH=num(n('K19_CarBH'));
    if(Number.isFinite(minimumBH)&&minimumBH>0&&Number.isFinite(actualBH)){
     const maxBH=strictTwoGate?minimumBH:minimumBH+440;
     if(strictTwoGate&&actualBH!==minimumBH){
      errors.push({id:'MRL!2G_BH_Exact',cell:namedAddr('K19_CarBH')||'D34',
       message:`1D-2G: BH ${mm(actualBH)} phải bằng đúng DataTable ${mm(minimumBH)} (không áp dụng +440 mm).`,
       detail:'MRL Version02: 1D-2G exact BH from the selected CWT layout'});
     }else if(standardSingleGate&&actualBH>maxBH){
      errors.push({id:'MRL!Standard1G_BHMax',cell:namedAddr('K19_CarBH')||'D34',
       message:`BH ${mm(actualBH)} vượt BH Max ${mm(maxBH)} (BH Min DataTable ${mm(minimumBH)} + 440 mm). Khuyến nghị bổ sung lưới/vách phía sau để giảm BH khai báo đặt hàng, đảm bảo điều kiện an toàn theo mục 5.4.7.4 TCVN 6396-20. Phương án kết cấu và kích thước BH hữu dụng phải được xác nhận trước khi đặt hàng.`,
       detail:'MRL Version02 / Mitsubishi Standard / 1D-1G: BH Max = CWT DataTable BH Min + 440 mm; safety illustration attached'});
     }
     verified.push('MRL Version02 '+entrance+': BH Min '+minimumBH+' mm, BH Max '+maxBH+' mm (source CWT dataset)');
    }else unverified.push('MRL '+entrance+': cannot determine CWT-specific BH minimum for BH maximum/exact validation');
   }
   // DataTable has AH-Max / BH-Max (and fire-door-specific maxima) even when
   // AH-Valid/BH-Valid returns MANUAL GQ LAYOUT. Enforce the explicitly declared
   // table envelope; exceeding it requires a different factory layout.
   const isFire=!/^Not suppl(?:ied|y)$/i.test(str(n('K36_FireproofDoor')));
   for(const [code,dimension] of [['AH','K19_CarAH'],['BH','K19_CarBH']]){
    if(code==='BH'&&(strictTwoGate||standardSingleGate))continue;
    const column=code+'-Max'+(isFire?'-FIRE':'');
    const max=num(row[column]),actual=num(n(dimension));
    if(Number.isFinite(max)&&max>0&&Number.isFinite(actual)&&actual>max){
     mark(dimension,`${shaftName[code]} ${mm(actual)} vượt quá giới hạn tối đa ${mm(max)} theo DataTable (${isFire?'cửa chống cháy':'cửa thường'}).`);
    }
    // A blank AH-Max/BH-Max in the source is an intentionally unspecified
    // upper envelope, not a failed/unverified formula. Do not invent a max.
    // Minimum dimensions and counterweight-specific layouts remain enforced
    // by DataTable D522 and SPEC_CWT_ENGINE.
   }
  }
  if(model==='MR'){
   // NEXIEZ-MR only: side counterweight + BH above the source maximum.
   // Keep the existing hard BH maximum failure, and attach one safety recommendation
   // to that same failure. Never infer a maximum from blank workbook cells.
   const sideDrop=/CWT\s*Side|Side\s*Drop/i.test(capacity)||/Side/i.test(str(row['CWT-Position']));
   const bhActual=num(n('K19_CarBH'));
   const fireSelection=str(n('K36_FireproofDoor'));
   const hasFire=fireSelection!==''&&!/^(?:Not supplied|Not supply|No|Without|-)$/i.test(fireSelection);
   const maxColumn=hasFire?'BH-Max-FIRE':'BH-Max';
   const sourceMax=num(row[maxColumn]);
   const alternateMax=num(row['BH-Max']);
   const declaredMax=Number.isFinite(sourceMax)&&sourceMax>0?sourceMax:
    (!hasFire&&Number.isFinite(alternateMax)&&alternateMax>0?alternateMax:NaN);
   const state=str(row['BH-Valid']);
   const stateLimit=num(row['BH-Valid-Value']);
   const bhMax=Number.isFinite(declaredMax)?declaredMax:
    (state==='N-Max'&&Number.isFinite(stateLimit)&&stateLimit>0?stateLimit:NaN);
   if(sideDrop&&Number.isFinite(bhActual)&&Number.isFinite(bhMax)&&bhActual>bhMax){
    errors.push({id:'MR!SideDropBHMax',cell:namedAddr('K19_CarBH')||'D35',
     message:`BH ${mm(bhActual)} vượt BH Max ${mm(bhMax)}. Khuyến nghị bổ sung lưới/vách phía sau để giảm BH khai báo đặt hàng, đảm bảo điều kiện an toàn theo mục 5.4.7.4 TCVN 6396-20. Phương án kết cấu và kích thước BH hữu dụng phải được xác nhận trước khi đặt hàng.`,
     detail:'MR Side Drop + BH > BH Max | reference image: balustrade 700 mm min; toe-board 100 mm min; engineering verification required'});
   }
   if(capKg===450&&speed>60)mark('K8_Speed','450 kg không áp dụng tốc độ 90/105 m/min / Restricted speed for 450 kg');
   const minTR=speed===60?3400:speed===90?5700:speed===105?7300:null,maxTR=speed===60?60000:80000;
   if(minTR!==null&&travel<minTR)mark('K21_CarTR',`Hành trình nâng ${travel} thấp hơn ${minTR} mm / Travel below minimum`);
   if(travel>maxTR)mark('K21_CarTR',`Hành trình nâng ${travel} vượt ${maxTR} mm / Travel exceeds maximum`);
   const pit=num(n('K21_CarPD'));
   const pitIssue=(min,max)=>{
    if(!Number.isFinite(pit))return;
    if(pit<min)mark('K21_CarPD',`Chiều sâu hố PIT (PD) ${mm(pit)} nhỏ hơn tiêu chuẩn tối thiểu ${mm(min)}.`);
    if(pit>max)mark('K21_CarPD',`Chiều sâu hố PIT (PD) ${mm(pit)} vượt quá giới hạn tối đa ${mm(max)}.`);
   };
   if(speed===60){const min=(capKg===1350&&!capacity.includes('P20C4'))?1300:1250;pitIssue(min,2500);}
   if(speed===90&&capKg!==450)pitIssue(1340,2800);
   if(speed===105&&capKg!==450){const min=(capKg===1350&&!capacity.includes('P20C4'))?1410:1380;pitIssue(min,2900);}
   // Decoration weight: DataTable AP5 is formula-derived and already included in source-check.
  }else{
   let minTR=null;
   if(capKg<=1050){if(group==='2C-2BC')minTR=10000;else if(group==='3C-AI22')minTR=10500;else if(group==='1C-2BC')minTR=speed===60?3400:[96,105].includes(speed)?7300:null;}
   else minTR=speed===60?3400:[96,105].includes(speed)?7300:null;
   if(minTR!==null&&travel<minTR)mark('K21_CarTR',`Hành trình nâng ${travel} thấp hơn ${minTR} mm / Travel below minimum for group operation`);
   const pitState=str(row['PD-Valid']),pitLimit=num(row['PD-Valid-Value']);
   if(pitState==='N-Max'&&Number.isFinite(pitLimit))mark('K21_CarPD',`Chiều sâu hố PIT (PD) ${mm(num(n('K21_CarPD')))} vượt quá giới hạn tối đa ${mm(pitLimit)}.`);
   else if(pitState==='N-Min'&&Number.isFinite(pitLimit))mark('K21_CarPD',`Chiều sâu hố PIT (PD) ${mm(num(n('K21_CarPD')))} nhỏ hơn tiêu chuẩn tối thiểu ${mm(pitLimit)}.`);
   // The new workbook handles F36/F37, D523; decoration is a C# Step1-only branch.
   // Use the workbook's *calculated* TblK5_Capacity columns rather than fixed engineering limits.
   const decorName=db.names?.local?.DataTable?.K29_DecorWeightTotal||db.names?.global?.K29_DecorWeightTotal;
   const decorRef=/DataTable!\$?([A-Z]+)\$?(\d+)/.exec(decorName||'');
   const decorValue=decorRef?calc.safe('DataTable',decorRef[1]+decorRef[2]):{ok:false,error:'Missing K29_DecorWeightTotal'};
   const decor=decorValue.ok?Number(decorValue.value):NaN;
   const cwt1=str(row['Check-DecorWeight-CWT1']),cwt2=str(row['Check-DecorWeight-CWT2']),inc=str(row['Check-DecorWeight-Inc']);
   const shaftCWT2=str(row['AH-Valid'])==='AH2'&&str(row['BH-Valid'])==='BH2';
   const ceilingFamily1=['S00','L210','L210S'];const ceilingFamily2=['S300','N300','N300S','L310','N120','N130','N140'];
   const limit=ceilingFamily1.includes(ceiling)?num(row['DecorWeight-CWT1-S00']):ceilingFamily2.includes(ceiling)?num(row['DecorWeight-CWT1-S300']):NaN;
   const incLimit=num(row['Increase-DecorWeight']);
   const desc='Tool.exe CheckSpecStep1 — CWT decoration calculation from TblK5_Capacity';
   const problem=reason=>errors.push({id:desc,cell:namedAddr('K29_DecorWeight')||'D57',message:reason,detail:'FormSpecGQXL2P.cs lines 837–968'});
   const warning=reason=>warnings.push({cell:namedAddr('K29_DecorWeight')||'D57',message:reason,detail:desc});
   const overweight=()=>problem(`Khối lượng trang trí ${decor} kg vượt giới hạn / Total decoration weight is over limit`);
   const needDeeperPD=optional=>problem('Khối lượng trang trí yêu cầu tăng độ sâu hố PIT'+(optional?' và đối trọng tùy chọn':'')+' / Decoration requires deeper pit'+(optional?' with optional CWT':''));
   const needOptionalCWT=()=>problem('Cần đối trọng tùy chọn nhưng kích thước AH/BH không phù hợp / Optional CWT requires different AH/BH layout');
   const optionalWarn=()=>warning('Khối lượng trang trí yêu cầu đối trọng tùy chọn / Optional counterweight required for decoration');
   if(!Number.isFinite(decor)||!cwt1||!cwt2){unverified.push('Tool.exe Decoration Weight: missing formula result / CWT1/CWT2 classification');}
   else{
    const possiblePD=Number.isFinite(limit)&&decor<=limit;
    const checkPD=()=>{if(!Number.isFinite(limit))unverified.push('Tool.exe Decoration Weight: missing allowable decoration threshold for '+ceiling);return possiblePD;};
    if(cwt1==='X'&&cwt2==='N')overweight();
    else if(cwt1==='N-PD'){
     if(cwt2==='Y'){
      if(shaftCWT2){optionalWarn();if(checkPD())warning('Cần tăng chiều sâu pit với đối trọng tiêu chuẩn / Deeper pit required with standard CWT');}
      else{needOptionalCWT();if(checkPD())needDeeperPD(false);}
     }else if(cwt2==='N'){
      if(shaftCWT2){if(inc==='N-PD'){
       if(!Number.isFinite(incLimit))unverified.push('Tool.exe Decoration Weight: missing increase limit');
       else if(decor<=incLimit)needDeeperPD(true);else overweight();
      }else overweight();}
      else needOptionalCWT();
      if(checkPD())needDeeperPD(false);
     }
    }else if(cwt1==='N'){
     if(capKg===825||capKg===1350){
      if(inc==='N')overweight();
      else if(inc==='N-PD'){
       if(!Number.isFinite(incLimit))unverified.push('Tool.exe Decoration Weight: missing increase limit');
       else if(cwt2==='Y'){
        if(shaftCWT2){optionalWarn();if(decor<=incLimit)warning('Cần tăng chiều sâu pit / Deeper pit required');}
        else{needOptionalCWT();if(decor>incLimit)overweight();else needDeeperPD(false);}
       }else if(cwt2==='N'){if(decor>incLimit)overweight();else needDeeperPD(false);}
      }
     }else{
      if(cwt2==='X')overweight();
      else if(cwt2==='Y'){if(shaftCWT2)optionalWarn();else needOptionalCWT();}
      else if(cwt2==='N'){
       if(shaftCWT2){if(inc==='X'||inc==='N')overweight();else if(inc==='Y')optionalWarn();}
       else needOptionalCWT();
      }
     }
    }
    verified.push('Tool.exe: Decoration weight CWT1/CWT2/Increase classification (TblK5_Capacity)');
   }

  }
  verified.push('Tool.exe: Capacity/shaft dimensions, travel and minimum pit (model-specific)');
 }
 // 3. Fixture checks per original Tool.exe.
 const pivKeys=model==='MRL'?['TopFloor','MainFloor','OtherFloor']:['MainFloor','OtherFloor'];
 const notSup=s=>s.includes('Not supplied');
 
 for(const key of pivKeys){
  const P='K36_PIV'+key,H='K36_HBV'+key,I='K36_PIH_PID'+key;
  const piv=f(P),hbv=f(H),pih=f(I);
  if(!notSup(piv)&&!window.SPEC_PIV_RULES.allowed(model,group,piv)){const info=window.SPEC_PIV_RULES.classify(piv);mark(P,info?`PIV ${piv}: ${info.count} màn hình; hệ điều khiển ${group} yêu cầu 1 màn hình / Single indicator required.`:`PIV ${piv}: chưa xác định số màn hình / Unknown PIV family.`);}
  // Exception confirmed by Sales Engineering: PIV may share a landing with
  // HBV1-F716N / HBV2-F716N / HBV3-F716N / HBV4-F716N only.
  // Attribute any other simultaneous PIV/HBV failure to HBV, the conflicting fixture.
  const pivWithHBVException=/^HBV[1-4]-F716N$/i.test(hbv);
  if(!notSup(piv)&&!notSup(hbv)&&!pivWithHBVException)
   mark(H,`Tầng ${key}: PIV không dùng đồng thời với HBV ${hbv}. Chỉ cho phép HBV1/2/3/4-F716N. / PIV cannot share this floor with HBV except HBV1/2/3/4-F716N.`);
  if(!notSup(piv)&&!notSup(pih)&&(model==='MRL'||key==='MainFloor'))mark(P,`Không được đồng thời sử dụng PIV và PIH/PID tại ${key} / Incompatible hall fixtures`);
  if(notSup(piv)&&notSup(hbv))mark(P,`Phải chọn PIV hoặc HBV tại ${key} / At least one hall fixture is required`);
 }
 if(carHeight>2300&&['K23_WallFinish','K24_CarDoorFinish','K25_CarTransomFinish'].some(nm=>f(nm)==='Pattern-printed steel sheet PPS'))
  mark('K17_CarH','Không dùng thép PPS khi HL > 2300 mm / PPS is unavailable above 2300 mm');
 verified.push('Tool.exe: Hall fixture coexistence, Group Control fixtures, PPS ceiling restrictions');
 // Original C# checked black hatch pattern / required named ranges with DevExpress layout state.
 // This state depends on live conditional formatting and is not serialized as a straightforward formula.
 unverified.push('Tool.exe: dynamic black-hatch (LightUp) input lock and required-cell paint state is not reconstructed in HTML; static input rules remain enforced in the UI');
 return{errors,verified,unverified,warnings};
};
