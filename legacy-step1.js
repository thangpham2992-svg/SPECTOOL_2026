'use strict';
/* Port of the non-formula rules in FormSpecGXLM2P / FormSpecGQXL2P CheckSpecStep1_ItemClick.
   Excel formulas and DataTable conditions are evaluated separately in source-check.js.
   UI/conditional-fill-only checks are not treated as certified unless we can observe their state. */
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
 const tableFloors=tbl('TblK10_SFloor');
 // 1. Validate floor names, count, COP indicator compatibility and special-floor groups.
 if(!tableFloors){unverified.push('Tool.exe: TblK10_SFloor not found');}
 else{
  const service=f('K10_SFloor').replace(/\s+/g,'');
  let floorCount=0,valid=true,groupName='';
  const byName=new Map(tableFloors.map(r=>[str(r.FloorNameButton),r]));
  const bracket=(s)=>/^([1-9]|[12]\d|3[0-6])$/.test(s);
  const pivs=model==='MRL'?['K36_PIVTopFloor','K36_PIVMainFloor','K36_PIVOtherFloor']:['K36_PIVMainFloor','K36_PIVOtherFloor'];
  for(const token of service.split(',')){
   const range=token.split('-');
   if(!token){valid=false;break;}
   if(range.length===1){
    const record=byName.get(token);
    if(!record&&!bracket(token)){valid=false;break;}
    floorCount++;
    if(!record)continue;
    const copCode=f('K31_COPCode'),copButton=f('K31_COPButton');
    const no=k=>str(record[k]).toUpperCase()==='N';
    if((['760','710'].includes(copCode)&&no('SegmentLED'))||(['730','780'].includes(copCode)&&no('LCD')))mark('K31_COPCode',`Loại COP ${copCode} không dùng được tại tầng ${token} / COP type incompatible with floor label`);
    if(no('TactileButton')&&['CBV1','CBV3','CBV5'].some(c=>copButton.includes(c)))mark('K31_COPButton',`Nút Braille/tactile ${copButton} không áp dụng tầng ${token} / Tactile button unavailable`);
    if(no('FlatButton')&&['CBV2','CBV4','CBV6'].some(c=>copButton.includes(c)))mark('K31_COPButton',`Nút phẳng ${copButton} không áp dụng tầng ${token} / Flat button unavailable`);
    for(const piv of pivs){const p=f(piv);
     if((['71','72','1010','1020'].some(x=>p.includes(x))&&no('SegmentLED'))||(['766','776'].some(x=>p.includes(x))&&no('LCD')))
      mark(piv,`PIV ${p} không áp dụng tầng ${token} / Hall indicator incompatible with floor label`);
    }
    if(f('K31_BrailleLetter')==='Supplied'&&no('Braille'))mark('K31_BrailleLetter',`Braille không hỗ trợ tầng ${token} / Unsupported floor label`);
    const grp=str(record.Group);if(grp){if(groupName&&groupName!==grp)mark('K10_SFloor',`Không cùng nhóm tầng đặc biệt tại ${token} / Mixed special-floor groups`);else groupName=grp;}
    if(model==='MR'&&no('AAN-G')){
     const sig=tbl('TblFeature_Signal');if(sig?.some(r=>str(r.Feature)==='AAN-G'&&Number(r.Apply)===1&&str(r.Option)!=='X'))mark('K10_SFloor',`Feature AAN-G không áp dụng tầng ${token} / AAN-G inapplicable`);
    }
   }else if(range.length===2&&bracket(range[0])&&bracket(range[1])&&Number(range[0])<Number(range[1]))floorCount+=Number(range[1])-Number(range[0])+1;
   else {valid=false;break;}
  }
  if(!valid)mark('K10_SFloor','Tên tầng không hợp lệ theo Tool.exe / Invalid floor names or range syntax');
  if(Number.isFinite(stops)&&floorCount!==stops)mark('K9_Stops',`Số tầng phục vụ ${floorCount} khác số điểm dừng ${stops} / Floor count differs from stops`);
  verified.push('Tool.exe: Floor names, COP/LED/Braille, service-floor count');
 }
 // 2. Core dimension and travel checks. Workbook's newest formulas have priority for OH/PD in MRL.
 const capacities=tbl('TblK5_Capacity');
 const row=capacities?.find(x=>str(x.EleCapacityName)===capacity);
 if(!row)unverified.push('Tool.exe: no matching TblK5_Capacity row for '+capacity);
 else{
  const capKg=num(row.Capacity),family=['S300','N300','N300S','L310','N120','N130','N140','FEVN10'];
  const ceiling=f('K22_CeilingModel');
  if(carHeight===2500&&family.includes(ceiling))mark('K22_CeilingModel',`HL=2500 không dùng được với trần ${ceiling} / Ceiling not compatible with HL=2500`);
  for(const [code,dimension] of [['AH','K19_CarAH'],['BH','K19_CarBH']]){
   const state=str(row[code+'-Valid']),limit=num(row[code+'-Valid-Value']),actual=num(n(dimension));
   if(state==='N-Max'&&Number.isFinite(limit))mark(dimension,`${code} = ${actual} mm vượt giới hạn Max ${limit} mm / Hoistway dimension exceeds maximum`);
   if(state==='N-Min'&&Number.isFinite(limit))mark(dimension,`${code} = ${actual} mm thấp hơn Min ${limit} mm / Hoistway dimension below minimum`);
  }
  if(model==='MR'){
   if(capKg===450&&speed>60)mark('K8_Speed','450 kg không áp dụng tốc độ 90/105 m/min / Restricted speed for 450 kg');
   const minTR=speed===60?3400:speed===90?5700:speed===105?7300:null,maxTR=speed===60?60000:80000;
   if(minTR!==null&&travel<minTR)mark('K21_CarTR',`Hành trình nâng ${travel} thấp hơn ${minTR} mm / Travel below minimum`);
   if(travel>maxTR)mark('K21_CarTR',`Hành trình nâng ${travel} vượt ${maxTR} mm / Travel exceeds maximum`);
   const pit=num(n('K21_CarPD'));
   if(speed===60){const min=(capKg===1350&&!capacity.includes('P20C4'))?1300:1250;if(pit<min||pit>2500)mark('K21_CarPD',`PD ${pit} mm ngoài khoảng ${min}–2500 mm / Pit depth outside allowed range`);}
   if(speed===90&&capKg!==450){if(pit<1340||pit>2800)mark('K21_CarPD',`PD ${pit} mm ngoài khoảng 1340–2800 mm / Pit depth outside allowed range`);}
   if(speed===105&&capKg!==450){const min=(capKg===1350&&!capacity.includes('P20C4'))?1410:1380;if(pit<min||pit>2900)mark('K21_CarPD',`PD ${pit} mm ngoài khoảng ${min}–2900 mm / Pit depth outside allowed range`);}
   // Decoration weight: DataTable AP5 is formula-derived and already included in source-check.
  }else{
   let minTR=null;
   if(capKg<=1050){if(group==='2C-2BC')minTR=10000;else if(group==='3C-AI22')minTR=10500;else if(group==='1C-2BC')minTR=speed===60?3400:[96,105].includes(speed)?7300:null;}
   else minTR=speed===60?3400:[96,105].includes(speed)?7300:null;
   if(minTR!==null&&travel<minTR)mark('K21_CarTR',`Hành trình nâng ${travel} thấp hơn ${minTR} mm / Travel below minimum for group operation`);
   const pitState=str(row['PD-Valid']),pitLimit=num(row['PD-Valid-Value']);
   if(pitState==='N-Max'&&Number.isFinite(pitLimit))mark('K21_CarPD',`PD vượt Max ${pitLimit} mm / Pit depth exceeds limit`);
   else if(pitState==='N-Min'&&Number.isFinite(pitLimit))mark('K21_CarPD',`PD thấp hơn Min ${pitLimit} mm / Pit depth below limit`);
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
 const requireSingle=model==='MRL'?['1C-2BC','3C-AI22'].includes(group):group==='1C-2BC';
 for(const key of pivKeys){
  const P='K36_PIV'+key,H='K36_HBV'+key,I='K36_PIH_PID'+key;
  const piv=f(P),hbv=f(H),pih=f(I);
  if(requireSingle&&!notSup(piv)&&!['71','73','766','1010','1030'].some(c=>piv.includes(c)))mark(P,`PIV ${piv} không áp dụng điều khiển nhóm ${group} / Invalid hall fixture for control group`);
  if(!notSup(piv)&&!notSup(hbv))mark(P,`Không được đồng thời sử dụng PIV và HBV tại ${key} / PIV and HBV cannot coexist`);
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
