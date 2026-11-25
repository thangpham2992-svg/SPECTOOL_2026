 'use strict';
/* Source-aligned, deterministic SpecFormExport builder for NEXIEZ ALPHA workbooks.
   Source C#: FormSpecGXLM2P.cs and FormSpecGQXL2P.cs, barButtonExportSpec_ItemClick,
   DeleteFeature. Evaluate formulas against currently loaded workbook; no cached pass values.
   The only user-facing worksheet is SpecFormExport. A4 PDF is produced using browser
   native print / Save as PDF from the SAME frozen workbook export model. */
(()=>{
 const n='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 const r='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
 const pkg='http://schemas.openxmlformats.org/package/2006/relationships';
 const text=v=>v==null?'':typeof v==='boolean'?(v?'TRUE':'FALSE'):String(v);
 const parse=s=>new DOMParser().parseFromString(String(s).replace(/^\uFEFF/, '').replace(/^\s+(?=<\?xml\b)/, ''),'application/xml');
 const serial=d=>new XMLSerializer().serializeToString(d);
 // Namespace-aware, WebKit-safe traversal. Do not rebuild any format/style XML.
 const all=(node,tag)=>{
   if(!node)return [];
   const list=Array.from(node.getElementsByTagNameNS?.(n,tag)||[]);
   return list.length?list:Array.from(node.getElementsByTagName('*')).filter(el=>el.localName===tag);
 };
 const child=(node,tag)=>Array.from(node?.childNodes||[]).find(x=>x.localName===tag)||null;
 const cellRow=address=>Number(/\d+$/.exec(address)?.[0]||0);
 const safe=(obj,path)=>obj?.[path];
 const excelName=x=>String(x||'').replace(/[<>:"/\\|?*]/g,'_').slice(0,85);
 function nameAddr(db,scope,name){
   const reference=db.names.local?.[scope]?.[name]||db.names.global?.[name];
   const match=/!?\$?([A-Z]+)\$?(\d+)/.exec(reference||'');
   return match?match[1]+match[2]:null;
 }
 function nameRange(db,scope,name){
   const reference=db.names.local?.[scope]?.[name]||db.names.global?.[name];
   const match=/!?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/.exec(reference||'');
   return match?[+match[2],+match[4]]:null;
 }
 const on=v=>v===1||v===true||['1','V','✓','TRUE'].includes(text(v).toUpperCase());
 // User-provided NEXIEZ_MR_Duty_Table.xlsx: Type A, Capacity B, Speed C, outputs E/F/G.
 // Distinguish B750 from Passenger at 750kg/60mpm: rope counts differ.
 const MR_DUTY_TABLE={"Passenger|450|60":{"sheave":410,"motor":2.8,"rope":"Ø10 × 3"},"Passenger|550|60":{"sheave":410,"motor":3.7,"rope":"Ø10 × 3"},"Passenger|550|90":{"sheave":410,"motor":5.6,"rope":"Ø10 × 3"},"Passenger|550|105":{"sheave":410,"motor":6.5,"rope":"Ø10 × 3"},"Passenger|600|60":{"sheave":410,"motor":3.7,"rope":"Ø10 × 3"},"Passenger|600|90":{"sheave":410,"motor":5.6,"rope":"Ø10 × 3"},"Passenger|600|105":{"sheave":410,"motor":6.5,"rope":"Ø10 × 3"},"Passenger|700|60":{"sheave":410,"motor":4.6,"rope":"Ø10 × 3"},"Passenger|700|90":{"sheave":410,"motor":6.9,"rope":"Ø10 × 3"},"Passenger|700|105":{"sheave":410,"motor":8.1,"rope":"Ø10 × 3"},"Passenger|750|60":{"sheave":410,"motor":4.6,"rope":"Ø10 × 3"},"Passenger|750|90":{"sheave":410,"motor":6.9,"rope":"Ø10 × 4"},"Passenger|750|105":{"sheave":410,"motor":8.1,"rope":"Ø10 × 4"},"Passenger|900|60":{"sheave":410,"motor":5.6,"rope":"Ø10 × 4"},"Passenger|900|90":{"sheave":410,"motor":8.3,"rope":"Ø10 × 4"},"Passenger|900|105":{"sheave":410,"motor":9.7,"rope":"Ø10 × 4"},"Passenger|1000|60":{"sheave":410,"motor":6.2,"rope":"Ø10 × 4"},"Passenger|1000|90":{"sheave":410,"motor":9.2,"rope":"Ø10 × 4"},"Passenger|1000|105":{"sheave":410,"motor":11,"rope":"Ø10 × 4"},"Passenger|1150|60":{"sheave":500,"motor":7.1,"rope":"Ø10 × 4"},"Passenger|1150|90":{"sheave":500,"motor":11,"rope":"Ø10 × 4"},"Passenger|1150|105":{"sheave":500,"motor":13,"rope":"Ø10 × 4"},"Passenger|1350|60":{"sheave":500,"motor":8.3,"rope":"Ø10 × 5"},"Passenger|1350|90":{"sheave":500,"motor":13,"rope":"Ø10 × 5"},"Passenger|1350|105":{"sheave":500,"motor":15,"rope":"Ø10 × 5"},"B750|750|60":{"sheave":410,"motor":4.6,"rope":"Ø10 × 4"},"B750|750|90":{"sheave":410,"motor":6.9,"rope":"Ø10 × 4"},"B750|750|105":{"sheave":410,"motor":8.1,"rope":"Ø10 × 4"},"B1000|1000|60":{"sheave":410,"motor":6.2,"rope":"Ø10 × 4"},"B1000|1000|90":{"sheave":410,"motor":9.2,"rope":"Ø10 × 4"},"B1000|1000|105":{"sheave":410,"motor":11,"rope":"Ø10 × 4"}};
 function mrDutyBySpec(eng){
   const carCode=text(eng.safe('SpecForm','D13').value);
   const kgMatch=/\((\d+(?:[.,]\d+)?)\s*KG\)/i.exec(carCode);
   const capacity=kgMatch?Number(kgMatch[1].replace(',','.')):NaN;
   const speed=Number(eng.safe('SpecForm','D15').value);
   const special=/\bB(?:750|1000)\b/i.exec(carCode);
   const kind=special?special[0].toUpperCase():'Passenger';
   const found=Number.isFinite(capacity)&&Number.isFinite(speed)?MR_DUTY_TABLE[`${kind}|${capacity}|${speed}`]:null;
   return {kind,capacity,speed,found:found||null};
 }

 // EN/MELCO-only motor powers from G11L1_Motor_Power_All_Capacity_Speed.xlsx.
 // Exact Capacity + Speed lookup; no India or Regulation fallback.
 const MRL_MOTOR_EN_MELCO={"450|60":2.8,"450|90":4.2,"450|96":4.5,"450|105":4.9,"550|60":3.7,"550|90":5.6,"550|96":6.2,"550|105":6.5,"630|60":3.7,"630|90":5.6,"630|96":6.2,"630|105":6.5,"825|60":5.1,"825|90":7.6,"825|96":8.1,"825|105":8.9,"825|120":11,"825|150":13,"1000|60":6.2,"1000|90":9.2,"1000|96":11,"1000|105":11,"1000|120":13,"1000|150":16,"1050|60":6.2,"1050|90":9.2,"1050|96":11,"1050|105":11,"1050|120":13,"1050|150":17,"1275|60":7.9,"1275|90":12,"1275|96":13,"1275|105":14,"1275|120":16,"1275|150":20,"1350|60":8.9,"1350|90":13,"1350|96":15,"1350|105":15,"1350|120":17,"1350|150":21,"1600|60":9.9,"1600|90":15,"1600|96":16,"1600|105":18,"1600|120":20,"1600|150":25,"1800|60":12,"1800|90":17,"1800|96":18,"1800|105":20};
 function mrlMotorBySpec(eng){
   const carCode=text(eng.safe('SpecForm','D13').value);
   const kgMatch=/\((\d+(?:[.,]\d+)?)\s*KG\)/i.exec(carCode);
   const capacity=kgMatch?Number(kgMatch[1].replace(',','.')):NaN;
   // MRL V02: D15 is Door type; D17 is Speed (MPM). MR uses D15.
   const speed=Number(eng.safe('SpecForm','D17').value);
   const key=`${capacity}|${speed}`;
   const motor=MRL_MOTOR_EN_MELCO[key];
   return {capacity,speed,found:Number.isFinite(motor)?motor:null};
 }

 const isNotSupplied=v=>/^\s*Not suppl(?:ied|y)\s*$/i.test(text(v));
 function currentResult(model,changes,db){
   if(!window.SPEC_SOURCE_CHECK)throw Error('Thiếu Source Check Engine / Validation engine not loaded');
   const result=window.SPEC_SOURCE_CHECK(model,changes,db);
   if(result.status!=='PASS')throw Error('Chỉ xuất khi Check Spec = PASS / Export blocked: '+result.status+'. '+result.errors.slice(0,3).map(e=>e.message).join(' | ')+' '+result.unverified.slice(0,2).join(' | '));
   return result;
 }
 function exportModel(model,changes,db){
   if(!db?.sheets?.SpecFormExport||!db?.tables)throw Error('Thiếu SpecFormExport hoặc DataTable / Workbook source missing');
   currentResult(model,changes,db);
   const eng=new window.SpecExcelSourceEngine.SourceEngine(db,changes);
   // JD is a Webtool-only supplemental field, not an Excel DataTable formula.
   // Validate it before any export is built, even when exportModel is invoked directly.
   if(['MR','MRL'].includes(model)&&window.SPEC_JAMB_DEPTH){
     const jdErrors=window.SPEC_JAMB_DEPTH.validate(model,key=>{
       if(key==='JD_MAIN'||key==='JD_OTHER')return changes[key];
       return eng.safe('SpecForm',key).value;
     });
     if(jdErrors.length)throw Error(jdErrors.map(e=>e.message).join(' | '));
   }

   const cells={},hiddenRows=new Set(),overrides={},sections={},unresolved=[];
   const output=db.sheets.SpecFormExport;
   for(const [addr,entry] of Object.entries(output)){
     if(entry&&typeof entry==='object'&&Object.hasOwn(entry,'f')){
       const result=eng.safe('SpecFormExport',addr);
       if(result.ok)cells[addr]=result.value;
       else {cells[addr]='';unresolved.push({addr,formula:entry.f,error:result.error})}
     }else cells[addr]=entry?.v??entry??'';
   }
   // v5.8.152: B750/B1000 special bed-lift handrail YH-56S (SUS304)
   // is valid in Check Spec but absent from the legacy TblK32_Handrail lookup.
   // Resolve only that known exception, leaving other #N/A errors visible.
   if(model==='MR'){
     const rail=String(eng.safe('SpecForm','D67').value??'').trim();
     const car=String(eng.safe('SpecForm','D13').value??'');
     const missing=unresolved.find(e=>e.addr==='C62'&&/VLOOKUP failed:\s*YH-56S/i.test(e.error));
     if(missing&&/^B(?:750|1000)-2S\b/i.test(car)&&rail==='YH-56S (SUS304)'){
       const location=String(eng.safe('SpecForm','D68').value??'').trim();
       const loc={
         '1 pc on rear wall':['1 tay vịn trên vách sau','1 pc on rear wall'],
         '1 pc on right side wall':['1 tay vịn trên vách phải','1 pc on right side wall'],
         '1 pc on left side wall':['1 tay vịn trên vách trái','1 pc on left side wall'],
         '2 pcs on both side walls':['2 tay vịn trên hai vách bên','2 pcs on both side walls'],
         '2 pcs on rear & right side wall':['2 tay vịn trên vách sau và vách phải','2 pcs on rear and right side walls'],
         '2 pcs on rear & left side wall':['2 tay vịn trên vách sau và vách trái','2 pcs on rear and left side walls'],
         '3 pcs on 3 walls':['3 tay vịn trên ba vách','3 pcs on three walls']
       }[location];
       if(loc){
         cells.C62='Tay vịn inox SUS304, mã YH-56S, '+loc[0]+'.\n'+
                   '(SUS304 stainless steel handrail, model YH-56S, '+loc[1]+').';
         unresolved.splice(unresolved.indexOf(missing),1);
       }
     }
   }
   // v5.8.125: use the same corrected bilingual label in MR/MRL Excel
   // and the PDF, which is generated from the same export worksheet.
   // Match the exact old heading only; do not rewrite technical fire options.
   for(const addr of Object.keys(cells)){
     if(/^Cửa chống cháy\s*\(Fireproof door\)\s*$/i.test(text(cells[addr]).trim()))
       cells[addr]='Giới hạn chịu lửa cửa tầng\n(Entrance Fire resistance / Fire rating)';
   }
   // Explicit rated duration requested for the EN81-58(II) option in both
   // NEXIEZ-MR and NEXIEZ-MRL Version2. The source SpecFormExport omits
   // the duration; keep the remaining options (No, GHCL E30) unchanged.
   const fireInputAddr=nameAddr(db,'SpecForm','K36_FireproofDoor');
   const fireExportAddr=nameAddr(db,'SpecFormExport','K36_FireproofDoor');
   if(fireInputAddr&&fireExportAddr){
     const fireSelected=text(eng.safe('SpecForm',fireInputAddr).value).trim();
     if(/^EN81[-\s]*58\s*\(II\)$/i.test(fireSelected)){
       cells[fireExportAddr]=
        'Cửa chống cháy EN81-58(II), khả năng chịu lửa 120 phút.\n'+
        '(Fire-rated landing door EN81-58(II), 120-minute fire resistance.)';
     }
   }
   const featureGroups=['Emergency','Door','Operation','Signal','Traffic'];
   const removedFeatures=[],includedFeatures=[];
   for(let index=0;index<featureGroups.length;index++){
     const group=featureGroups[index],input=db.tables['TblFeature_'+group];
     const xrange=nameRange(db,'SpecFormExport','K'+(37+index)+'_'+group+'Features');
     if(!input||!xrange)throw Error('Thiếu Feature Table hoặc Named Range: '+group+' / Missing export mapping');
     const bounds=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(input.ref);
     if(!bounds)throw Error('Invalid Feature Table range '+input.ref);
     let visibleCount=0;
     for(let row=+bounds[2]+1;row<=+bounds[4];row++){
       const code=text(eng.safe('SpecForm','B'+row).value).trim();if(!code)continue;
       // MBSS signal-type door sensor is discontinued. Never include it in exports.
       if(code.toUpperCase()==='MBSS'){for(let er=xrange[0];er<=305;er++)if(text(cells['D'+er]).trim().toUpperCase()==='MBSS')hiddenRows.add(er);removedFeatures.push({group,code,row});continue;}
       const apply=eng.safe('SpecForm','A'+row),option=eng.safe('SpecForm','E'+row);
       if(!apply.ok||!option.ok)throw Error('Không tính được Feature '+code+' tại dòng '+row+' / Cannot evaluate feature state');
       const present=on(apply.value)&&text(option.value).toUpperCase()!=='X';
       const featureRows=[];
       for(let er=xrange[0];er<=xrange[1];er++)if(text(cells['D'+er]).trim()===code)featureRows.push(er);
       if(!featureRows.length){
         // Some updated Excel templates extend a feature table by two rows but
         // leave its old export Named Range unchanged (e.g., MRL VIP-S).
         // Match the feature code in the export block to prevent leakage of Option X.
         for(let er=xrange[0];er<=305;er++)if(text(cells['D'+er]).trim()===code)featureRows.push(er);
       }
       if(!featureRows.length)continue; // Code has no export row in this workbook.
       if(present){visibleCount++;includedFeatures.push({group,code,row,option:text(option.value),exportRows:featureRows});}
       else{removedFeatures.push({group,code,row});for(const er of featureRows)hiddenRows.add(er);}
       // Retain D/E values for all features, including hidden ones, so Excel Unhide preserves the source rows.
     }
     if(!visibleCount){const heading=nameAddr(db,'SpecFormExport','K'+(37+index)+'_Title_'+group+'Features');if(heading)hiddenRows.add(cellRow(heading));}
     sections[group]={shown:visibleCount};
   }
   // Hall indicators/buttons: follow the exact K36 named inputs and export row mappings in Tool.exe.
   const halls=model==='MRL'?['TopFloor','MainFloor','OtherFloor']:['MainFloor','OtherFloor'];
   let hiddenHall=0;
   for(let i=0;i<halls.length;i++){
     const floor=halls[i];
     for(const type of (model==='MR'&&floor==='OtherFloor'?['HBV','PIV']:['PIH_PID','HBV','PIV'])){
       const inputAddr=nameAddr(db,'SpecForm','K36_'+type+floor);
       const exportAddr=nameAddr(db,'SpecFormExport','K36_'+type+floor);
       if(!inputAddr||!exportAddr)throw Error('Hall indicator mapping missing: '+type+floor);
       const result=eng.safe('SpecForm',inputAddr);
       if(!result.ok)throw Error('Cannot evaluate hall indicator '+inputAddr+': '+result.error);
       if(isNotSupplied(result.value)){
         hiddenRows.add(cellRow(exportAddr));hiddenHall++;
         if(type==='PIV'){
           // K36_Title... shifts label after the PIV row is removed (Tool.exe behavior).
           const titleAddr=nameAddr(db,'SpecFormExport','K36_TitleHallIndicatorButtons'+floor);
           const valueNum=model==='MRL'?i+2:i+2;
           const helperAddr=nameAddr(db,'SpecFormExport','K36_Value'+valueNum);
           if(titleAddr&&helperAddr)overrides[titleAddr]=cells[helperAddr];
         }
       }
     }
   }
   for(let i=1;i<=(model==='MRL'?4:3);i++){
     const a=nameAddr(db,'SpecFormExport','K36_Value'+i);
     if(a)cells[a]='';
   }
   Object.assign(cells,overrides);
   // Include selected Hall Lanterns in the SAME entrance-design floor groups as
   // PIV/HBV/PIH, with complete Vietnamese and English faceplate descriptions.
   // Both XLSX and A4 PDF use these frozen SpecFormExport cells (no separate PDF text).
   const hallLanternExportRows=[];
   if(model==='MRL'&&window.SPEC_HALL_LANTERN){
     const hall=window.SPEC_HALL_LANTERN;
     for(const [floor,key,row] of [['TopFloor','HL_TOP',76],['OtherFloor','HL_OTHER',82],['MainFloor','HL_MAIN',79]]){
       const choice=hall.selection(changes,key);
       if(choice===hall.none)continue;
       const description=hall.exportDescription(choice);
       if(!description)throw Error('Hall Lantern material information missing for '+choice);
       const addr='C'+row;
       const original=text(cells[addr]).trim();
       // The floor's PIV can be Not supplied even while a Hall Lantern is present.
       // Do not show an inaccurate "Not supplied" immediately before the lantern.
       const keepOriginal=!/^(?:Không cung cấp\s*\(Not supplied\)|Not supplied)\.?$/i.test(original);
       cells[addr]=[keepOriginal?original:'',description.vi,description.en].filter(Boolean).join('\n');
       hiddenRows.delete(row);
       hallLanternExportRows.push(row);
     }
   }
   // SUS-HL/SUS-M are material FINISHES, never paint codes.
   // Override only ceiling bilingual description rows; leave all paint colors intact.
   const finishCell=model==='MR'?'D46':'D42',ceilingCell=model==='MR'?'D45':'D41';
   const ceilingName=text(eng.safe('SpecForm',ceilingCell).value).trim();
   const ceilingFinish=text(eng.safe('SpecForm',finishCell).value).trim().toUpperCase();
   const ceilingOutputRow=model==='MR'?46:45;
   if(ceilingName && /^(SUS-HL|SUS-M)$/.test(ceilingFinish)){
     const mirror=ceilingFinish==='SUS-M';
     cells['C'+ceilingOutputRow]=`Trần phòng thang kiểu ${ceilingName}, ${mirror?'inox gương':'inox sọc nhuyễn'} (${ceilingFinish}).`;
     cells['C'+(ceilingOutputRow+1)]=`(Car ceiling model ${ceilingName}, ${mirror?'mirror-finish':'hairline-finish'} stainless steel ceiling (${ceilingFinish})).`;
   }
   // Factory exception for both models: DSJ1, DS1 and Window door assemblies
   // use SUS304, independent of the other car/landing-door finishes.
   // Only the door's own bilingual export text changes; never affect jambs,
   // transom panels, or front returns.
   {
     const read=addr=>text(eng.safe('SpecForm',addr).value);
     const doorInputs=model==='MRL'?
       [['D46',[49,50]],['E73',[72,73]],['E74',[74,75]]]:
       [['D50',[50,51]],['E76',[72,73]],['E77',[74,75]]];
     for(const [input,rows] of doorInputs){
       if(!/\b(?:DSJ1|DS1|WINDOW)\b/i.test(read(input)))continue;
       for(const row of rows){
         const key='C'+row;
         cells[key]=text(cells[key]).replace(/SUS430J1L|SUS430/gi,'SUS304');
       }
     }
   }
   // MRL Version02: export the live (and manually overridable) granite weight,
   // not a cached SpecFormExport result from the source XLSX.
   if(model==='MRL'){
     const floorType=text(eng.safe('SpecForm','D55').value).trim();
     if(floorType==='Local granite (20 mm)'){
       const weightResult=eng.safe('SpecForm','D56');
       // SourceEngine reads user overrides first. Retain a manually entered weight.
       const AA=Number(eng.safe('SpecForm','D28').value),BB=Number(eng.safe('SpecForm','D29').value),JJ=Number(eng.safe('SpecForm','D31').value);
       const computed=Number.isFinite(AA)&&Number.isFinite(BB)&&Number.isFinite(JJ)?Math.floor(AA*BB*0.02*2.7*0.001+0.5)+(JJ<1000?5:10):NaN;
       const userWeight=Object.hasOwn(changes,'D56')?Number(changes.D56):(Number.isFinite(Number(weightResult.value))?Number(weightResult.value):computed);
       if(Number.isFinite(userWeight)&&userWeight>=0){
         const w=Number.isInteger(userWeight)?String(userWeight):String(userWeight);
         cells.C56=`Sàn đá granite cung cấp trong nước, chiều dày hoàn thiện sau thi công 20mm, khối lượng ${w} Kg (Local natural granite, finishing thickness 20 mm, weight ${w} Kg).`;
       }
     }
   }
   // Mirror weight adjustment is calculated in Excel only; do not append note in Export.
   // MR FE-Hongkong (FE-C30): the historical template omits a dedicated row.
   // In this variant the existing FE bilingual export slot represents FE-C30.
   // Avoid adding rows to the XLSX layout or mislabeling standard FE.
   if(model==='MR' && ['1','TRUE','V','✓'].includes(text(changes.FE_C30).toUpperCase())){
     hiddenRows.delete(93);hiddenRows.delete(94);
     cells.A93='(FE-C30) FE-Hongkong with MELCO switch';
     cells.C93='Firefighter emergency operation — Hong Kong configuration with MELCO switch (FE-C30).';
     cells.A94='Vận hành khẩn cấp chữa cháy FE Hongkong (FE-C30)';
     cells.C94='Vận hành cứu hỏa theo cấu hình FE Hongkong, sử dụng công tắc MELCO (FE-C30).';
     includedFeatures.push({group:'Emergency',code:'FE-C30',row:null,option:'A',exportRows:[93,94]});
   }
   // Hidden/manual-only Version02 feature definitions are used for source checks and
   // availability, not shown in the customer-facing SpecFormExport appendix.
   // Explicitly selected customer deliverables (e.g. ACC or NSCR-H) are handled below.
   // ACC: supplementary ITV-dependent traveling power cable, for both models.
   // Existing SpecFormExport rows, merges, source styles and feature categories are untouched.
   if(['1','TRUE','V','✓'].includes(text(changes.AC_TRAVEL_CABLE).toUpperCase())){
     const source=db.sheets.SpecForm;
     const itv=Object.entries(source).find(([k,v])=>/^B\d+$/.test(k)&&/^(ITV|ITV Cable)$/.test(text(v)));
     const itvOn=itv&&on(changes['A'+itv[0].slice(1)]??source['A'+itv[0].slice(1)]);
     if(!itvOn)throw Error('ACC requires ITV / Cáp động điều hòa yêu cầu ITV');
     cells.A296='VI. Tính năng bổ sung / Additional feature';
     cells.A297='ACC – Aircondition Cabin';
     cells.C297='Power supply traveling cable for the elevator car air conditioner, inside the hoistway.';
     cells.A298='Cáp động cấp nguồn điều hòa cabin';
     cells.C298='Cáp động cấp nguồn cho điều hòa cabin, nằm bên trong giếng thang (hố thang máy).';
     includedFeatures.push({group:'Signal',code:'ACC',row:null,option:'A',exportRows:[297,298]});
   }
   // NSCR-H supplementary feature: preserve original source export rows and styles.
   // Rows 299–300 are unused in the standard feature print range; show only when selected.
   if(on(changes.NSCR_H)){
     if(!cells.A296)cells.A296='VI. Tính năng bổ sung / Additional features';
     cells.A299='NSCR-H – Non-service Temporary Release for Hall Call - Card Reader Type';
     cells.C299='Hall calls at restricted floors can be registered only by placing a card over a card reader. This function is automatically deactivated during emergency operation. Only dry contact is supplied; card reader and card are excluded.';
     cells.A300='Hạn chế phục vụ cuộc gọi ở tầng – Dạng thẻ từ';
     cells.C300='Các cuộc gọi ở tầng bị hạn chế chỉ được chấp nhận khi sử dụng thẻ từ. Chức năng tự động vô hiệu hóa trong chế độ khẩn cấp. Chỉ cung cấp tiếp điểm khô, không bao gồm đầu đọc thẻ và thẻ từ.';
     includedFeatures.push({group:'Operation',code:'NSCR-H',row:null,option:'A',exportRows:[299,300]});
   }
   // v5.8.106 — MRL Version02 Observation: the source workbook maps
   // SpecFormExport!C16 to SpecForm!D12 (Use=Passenger), so an NY-40
   // observation selection in D14 is omitted from both PDF and XLSX.
   // Override only the export presentation; do not change underlying D12.
   if(model==='MRL'){
     const observation=text(eng.safe('SpecForm','D14').value).trim();
     const matched=/^NY[\s-]*40\s*([A-E])$/i.exec(observation);
     if(matched){
       cells.C16=`Thang máy quan sát (Observation) kiểu NY40${matched[1].toUpperCase()}`;
     }
   }
   // Export presentation only (5.8.87): model name is a product name, not the
   // workbook's detailed configuration code. Both XLSX and PDF use doc.cells.
   const exportModelRow=model==='MR'?13:14;
   cells['C'+exportModelRow]=model==='MR'?'NEXIEZ-MR':'NEXIEZ-MRL Version2';
   // v5.8.91 — The MRL SpecFormExport speed lookup is occasionally evaluated
   // as a bare number by the browser Excel interpreter. Preserve the complete
   // bilingual wording defined by G11L1P DataTable TblK8_Speed[SpeedVN].
   // For non-listed future speeds use mpm/60 (m/s) without changing Spec inputs.
   if(model==='MRL'){
     const raw=eng.safe('SpecForm','D17');
     const mpm=Number(raw?.value);
     const sourceSpeed={
       60:'60 m/phút - 1 m/giây (60 mpm - 1 m/s)',
       96:'96 m/phút - 1.6 m/giây (96 mpm - 1.6 m/s)',
       105:'105 m/phút - 1.75 m/giây (105 mpm - 1.75 m/s)',
       120:'120 m/phút - 2 m/giây (120 mpm - 2 m/s)',
       150:'150 m/phút - 2.5 m/giây (150 mpm - 2.5 m/s)'
     };
     if(raw?.ok&&Number.isFinite(mpm)&&mpm>0){
       const rate=Number((mpm/60).toFixed(4)).toString();
       cells.C20=sourceSpeed[mpm]||`${mpm} m/phút - ${rate} m/giây (${mpm} mpm - ${rate} m/s)`;
     }
   }
   // Obsolete MBSS: hide the bilingual pair even if the historic MR workbook
   // contains it in SpecFormExport but no longer lists it as an active feature.
   // Do not remove unrelated door-sensor features.
   if(model==='MR'){
     for(let rn=1;rn<=305;rn++){
       if(text(cells['D'+rn]).trim().toUpperCase()==='MBSS'||
          /^\s*\(MBSS\)(?:\s|$)/i.test(text(cells['A'+rn]))){
         hiddenRows.add(rn);
         if(text(cells['D'+(rn+1)]).trim().toUpperCase()==='MBSS')hiddenRows.add(rn+1);
       }
     }
   }
   // v4.2: Remove the legacy decorative-weight exclusion note from any export.
   for(const addr of Object.keys(cells))if(typeof cells[addr]==='string')cells[addr]=cells[addr]
     .replace(/\s*[-–]?\s*Trừ\s*:\s*Handrail\s*&\s*mirror\s*nhập khẩu\s*;\s*sàn\s*\(By local\)/gi,'');
   // MRL entrance design: use the same simple device/floor hierarchy as MR,
   // instead of multi-row vertical merges that spill unpredictably across A4 pages.
   // All values still come from the selected SpecFormExport workbook (not guesses).
   if(model==='MRL'){
     const read=a=>text(eng.safe('SpecForm',a).value).trim();
     const obs=read('D14');
     // Observation NY-40A: stainless-steel hairline car wall with aluminum
     // framed glazing. Do not describe the observation-glass frame as SUS-HL.
     // Other NY-40 types retain their own source-workbook description.
     if(obs==='NY-40A'){
       cells.C47='Vách phòng thang kiểu NY-40A bằng inox sọc nhuyễn (SUS-HL), kính quan sát với khung/nẹp nhôm.';
       cells.C48='(NY-40A observation car wall in hairline-finish stainless steel (SUS-HL), with aluminum-framed observation glass.)';
     }
     // v5.8.91 — Human-readable bilingual handrail specification. Build one
     // description per source K32_Handrail and K32_HandrailLocation; do not
     // duplicate the code in parentheses or strip the Vietnamese description.
     const handrail=read('D64'),position=read('D65');
     const positionVi={
       '1 pc on rear wall':'1 tay vịn trên vách sau',
       '1 pc on right side wall':'1 tay vịn trên vách phải',
       '1 pc on left side wall':'1 tay vịn trên vách trái',
       '2 pcs on both side walls':'2 tay vịn trên 2 vách bên',
       '2 pcs on rear & right side wall':'2 tay vịn trên vách sau và bên phải',
       '2 pcs on rear & left side wall':'2 tay vịn trên vách sau và bên trái',
       '3 pcs on 3 walls':'3 tay vịn trên 3 vách'
     };
     const kind=/^YH-100SL(?:\s|\(|$)/i.test(handrail)?'round':
       /^YH-57S(?:\s|\(|$)/i.test(handrail)?'flat':null;
     if(kind){
       const code=kind==='round'?'YH-100SL':'YH-57S';
       const viType=kind==='round'?'Tay vịn tròn':'Tay vịn dẹt';
       const enType=kind==='round'?'Round':'Flat';
       const viPlace=positionVi[position]||position;
       const vi=viType+' inox sọc nhuyễn SUS304 mã hiệu '+code+(viPlace?', '+viPlace:'');
       const en=enType+' hairline-finish stainless steel SUS304 handrail type '+code+(position?', '+position:'');
       cells.C62=vi+' ('+en+').';
     }else if(/^(?:Not supplied|No)$/i.test(handrail)){
       cells.C62='Không cung cấp (Not supplied).';
     }else if(handrail){
       // Keep any other option's source-workbook bilingual description;
       // never assume other models have the same material or cross-section.
       const prior=text(cells.C62).trim();
       if(!prior||prior===handrail)cells.C62=handrail+(position?', '+position:'')+'.';
     }
     const mirror=read('D66');
     if(mirror&&!/^(?:Not supplied|No)$/i.test(mirror)&&
        /^\s*([^\n]+?)\s*\(\s*\1\s*\)\.?\s*$/.test(text(cells.C63)))cells.C63=mirror+'.';
     if(/^(?:Not supplied|No)$/i.test(mirror)&&/Not supplied\s*\(Not supplied\)/i.test(text(cells.C63)))
       cells.C63='Không cung cấp (Not supplied).';

     // Each jamb/landing entry occupies one bilingual row, without a merged
     // cell that spans subsequent rows or page boundaries.
     const body=(vi,en)=>[text(vi).trim(),text(en).trim()].filter(Boolean).join('\n');
     for(const [rn,following,heading,floor] of [
       [68,69,'Khung cửa tầng (Jamb Model)','Tầng chính (Main floor)'],
       [70,71,'Khung cửa tầng (Jamb Model)','Tầng khác (Other floors)'],
       [72,73,'Cửa tầng (Landing door)','Tầng chính (Main floor)'],
       [74,75,'Cửa tầng (Landing door)','Tầng khác (Other floors)']]){
       cells['C'+rn]=body(cells['C'+rn],cells['C'+following]);
       cells['A'+rn]=heading;cells['B'+rn]=floor;
       hiddenRows.add(following);
     }
     // Append JD in both languages to the selected wide jamb, independently by floor.
     // E-102 must not show a JD value or a JD text line.
     if(window.SPEC_JAMB_DEPTH){
       for(const floor of window.SPEC_JAMB_DEPTH.floorsForModel(model)){
         const jamb=window.SPEC_JAMB_DEPTH.jambType(eng.safe('SpecForm',floor.cell).value);
         if(!window.SPEC_JAMB_DEPTH.limits(jamb))continue;
         const jd=String(changes[floor.key]);
         cells['C'+floor.exportRow]=text(cells['C'+floor.exportRow]).trimEnd()+'\n'+window.SPEC_JAMB_DEPTH.description(jd);
       }
     }
     // Three independent floor groups. Combine PIV, HBV, PIH and optional
     // Hall Lantern per floor only once; keep every selected device description.
     const hallGroups=[
       [76,[76,77,78],'Tầng trên cùng (Top floor)'],
       [79,[79,80,81],'Tầng chính (Main floor)'],
       [82,[82,83,84],'Tầng khác (Other floors)']];
     for(const [main,members,floor] of hallGroups){
       const description=members.filter(rn=>!hiddenRows.has(rn))
         .map(rn=>text(cells['C'+rn]).trim())
         .filter(v=>v&&!/^(?:Not supplied|Không cung cấp(?: \(Not supplied\))?)\.?$/i.test(v))
         .filter((v,i,array)=>array.indexOf(v)===i).join('\n');
       for(const rn of members.slice(1))hiddenRows.add(rn);
       cells['A'+main]='Hiển thị và nút bấm (Hall indicators & buttons)';
       cells['B'+main]=floor;
       cells['C'+main]=description;
       if(description)hiddenRows.delete(main);else hiddenRows.add(main);
     }
     cells.A87='V. Các tính năng (Features) :';
   }
   // NEXIEZ-MR uses the same Jamb output rows C68/C70 but does not enter the
   // MRL-specific combined Entrance Design layout above. Append JD here for MR.
   if(model==='MR'&&window.SPEC_JAMB_DEPTH){
     for(const floor of window.SPEC_JAMB_DEPTH.floorsForModel(model)){
       const jamb=window.SPEC_JAMB_DEPTH.jambType(eng.safe('SpecForm',floor.cell).value);
       if(!window.SPEC_JAMB_DEPTH.limits(jamb))continue;
       const jd=text(changes[floor.key]).trim();
       cells['C'+floor.exportRow]=text(cells['C'+floor.exportRow]).trimEnd()+'\n'+window.SPEC_JAMB_DEPTH.description(jd);
     }
   }
   // v5.8.109: Selected hall-panel quantity is exported below its corresponding
   // panel model on both Excel and PDF. Counts are NOT written into SpecForm.
   // Do this AFTER MRL's three-floor consolidation so the note survives it.
   const qtyHalls=model==='MRL'?['TopFloor','MainFloor','OtherFloor']:['MainFloor','OtherFloor'];
   const opValue=text(eng.safe('SpecForm',model==='MR'?'D21':'D22').value).toUpperCase();
   const groupMatch=opValue.match(/(?:^|\D)([1-4])\s*C(?:\D|$)/);
   const liftCount=groupMatch?Number(groupMatch[1]):1;
   const qtyDetails=[];
   for(const floor of qtyHalls){
     const notes=[];
     for(const kind of ['PIV','HBV']){
       const inputAddr=nameAddr(db,'SpecForm','K36_'+kind+floor);
       if(!inputAddr)continue;
       const raw=text(eng.safe('SpecForm',inputAddr).value);
       if(!raw||isNotSupplied(raw))continue;
       const count=Number(changes['HALL_QTY_'+inputAddr]??1);
       const family=raw.match(/PIV\d*[- ]?([A-Z]\d{3,4})/i)?.[1]?.toUpperCase()||'';
       const two=['A1020','A1040','A1090','C720','C740','C776','C790'].includes(family)&&kind==='PIV';
       const max=two&&liftCount>=2?liftCount-1:liftCount;
       if(!Number.isInteger(count)||count<1||count>Math.max(1,max))throw Error('Số hộp không hợp lệ / Invalid hall panel quantity: '+raw+' = '+count+', maximum '+max);
       notes.push(count+' hộp / '+liftCount+' thang / 1 tầng');
       qtyDetails.push({floor,kind,code:raw,count,cars:liftCount});
     }
     if(!notes.length)continue;
     const target=model==='MRL'?({TopFloor:76,MainFloor:79,OtherFloor:82}[floor]):null;
     if(target){
       cells['C'+target]=[text(cells['C'+target]).trim(),...notes].filter(Boolean).join('\n');
       hiddenRows.delete(target);
     }else{
       // MR keeps separate PIV/HBV export lines. Append to the corresponding cell.
       for(const item of qtyDetails.filter(x=>x.floor===floor)){
         const outputAddr=nameAddr(db,'SpecFormExport','K36_'+item.kind+floor);
         if(!outputAddr)continue;
         cells[outputAddr]=text(cells[outputAddr]).trim()+'\n'+item.count+' hộp / '+liftCount+' thang / 1 tầng';
       }
     }
   }
   // MR source has two historical MBSS #N/A lookups on feature rows that are removed
   // when MBSS does not apply; only unresolved visible output cells block export.
   // Correct spelling in every customer-facing output cell, including formula results.
   // This must run after all layout transformations; source workbook data stays intact.
   if(window.SPEC_PROOFREAD){
     for(const addr of Object.keys(cells)){if(typeof cells[addr]==='string')cells[addr]=window.SPEC_PROOFREAD(cells[addr]);}
   }
   // v5.8.113: remove obsolete 1800 mm MBS beam-height claim in both languages.
   // Keep all other sensor functions and SR/MBSS incompatibility wording unchanged.
   for(const addr of Object.keys(cells)){
     if(typeof cells[addr]!=='string')continue;
     cells[addr]=cells[addr]
       .replace(/suốt khoảng\s*1800\s*mm\s*/gi,'')
       .replace(/cover a door height of approximately\s*1800\s*mm\s*/gi,'are arranged along the door height ')
       .replace(/\(Overhead height\)/gi,'(Overhead)');
   }
   // v5.8.203: the standard COP and wheelchair COP boxes are
   // SUS430 hairline stainless steel. The grade becomes SUS304 ONLY when
   // the global SUS selection is SUS304 and every relevant cabin material uses
   // the SUS-HL finish. Apply regardless of the optional details checkbox,
   // so Excel and PDF generated from these cells always agree.
   if(model==='MR'||model==='MRL'){
     const materialNames=['K23_WallFinish','K24_CarDoorFinish','K25_CarTransomFinish','K26_CarFrontFinish'];
     const readName=name=>{
       const addr=nameAddr(db,'SpecForm',name);
       return addr?text(eng.safe('SpecForm',addr).value).trim():'';
     };
     const allSUS304=readName('K22_SUSType').toUpperCase()==='SUS304'&&
       materialNames.every(name=>/^SUS-HL$/i.test(readName(name)));
     const grade=allSUS304?'SUS304':'SUS430';
     for(const row of model==='MR'?[60]:[59,60]){
       const label=text(cells['A'+row]).trim();
       const key='C'+row,source=text(cells[key]).trim();
       if(!/(?:bảng điều khiển|car operat(?:ing|ion) panel|wheelchair COP|người khuyết tật)/i.test(label))continue;
       if(!source||/\b(?:Not supplied|Không cung cấp)\b/i.test(source))continue;
       // Replace only an existing explicit COP-box material; retain models,
       // buttons, thickness and other technical descriptions unchanged.
       const withoutOld=source.replace(/\n?Vật liệu hộp COP(?:\/COP Wheelchair)?:[^\n]*/gi,'')
         .replace(/\n?\(COP(?:\/wheelchair COP)? box material:[^\n]*\)/gi,'')
         .replace(/\n?Bề mặt inox sọc nhuyễn SUS(?:430|304)\.?/gi,'').trim();
       // Concise material note for COP and wheelchair COP; omit duplicate
       // bilingual material wording while retaining the model description.
       cells[key]=withoutOld+'\nBề mặt inox sọc nhuyễn '+grade+'.';
     }
   }
   // v5.8.204: explicitly identify the standard SUS430 ceiling material
   // in exported MR/MRL specifications. Other colors/materials are unchanged.
   {
     const readCeiling=name=>{
       const addr=nameAddr(db,'SpecForm',name);
       return addr?text(eng.safe('SpecForm',addr).value).trim().toUpperCase():'';
     };
     const ceiling=readCeiling('K22_CeilingModel');
     const ceilingColor=readCeiling('K22_CeilingColor');
     const standardCeiling=(model==='MR'&&/^(?:L210S|L200S|N300S)(?:$|[-\s])/.test(ceiling))||
       (model==='MRL'&&ceilingColor==='SUS-HL');
     if(standardCeiling){
       const viKey='C'+(model==='MR'?45:46),enKey='C'+(model==='MR'?46:47);
       if(cells[viKey]&&!/SUS430/i.test(text(cells[viKey])))cells[viKey]=text(cells[viKey]).trimEnd()+' — Inox sọc nhuyễn SUS430.';
       if(cells[enKey]&&!/SUS430/i.test(text(cells[enKey])))cells[enKey]=text(cells[enKey]).trimEnd()+' (SUS430 hairline stainless steel).';
     }
   }
   cells['C'+(model==='MR'?11:12)]='MITSUBISHI ELECTRIC';
   const fatal=unresolved.filter(e=>!hiddenRows.has(cellRow(e.addr))&&/^[ABC]/.test(e.addr));
   if(fatal.length)throw Error('Công thức Export chưa tính được / Cannot evaluate '+fatal.slice(0,5).map(x=>x.addr+' '+x.error).join(' | '));
   return {model,cells,hiddenRows,removedFeatures,includedFeatures,sections,hiddenHall,hallLanternExportRows,totalFormulaUnresolved:unresolved.length,fatal,traction:model==='MR'?mrDutyBySpec(eng):null,mrlMotor:model==='MRL'?mrlMotorBySpec(eng):null};
 }
 // Material thickness for MR and MRL Version2; applied only when technical export is enabled.
 // Laminated steel sheet corresponds to the PPS column of the supplied material-thickness chart.
 // v5.8.135: single-tool export descriptions, using distinct MR/MRL reference charts.
 // Do not overwrite an unverified door/wall thickness with an assumed value.
 function attachMRMaterialThickness(model,changes,db,doc){
   if(!['MR','MRL'].includes(model)||!doc?.cells)return;
   const eng=new window.SpecExcelSourceEngine.SourceEngine(db,changes);
   const read=a=>text(eng.safe('SpecForm',a).value).trim();
   const mrl=model==='MRL';
   const isUnavailable=v=>/^(?:Not suppl(?:ied|y)|Không cung cấp)(?:\s*\(Not supplied\))?\.?$/i.test(text(v).trim());
   const endsWithThickness=v=>/(?:\bdày\s*\d+(?:[.,]\d+)?\s*mm|\bthickness\s*\d+(?:[.,]\d+)?\s*mm)/i.test(v);
   // Panel/model descriptions in Section IV: 1.5 mm, only in detailed Export.
   // MR: separate PIV/HBV/PIH rows; MRL: consolidated 3-floor blocks.
   // Keep the standard Export descriptions unchanged when detail is disabled.
   if(!doc.includeTechDetails)return;
   // v5.8.139: Every supplied car operating panel (COP), including wheelchair COP,
   // has 2.0 mm panel thickness in the detailed Excel/PDF export. Do not alter
   // other feature descriptions that incidentally mention a COP.
   const copExportRows=mrl?[60]:[59,60];
   // In the MRL workbook, row 60 is the principal COP.
   // In the MR workbook, row 59 is the principal COP and row 60 the wheelchair COP.
   const effectiveCopRows=copExportRows;
   const appendCopThickness=(raw)=>{
     const source=text(raw).trim();
     if(!source||isUnavailable(source)||/\b(?:Not supplied|Không cung cấp)\b/i.test(source))return raw;
     if(/(?:\bdày\s*2(?:[.,]0)?\s*mm|\bthickness\s*2(?:[.,]0)?\s*mm)/i.test(source))return raw;
     const m=source.match(/^([\s\S]*?)\s*\(([^()]*)\)\.?\s*$/);
     if(m){
       const vi=m[1].replace(/[.;]\s*$/,'').trim();
       const en=m[2].replace(/[.;]\s*$/,'').trim();
       return vi+', dày 2.0 mm. ('+en+', thickness 2.0 mm).';
     }
     return source.replace(/[.;]\s*$/,'')+', dày 2.0 mm (thickness 2.0 mm).';
   };
   // v5.8.144: Validate the EXPORT ROW LABEL, not only a fixed row number.
   // In some model/layout combinations the preceding car sill (ngưỡng cửa)
   // occupies a nearby row. Only real car operating panels may receive 2 mm.
   for(const row of effectiveCopRows){
     const label=text(doc.cells['A'+row]).trim();
     if(/(?:ngưỡng cửa|car sill)/i.test(label))continue;
     if(!/(?:bảng điều khiển|car operat(?:ing|ion) panel|wheelchair COP)/i.test(label))continue;
     const addr='C'+row;
     if(Object.hasOwn(doc.cells,addr))doc.cells[addr]=appendCopThickness(doc.cells[addr]);
   }
   // Car sill has no specified sheet thickness: do not export a COP thickness
   // into this field even if the spreadsheet layout changes in the future.
   for(const [addr,label] of Object.entries(doc.cells)){
     if(!/^A\d+$/.test(addr)||!/(?:ngưỡng cửa phòng thang|car sill)/i.test(text(label)))continue;
     const result='C'+addr.slice(1);
     if(typeof doc.cells[result]==='string')doc.cells[result]=doc.cells[result]
       .replace(/,\s*dày\s*2(?:[.,]0)?\s*mm(?=\.|\s*\()/gi,'')
       .replace(/,\s*thickness\s*2(?:[.,]0)?\s*mm(?=\)|\.)/gi,'');
   }
   const rows=mrl?[76,79,82]:[76,77,78,79,80,81];
   const panelCode=/(?:\b(?:PIV|HBV|PIH|PID)\s*\d*\s*[-–]?\s*[A-Z]?\d+[A-Z0-9-]*\b)/i;
   const panelLanguage=/(?:hiển thị|hộp gọi|nút gọi|hall indicators|hall buttons|hall position indicators|hall calls)/i;
   const amendPanelLine=(line)=>{
     if(!panelCode.test(line)||!panelLanguage.test(line)||/\b(?:Not supplied|Không cung cấp)\b/i.test(line))return line;
     if(endsWithThickness(line))return line;
     if(/^\s*\(/.test(line)){
       let body=line.trim();const match=body.match(/^(\s*\()(.*?)(\)\.?\s*)$/);
       if(match)return match[1]+match[2].replace(/[.,;]\s*$/,'')+', thickness 1.5 mm'+match[3];
       return line+' (thickness 1.5 mm)';
     }
     return line.replace(/[.;]\s*$/,'')+', dày 1.5 mm';
   };
   for(const row of rows){
     const addr='C'+row, original=text(doc.cells[addr]);
     if(!original||isUnavailable(original))continue;
     doc.cells[addr]=original.split('\n').map(amendPanelLine).join('\n');
   }
   const hl=Number(read(mrl?'D30':'D31'));
   const hh=Number(read(mrl?'D32':'D33'));
   const k=hl-25.5; // K = HL - 25.5 mm, applies to MR and MRL Version02
   const fire=/EN81[-\s]*58\s*\(II\)/i.test(read(mrl?'D84':'D83'));
   const hhk=hh+(fire?10:0);
   const opening=read(mrl?'D15':'D14').toUpperCase();
   const windowType=[read(mrl?'D46':'D50'),read(mrl?'D48':'D52'),read('D50'),read('D51')].join(' ');
   const hasWindow=/DSJ1|DS1|LARGE\s*WINDOW/i.test(windowType);
   const material=v=>{
     const t=text(v).toUpperCase();
     if(/LAMINATED|LAMINATE/.test(t))return 'LAMINATED';
     if(/SUS[-\s]*HE/.test(t))return 'SUS-HE';
     if(/SUS[-\s]*HL|SUS\s*430(?:J1L)?|STAINLESS STEEL HAIRLINE/.test(t))return 'SUS-HL';
     if(/SUS[-\s]*M\b/.test(t))return 'SUS-M';
     if(/\bPPS\b/.test(t))return mrl?'LAMINATED':'PPS';
     if(/\bPP\b|PAINTED STEEL|PAINTED/.test(t))return 'PP';
     return null;
   };
   const wallThickness=m=>{
     if(m==='SUS-HL'||(mrl&&m==='LAMINATED')){
       if(k>=2174.5&&k<=2274.5)return 1.0;
       if(k>2274.5&&k<2775)return 1.2;
       return null;
     }
     return (mrl?{'PP':1.2}:{'PP':1.2,'SUS-HE':1.2,'SUS-M':1.5,'PPS':0.8})[m]??null;
   };
   const carDoorThickness=m=>{
     if(hasWindow)return 1.5;
     if(mrl)return ({'PP':1.2,'SUS-HL':1.2,'LAMINATED':1.2})[m]??null;
     if(['PP','SUS-HL','SUS-HE'].includes(m)){
       if(hhk>=2100&&hhk<2401)return 1.2;
       if(hhk>=2401&&hhk<2601)return m==='PP'?1.6:1.5;
       return null;
     }
     return ({'SUS-M':1.5,'PPS':0.8})[m]??null;
   };
   const transomThickness=m=>mrl?({'PP':1.2,'SUS-HL':1.2,'LAMINATED':1.2})[m]??null:({'PP':1.2,'SUS-HL':1.2,'SUS-HE':1.2,'SUS-M':1.5,'PPS':0.8})[m]??null;
   const frontThickness=m=>mrl?({'PP':1.2,'SUS-HL':1.2,'LAMINATED':1.2})[m]??null:({'PP':1.2,'SUS-HL':1.2,'SUS-M':1.5})[m]??null;
   const entranceThickness=m=>{
     if(hasWindow)return 1.5;
     if(fire){if(opening==='CO')return 1.5;if(opening==='2S')return 1.2;return null;}
     if(mrl)return ({'PP':1.6,'SUS-HL':1.2,'LAMINATED':1.2})[m]??null;
     return ({'PP':1.2,'SUS-HL':1.2,'SUS-HE':1.5,'SUS-M':1.5,'PPS':0.8})[m]??null;
   };
   const jambThickness=m=>mrl?({'PP':1.2,'SUS-HL':1.2,'LAMINATED':1.2})[m]??null:({'PP':1.2,'SUS-HL':1.2})[m]??null;
   const supplement=(addr,thickness,english,unverified=false)=>{
     const original=text(doc.cells[addr]).trim();
     if(!original||isUnavailable(original)||endsWithThickness(original))return;
     if(unverified){
       const phrase=english?'thickness to be confirmed':'chiều dày cần xác minh';
       const stripped=original.replace(/^\(/,'').replace(/\)\.?\s*$/,'').replace(/[.;]\s*$/,'');
       doc.cells[addr]=english?'('+stripped+', '+phrase+').':stripped+', '+phrase+'.';
       return;
     }
     if(!Number.isFinite(thickness))return;
     const val=Number(thickness).toFixed(1)+' mm';
     const stripped=original.replace(/^\(/,'').replace(/\)\.?\s*$/,'').replace(/[.;]\s*$/,'');
     doc.cells[addr]=english?'('+stripped+', thickness '+val+').':stripped+', dày '+val+'.';
   };
   const components=mrl?[
     ['Car wall',read('D43'),wallThickness,47,48],
     ['Car door',read('D46'),carDoorThickness,49,50],
     ['Transom panel',read('D49'),transomThickness,51,52],
     ['Front return panel',read('D52'),frontThickness,53,54],
     ['Entrance door – main',read('E73'),entranceThickness,72,73],
     ['Entrance door – other',read('E74'),entranceThickness,74,75],
     ['Jamb w/o TP – main',read('E71'),jambThickness,68,69],
     ['Jamb w/o TP – other',read('E72'),jambThickness,70,71]
   ]:[
     ['Car wall',read('D47'),wallThickness,48,49],
     ['Car door',read('D50'),carDoorThickness,50,51],
     ['Transom panel',read('D53'),transomThickness,52,53],
     ['Front return panel',read('D56'),frontThickness,54,55],
     ['Entrance door – main',read('E76'),entranceThickness,72,73],
     ['Entrance door – other',read('E77'),entranceThickness,74,75],
     ['Jamb w/o TP – main',read('E74'),jambThickness,68,69],
     ['Jamb w/o TP – other',read('E75'),jambThickness,70,71]
   ];
   doc.materialThicknessResults=[];
   for(const [name,raw,fn,viRow,enRow] of components){
     if(name.startsWith('Jamb')){
       const style=read(mrl?(name.endsWith('main')?'D71':'D72'):(name.endsWith('main')?'D74':'D75'));
       if(style&&!/E-102|WITHOUT\s*TP|W\/O\s*TP/i.test(style))continue;
     }
     const mat=material(raw),thickness=mat?fn(mat):null;
     doc.materialThicknessResults.push({name,material:mat,thickness,k});
     const unverified=name==='Car wall'&&mat==='SUS-HL'&&thickness===null&&Number.isFinite(k);
     const unverifiedLam=name==='Car wall'&&mrl&&mat==='LAMINATED'&&thickness===null&&Number.isFinite(k);
     if(!Number.isFinite(thickness)&&!(unverified||unverifiedLam))continue;
     supplement('C'+viRow,thickness,false,unverified||unverifiedLam);
     supplement('C'+enRow,thickness,true,unverified||unverifiedLam);
   }
   // Kickplate source contains Vietnamese and English within the SAME cell.
   // Preserve both descriptions and add thickness to each language separately.
   const kickAddr='C'+(mrl?55:56);
   const kickRaw=text(doc.cells[kickAddr]).trim();
   const kickMatch=kickRaw.match(/^([\s\S]*?)\s*\(([^()]*(?:\([^()]*\)[^()]*)*)\)\.?\s*$/);
   if(kickMatch&&!endsWithThickness(kickRaw)){
     doc.cells[kickAddr]=kickMatch[1].replace(/[.;]\s*$/,'')+', dày 1.2 mm.\n('+kickMatch[2].replace(/[.;]\s*$/,'')+', thickness 1.2 mm).';
   }else supplement(kickAddr,1.2,false);
   doc.materialThicknessResults.push({name:'Kickplate',thickness:1.2});
 }
 function updateCell(document, c, v){
   for(const old of Array.from(c.childNodes))if(['f','v','is'].includes(old.localName))c.removeChild(old);
   if(v===null||v===undefined||v===''){c.removeAttribute('t');return;}
   if(typeof v==='number'&&Number.isFinite(v)){
     c.removeAttribute('t');const val=document.createElementNS(n,'v');val.textContent=String(v);c.appendChild(val);return;
   }
   if(typeof v==='boolean'){c.setAttribute('t','b');const val=document.createElementNS(n,'v');val.textContent=v?'1':'0';c.appendChild(val);return;}
   const inline=document.createElementNS(n,'is'),t=document.createElementNS(n,'t');
   c.setAttribute('t','inlineStr');if(/^\s|\s$|\n/.test(String(v)))t.setAttribute('xml:space','preserve');t.textContent=String(v);inline.appendChild(t);c.appendChild(inline);
 }
 async function xlsxFromModel(zip, model, db, doc){
   const workbookXML=await zip.file('xl/workbook.xml').async('string');
   const relsXML=await zip.file('xl/_rels/workbook.xml.rels').async('string');
   const workbook=parse(workbookXML);
   const rels=parse(relsXML);
   if(workbook.getElementsByTagName('parsererror').length||rels.getElementsByTagName('parsererror').length){
     throw Error('Workbook XML parsing failed');
   }
   const sheetnodes=all(workbook,'sheet');
   const sheet=sheetnodes.find(x=>x.getAttribute('name')==='SpecFormExport');
   if(!sheet)throw Error('Không có SpecFormExport / Missing sheet');
   // iOS WebKit occasionally fails to expose an r:id in a namespaced XML node.
   // Look up the ID from raw source XML as a fallback; retain original workbook.
   let exportRid=sheet.getAttributeNS(r,'id')||sheet.getAttribute('r:id');
   if(!exportRid){
     const tag=workbookXML.match(/<(?:[\w.-]+:)?sheet\b[^>]*\bname=["']SpecFormExport["'][^>]*>/i)?.[0];
     exportRid=tag?.match(/\br:id=["']([^"']+)["']/i)?.[1]||'';
   }
   const relationships=Array.from(rels.documentElement.childNodes).filter(x=>x.nodeType===1);
   let ref=relationships.find(x=>x.getAttribute('Id')===exportRid);
   if(!ref && exportRid){
     const tag=relsXML.match(/<Relationship\b[^>]*\bId=["']([^"']+)["'][^>]*\/?>(?:<\/Relationship>)?/gi)?.find(x=>x.includes(exportRid));
     const target=tag?.match(/\bTarget=["']([^"']+)["']/i)?.[1];
     if(target)ref=relationships.find(x=>x.getAttribute('Target')===target);
   }
   if(!ref)throw Error('No workbook relationship for Export sheet: '+(exportRid||'missing r:id'));
   const sheetpath=ref.getAttribute('Target').replace(/^\//,'').replace(/^xl\//,'');
   const finalpath='xl/'+sheetpath;
   const worksheet=zip.file(finalpath);if(!worksheet)throw Error('Missing export XML: '+finalpath);
   const x=parse(await worksheet.async('string'));
   // v5.8.114: Excel exports must open without Freeze Panes for BOTH models.
   // Remove pane configuration from the exported worksheet, leaving the source database untouched.
   for(const sheetView of all(x,'sheetView')){
     for(const pane of Array.from(sheetView.children).filter(node=>node.localName==='pane'))pane.remove();
     for(const selection of Array.from(sheetView.children).filter(node=>node.localName==='selection')){
       selection.removeAttribute('pane');
     }
   }
   const cellNodes=all(x,'c');let freezeCount=0;
   for(const c of cellNodes){const addr=c.getAttribute('r');if(!addr)continue;
     if(Object.hasOwn(doc.cells,addr)){updateCell(x,c,doc.cells[addr]);freezeCount++}
     if(doc.hiddenRows.has(cellRow(addr))){/* Source C# deletes these two-row features. Hide rows
       instead of destructive shift, to preserve all merged cell geometry and print areas. */}
   }
   // Include new, bilingual Version02 manual appendix rows, preserving original sheet content.
   const sheetData=all(x,'sheetData')[0];
   const existing=new Set(cellNodes.map(c=>c.getAttribute('r')));
   const addByRow=new Map();
   for(const [addr,val] of Object.entries(doc.cells)){
      if(!/^[ABC]\d+$/.test(addr)||existing.has(addr)||val===null||val===undefined||val==='')continue;
      const rn=cellRow(addr);if(!addByRow.has(rn))addByRow.set(rn,[]);addByRow.get(rn).push([addr,val]);
   }
   if(sheetData){
     for(const rn of [...addByRow.keys()].sort((a,b)=>a-b)){
       let row=Array.from(sheetData.children).find(el=>el.localName==='row'&&+el.getAttribute('r')===rn);
       if(!row){row=x.createElementNS(n,'row');row.setAttribute('r',String(rn));
          const later=Array.from(sheetData.children).find(el=>el.localName==='row'&&+el.getAttribute('r')>rn);
          if(later)sheetData.insertBefore(row,later);else sheetData.appendChild(row);
       }
       for(const [addr,val] of addByRow.get(rn)){
         const cell=x.createElementNS(n,'c');cell.setAttribute('r',addr);cell.setAttribute('s','0');updateCell(x,cell,val);row.appendChild(cell);
       }
     }
   }
   // Supplemental Features (ACC, NSCR-H, and Version02-only options) must have
   // exactly the same two-column label + description formatting as ordinary
   // bilingual Feature rows. Blank rows 296+ in the source have style 0/25.
   const extraPairs=(doc.includedFeatures||[]).filter(f=>Array.isArray(f.exportRows)&&f.exportRows.length===2&&f.exportRows[0]>=296);
   const extraFeatureRows=new Set(extraPairs.flatMap(f=>f.exportRows));
   const extraHeaders=new Set(extraFeatureRows.size&&doc.cells.A296?[296]:[]);
   if(extraFeatureRows.size&&doc.cells.A312)extraHeaders.add(312);
   if(extraFeatureRows.size){
     const existingRows=new Map(all(x,'row').map(row=>[Number(row.getAttribute('r')),row]));
     const templateEnglish=existingRows.get(171),templateVietnamese=existingRows.get(172);
     const styleFor=(src,col)=>Array.from(src?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+src.getAttribute('r'))?.getAttribute('s')||'0';
     const merged=all(x,'mergeCell');
     let merges=all(x,'mergeCells')[0];
     if(!merges){
       merges=x.createElementNS(n,'mergeCells');
       const anchor=Array.from(x.documentElement.children).find(el=>['conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup'].includes(el.localName));
       if(anchor)x.documentElement.insertBefore(merges,anchor);else x.documentElement.appendChild(merges);
     }
     const overlaps=(rn)=>merged.some(m=>{
       const match=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m.getAttribute('ref')||'');
       if(!match)return false;
       const a=Number(match[2]),b=Number(match[4]);
       return rn>=a&&rn<=b&&!(match[1]==='D'||match[3]==='D');
     });
     function ensureCell(row,rn,col,style){
       let c=Array.from(row.children).find(el=>el.localName==='c'&&el.getAttribute('r')===col+rn);
       if(!c){c=x.createElementNS(n,'c');c.setAttribute('r',col+rn);
         const next=Array.from(row.children).find(el=>el.localName==='c'&&String(el.getAttribute('r')||'').charCodeAt(0)>col.charCodeAt(0));
         if(next)row.insertBefore(c,next);else row.appendChild(c);
       }
       c.setAttribute('s',style);
       if(Object.hasOwn(doc.cells,col+rn))updateCell(x,c,doc.cells[col+rn]);
       return c;
     }
     for(const pair of extraPairs){
       for(let i=0;i<2;i++){
         const rn=pair.exportRows[i];const row=existingRows.get(rn);
         if(!row)continue;
         const src=i?templateVietnamese:templateEnglish;
         for(const col of ['A','B','C'])ensureCell(row,rn,col,styleFor(src,col));
         const aText=String(doc.cells['A'+rn]||''),cText=String(doc.cells['C'+rn]||'');
         // Widths A+B and C are ~46/70 character units. Respect wrapped bilingual text.
         const lines=Math.max(1,Math.ceil(aText.length/36),Math.ceil(cText.length/67));
         row.setAttribute('ht',String(Math.min(140,Math.max(27.6,lines*14+7))));
         row.setAttribute('customHeight','1');
         if(!overlaps(rn)){
           const m=x.createElementNS(n,'mergeCell');m.setAttribute('ref',`A${rn}:B${rn}`);merges.appendChild(m);merged.push(m);
         }
       }
     }
     // Section headings use a full-width merged blue title, consistent with I–V.
     for(const rn of extraHeaders){
       const row=existingRows.get(rn);if(!row)continue;
       for(const col of ['A','B','C'])ensureCell(row,rn,col,'25');
       if(!overlaps(rn)){
         const m=x.createElementNS(n,'mergeCell');m.setAttribute('ref',`A${rn}:C${rn}`);merges.appendChild(m);merged.push(m);
       }
       row.setAttribute('ht','20');row.setAttribute('customHeight','1');
     }
     merges.setAttribute('count',String(all(merges,'mergeCell').length));
   }
   // v5.8.111: merge the full width of every main/subtitle in the Features
   // section (MR and MRL have different source row positions). The factory
   // worksheet kept those headings only in A and left B/C blank.
   // Do this before model-specific row shifting; the later transforms adjust
   // the merge references consistently with their data rows.
   {
     const featureHeadingRows=model==='MR'
       ? [83,84,110,146,212,254]
       : [87,88,114,150,217,259];
     let merges=all(x,'mergeCells')[0];
     if(!merges){
       merges=x.createElementNS(n,'mergeCells');
       const after=all(x,'sheetData')[0];
       const anchor=Array.from(x.documentElement.children).find(el=>['conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup'].includes(el.localName));
       if(anchor)x.documentElement.insertBefore(merges,anchor);
       else if(after?.nextSibling)x.documentElement.insertBefore(merges,after.nextSibling);
       else x.documentElement.appendChild(merges);
     }
     const toIndex=str=>[...str].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
     const spans=(ref)=>{
       const m=/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref||'');
       return m?{a:toIndex(m[1]),b:toIndex(m[3]||m[1]),top:+m[2],bottom:+(m[4]||m[2])}:null;
     };
     for(const rn of featureHeadingRows){
       if(doc.hiddenRows.has(rn))continue;
       const blockers=all(merges,'mergeCell').filter(el=>{
         const q=spans(el.getAttribute('ref'));
         return q&&q.top<=rn&&q.bottom>=rn&&q.a<=3&&q.b>=1;
       });
       if(blockers.some(el=>{
         const q=spans(el.getAttribute('ref'));
         return q.top!==rn||q.bottom!==rn||q.a!==1||q.b>3;
       }))continue; // preserve any overlapping factory vertical/other merge
       for(const el of blockers)el.remove();
       const el=x.createElementNS(n,'mergeCell');
       el.setAttribute('ref',`A${rn}:C${rn}`);merges.appendChild(el);
     }
     merges.setAttribute('count',String(all(merges,'mergeCell').length));
   }
   // v5.8.113: the Version02 source omitted the label-cell merges for Regulation and Fireproof Door.
   if(model==='MRL'){
     let merges=all(x,'mergeCells')[0];
     if(!merges)throw Error('Missing mergeCells in MRL export worksheet');
     for(const rn of [33,85]){
       const wanted=`A${rn}:B${rn}`;
       const overlaps=all(merges,'mergeCell').some(m=>{
         const ref=m.getAttribute('ref')||'';
         const mm=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
         if(!mm)return false;
         const col=k=>[...k].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0);
         return +mm[2]<=rn&&+mm[4]>=rn&&col(mm[1])<=2&&col(mm[3])>=1;
       });
       if(!overlaps){const m=x.createElementNS(n,'mergeCell');m.setAttribute('ref',wanted);merges.appendChild(m);}
     }
     merges.setAttribute('count',String(all(merges,'mergeCell').length));
   }
   for(const row of all(x,'row')){
     const num=+(row.getAttribute('r')||0);
     if(doc.hiddenRows.has(num)){row.setAttribute('hidden','1');row.setAttribute('ht','0');row.setAttribute('customHeight','1');}
     // Expanded bilingual Hall Lantern description needs a taller Excel row;
     // all existing wrapped-cell and merged-cell geometry is retained.
     else if(doc.hallLanternExportRows?.includes(num)){
       const base=Number(row.getAttribute('ht'))||24;
       const description=String(doc.cells['C'+num]||'');
       const lines=description.split(/\r?\n/).reduce((total,line)=>total+Math.max(1,Math.ceil(line.length/69)),0);
       row.setAttribute('ht',String(Math.max(base,Math.min(185,lines*16+14))));
       row.setAttribute('customHeight','1');
     }
   }
   // Make the exact Model value bold in the *actual* XLSX cell, preserving the
   // workbook font family, color, size, cell borders, number format and merges.
   // The PDF renderer reads this XLSX font style so both export formats agree.
   {
     const address='C'+(model==='MR'?13:14);
     const modelCell=all(x,'c').find(c=>c.getAttribute('r')===address);
     if(!modelCell)throw Error('Missing Model export cell '+address);
     const stylesPath='xl/styles.xml',part=zip.file(stylesPath);
     if(!part)throw Error('Missing XLSX styles for Model emphasis');
     const sd=parse(await part.async('string'));
     const fonts=all(sd,'fonts')[0],xfs=all(sd,'cellXfs')[0];
     if(!fonts||!xfs)throw Error('Missing fonts/cellXfs');
     const originalXfs=Array.from(xfs.children).filter(c=>c.localName==='xf');
     const oldStyle=Number(modelCell.getAttribute('s')||0);
     const oldXf=originalXfs[oldStyle]||originalXfs[0];
     const fontList=Array.from(fonts.children).filter(c=>c.localName==='font');
     const oldFont=fontList[Number(oldXf.getAttribute('fontId')||0)]||fontList[0];
     const boldFont=oldFont.cloneNode(true);
     let b=Array.from(boldFont.children).find(c=>c.localName==='b');
     if(!b){b=sd.createElementNS(n,'b');boldFont.appendChild(b);}
     b.setAttribute('val','1');
     const fontId=fontList.length;fonts.appendChild(boldFont);
     fonts.setAttribute('count',String(fontId+1));
     const newXf=oldXf.cloneNode(true);
     newXf.setAttribute('fontId',String(fontId));
     newXf.setAttribute('applyFont','1');
     const styleId=originalXfs.length;
     xfs.appendChild(newXf);xfs.setAttribute('count',String(styleId+1));
     modelCell.setAttribute('s',String(styleId));
     const manufacturerAddr='C'+(model==='MR'?11:12);
     const manufacturerCell=all(x,'c').find(c=>c.getAttribute('r')===manufacturerAddr);
     if(!manufacturerCell)throw Error('Missing Manufacturer export cell '+manufacturerAddr);
     const oldManufacturerXf=originalXfs[Number(manufacturerCell.getAttribute('s')||0)]||originalXfs[0];
     const oldManufacturerFont=fontList[Number(oldManufacturerXf.getAttribute('fontId')||0)]||fontList[0];
     const manufacturerFont=oldManufacturerFont.cloneNode(true);
     let bold=Array.from(manufacturerFont.children).find(c=>c.localName==='b');
     if(!bold){bold=sd.createElementNS(n,'b');manufacturerFont.appendChild(bold);}
     bold.setAttribute('val','1');
     fonts.appendChild(manufacturerFont);
     const manufacturerStyle=oldManufacturerXf.cloneNode(true);
     manufacturerStyle.setAttribute('fontId',String(fontId+1));
     manufacturerStyle.setAttribute('applyFont','1');
     xfs.appendChild(manufacturerStyle);
     manufacturerCell.setAttribute('s',String(styleId+1));
     fonts.setAttribute('count',String(fontId+2));
     xfs.setAttribute('count',String(styleId+2));
     zip.file(stylesPath,serial(sd));
   }
   // v3.8: highlight only Roman-numbered specification section headings (I., II., III., IV., etc.).
   // Style the frozen export worksheet itself, so both XLSX and its PDF print share
   // exactly the same blue sections regardless of the source workbook theme/fill.
   const sectionTitle=/^\s*[IVXLCDM]+\.\s+/i;
   const blueRows=[...new Set(Object.entries(doc.cells)
     .filter(([addr,value])=>/^A\d+$/.test(addr)&&cellRow(addr)<=480&&!doc.hiddenRows.has(cellRow(addr))&&sectionTitle.test(text(value)))
     .map(([addr])=>cellRow(addr)))];
   if(blueRows.length){
     const stylePath='xl/styles.xml';const styleFile=zip.file(stylePath);
     if(!styleFile)throw Error('Missing Excel styles file for section formatting');
     const styles=parse(await styleFile.async('string'));
     const fills=all(styles,'fills')[0],xfs=all(styles,'cellXfs')[0];
     if(!fills||!xfs)throw Error('Missing Excel fills/cellXfs for section formatting');
     const fill=styles.createElementNS(n,'fill'),pattern=styles.createElementNS(n,'patternFill');
     pattern.setAttribute('patternType','solid');
     const fg=styles.createElementNS(n,'fgColor');fg.setAttribute('rgb','FFD9EDF7');
     const bg=styles.createElementNS(n,'bgColor');bg.setAttribute('indexed','64');
     pattern.append(fg,bg);fill.append(pattern);
     const newFillId=Array.from(fills.children).filter(el=>el.localName==='fill').length;
     fills.appendChild(fill);fills.setAttribute('count',String(newFillId+1));
     const existingStyles=Array.from(xfs.children).filter(el=>el.localName==='xf');
     const baseStyleCount=existingStyles.length;const blueStyleCache=new Map();
     function blueStyle(baseId){
       baseId=baseId>=0&&baseId<baseStyleCount?baseId:0;
       if(blueStyleCache.has(baseId))return blueStyleCache.get(baseId);
       const xf=existingStyles[baseId].cloneNode(true);xf.setAttribute('fillId',String(newFillId));xf.setAttribute('applyFill','1');
       const id=baseStyleCount+blueStyleCache.size;xfs.appendChild(xf);blueStyleCache.set(baseId,id);return id;
     }
     const rowMap=new Map(all(x,'row').map(row=>[Number(row.getAttribute('r')),row]));
     for(const rowNum of blueRows){
       const row=rowMap.get(rowNum);if(!row)continue;
       const origA=Array.from(row.children).find(el=>el.localName==='c'&&el.getAttribute('r')==='A'+rowNum);
       const aStyle=Number(origA?.getAttribute('s')||0);
       for(const letter of ['A','B','C']){
         const addr=letter+rowNum;
         let cell=Array.from(row.children).find(el=>el.localName==='c'&&el.getAttribute('r')===addr);
         if(!cell){
           cell=x.createElementNS(n,'c');cell.setAttribute('r',addr);
           const next=Array.from(row.children).find(el=>el.localName==='c'&&String(el.getAttribute('r')||'').charCodeAt(0)>letter.charCodeAt(0));
           if(next)row.insertBefore(cell,next);else row.appendChild(cell);
         }
         const oldStyle=Number(cell.getAttribute('s')||aStyle);
         cell.setAttribute('s',String(blueStyle(oldStyle)));
       }
     }
     xfs.setAttribute('count',String(baseStyleCount+blueStyleCache.size));
     zip.file(stylePath,serial(styles));
   }
   // Normal feature rows remain neutral: no tan/brown highlighting is added.
   // Hide helper columns D/E; Source C# blanks them after evaluating Option status.
   const cols=all(x,'cols')[0];if(cols){for(const col of all(cols,'col')){
       let min=+col.getAttribute('min'),max=+col.getAttribute('max');
       if(min>=4&&max<=5)col.setAttribute('hidden','1');
     }}
   // Ensure print settings: A4 portrait and fit one page in width, multiple pages in height.
   let sheetPr=child(x.documentElement,'sheetPr');if(!sheetPr){sheetPr=x.createElementNS(n,'sheetPr');x.documentElement.insertBefore(sheetPr,x.documentElement.firstChild)}
   let pagePr=child(sheetPr,'pageSetUpPr');if(!pagePr){pagePr=x.createElementNS(n,'pageSetUpPr');sheetPr.appendChild(pagePr)}pagePr.setAttribute('fitToPage','1');
   let pageSetup=child(x.documentElement,'pageSetup');if(!pageSetup){pageSetup=x.createElementNS(n,'pageSetup');let margins=child(x.documentElement,'pageMargins');if(margins)margins.after(pageSetup);else x.documentElement.appendChild(pageSetup)}
   pageSetup.setAttribute('paperSize','9');pageSetup.setAttribute('orientation','portrait');pageSetup.setAttribute('fitToWidth','1');pageSetup.setAttribute('fitToHeight','0');
   sheet.removeAttribute('state');sheet.setAttribute('name','Specification');
   for(const old of sheetnodes)if(old!==sheet)old.remove();
   for(const old of relationships)if(old!==ref&&/\/(?:worksheet|calcChain)$/.test(old.getAttribute('Type')||''))old.remove();
   // Formulas have been frozen and no longer need named ranges from removed source sheets.
   const defs=all(workbook,'definedNames')[0];if(defs){for(const old of Array.from(defs.childNodes))defs.removeChild(old);
     const print=x.createElementNS(n,'definedName');print.setAttribute('name','_xlnm.Print_Area');print.setAttribute('localSheetId','0');print.textContent="'Specification'!$A$1:$C$"+Math.max(...Object.keys(doc.cells).filter(a=>/^[ABC]\d+$/.test(a)).map(cellRow));defs.appendChild(print);
   }
   const calc=all(workbook,'calcPr')[0];if(calc){calc.setAttribute('fullCalcOnLoad','0');calc.setAttribute('calcMode','manual');}
   // Export-only rail rows for MRL Version02. Insert after the hoistway structure
   // without modifying SpecForm or any technical DataTable formulas. This operates
   // on the frozen export worksheet, preserving all existing styles and merges.
   if(model==='MRL'){
     const pivot=29,offset=2;
     const shiftedRef=ref=>String(ref).replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g,(match,col,num)=>col+(Number(num)>pivot?Number(num)+offset:num));
     // Move existing rows and cell references; preserve each cell's original style.
     for(const row of all(x,'row')){
       const old=Number(row.getAttribute('r'));
       if(old>pivot){row.setAttribute('r',String(old+offset));
         for(const c of Array.from(row.children).filter(el=>el.localName==='c')){
           const addr=c.getAttribute('r');if(addr)c.setAttribute('r',shiftedRef(addr));
         }
       }
     }
     // Shift merged ranges and print/repeat references that are sheet-local.
     for(const m of all(x,'mergeCell')){const ref=m.getAttribute('ref');if(ref)m.setAttribute('ref',shiftedRef(ref));}
     const filt=all(x,'autoFilter')[0];if(filt?.getAttribute('ref'))filt.setAttribute('ref',shiftedRef(filt.getAttribute('ref')));
     for(const br of all(x,'brk'))if(Number(br.getAttribute('id'))>pivot)br.setAttribute('id',String(Number(br.getAttribute('id'))+offset));
     const dim=all(x,'dimension')[0];if(dim?.getAttribute('ref'))dim.setAttribute('ref',shiftedRef(dim.getAttribute('ref')));
     // Reuse proven header/detail border styles from this exact workbook.
     const source=all(x,'row').find(row=>Number(row.getAttribute('r'))===29);
     const findStyle=(col,fall='0')=>Array.from(source?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+'29')?.getAttribute('s')||fall;
     function addRailRow(num,label,top){
       const row=x.createElementNS(n,'row');row.setAttribute('r',String(num));row.setAttribute('ht','17');row.setAttribute('customHeight','1');
       for(const [col,value,style] of [
         ['A',top?'Thanh dẫn hướng (Rail)':'',findStyle('A')],
         ['B',label,findStyle('B')],
         ['C','Mitsubishi Standard',findStyle('C')]]){
         const c=x.createElementNS(n,'c');c.setAttribute('r',col+num);c.setAttribute('s',style);updateCell(x,c,value);row.appendChild(c);
       }
       const following=Array.from(sheetData.children).find(el=>el.localName==='row'&&Number(el.getAttribute('r'))>num);
       if(following)sheetData.insertBefore(row,following);else sheetData.appendChild(row);
     }
     addRailRow(30,'Cabin (CAR)',true);
     addRailRow(31,'Rail đối trọng (CWT)',false);
     let merges=all(x,'mergeCells')[0];if(!merges){merges=x.createElementNS(n,'mergeCells');
       const after=Array.from(x.documentElement.children).find(el=>['conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup'].includes(el.localName));
       if(after)x.documentElement.insertBefore(merges,after);else x.documentElement.appendChild(merges);
     }
     const merge=x.createElementNS(n,'mergeCell');merge.setAttribute('ref','A30:A31');merges.appendChild(merge);
     merges.setAttribute('count',String(all(merges,'mergeCell').length));
     // The new rail lines are part of the specification's full print area.
     const printName=all(workbook,'definedName').find(el=>el.getAttribute('name')==='_xlnm.Print_Area');
     if(printName)printName.textContent=printName.textContent.replace(/\$C\$(\d+)/,(_,num)=>'$C$'+(Number(num)+offset));
   }


   // Preserve the MR Rail row-height fix regardless of optional motor-detail export.
   if(model==='MR'){
     for(const rn of [29,30]){
       const row=all(x,'row').find(y=>Number(y.getAttribute('r'))===rn);
       if(row){row.setAttribute('ht','28');row.setAttribute('customHeight','1');}
     }
   }

   // NEXIEZ-MR: insert three duty lines directly after Control system (row 33).
   // Existing Main code becomes row 37. Preserve merged cells and workbook print area.
   if(model==='MR' && doc.includeTechDetails===true){
     const pivot=33, offset=3;
     const shiftedRef=ref=>String(ref).replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g,(match,col,num)=>col+(Number(num)>pivot?Number(num)+offset:num));
     const styleFrom=(rn,col,fallback='0')=>{
       const row=all(x,'row').find(y=>Number(y.getAttribute('r'))===rn);
       return Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+rn)?.getAttribute('s')||fallback;
     };
     const styleA=styleFrom(34,'A','230'),styleB=styleFrom(34,'B','231'),styleC=styleFrom(34,'C','172');
     for(const row of all(x,'row')){
       const old=Number(row.getAttribute('r'));
       if(old>pivot){row.setAttribute('r',String(old+offset));
         for(const c of Array.from(row.children).filter(el=>el.localName==='c')){
           const addr=c.getAttribute('r');if(addr)c.setAttribute('r',shiftedRef(addr));
         }
       }
     }
     for(const m of all(x,'mergeCell')){const ref=m.getAttribute('ref');if(ref)m.setAttribute('ref',shiftedRef(ref));}
     const filt=all(x,'autoFilter')[0];if(filt?.getAttribute('ref'))filt.setAttribute('ref',shiftedRef(filt.getAttribute('ref')));
     for(const br of all(x,'brk'))if(Number(br.getAttribute('id'))>pivot)br.setAttribute('id',String(Number(br.getAttribute('id'))+offset));
     const dim=all(x,'dimension')[0];if(dim?.getAttribute('ref'))dim.setAttribute('ref',shiftedRef(dim.getAttribute('ref')));
     const v=doc.traction?.found;
     const unavailable='Chưa có dữ liệu trong Duty Table / Not available';
     const ropeMatch=v?/^(?:[ØΦφ]\s*)?(\d+(?:[.,]\d+)?)\s*[×xX]\s*(\d+)$/i.exec(String(v.rope||'').trim()):null;
     const rope=ropeMatch?`${ropeMatch[2]}×Φ${ropeMatch[1]}`:null;
     const entries=[
       [34,'Công suất máy kéo (Motor) – kW',v?`Hệ thống máy kéo thang máy không hộp số, vận hành bằng động cơ đồng bộ nam châm vĩnh cửu ${v.motor} kW.`:unavailable,38],
       [35,'Đường kính Puli (Pulley) động cơ (mm)',v?`Φ${v.sheave} mm`:unavailable,22],
       [36,'Cáp tải (Hoisting rope)',v&&rope?`Cáp tải chuyên dụng thang máy, cấu trúc cáp thép bện tròn lõi tẩm dầu ${rope} mm.`:unavailable,38]
     ];
     const rowsNode=all(x,'sheetData')[0];
     for(const [rn,label,value,height] of entries){
       const row=x.createElementNS(n,'row');row.setAttribute('r',String(rn));row.setAttribute('ht',String(height));row.setAttribute('customHeight','1');
       for(const [col,content,style] of [['A',label,styleA],['B','',styleB],['C',value,styleC]]){
         const c=x.createElementNS(n,'c');c.setAttribute('r',col+rn);c.setAttribute('s',style);
         updateCell(x,c,content);row.appendChild(c);
       }
       const following=Array.from(rowsNode.children).find(el=>el.localName==='row'&&Number(el.getAttribute('r'))>rn);
       if(following)rowsNode.insertBefore(row,following);else rowsNode.appendChild(row);
     }
     let merges=all(x,'mergeCells')[0];
     if(!merges){merges=x.createElementNS(n,'mergeCells');
       const before=Array.from(x.documentElement.children).find(el=>['conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup'].includes(el.localName));
       if(before)x.documentElement.insertBefore(merges,before);else x.documentElement.appendChild(merges);
     }
     for(let rn=34;rn<=36;rn++){
       const m=x.createElementNS(n,'mergeCell');m.setAttribute('ref',`A${rn}:B${rn}`);merges.appendChild(m);
     }
     merges.setAttribute('count',String(all(merges,'mergeCell').length));
     const printName=all(workbook,'definedName').find(el=>el.getAttribute('name')==='_xlnm.Print_Area');
     if(printName)printName.textContent=shiftedRef(printName.textContent);
   }

   // MRL Version02: insert Motor only after Control system row 34 (already
   // shifted +2 by the MRL Rail insert). The Regulation / Main code moves to row 36.
   // No Pulley / Hoisting rope for MRL: values not supplied by source.
   if(model==='MRL' && doc.includeTechDetails===true){
     const pivot=34,offset=1;
     const shiftedRef=ref=>String(ref).replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g,(match,col,num)=>col+(Number(num)>pivot?Number(num)+offset:num));
     const styleFrom=(rn,col,fallback='0')=>{
       const row=all(x,'row').find(y=>Number(y.getAttribute('r'))===rn);
       return Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+rn)?.getAttribute('s')||fallback;
     };
     const sa=styleFrom(35,'A','230'),sb=styleFrom(35,'B','231'),sc=styleFrom(35,'C','172');
     for(const row of all(x,'row')){
       const old=Number(row.getAttribute('r'));
       if(old>pivot){row.setAttribute('r',String(old+offset));
         for(const c of Array.from(row.children).filter(el=>el.localName==='c')){
           const addr=c.getAttribute('r');if(addr)c.setAttribute('r',shiftedRef(addr));
         }
       }
     }
     for(const m of all(x,'mergeCell')){const ref=m.getAttribute('ref');if(ref)m.setAttribute('ref',shiftedRef(ref));}
     const dim=all(x,'dimension')[0];if(dim?.getAttribute('ref'))dim.setAttribute('ref',shiftedRef(dim.getAttribute('ref')));
     const filter=all(x,'autoFilter')[0];if(filter?.getAttribute('ref'))filter.setAttribute('ref',shiftedRef(filter.getAttribute('ref')));
     for(const br of all(x,'brk'))if(Number(br.getAttribute('id'))>pivot)br.setAttribute('id',String(Number(br.getAttribute('id'))+offset));
     const motor=doc.mrlMotor?.found;
     const val=motor!==null&&Number.isFinite(motor)
       ?`Hệ thống máy kéo thang máy không hộp số, vận hành bằng động cơ đồng bộ nam châm vĩnh cửu ${motor} kW.`
       :'Chưa có dữ liệu công suất máy kéo EN/MELCO cho tải trọng và tốc độ này / Motor rating unavailable';
     const newRow=x.createElementNS(n,'row');newRow.setAttribute('r','35');newRow.setAttribute('ht','38');newRow.setAttribute('customHeight','1');
     for(const [col,value,style] of [['A','Công suất máy kéo (Motor) – kW',sa],['B','',sb],['C',val,sc]]){
       const c=x.createElementNS(n,'c');c.setAttribute('r',col+'35');c.setAttribute('s',style);updateCell(x,c,value);newRow.appendChild(c);
     }
     const sd=all(x,'sheetData')[0];const following=Array.from(sd.children).find(el=>el.localName==='row'&&Number(el.getAttribute('r'))>35);
     if(following)sd.insertBefore(newRow,following);else sd.appendChild(newRow);
     let merges=all(x,'mergeCells')[0];if(!merges){merges=x.createElementNS(n,'mergeCells');const before=Array.from(x.documentElement.children).find(el=>['conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup'].includes(el.localName));if(before)x.documentElement.insertBefore(merges,before);else x.documentElement.appendChild(merges);}
     const merge=x.createElementNS(n,'mergeCell');merge.setAttribute('ref','A35:B35');merges.appendChild(merge);merges.setAttribute('count',String(all(merges,'mergeCell').length));
     const printName=all(workbook,'definedName').find(el=>el.getAttribute('name')==='_xlnm.Print_Area');if(printName)printName.textContent=shiftedRef(printName.textContent);
   }

   // 5.8.52: make optional, recently added technical export rows easy to identify.
   // Light neutral gray on A:C only; preserve each cell's original borders,
   // font, wrapping and alignment by cloning its existing cellXf.
   // These rows remain editable in Excel and can be hidden or deleted by Sales.
   const optionalTechRows=doc.includeTechDetails===true?(model==='MR'?[34,35,36]:model==='MRL'?[35]:[]):[];
   if(optionalTechRows.length){
     const stylePath='xl/styles.xml',stylePart=zip.file(stylePath);
     if(!stylePart)throw Error('Missing Excel styles for optional technical rows');
     const sd=parse(await stylePart.async('string'));
     const fills=all(sd,'fills')[0],xfs=all(sd,'cellXfs')[0];
     if(!fills||!xfs)throw Error('Missing Excel fill/style records');
     const fill=sd.createElementNS(n,'fill'),pattern=sd.createElementNS(n,'patternFill');
     pattern.setAttribute('patternType','solid');
     const fg=sd.createElementNS(n,'fgColor');fg.setAttribute('rgb','FFF0F0F0');
     const bg=sd.createElementNS(n,'bgColor');bg.setAttribute('indexed','64');
     pattern.append(fg,bg);fill.append(pattern);
     const newFillId=Array.from(fills.children).filter(el=>el.localName==='fill').length;
     fills.appendChild(fill);fills.setAttribute('count',String(newFillId+1));
     const baseXfs=Array.from(xfs.children).filter(el=>el.localName==='xf');
     const initialCount=baseXfs.length,cache=new Map();
     function grayStyle(styleId){
       const base=Number.isInteger(styleId)&&styleId>=0&&styleId<initialCount?styleId:0;
       if(cache.has(base))return cache.get(base);
       const xf=baseXfs[base].cloneNode(true);
       xf.setAttribute('fillId',String(newFillId));xf.setAttribute('applyFill','1');
       const newId=initialCount+cache.size;
       xfs.appendChild(xf);cache.set(base,newId);
       return newId;
     }
     for(const rn of optionalTechRows){
       const row=all(x,'row').find(el=>Number(el.getAttribute('r'))===rn);
       if(!row)throw Error('Missing optional technical row '+rn);
       for(const col of ['A','B','C']){
         const cell=Array.from(row.children).find(el=>el.localName==='c'&&el.getAttribute('r')===col+rn);
         if(!cell)throw Error('Missing optional technical cell '+col+rn);
         cell.setAttribute('s',String(grayStyle(Number(cell.getAttribute('s')||0))));
       }
     }
     xfs.setAttribute('count',String(initialCount+cache.size));
     zip.file(stylePath,serial(sd));
   }

   // v5.8.115: NON-DESTRUCTIVE MRL border repair.
   // Preserve EVERY existing merge (including multi-row Hall indicator groups,
   // A:B label merges, and main/subheadings). Never paint a rectangle over
   // all merged cells: that introduces visible internal lines.
   if(model==='MRL'){
     const stylePath='xl/styles.xml',styleFile=zip.file(stylePath);
     if(!styleFile)throw Error('Missing XLSX styles for MRL border correction');
     const styles=parse(await styleFile.async('string'));
     const borders=all(styles,'borders')[0],xfs=all(styles,'cellXfs')[0];
     if(!borders||!xfs)throw Error('Missing XLSX borders / cellXfs');
     const borderList=Array.from(borders.children).filter(e=>e.localName==='border');
     const originalXfs=Array.from(xfs.children).filter(e=>e.localName==='xf');
     const blackBorderById=new Map(),blackStyleById=new Map();
     function blackBorder(originalId){
       if(blackBorderById.has(originalId))return blackBorderById.get(originalId);
       const base=borderList[originalId]||borderList[0];
       const clone=base.cloneNode(true);
       for(const side of Array.from(clone.children)){
         if(!['left','right','top','bottom','horizontal','vertical'].includes(side.localName))continue;
         // Missing style = intentionally absent border; do not add a line.
         if(!side.getAttribute('style'))continue;
         let color=Array.from(side.children).find(x=>x.localName==='color');
         if(!color){color=styles.createElementNS(n,'color');side.appendChild(color);}
         for(const attr of ['theme','indexed','tint','auto'])color.removeAttribute(attr);
         color.setAttribute('rgb','FF000000');
       }
       const index=all(borders,'border').length;
       borders.appendChild(clone);borders.setAttribute('count',String(index+1));
       blackBorderById.set(originalId,index);return index;
     }
     function blackStyle(id){
       if(blackStyleById.has(id))return blackStyleById.get(id);
       const base=originalXfs[id]||originalXfs[0],xf=base.cloneNode(true);
       xf.setAttribute('borderId',String(blackBorder(Number(base.getAttribute('borderId')||0))));
       xf.setAttribute('applyBorder','1');
       const index=all(xfs,'xf').length;
       xfs.appendChild(xf);blackStyleById.set(id,index);return index;
     }
     for(const row of all(x,'row')){
       const rn=Number(row.getAttribute('r')||0);
       if(rn<13)continue;
       for(const c of Array.from(row.children).filter(e=>e.localName==='c'&&/^[ABC]\d+$/.test(e.getAttribute('r')||''))){
         c.setAttribute('s',String(blackStyle(Number(c.getAttribute('s')||0))));
       }
     }
     // Suppress only the separator between a genuine English/Vietnamese
     // description pair in column C, not the outer borders or merged labels.
     // Use explicit source structure to avoid crossing section/heading rows.
     const rowMap=new Map(all(x,'row').map(row=>[Number(row.getAttribute('r')||0),row]));
     const stringFile=zip.file('xl/sharedStrings.xml');
     const strings=stringFile?all(parse(await stringFile.async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent||'').join('')):[];
     function cellAt(rn,col){return Array.from(rowMap.get(rn)?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+rn);}
     function cellText(rn,col){const c=cellAt(rn,col);if(!c)return '';
       if(c.getAttribute('t')==='inlineStr')return all(c,'t').map(t=>t.textContent||'').join('');
       const v=all(c,'v')[0]?.textContent||'';
       return c.getAttribute('t')==='s'?(strings[Number(v)]||''):v;
     }
     const pairStarts=new Set();
     const exportShift=rn=>rn+(rn>29?2:0)+(doc.includeTechDetails===true&&rn>29?1:0);
     for(const feature of doc.includedFeatures||[]){
       const rr=feature.exportRows;
       if(Array.isArray(rr)&&rr.length===2&&rr[1]===rr[0]+1)pairStarts.add(exportShift(rr[0]));
     }
     for(const rn of [...rowMap.keys()].sort((a,b)=>a-b)){
       if(rn<13||rn>=exportShift(87)||!rowMap.has(rn+1))continue;
       if(cellText(rn,'A').trim()&&cellText(rn,'C').trim()
          &&!cellText(rn+1,'A').trim()&&!cellText(rn+1,'B').trim()
          &&/^\s*\(/.test(cellText(rn+1,'C')))pairStarts.add(rn);
     }
     const pairStyles=new Map();
     function suppressBorder(styleId,edge){
       const key=styleId+':'+edge;
       if(pairStyles.has(key))return pairStyles.get(key);
       const current=all(xfs,'xf'),base=current[styleId];if(!base)return styleId;
       const xf=base.cloneNode(true);
       const border=all(borders,'border')[Number(xf.getAttribute('borderId')||0)];if(!border)return styleId;
       const clone=border.cloneNode(true);
       const side=Array.from(clone.children).find(el=>el.localName===edge);
       if(side){side.removeAttribute('style');for(const child of Array.from(side.children))side.removeChild(child);}
       const borderId=all(borders,'border').length;borders.appendChild(clone);
       xf.setAttribute('borderId',String(borderId));xf.setAttribute('applyBorder','1');
       const nextId=all(xfs,'xf').length;xfs.appendChild(xf);
       pairStyles.set(key,nextId);return nextId;
     }
     for(const top of pairStarts){
       if(rowMap.get(top)?.getAttribute('hidden')==='1'||rowMap.get(top+1)?.getAttribute('hidden')==='1')continue;
       if(!cellText(top,'C').trim()||!cellText(top+1,'C').trim())continue;
       const upper=cellAt(top,'C'),lower=cellAt(top+1,'C');
       if(upper)upper.setAttribute('s',String(suppressBorder(Number(upper.getAttribute('s')||0),'bottom')));
       if(lower)lower.setAttribute('s',String(suppressBorder(Number(lower.getAttribute('s')||0),'top')));
     }
     borders.setAttribute('count',String(all(borders,'border').length));
     xfs.setAttribute('count',String(all(xfs,'xf').length));
     zip.file(stylePath,serial(styles));
     // Fail visibly if future changes produce overlapping merge ranges.
     const ranges=all(x,'mergeCell').map(m=>m.getAttribute('ref')||'');
     const toNum=k=>[...k].reduce((v,ch)=>v*26+ch.charCodeAt(0)-64,0);
     const rect=ref=>{const m=/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref);return m?{
       x1:toNum(m[1]),y1:+m[2],x2:toNum(m[3]||m[1]),y2:+(m[4]||m[2])}:null;};
     for(let i=0;i<ranges.length;i++)for(let j=i+1;j<ranges.length;j++){
       const a=rect(ranges[i]),b=rect(ranges[j]);if(!a||!b)continue;
       if(a.x1<=b.x2&&b.x1<=a.x2&&a.y1<=b.y2&&b.y1<=a.y2)
         throw Error('Overlapping export merges: '+ranges[i]+' / '+ranges[j]);
     }
   }
   // v5.8.117: Keep multi-line hall-panel descriptions readable without
   // over-expanding the merged floor rows. The estimate is applied to the
   // combined height of the existing top/main/other merged blocks.
   if(model==='MRL'){
     const rowMap=new Map(all(x,'row').map(row=>[Number(row.getAttribute('r')||0),row]));
     const actualRow=rn=>rn+(rn>29?2:0)+(doc.includeTechDetails===true&&rn>34?1:0);
     for(const raw of [76,79,82]){
       const rn=actualRow(raw),row=rowMap.get(rn);
       if(!row||row.getAttribute('hidden')==='1')continue;
       const value=String(doc.cells['C'+raw]||'').trim();
       if(!value)continue;
       const lines=bilingualRuns(value).map(run=>run.text).join('').split(/\r?\n/)
         .reduce((sum,line)=>sum+Math.max(1,Math.ceil(line.trim().length/74)),0);
       // Merged floor panels are 3 source rows; set the combined height on
       // the leading row while preserving the other rows' source dimensions.
       const following=[rowMap.get(rn+1),rowMap.get(rn+2)].filter(Boolean);
       const remaining=following.reduce((sum,r)=>sum+Number(r.getAttribute('ht')||15),0);
       const desired=Math.max(66,Math.min(130,lines*15+11));
       row.setAttribute('ht',String(Math.max(16,Math.round(desired-remaining))));
       row.setAttribute('customHeight','1');
     }
   }
   // v5.8.117: Arial for both MR/MRL. The XLSX rich text below uses the
   // same family and preserves bold/italics from the cell's base style.
   {
     const stylesPath='xl/styles.xml',part=zip.file(stylesPath);
     if(!part)throw Error('Missing Excel styles for Arial normalization');
     const sd=parse(await part.async('string'));
     const fonts=all(sd,'fonts')[0];
     if(!fonts)throw Error('Missing Excel fonts');
     for(const font of Array.from(fonts.children).filter(node=>node.localName==='font')){
       let name=Array.from(font.children).find(node=>node.localName==='name');
       if(!name){name=sd.createElementNS(n,'name');font.insertBefore(name,font.firstChild);}
       name.setAttribute('val','Arial');
     }
     zip.file(stylesPath,serial(sd));
   }
   // v5.8.117: Rich-text italic on English descriptions, not just the
   // entire cell font. Existing English on the same line is moved below the
   // Vietnamese text; technical tokens (SUS304, kg, PIV) are untouched.
   for(const c of all(x,'c')){
     const addr=c.getAttribute('r')||'';
     if(!/^C\d+$/.test(addr)||c.getAttribute('t')!=='inlineStr')continue;
     const source=all(c,'t').map(t=>t.textContent||'').join('');
     if(!source||!/[À-ỹ]/.test(source))continue;
     const runs=bilingualRuns(source);
     if(!runs.some(t=>t.english))continue;
     const is=child(c,'is');if(!is)continue;
     while(is.firstChild)is.removeChild(is.firstChild);
     for(const part of runs){
       if(!part.text)continue;
       const rNode=x.createElementNS(n,'r');
       const prop=x.createElementNS(n,'rPr');
       const family=x.createElementNS(n,'rFont');family.setAttribute('val','Arial');prop.appendChild(family);
       if(part.english)prop.appendChild(x.createElementNS(n,'i'));
       rNode.appendChild(prop);
       const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');
       t.textContent=part.text;rNode.appendChild(t);is.appendChild(rNode);
     }
   }
   // v5.8.119: split single-row bilingual descriptions into TWO Excel rows.
   // Feature descriptions already encoded as two source rows remain untouched.
   // Preserve the worksheet's original outside borders; the only new internal
   // separator is a thin pale-grey rule above the new English row.
   {
     const data=all(x,'sheetData')[0];
     const stylePart=zip.file('xl/styles.xml');
     if(!data||!stylePart)throw Error('Missing worksheet/styles for bilingual row splitting');
     const sd=parse(await stylePart.async('string'));
     const fonts=all(sd,'fonts')[0],xfs=all(sd,'cellXfs')[0],borders=all(sd,'borders')[0];
     if(!fonts||!xfs||!borders)throw Error('Missing styles for bilingual row splitting');
     const styleCache=new Map(),vietnameseStyleCache=new Map();
     function vietnameseStyle(baseId){
       if(vietnameseStyleCache.has(baseId))return vietnameseStyleCache.get(baseId);
       const original=all(xfs,'xf')[baseId]||all(xfs,'xf')[0],xf=original.cloneNode(true);
       const originalBorder=all(borders,'border')[Number(original.getAttribute('borderId')||0)]||all(borders,'border')[0];
       const b=originalBorder.cloneNode(true);
       const bottom=Array.from(b.children).find(e=>e.localName==='bottom');
       if(bottom){bottom.removeAttribute('style');while(bottom.firstChild)bottom.removeChild(bottom.firstChild);}
       const id=all(borders,'border').length;borders.appendChild(b);borders.setAttribute('count',String(id+1));
       xf.setAttribute('borderId',String(id));xf.setAttribute('applyBorder','1');
       const sid=all(xfs,'xf').length;xfs.appendChild(xf);xfs.setAttribute('count',String(sid+1));
       vietnameseStyleCache.set(baseId,sid);return sid;
     }
     function englishStyle(baseId){
       if(styleCache.has(baseId))return styleCache.get(baseId);
       const original=all(xfs,'xf')[baseId]||all(xfs,'xf')[0];
       const xf=original.cloneNode(true);
       const fontId=Number(original.getAttribute('fontId')||0);
       const baseFont=all(fonts,'font')[fontId]||all(fonts,'font')[0];
       const font=baseFont.cloneNode(true);
       if(!Array.from(font.children).some(e=>e.localName==='i'))font.appendChild(sd.createElementNS(n,'i'));
       const nf=all(fonts,'font').length;fonts.appendChild(font);fonts.setAttribute('count',String(nf+1));
       xf.setAttribute('fontId',String(nf));xf.setAttribute('applyFont','1');
       const baseBorder=all(borders,'border')[Number(original.getAttribute('borderId')||0)]||all(borders,'border')[0];
       const b=baseBorder.cloneNode(true);
       let top=Array.from(b.children).find(e=>e.localName==='top');
       if(!top){top=sd.createElementNS(n,'top');b.appendChild(top);}
       top.setAttribute('style','thin');while(top.firstChild)top.removeChild(top.firstChild);
       const col=sd.createElementNS(n,'color');col.setAttribute('rgb','FFE2E7EC');top.appendChild(col);
       const nb=all(borders,'border').length;borders.appendChild(b);borders.setAttribute('count',String(nb+1));
       xf.setAttribute('borderId',String(nb));xf.setAttribute('applyBorder','1');
       const ns=all(xfs,'xf').length;xfs.appendChild(xf);xfs.setAttribute('count',String(ns+1));
       styleCache.set(baseId,ns);return ns;
     }
     const content=c=>all(c,'t').map(t=>t.textContent||'').join('');
     const getCell=(row,col)=>Array.from(row?.children||[]).find(e=>e.localName==='c' && e.getAttribute('r')===col+row.getAttribute('r'));
     const rows=()=>Array.from(data.children).filter(e=>e.localName==='row');
     const cellMap=()=>new Map(rows().map(row=>[Number(row.getAttribute('r')),row]));
     const mergeRoot=all(x,'mergeCells')[0];
     const mergeEls=()=>mergeRoot?Array.from(mergeRoot.children).filter(e=>e.localName==='mergeCell'):[];
     const parseRange=ref=>{const m=/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref||'');return m?{a:m[1],r1:+m[2],b:m[3]||m[1],r2:+(m[4]||m[2])}:null};
     const writeRange=q=>q.a+q.r1+':'+q.b+q.r2;
     const colNum=v=>[...v].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0);
     function makeMerge(ref){if(!mergeRoot)return;
       const q=parseRange(ref);if(!q)return;
       if(mergeEls().some(el=>{const p=parseRange(el.getAttribute('ref'));return p&&colNum(p.a)<=colNum(q.b)&&colNum(q.a)<=colNum(p.b)&&p.r1<=q.r2&&q.r1<=p.r2;}))return;
       const m=x.createElementNS(n,'mergeCell');m.setAttribute('ref',ref);mergeRoot.appendChild(m);
     }
     function boundaryAdjust(at){
       // Every existing merge is shifted or vertically extended through insertion.
       for(const m of mergeEls()){
         const q=parseRange(m.getAttribute('ref'));if(!q)continue;
         if(q.r1>at){q.r1++;q.r2++;}
         else if(q.r2>=at){q.r2++;}
         m.setAttribute('ref',writeRange(q));
       }
       // Row and cell references, plus sheet-local print/page references.
       for(const row of rows()){
         const old=Number(row.getAttribute('r'));
         if(old<=at)continue;
         row.setAttribute('r',String(old+1));
         for(const cell of Array.from(row.children).filter(e=>e.localName==='c')){
           cell.setAttribute('r',(cell.getAttribute('r')||'').replace(/\d+$/,String(old+1)));
         }
       }
       const shiftAddress=v=>v.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g,(whole,col,num)=>col+(Number(num)>at?Number(num)+1:num));
       // The exporter retains a single printable sheet, but its page breaks,
       // drawing anchors, and print areas must follow the inserted rows.
       for(const el of all(x,'hyperlink'))if(el.hasAttribute('ref'))el.setAttribute('ref',shiftAddress(el.getAttribute('ref')));
       for(const el of all(x,'rowBreaks'))for(const br of Array.from(el.children))if(br.localName==='brk'&&+br.getAttribute('id')>at)br.setAttribute('id',String(+br.getAttribute('id')+1));
       for(const el of all(x,'dimension'))if(el.hasAttribute('ref'))el.setAttribute('ref',shiftAddress(el.getAttribute('ref')));
       for(const el of all(x,'conditionalFormatting'))if(el.hasAttribute('sqref'))el.setAttribute('sqref',shiftAddress(el.getAttribute('sqref')));
       for(const el of all(x,'dataValidation'))if(el.hasAttribute('sqref'))el.setAttribute('sqref',shiftAddress(el.getAttribute('sqref')));
     }
     function splitPair(str){
       const rr=bilingualRuns(str);
       const english=rr.filter(t=>t.english);
       if(english.length!==1)return null;
       const idx=rr.findIndex(t=>t.english);
       const before=rr.slice(0,idx).map(t=>t.text).join('').trim();
       const after=rr.slice(idx+1).map(t=>t.text).join('').trim();
       const en=english[0].text.trim();
       if(!before||!/[À-ỹ]/.test(before)||!en||after)return null;
       return [before,en];
     }
     // v5.8.120: Restore the THREE actual source rows of each MRL hall group.
     // Previously two rows were hidden and Vietnamese, English and quantity
     // were squeezed into one big wrapped cell. Keep the factory A/B merges.
     const restoredHallRows=new Set();
     if(model==='MRL'){
       const floorNames=['Tầng trên cùng','Tầng chính','Tầng khác'];
       function ensureCell(row,col){
         let cell=getCell(row,col);
         if(!cell){cell=x.createElementNS(n,'c');cell.setAttribute('r',col+row.getAttribute('r'));row.appendChild(cell);}
         return cell;
       }
       function displayLines(value){
         const runs=bilingualRuns(value);
         const notes=[];
         const vi=runs.filter(part=>!part.english).map(part=>part.text).join('')
           .split(/\r?\n/).map(line=>line.trim()).filter(line=>{
             if(!line)return false;
             if(/^\d+\s*hộp\s*\//i.test(line)){notes.push(line);return false;}
             return true;
           });
         const en=runs.filter(part=>part.english).map(part=>part.text.trim()).filter(Boolean);
         return {vi:vi.join('\n'),en:en.join('\n'),notes:notes.join('\n')};
       }
       for(const name of floorNames){
         const top=rows().find(row=>{
           const label=content(getCell(row,'B')||x.createElementNS(n,'c'));
           return label.includes(name) && !!content(getCell(row,'C')||x.createElementNS(n,'c')).trim() &&
             content(getCell(row,'A')||x.createElementNS(n,'c')).includes('Hiển thị');
         }) || rows().find(row=>content(getCell(row,'B')||x.createElementNS(n,'c')).includes(name)
           && content(getCell(row,'C')||x.createElementNS(n,'c')).includes('Hall indicators'));
         if(!top)continue;
         const rn=+top.getAttribute('r'), rowMap=cellMap();
         const englishRow=rowMap.get(rn+1),quantityRow=rowMap.get(rn+2);
         if(!englishRow||!quantityRow)continue;
         const primary=ensureCell(top,'C');
         const parts=displayLines(content(primary));
         if(!parts.vi||!parts.en)continue; // preserve original layout if unrecognizable
         const baseStyle=Number(primary.getAttribute('s')||0);
         // Thin pale separator only between VI and EN. Do not add thick borders.
         updateCell(x,primary,parts.vi);primary.setAttribute('s',String(vietnameseStyle(baseStyle)));
         const enCell=ensureCell(englishRow,'C');updateCell(x,enCell,parts.en);
         enCell.setAttribute('s',String(englishStyle(vietnameseStyle(baseStyle))));
         const qtyCell=ensureCell(quantityRow,'C');updateCell(x,qtyCell,parts.notes);
         qtyCell.setAttribute('s',String(baseStyle));
         const estimate=(value,chars)=>String(value||'').split('\n').reduce((sum,line)=>sum+Math.max(1,Math.ceil(line.length/chars)),0);
         for(const [rr,value] of [[top,parts.vi],[englishRow,parts.en],[quantityRow,parts.notes]]){
           rr.removeAttribute('hidden');rr.setAttribute('customHeight','1');
           rr.setAttribute('ht',String(Math.max(17,Math.min(145,estimate(value,80)*15+5))));
           restoredHallRows.add(+rr.getAttribute('r'));
         }
       }
     }
     const initial=rows().map(row=>({row,number:+row.getAttribute('r')}));
     // Bilingual form body, excluding already-split lines and existing
     // multi-level hall-panel 3-row blocks (which have a separate layout).
     const selected=[];
     const originalMap=cellMap();
     for(const {row,number:rn} of initial){
       if(rn<13||rn>340||row.getAttribute('hidden')==='1')continue;
       const c=getCell(row,'C');if(!c||c.getAttribute('t')!=='inlineStr')continue;
       const next=originalMap.get(rn+1),nextC=next&&getCell(next,'C');
       if(nextC && /^\s*\(/.test(content(nextC)) && !(getCell(next,'A')&&content(getCell(next,'A')).trim()) && !(getCell(next,'B')&&content(getCell(next,'B')).trim()))continue;
       const value=content(c),pair=splitPair(value);if(!pair)continue;
       // Hall indicator blocks are already special merged 3-row records;
       // avoid splitting the combined count/description into an invalid merge.
       if(/\b(?:PIV\d|HBV\d|PIH[-\d]|\d+\s*hộp\s*\/)/i.test(value)&&rn>=70&&rn<110)continue;
       if(mergeEls().some(el=>{const q=parseRange(el.getAttribute('ref'));return q&&colNum(q.a)<=3&&colNum(q.b)>=3&&q.r1<=rn&&q.r2>rn;}))continue;
       selected.push({row,cell:c,pair});
     }
     // Work from the bottom up: each insertion preserves prior row addresses.
     for(const item of selected.reverse()){
       const row=item.row,rn=+row.getAttribute('r'),oldStyle=Number(item.cell.getAttribute('s')||0);
       boundaryAdjust(rn);
       const extra=x.createElementNS(n,'row');extra.setAttribute('r',String(rn+1));
       extra.setAttribute('customHeight','1');
       const enLines=item.pair[1].split('\n').reduce((sum,line)=>sum+Math.max(1,Math.ceil(line.length/77)),0);
       const viLines=item.pair[0].split('\n').reduce((sum,line)=>sum+Math.max(1,Math.ceil(line.length/77)),0);
       extra.setAttribute('ht',String(Math.max(16,Math.min(85,enLines*14+3))));
       row.setAttribute('ht',String(Math.max(16,Math.min(85,viLines*14+3))));row.setAttribute('customHeight','1');
       updateCell(x,item.cell,item.pair[0]);item.cell.setAttribute('s',String(vietnameseStyle(oldStyle)));
       const en=x.createElementNS(n,'c');en.setAttribute('r','C'+(rn+1));
       en.setAttribute('s',String(englishStyle(oldStyle)));updateCell(x,en,item.pair[1]);extra.appendChild(en);
       data.insertBefore(extra,row.nextSibling);
       // If label was not already merged over the two rows, merge it now.
       for(const col of ['A','B']){
         const label=getCell(row,col);if(!label||!content(label).trim())continue;
         const covered=mergeEls().some(el=>{const q=parseRange(el.getAttribute('ref'));return q&&q.r1<=rn&&q.r2>=rn+1&&colNum(q.a)<=colNum(col)&&colNum(q.b)>=colNum(col);});
         if(!covered)makeMerge(col+rn+':'+col+(rn+1));
       }
     }
     // v5.8.121: normalize EVERY already-separated bilingual C-cell pair
     // in BOTH models, not only a hard-coded English phrase whitelist.
     // The upper Vietnamese row loses its bottom edge and the lower English
     // row receives one SOLID THIN pale-grey top edge (never dotted "hair").
     const afterSplit=cellMap();
     // Factory MRL entrance rows sometimes repeat English inside the
     // Vietnamese cell while the actual English row remains hidden. Restore
     // those existing lines instead of leaving a false single-row bilingual
     // cell (no row insertions, so all factory merges stay intact).
     if(model==='MRL')for(const [rn,viRow] of afterSplit){
       const enRow=afterSplit.get(rn+1);
       if(!enRow||enRow.getAttribute('hidden')!=='1')continue;
       const viCell=getCell(viRow,'C'),enCell=getCell(enRow,'C');
       if(!viCell||!enCell)continue;
       let pair=splitPair(content(viCell));
       if(!pair){
         const current=content(viCell),alreadyEnglish=content(enCell).trim();
         if(current.endsWith(alreadyEnglish)&&/[À-ỹ]/.test(current))
           pair=[current.slice(0,-alreadyEnglish.length).trim(),alreadyEnglish];
       }
       if(!pair||!/^\s*\([A-Za-z]/.test(content(enCell)))continue;
       const normalize=t=>String(t||'').replace(/\s+/g,' ').trim().replace(/[. ]+$/,'');
       if(normalize(pair[1])!==normalize(content(enCell)))continue;
       if(!/Jamb|Landing door|Khung cửa|Cửa tầng/i.test(
         content(getCell(viRow,'A')||x.createElementNS(n,'c'))))continue;
       updateCell(x,viCell,pair[0]);updateCell(x,enCell,pair[1]);
       enRow.removeAttribute('hidden');
       for(const [rr,v] of [[viRow,pair[0]],[enRow,pair[1]]]){
         rr.setAttribute('customHeight','1');
         rr.setAttribute('ht',String(Math.max(17,Math.min(64,
           v.split('\n').reduce((a,line)=>a+Math.max(1,Math.ceil(line.length/72)),0)*15+3))));
       }
     }
     const isEnglishDescription=value=>{
       const str=String(value||'').trim().replace(/^\(/,'').replace(/\)$/,'').trim();
       if(!str||/[À-ỹ]/.test(str)||/^[0-9]+\s*hộp\s*\//i.test(str))return false;
       return /\b(?:provided|not supplied|not provided|stainless|steel|hairline|glass|aluminum|aluminium|ceiling|door|floor|hall|indicators?|buttons?|emergency|ventilation|lighting|model|car|motor|frame|painted|granite|with|without|landing|machine|power|traction|safety|panel|handrail|mirror|fireproof|sensor|finish|material|side|operation|supply|flooring|handrails?|operating|buttons|braille|marble|stainless|steel|wheelchair)\b/i.test(str);
     };
     for(const [rn,row] of afterSplit){
       const next=afterSplit.get(rn+1);
       if(!next||row.getAttribute('hidden')==='1'||next.getAttribute('hidden')==='1')continue;
       // A hall triple is detected by its content, not stale pre-insertion row numbers.
       const qtyRow=afterSplit.get(rn+2);
       if(/Hiển thị & nút gọi tầng/i.test(content(getCell(row,'C')||x.createElementNS(n,'c'))) &&
         /Hall indicators/i.test(content(getCell(next,'C')||x.createElementNS(n,'c'))) &&
         /^\s*\d+\s*hộp\s*\//i.test(content(getCell(qtyRow,'C')||x.createElementNS(n,'c'))))continue;
       const upper=getCell(row,'C'),lower=getCell(next,'C');
       if(!upper||!lower)continue;
       const vi=content(upper).trim(),en=content(lower).trim();
       if(!vi||!en||!/[À-ỹ]/.test(vi)||!(isEnglishDescription(en)||(/^\s*\([A-Za-z]/.test(en)&&en.length>12)))continue;
       // A new label in the following row marks the start of a distinct item.
       if(content(getCell(next,'A')||x.createElementNS(n,'c')).trim()||
          content(getCell(next,'B')||x.createElementNS(n,'c')).trim())continue;
       const upperBase=Number(upper.getAttribute('s')||0);
       // Never derive the English cell from the Vietnamese style: the latter
       // intentionally has NO bottom border. Using it for both rows erased
       // the OUTER bottom edge (MR rows 71-72/88-89 and MRL equivalents).
       lower.setAttribute('s',String(englishStyle(Number(lower.getAttribute('s')||0))));
       upper.setAttribute('s',String(vietnameseStyle(upperBase)));
     }
     // v5.8.123: final perimeter repair for BOTH source formats.
     // Separate VI/EN rows use exactly one light separator BETWEEN the rows,
     // while the bottom of the English row remains a normal black table edge.
     // Do not paint generic borders on merged regions or introduce thick rules.
     const perimeterStyles=new Map();
     function perimeterStyle(styleId, edits){
       const key=styleId+':'+Object.entries(edits).map(([k,v])=>k+'='+v).join(';');
       if(perimeterStyles.has(key))return perimeterStyles.get(key);
       const base=all(xfs,'xf')[styleId]||all(xfs,'xf')[0],xf=base.cloneNode(true);
       const baseBorder=all(borders,'border')[Number(base.getAttribute('borderId')||0)]||all(borders,'border')[0];
       const border=baseBorder.cloneNode(true);
       for(const [sideName,value] of Object.entries(edits)){
         let side=Array.from(border.children).find(e=>e.localName===sideName);
         if(!side){side=sd.createElementNS(n,sideName);border.appendChild(side);}
         while(side.firstChild)side.removeChild(side.firstChild);
         if(value==='none'){side.removeAttribute('style');continue;}
         side.setAttribute('style','thin');
         const color=sd.createElementNS(n,'color');
         color.setAttribute('rgb',value==='separator'?'FFE2E7EC':'FF000000');
         side.appendChild(color);
       }
       const bid=all(borders,'border').length;borders.appendChild(border);
       xf.setAttribute('borderId',String(bid));xf.setAttribute('applyBorder','1');
       const id=all(xfs,'xf').length;xfs.appendChild(xf);
       perimeterStyles.set(key,id);return id;
     }
     function hasBilingualPair(rn){
       const upperRow=afterSplit.get(rn),lowerRow=afterSplit.get(rn+1);
       if(!upperRow||!lowerRow||upperRow.getAttribute('hidden')==='1'||lowerRow.getAttribute('hidden')==='1')return false;
       const qtyRow=afterSplit.get(rn+2);
       if(/Hiển thị & nút gọi tầng/i.test(content(getCell(upperRow,'C')||x.createElementNS(n,'c'))) &&
         /Hall indicators/i.test(content(getCell(lowerRow,'C')||x.createElementNS(n,'c'))) &&
         /^\s*\d+\s*hộp\s*\//i.test(content(getCell(qtyRow,'C')||x.createElementNS(n,'c'))))return false;
       const vi=content(getCell(upperRow,'C')||x.createElementNS(n,'c')).trim();
       const en=content(getCell(lowerRow,'C')||x.createElementNS(n,'c')).trim();
       if(!vi||!en||!/[À-ỹ]/.test(vi)||!(isEnglishDescription(en)||(/^\s*\([A-Za-z]/.test(en)&&en.length>12)))return false;
       return !(content(getCell(lowerRow,'A')||x.createElementNS(n,'c')).trim()||
         content(getCell(lowerRow,'B')||x.createElementNS(n,'c')).trim());
     }
     for(const [rn,row] of afterSplit){
       if(!hasBilingualPair(rn))continue;
       const bottomRow=afterSplit.get(rn+1),vi=getCell(row,'C'),en=getCell(bottomRow,'C');
       vi.setAttribute('s',String(perimeterStyle(Number(vi.getAttribute('s')||0),
         {left:'table',right:'table',bottom:'none'})));
       en.setAttribute('s',String(perimeterStyle(Number(en.getAttribute('s')||0),
         {left:'table',right:'table',top:'separator',bottom:'table'})));
       // The fireproof-door label must span A:B *and* both language rows.
       // MR's source template had A88:A89 only; MRL already used A:B.
       const label=content(getCell(row,'A')||x.createElementNS(n,'c')).trim();
       if(/Cửa chống cháy|Fireproof door|Giới hạn chịu lửa cửa tầng|Entrance Fire resistance/i.test(label)){
         const candidates=mergeEls().filter(el=>{const q=parseRange(el.getAttribute('ref'));
           return q&&q.a==='A'&&q.r1===rn&&q.r2===rn+1;});
         if(candidates.length===1){
           const q=parseRange(candidates[0].getAttribute('ref'));
           if(q.b==='A'){
             const overlaps=mergeEls().some(el=>el!==candidates[0]&&(()=>{
               const p=parseRange(el.getAttribute('ref'));
               return p&&colNum(p.a)<=2&&colNum(p.b)>=1&&p.r1<=rn+1&&p.r2>=rn;})());
             if(!overlaps && !content(getCell(row,'B')||x.createElementNS(n,'c')).trim())
               candidates[0].setAttribute('ref','A'+rn+':B'+(rn+1));
           }
         }
         // Every side of a merged 2x2 perimeter needs an edge on the last
         // row; otherwise Excel displays a missing/dotted bottom border.
         const makeOrGet=(target,col)=>{
           let cell=getCell(target,col);
           if(!cell){cell=x.createElementNS(n,'c');cell.setAttribute('r',col+target.getAttribute('r'));target.appendChild(cell);}
           return cell;
         };
         for(const col of ['A','B']){
           const first=makeOrGet(row,col),last=makeOrGet(bottomRow,col);
           const originalId=Number(first.getAttribute('s')||0);
           first.setAttribute('s',String(perimeterStyle(originalId,
             {left:col==='A'?'table':'none',right:col==='B'?'table':'none',bottom:'none'})));
           last.setAttribute('s',String(perimeterStyle(originalId,
             {left:col==='A'?'table':'none',right:col==='B'?'table':'none',top:'none',bottom:'table'})));
         }
         // v5.8.126: The 2x2 merged label A:B must WRAP as one Excel cell.
         // Previously A had no wrapText and the extended Fire Rating label
         // was truncated, even though the merge itself was valid.
         const firstLabel=makeOrGet(row,'A');
         const originalStyle=all(xfs,'xf')[Number(firstLabel.getAttribute('s')||0)];
         if(originalStyle){
           const wrapped=originalStyle.cloneNode(true);
           let align=Array.from(wrapped.children).find(e=>e.localName==='alignment');
           if(!align){align=sd.createElementNS(n,'alignment');
             const next=Array.from(wrapped.children).find(e=>['protection','extLst'].includes(e.localName));
             if(next)wrapped.insertBefore(align,next);else wrapped.appendChild(align);}
           align.setAttribute('wrapText','1');align.setAttribute('vertical','center');
           wrapped.setAttribute('applyAlignment','1');
           const newStyleId=all(xfs,'xf').length;xfs.appendChild(wrapped);
           firstLabel.setAttribute('s',String(newStyleId));
         }
         // The title occupies about three Arial 11pt lines at A:B width.
         // Give the two merged rows enough combined height without widening
         // all other descriptions or introducing extra blank lines.
         for(const [r,minHeight] of [[row,29],[bottomRow,25]]){
           const existing=Number(r.getAttribute('ht')||0);
           r.setAttribute('ht',String(Math.max(existing,minHeight)));
           r.setAttribute('customHeight','1');
         }
       }
     }
     // MRL hall groups have 3 physical rows (VI / EN / quantity).
     // Find them by CONTENT after all row insertions (stored row numbers can
     // shift when earlier bilingual cells were split into two Excel rows).
     // English is the middle row; ONLY the quantity row closes the box.
     if(model==='MRL')for(const [rn,viRow] of afterSplit){
       const enRow=afterSplit.get(rn+1),qtyRow=afterSplit.get(rn+2);
       if(!enRow||!qtyRow)continue;
       const vi=getCell(viRow,'C'),en=getCell(enRow,'C'),qty=getCell(qtyRow,'C');
       if(!vi||!en||!qty)continue;
       if(!/Hiển thị/i.test(content(vi)) || !/Hall indicators/i.test(content(en)) ||
          !/^\s*\d+\s*hộp\s*\//i.test(content(qty)))continue;
       vi.setAttribute('s',String(perimeterStyle(Number(vi.getAttribute('s')||0),
         {left:'table',right:'table',bottom:'none'})));
       en.setAttribute('s',String(perimeterStyle(Number(en.getAttribute('s')||0),
         {left:'table',right:'table',top:'separator',bottom:'none'})));
       qty.setAttribute('s',String(perimeterStyle(Number(qty.getAttribute('s')||0),
         {left:'table',right:'table',top:'none',bottom:'table'})));
     }
     if(mergeRoot)mergeRoot.setAttribute('count',String(mergeEls().length));
     const print=all(workbook,'definedName').find(el=>el.getAttribute('name')==='_xlnm.Print_Area');
     if(print){const last=rows().reduce((m,row)=>Math.max(m,+row.getAttribute('r')),1);print.textContent=print.textContent.replace(/(\$[A-Z]+\$)\d+\s*$/,(_,col)=>col+last);}
     // v5.8.124: OOXML integrity guard. Styles created during the last
     // bilingual perimeter pass MUST be included in collection counts.
     // More importantly, CT_Border children have a fixed schema order:
     // <left>, <right>, <top>, <bottom>... Earlier code appended missing
     // sides at the end, producing invalid records that Excel repairs.
     const borderSequence=['left','right','top','bottom','diagonal','vertical','horizontal','start','end'];
     const seqIndex=name=>{const i=borderSequence.indexOf(name);return i<0?100:i;};
     for(const border of all(borders,'border')){
       const sides=Array.from(border.children);
       const seen=new Set();
       for(const node of sides){
         if(seen.has(node.localName))throw Error('XLSX: duplicate border edge '+node.localName);
         seen.add(node.localName);
       }
       const ordered=sides.slice().sort((a,b)=>seqIndex(a.localName)-seqIndex(b.localName));
       if(sides.some((node,i)=>node!==ordered[i]))for(const node of ordered)border.appendChild(node);
     }
     for(const [tag,member] of [
       ['numFmts','numFmt'],['fonts','font'],['fills','fill'],['borders','border'],
       ['cellStyleXfs','xf'],['cellXfs','xf'],['cellStyles','cellStyle'],['dxfs','dxf']
     ]){
       const collection=all(sd,tag)[0];
       if(collection)collection.setAttribute('count',String(Array.from(collection.children).filter(e=>e.localName===member).length));
     }
     const styleRecords=all(xfs,'xf'),borderRecords=all(borders,'border');
     const fontsCount=all(all(sd,'fonts')[0],'font').length;
     const fillsCount=all(all(sd,'fills')[0],'fill').length;
     for(const [i,xf] of styleRecords.entries()){
       for(const [attr,max] of [['borderId',borderRecords.length],['fontId',fontsCount],['fillId',fillsCount]]){
         const id=Number(xf.getAttribute(attr)||0);
         if(!Number.isInteger(id)||id<0||id>=max)throw Error('XLSX: invalid '+attr+' at cellXf '+i);
       }
     }
     for(const c of all(x,'c')){
       const id=Number(c.getAttribute('s')||0);
       if(!Number.isInteger(id)||id<0||id>=styleRecords.length)
         throw Error('XLSX: invalid cell style '+c.getAttribute('r')+' -> '+id);
     }
     // v5.8.126: Microsoft Excel repairs sheet2.xml when cells in a row
     // are written out-of-order. Previous Fire Rating perimeter repair
     // appended A/B after C, creating <c r="C93"/><c r="A93"/><c r="B93"/>.
     // Enforce strictly ascending cell references in EVERY row, independent
     // of the source model, and refuse duplicate/foreign-row cells.
     const excelColNumber=letters=>[...letters].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0);
     for(const worksheetRow of all(x,'row')){
       const rowNumber=Number(worksheetRow.getAttribute('r')||0);
       const cells=Array.from(worksheetRow.children).filter(el=>el.localName==='c');
       const sorted=cells.map(el=>{
         const addr=el.getAttribute('r')||'';
         const match=/^([A-Z]+)([1-9][0-9]*)$/.exec(addr);
         if(!match||Number(match[2])!==rowNumber)
           throw Error('XLSX: incorrect cell coordinate '+addr+' on row '+rowNumber);
         return {el,addr,col:excelColNumber(match[1])};
       }).sort((a,b)=>a.col-b.col);
       for(let i=1;i<sorted.length;i++)if(sorted[i].col===sorted[i-1].col)
         throw Error('XLSX: duplicate cell '+sorted[i].addr);
       // Always place the cell sequence before the optional extLst node.
       const ext=Array.from(worksheetRow.children).find(el=>el.localName==='extLst')||null;
       for(const entry of sorted)worksheetRow.insertBefore(entry.el,ext);
     }
     zip.file('xl/styles.xml',serial(sd));
   }
   // v5.8.131: OPTIONAL Ride Comfort & Noise (Vietnamese-only, columns A:C).
   // Insert AFTER the motor/traction detail, before the Main Code. The data are
   // model-specific; MRL speed is read from SpecForm D17 (mpm), not another UI input.
   // The same XLSX is the source of the printable PDF, ensuring export parity.
   if(doc.includeTechDetails===true){
     const speedMpm=model==='MRL'?Number(doc.mrlMotor?.speed):Number(doc.traction?.speed);
     const highSpeed=model==='MRL'&&speedMpm>=120;
     // v5.8.138: concise grouping requested for both MR/MRL.
     // Peak-start/stop and constant-speed vertical/horizontal vibration share one row.
     const rideRows=[
       ['Độ giật khi khởi động / dừng (m/s²)','0–Pmax 0,30 (30 gal) hoặc thấp hơn'],
       ['Gia tốc (m/s²)','0,60 ± 0,10'],
       ['Gia tốc giảm tốc (m/s²)','0,60 ± 0,10'],
       ['Rung đứng / rung ngang – tốc độ không đổi (m/s²)',`Đứng: ${highSpeed?'0,25':'0,20'}; ngang: 0,30 (P–Pmax, hoặc thấp hơn)`],
       ['Độ chính xác dừng tầng (mm)','Trong phạm vi ±5 mm']
     ];
     const noiseRows=model==='MR'?[
       ['Độ ồn trong cabin [dB(A)]','≤ 50 (tối đa 55)'],
       ['Độ ồn tại sảnh tầng [dB(A)]','≤ 50'],
       ['Độ ồn khi cửa mở và đóng [dB(A)]','≤ 52'],
       ['Độ ồn phòng máy [dB(A)]','≤ 80']
     ]:[
       ['Độ ồn trong cabin [dB(A)]',highSpeed?'≤ 55 (tối đa 60)':'≤ 50 (tối đa 55)'],
       ['Độ ồn tại sảnh tầng [dB(A)]','≤ 50 (tối đa 55)'],
       ['Độ ồn khi cửa mở và đóng [dB(A)]','≤ 52']
     ];
     // Earlier bilingual splitting may insert rows BEFORE the Motor block. Locate
     // its LAST physical row by the actual label, never by the old row number.
     const anchorLabel=model==='MR'?/Cáp tải \(Hoisting rope\)/:/Công suất máy kéo \(Motor\)/;
     const anchorRow=all(x,'row').find(row=>Array.from(row.children).some(c=>c.localName==='c'&&/^A\d+$/.test(c.getAttribute('r')||'')&&anchorLabel.test(all(c,'t').map(t=>t.textContent||'').join(''))));
     if(!anchorRow)throw Error('Ride Comfort & Noise: cannot locate motor detail insertion point');
     const insertAfter=Number(anchorRow.getAttribute('r'));
     const total=rideRows.length+noiseRows.length;
     const targetRows=all(x,'row');
     const shiftReference=str=>String(str).replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g,(match,col,num)=>col+(Number(num)>insertAfter?Number(num)+total:num));
     for(const row of targetRows){
       const old=Number(row.getAttribute('r'));
       if(old<=insertAfter)continue;
       row.setAttribute('r',String(old+total));
       for(const c of Array.from(row.children).filter(el=>el.localName==='c')){
         if(c.getAttribute('r'))c.setAttribute('r',shiftReference(c.getAttribute('r')));
       }
     }
     for(const m of all(x,'mergeCell')){
       const ref=m.getAttribute('ref');if(ref)m.setAttribute('ref',shiftReference(ref));
     }
     const dim=all(x,'dimension')[0];if(dim?.getAttribute('ref'))dim.setAttribute('ref',shiftReference(dim.getAttribute('ref')));
     for(const br of all(x,'brk'))if(Number(br.getAttribute('id'))>insertAfter)br.setAttribute('id',String(Number(br.getAttribute('id'))+total));
     const printName=all(workbook,'definedName').find(el=>el.getAttribute('name')==='_xlnm.Print_Area');
     if(printName)printName.textContent=shiftReference(printName.textContent);
     const rowsRoot=all(x,'sheetData')[0];
     const prior=targetRows.find(row=>Number(row.getAttribute('r'))===insertAfter);
     const baseCell=col=>Array.from(prior?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+insertAfter);
     // Dedicated fully bordered, wrapped Excel styles. Using the motor-row
     // styles unchanged previously produced missing vertical lines and clipping.
     const stylesDoc=parse(await zip.file('xl/styles.xml').async('string'));
     const styleXfs=all(stylesDoc,'cellXfs')[0],styleBorders=all(stylesDoc,'borders')[0];
     if(!styleXfs||!styleBorders)throw Error('Ride/Noise: no stylesheet');
     const makeRideStyle=(baseId)=>{
       const base=all(styleXfs,'xf')[Number(baseId)]||all(styleXfs,'xf')[0];
       const xf=base.cloneNode(true);
       const border=stylesDoc.createElementNS(n,'border');
       for(const sideName of ['left','right','top','bottom']){
         const side=stylesDoc.createElementNS(n,sideName);
         side.setAttribute('style','thin');
         const color=stylesDoc.createElementNS(n,'color');color.setAttribute('rgb','FF777777');
         side.appendChild(color);border.appendChild(side);
       }
       const bid=all(styleBorders,'border').length;
       styleBorders.appendChild(border);styleBorders.setAttribute('count',String(bid+1));
       xf.setAttribute('borderId',String(bid));xf.setAttribute('applyBorder','1');
       let alignment=Array.from(xf.children).find(el=>el.localName==='alignment');
       if(!alignment){alignment=stylesDoc.createElementNS(n,'alignment');xf.appendChild(alignment);}
       alignment.setAttribute('wrapText','1');alignment.setAttribute('vertical','center');
       xf.setAttribute('applyAlignment','1');
       const id=all(styleXfs,'xf').length;styleXfs.appendChild(xf);
       styleXfs.setAttribute('count',String(id+1));return String(id);
     };
     const styleA=makeRideStyle(baseCell('A')?.getAttribute('s')||'0');
     const styleB=makeRideStyle(baseCell('B')?.getAttribute('s')||'0');
     const styleC=makeRideStyle(baseCell('C')?.getAttribute('s')||'0');
     zip.file('xl/styles.xml',serial(stylesDoc));
     const refRow=Array.from(rowsRoot.children).find(el=>el.localName==='row'&&Number(el.getAttribute('r'))>insertAfter);
     const addLine=(rn,group,item,value)=>{
       const row=x.createElementNS(n,'row');row.setAttribute('r',String(rn));
       // A short row for one-line text; taller for wrapped content in B/C.
       // Est. column B ~30 chars; C ~55 chars, 16 pt per visible text line.
       const lineCount=Math.max(
         1,Math.ceil(item.length/29),Math.ceil(value.length/49),
         ...[item,value].map(t=>t.split('\n').length)
       );
       const height=Math.max(27,19*lineCount+8);
       row.setAttribute('ht',String(height));row.setAttribute('customHeight','1');
       for(const [col,content,style] of [['A',group,styleA],['B',item,styleB],['C',value,styleC]]){
         const cell=x.createElementNS(n,'c');cell.setAttribute('r',col+rn);cell.setAttribute('s',style);
         updateCell(x,cell,content);row.appendChild(cell);
       }
       rowsRoot.insertBefore(row,refRow||null);
     };
     let rn=insertAfter+1;
     const rideStart=rn;
     for(const [i,[label,value]] of rideRows.entries())addLine(rn++,i===0?'ĐỘ ÊM VẬN HÀNH':'',label,value);
     const rideEnd=rn-1,noiseStart=rn;
     for(const [i,[label,value]] of noiseRows.entries())addLine(rn++,i===0?'ĐỘ ỒN':'',label,value);
     const noiseEnd=rn-1;
     let merges=all(x,'mergeCells')[0];
     if(!merges){merges=x.createElementNS(n,'mergeCells');const before=Array.from(x.documentElement.children).find(el=>['conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup'].includes(el.localName));if(before)x.documentElement.insertBefore(merges,before);else x.documentElement.appendChild(merges);}
     for(const range of [`A${rideStart}:A${rideEnd}`,`A${noiseStart}:A${noiseEnd}`]){
       const merge=x.createElementNS(n,'mergeCell');merge.setAttribute('ref',range);merges.appendChild(merge);
     }
     merges.setAttribute('count',String(all(merges,'mergeCell').length));
   }
   zip.file(finalpath,serial(x));zip.file('xl/workbook.xml',serial(workbook));zip.file('xl/_rels/workbook.xml.rels',serial(rels));
   // remove unused worksheets and their sheet relationships; keep style and drawing parts.
   for(const old of sheetnodes.filter(s=>s!==sheet)){
     const rid=old.getAttributeNS(r,'id')||old.getAttribute('r:id');
     const rel=relationships.find(x=>x.getAttribute('Id')===rid);
     if(rel){let target=rel.getAttribute('Target')||'';const path=target.startsWith('/xl/')?target.slice(1):'xl/'+target.replace(/^\//,'');zip.remove(path);}
   }
   zip.remove('xl/calcChain.xml');
   const ctPath='[Content_Types].xml',types=parse(await zip.file(ctPath).async('string'));
   for(const item of Array.from(types.documentElement.children)){
     const part=item.getAttribute('PartName')||'';
     if(part.endsWith('/calcChain.xml')||/^\/xl\/worksheets\/sheet(?:1|3)\.xml$/.test(part))item.remove();
   }
   zip.file(ctPath,serial(types));
   // Refuse to download a workbook if the OOXML package contains the same
   // structural faults that triggered Microsoft Excel Repair in v5.8.123.
   // This final pass runs AFTER ZIP serialization input has been fully edited.
   await verifyExportPackage(zip,finalpath);
   return {zip,freezeCount,hiddenRows:doc.hiddenRows.size,filename:doc.model};
 }
 // v5.8.127: Fail closed on malformed OOXML. Do not silently generate a
 // workbook that Microsoft Excel will need to repair by removing cell content.
 async function verifyExportPackage(zip,worksheetPath){
   const read=async path=>{const f=zip.file(path);if(!f)throw Error('XLSX integrity: missing '+path);
     const doc=parse(await f.async('string'));
     if(doc.getElementsByTagName('parsererror').length)throw Error('XLSX integrity: malformed '+path);
     return doc;
   };
   const [ws,style]=await Promise.all([read(worksheetPath),read('xl/styles.xml')]);
   const styleItems=type=>all(style,type)[0];
   const collect=(parent,type)=>parent?Array.from(parent.children).filter(e=>e.localName===type):[];
   const categories={fonts:'font',fills:'fill',borders:'border',cellStyleXfs:'xf',cellXfs:'xf',cellStyles:'cellStyle',numFmts:'numFmt',dxfs:'dxf'};
   for(const [name,childName] of Object.entries(categories)){
     const el=styleItems(name);if(!el)continue;
     const records=collect(el,childName);
     if(el.hasAttribute('count')&&Number(el.getAttribute('count'))!==records.length)
       throw Error('XLSX integrity: '+name+' count '+el.getAttribute('count')+' != '+records.length);
   }
   const borders=collect(styleItems('borders'),'border');
   const cellXfs=collect(styleItems('cellXfs'),'xf');
   const fonts=collect(styleItems('fonts'),'font');
   const fills=collect(styleItems('fills'),'fill');
   const bOrder=['left','right','top','bottom','diagonal','vertical','horizontal','start','end'];
   for(const [i,border] of borders.entries()){
     const names=Array.from(border.children).map(e=>e.localName);
     const sorted=[...names].sort((a,b)=>bOrder.indexOf(a)-bOrder.indexOf(b));
     if(names.some((value,j)=>value!==sorted[j])||new Set(names).size!==names.length)
       throw Error('XLSX integrity: invalid border '+i);
   }
   for(const [i,xf] of cellXfs.entries()){
     for(const [prop,limit] of [['fontId',fonts.length],['fillId',fills.length],['borderId',borders.length]]){
       const val=Number(xf.getAttribute(prop)||0);
       if(!Number.isInteger(val)||val<0||val>=limit)
         throw Error('XLSX integrity: '+prop+' of style '+i+' invalid');
     }
   }
   const toCol=letters=>[...letters].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0);
   let lastRow=0;const occupied=new Set();
   const data=all(ws,'sheetData')[0];if(!data)throw Error('XLSX integrity: missing sheetData');
   for(const row of collect(data,'row')){
     const r=Number(row.getAttribute('r'));
     if(!Number.isInteger(r)||r<=lastRow)throw Error('XLSX integrity: unordered row '+r);
     lastRow=r;let lastColumn=0;
     for(const cell of collect(row,'c')){
       const address=cell.getAttribute('r')||'';
       const m=/^([A-Z]+)([1-9][0-9]*)$/.exec(address);
       if(!m||Number(m[2])!==r)throw Error('XLSX integrity: misplaced cell '+address);
       const col=toCol(m[1]);if(col<=lastColumn)throw Error('XLSX integrity: duplicate/unsorted '+address);
       lastColumn=col;
       const id=Number(cell.getAttribute('s')||0);
       if(!Number.isInteger(id)||id<0||id>=cellXfs.length)
         throw Error('XLSX integrity: invalid style in '+address);
       if(cell.getAttribute('t')==='inlineStr' && collect(cell,'is').length!==1)
         throw Error('XLSX integrity: missing inline string '+address);
       if(cell.getAttribute('t')==='inlineStr' && collect(cell,'v').length)
         throw Error('XLSX integrity: unexpected numeric value '+address);
     }
   }
   const mroot=all(ws,'mergeCells')[0];
   for(const merge of collect(mroot,'mergeCell')){
     const ref=merge.getAttribute('ref')||'';
     const m=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
     if(!m)throw Error('XLSX integrity: invalid merge '+ref);
     const c1=toCol(m[1]),c2=toCol(m[3]),r1=Number(m[2]),r2=Number(m[4]);
     if(c1>c2||r1>r2)throw Error('XLSX integrity: reversed merge '+ref);
     for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++){
       const key=r+':'+c;if(occupied.has(key))throw Error('XLSX integrity: overlapping merge '+ref);
       occupied.add(key);
     }
   }
 }
 // Decompose bilingual specification descriptions into literal Vietnamese and
 // English fragments. Do not interpret short engineering codes as English.
 function bilingualRuns(source){
   const raw=String(source||'');
   const re=/\((?:[A-Za-z][^()\r\n]*|[A-Za-z][^()\r\n]*\([^()]*\)[^()\r\n]*)\)/g;
   const out=[];let last=0,m;
   const englishPhrase=part=>{
     const inside=part.slice(1,-1).trim();
     return inside.length>=9 && /(?:\s|[-–])/.test(inside) &&
       /\b(?:stainless|steel|hairline|provided|not supplied|passenger|hall|indicator|button|model|door|ceiling|car|floor|glass|painted|plastic|quantity|operation|landing|opening|weight|lighting|granite|emergency|round|safety|front|rear|with|without|side|frame|motor|power|finish|material|supply|sensor|travelling|traction|shaft|handrail|mirror|ventilation|fire|speed|design|structure|control|inspection|contact|automatic|type)\b/i.test(inside);
   };
   while((m=re.exec(raw))){
     const segment=m[0];
     if(!englishPhrase(segment))continue;
     const prefix=raw.slice(last,m.index);
     if(prefix)out.push({text:prefix,english:false});
     const earlier=raw.slice(0,m.index);
     const currentLine=earlier.slice(earlier.lastIndexOf('\n')+1);
     // Newline if English text follows Vietnamese on the same physical line.
     if(/[À-ỹ]/.test(currentLine)&&currentLine.trim() && !/\n$/.test(prefix))out.push({text:'\n',english:false});
     out.push({text:segment,english:true});last=m.index+segment.length;
   }
   if(last<raw.length)out.push({text:raw.slice(last),english:false});
   return out.length?out:[{text:raw,english:false}];
 }
 function printHtml(modelData,model){
   const v=(addr)=>text(modelData.cells[addr]);
   const escape=s=>text(s).replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
   const highest=Math.max(...Object.keys(modelData.cells).map(cellRow));let rows=[];
   for(let i=1;i<=highest;i++){
     if(modelData.hiddenRows.has(i))continue;
     let a=v('A'+i),b=v('B'+i),c=v('C'+i);
     if(i>300)continue; // post-spec helper cells, never printed by the original export
     if(!a&&!b&&!c)continue;
     let section=/^\s*(I{1,4}|V|VI)\.\s/.test(a)||/các tính năng.*Features|ĐẶC TÍNH KỸ THUẬT/i.test(a);
     let sub=/^\s*[-–]\s*Tính năng/.test(a);
     let feature=!!(/\(.*\)/.test(a)&&i>82&&c);
     let rowClass=section?'section':sub?'subsection':feature?'feature':'';
     const bilingualHtml=(v)=>bilingualRuns(v).map(part=>part.english?'<em>'+escape(part.text)+'</em>':escape(part.text)).join('');
     rows.push('<tr class="'+rowClass+'"><td>'+escape(a)+'</td><td>'+escape(b)+'</td><td>'+bilingualHtml(c)+'</td></tr>');
   }
   let title=model==='MR'?'NEXIEZ-MR':'NEXIEZ-MRL Version2';
   return `<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>SPEC ${title} - A4 PDF</title><style>
   @page{size:A4 portrait;margin:12mm 10mm 13mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;font-size:9pt;color:#1b1b1b;margin:0;padding:0}header{margin:0 0 12px;border-bottom:2px solid #b5161c;padding:7px 0 10px}header strong{font-size:13pt;color:#10345c}header small{display:block;margin-top:3px}table{border-collapse:collapse;width:100%;table-layout:fixed}td{border:1px solid #777;padding:4px 5px;vertical-align:top;white-space:pre-line;overflow-wrap:anywhere;line-height:1.33}td:nth-child(1){width:31%;font-weight:600}td:nth-child(2){width:21%}td:nth-child(3){width:48%}tr.section td{font-weight:700;background:#e9edf2;color:#132e4c;break-after:avoid}tr.subsection td{font-weight:700;background:#f3f4f5;break-after:avoid}tr.feature td:first-child{font-weight:500}tr{break-inside:avoid}.toolbar{position:sticky;top:0;padding:12px;background:#f7f8fa;border-bottom:1px solid #aaa;margin-bottom:12px;font-family:Arial,sans-serif}.toolbar button{padding:9px 16px;background:#ad1119;color:white;border:0;border-radius:5px;cursor:pointer}.toolbar small{margin-left:12px;color:#555}footer{color:#666;font-size:8pt;padding-top:8px} @media print{.toolbar{display:none}header{margin:0 0 10px}footer{display:none}} 
   </style></head><body><div class="toolbar"><button onclick="window.print()">In / Save as A4 PDF</button><small>Chọn máy in “Save as PDF” trong hộp thoại / Select Save as PDF.</small></div><header><strong>MITSUBISHI ELECTRIC VIETNAM — SPECIFICATION</strong><small>${escape(title)} | ${escape(v('B8'))} — ${escape(v('B9'))}</small></header><table><tbody>${rows.join('')}</tbody></table><footer>Source: SpecFormExport | A4 portrait | Excel Database</footer></body></html>`;
 }
 window.SPEC_EXPORT={exportModel,xlsxFromModel,printHtml,attachMRMaterialThickness};
})();
