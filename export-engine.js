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
   // v5.8.221: Correct the repeated IV heading in the MR factory template.
   // Features follow IV. Entrance design and therefore belong to section V.
   if(model==='MR'&&/^\s*IV\.\s*Các tính năng\b/i.test(text(cells.A84)))
     cells.A84='V. Các tính năng (Features) :';
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
       if(/^(?:MBS|Multi\s*-\s*Beam Door Sensor)$/i.test(code)){
         for(const er of featureRows)if(/^Multi\s*-\s*Beam Door Sensor$/i.test(text(cells['A'+er]).trim()))cells['A'+er]='(MBS) Multi - Beam Door Sensor';
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
   if(window.SPEC_CEILING_EXIT.sized(model,ceilingName)){
     cells.C64='Trên trần phòng thang kích thước 500x700mm (Provided on the ceiling 500x700mm)';
   }
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
   if(model==='MR' && window.SPEC_CEILING_EXIT.c30Selected(changes)){
     const heading=nameAddr(db,'SpecFormExport','K37_Title_EmergencyFeatures');
     if(heading)hiddenRows.delete(cellRow(heading));
     hiddenRows.delete(93);hiddenRows.delete(94);
     cells.D93='FE-C30';cells.D94='FE-C30';
     cells.A93='(FE-C30) FE-Hongkong with MELCO switch';
     cells.C93=window.SPEC_CEILING_EXIT.en;
     cells.A94='Vận hành khẩn cấp chữa cháy FE Hongkong (FE-C30)';
     cells.C94=window.SPEC_CEILING_EXIT.vi;
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
   // v5.8.239: preserve at least 1 decimal for the m/s display.
   // Match speed by VALUE rather than a fixed row index; final Excel row
   // positions differ from the source depending on split/inserted rows.
   {
     for(const [cell,rawValue] of Object.entries(cells)){
       if(!/^C\d+$/.test(cell))continue;
       const raw=String(rawValue??'');
       const match=/(\d+(?:[.,]\d+)?)\s*m\/phút/i.exec(raw);
       if(!match||!/(?:m\/giây|m\/s)/i.test(raw))continue;
       const mpm=Number(match[1].replace(',','.'));
       if(!(Number.isFinite(mpm)&&mpm>0))continue;
       const speed=(mpm/60).toFixed(2).replace(/0$/,'');
       const amount=String(mpm);
       cells[cell]=`${amount} m/phút - ${speed} m/giây (${amount} mpm - ${speed} m/s)`;
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
   // v5.8.218: MBSS was removed from BOTH product models. The factory
   // workbook still contains incompatibility footnotes mentioning MBSS;
   // remove only that obsolete part and retain any valid SR/MBS warning.
   function removeDiscontinuedMBSSNotes(source){
     return String(source).replace(/\(([^()]*\bMBSS\b[^()]*)\)/gi,(original,body,offset,sourceText)=>{
       if(!/(?:cannot be combined|không thể sử dụng|not together)/i.test(body))return original;
       const other=body.match(/(?:\bSR(?:-?[12])?\b|\bMBS(?:3)?\b|Multi\s*-?\s*Beam\s+Door\s+Sensor|Muti\s*-?\s*Beam\s+Door\s+Sensor)/gi)||[];
       if(!other.length)return '';
       let cleaned=body
         .replace(/\s*(?:,\s*|\b(?:or|hay|hoặc|and)\s+)MBSS(?:\s+(?:feature|chức năng))?/gi,'')
         .replace(/\bMBSS(?:\s+(?:feature|chức năng))?/gi,'')
         .replace(/\s+,/g,',').replace(/,\s*\./g,'.')
         .replace(/\s+(?:or|hay|hoặc|and)\s*[.,;]*\s*$/i,'').replace(/[\s,;.]+$/,'').trim();
       const outsidePeriod=/^\s*\./.test(sourceText.slice(offset+original.length));
       const insidePeriod=/\.\s*$/.test(body);
       return cleaned ? '('+cleaned+(insidePeriod&&!outsidePeriod?'.':'')+')' : '';
     }).replace(/\s+\.(?=\s|$)/g,'').replace(/[ \t]{2,}/g,' ').replace(/[ \t]+\n/g,'\n');
   }
   for(const addr of Object.keys(cells)){
     if(typeof cells[addr]!=='string')continue;
     cells[addr]=removeDiscontinuedMBSSNotes(cells[addr])
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
   // v5.8.213: Export each dimension in three INDEPENDENT columns:
   // A: Vietnamese item + English item; B: Vietnamese axes + English axes;
   // C: actual numeric result with mm. Do not merge A:B for these four rows.
   // Factory MR and MRL Version02 have different source row positions.
   const dimensionRows=model==='MR'
     ? [[37,'Kích thước phòng thang','Car internal size','R × S × C','W × D × H',3],
        [38,'Kích thước cửa','Door size','R × C','W × H',2],
        [39,'Kích thước hố thang','Hoistway size','R × S','W × D',2],
        [40,'Kích thước phòng máy','Machine room size','R × S × C','W × D × H',3]]
     : [[36,'Kích thước phòng thang','Car internal size','R × S × C','W × D × H',3],
        [37,'Kích thước cửa','Door size','R × C','W × H',2],
        [38,'Kích thước hố thang','Hoistway size','R × S','W × D',2],
        [39,'Kích thước phòng máy','Machine room size','R × S × C','W × D × H',3]];
   for(const [row,vi,en,axesVi,axesEn,axes] of dimensionRows){
     const label='A'+row,value='C'+row;
     if(!text(cells[label]).trim())continue;
     cells[label]=vi+'\n'+en;
     cells['B'+row]=axesVi+'\n'+axesEn;
     const current=text(cells[value]).trim();
     // Keep 'Machine roomless' unchanged; do not invent any machine room size.
     const clean=current.replace(/\s*\b(?:mm|millimeters?)\b\s*$/i,'').trim();
     const numeric=clean.match(/\d{2,5}(?:[.,]\d{1,2})?/g)||[];
     if(numeric.length===axes && !/[\r\n]/.test(clean) &&
        /[x×*]/i.test(clean) &&
        !/\b(?:not supplied|không cung cấp|machine roomless)\b/i.test(clean)){
       cells[value]=numeric.join('×')+' mm';
     }
   }
   // Remove legacy MR header/version metadata in the FINAL XLSX export.
   if(model==='MR'){
     cells.A1='MITSUBISHI ELECTRIC VIETNAM CO., LTD.';
     cells.C2=''; // Ver.: K0 (06.03.2018)
     // The factory MR evaluator can sometimes omit the final 'd'. Keep the
     // canonical bilingual origin in the customer-facing export in all cases.
     cells.C12='Thái Lan\nThailand';
   }
   cells.C1=''; // Factory form SE-F16.02 is obsolete on both models.
   cells['C'+(model==='MR'?11:12)]='MITSUBISHI ELECTRIC';
   // v5.8.217: MRL customer-facing wording must use the complete
   // factory DoorTypeVN lookup, not a bare CO/2S code. Keep the model
   // SpecForm input unchanged: this is EXPORT-PRESENTATION only.
   if(model==='MRL'){
     cells.C13=text(cells.C13).replace(/\bThailan\b/gi,'Thailand');
     const doorType=text(eng.safe('SpecForm','D15').value).trim().toUpperCase();
     const doorWording={
       CO:'CO - Hai cánh mở ra hai bên từ trung tâm (2-panel center opening doors)',
       '2S':'2S - Hai cánh mở về một phía (2-panel side sliding doors)'
     };
     if(doorWording[doorType])cells.C18=doorWording[doorType];
   }
   // v5.8.243: Export the customer's bilingual Control system description and
   // explain actual MRL group-operation codes without changing any SpecForm values.
   // Both XLSX and PDF are built from these same frozen customer-facing cells.
   // v5.8.244: Use SOURCE SpecFormExport coordinates, not final display
   // coordinates. Both models insert rail rows before rendering; raw C35 is
   // blank (MR) or a section heading (MRL). Writing VVVF to raw C35 used to
   // leak text into final C37/C38 and inflate MRL's merged section heading.
   const controlSourceAddress = model==='MR'?'C33':'C32';
   cells[controlSourceAddress]='Động cơ máy kéo và động cơ cửa cabin được điều khiển bằng công nghệ biến đổi điện áp và tần số (VVVF), kết hợp với hệ thống điều khiển vi xử lý.\n'+
     'The traction machine and cabin door motor are controlled by Variable Voltage Variable Frequency (VVVF) technology, in combination with a microprocessor-based control system';
   if(model==='MRL'){
     const code=text(cells.C25).trim().toUpperCase();
     const match=/^([1-4])C(?:[-/].+)?$/.exec(code);
     if(match){
       const types={
         1:['Điều khiển đơn','Simplex control'],
         2:['Điều khiển đôi','Duplex control'],
         3:['Điều khiển nhóm 3 thang','3-elevator group control'],
         4:['Điều khiển nhóm 4 thang','4-elevator group control']
       };
       const [vi,en]=types[Number(match[1])];
       // Three physical text lines in ONE cell; the English run is blue italic.
       cells.C25=code+'\n'+vi+'\n'+en;
     }
   }
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
   // The glass specification follows the selected door at each location.
   const glazedRows=new Set();
   const doors=mrl?[['D46',49],['E73',72],['E74',74]]:[['D50',50],['E76',72],['E77',74]];
   for(const [input,row] of doors){
     const option=read(input);
     const dsj=/\bDSJ1\b/i.test(option);
     if(!dsj&&!/large\s+window/i.test(option))continue;
     const name=dsj?'DSJ1':'Large window';
     const openingSize=!dsj&&/W?(\d+)\s*[x×]\s*H?(\d+)/i.exec(option);
     const size=doc.includeTechDetails&&openingSize?' '+openingSize[1]+'×'+openingSize[2]+' mm':'';
     const glass=dsj?'8 mm, 4+4 mm':'10 mm, 5+5 mm';
     const vi=text(doc.cells['C'+row]).split(/,\s*(?:kính quan sát|cửa sổ quan sát)|\s+với (?:kính|cửa sổ) quan sát/i)[0];
     const en=text(doc.cells['C'+(row+1)]).replace(/^\(/,'').split(/,\s*(?:see-through|large observation window)|\s+with (?:see-through|large observation window)/i)[0];
     doc.cells['C'+row]=vi+', '+name+size+(doc.includeTechDetails?', kính trong '+glass:'')+'.';
     doc.cells['C'+(row+1)]='('+en+', '+name+size+(doc.includeTechDetails?', clear laminated glass '+glass:'')+'.)';
     glazedRows.add(row);
   }
   // Panel/model descriptions in Section IV: 1.5 mm, only in detailed Export.
   // MR: separate PIV/HBV/PIH rows; MRL: consolidated 3-floor blocks.
   // Keep the standard Export descriptions unchanged when detail is disabled.
   if(!doc.includeTechDetails)return;
   // v5.8.236: Add verified handrail cross-section dimensions only in
   // optional detailed Export, not in the SpecForm/weight-calculation inputs.
   // MR YH-59S/YH-59M: Ø38 mm; MRL YH-100SL: Ø32 mm;
   // YH-57S in both products: 75 mm wide. Always output bilingual wording.
   {
     const rail=read(mrl?'D64':'D67');
     const position=read(mrl?'D65':'D68');
     const railModel=/^(YH-59S|YH-59M|YH-100SL|YH-57S)(?=\s|\(|$)/i.exec(rail)?.[1]?.toUpperCase();
     const railDim=(mrl?{'YH-100SL':['tròn','round','đường kính Ø32 mm','Ø32 mm diameter'],
                         'YH-57S':['dẹt','flat','bản rộng 75 mm','75 mm wide']}:
                        {'YH-59S':['tròn','round','đường kính Ø38 mm','Ø38 mm diameter'],
                         'YH-59M':['tròn','round','đường kính Ø38 mm','Ø38 mm diameter'],
                         'YH-57S':['dẹt','flat','bản rộng 75 mm','75 mm wide']})[railModel];
     if(railDim){
       const [typeVi,typeEn,dimensionVi,dimensionEn]=railDim;
       const locationVi={
         '1 pc on rear wall':'1 tay vịn trên vách sau',
         '1 pc on right side wall':'1 tay vịn trên vách phải',
         '1 pc on left side wall':'1 tay vịn trên vách trái',
         '2 pcs on both side walls':'2 tay vịn trên hai vách bên',
         '2 pcs on rear & right side wall':'2 tay vịn trên vách sau và vách phải',
         '2 pcs on rear & left side wall':'2 tay vịn trên vách sau và vách trái',
         '3 pcs on 3 walls':'3 tay vịn trên ba vách',
         'No':''
       }[position]||position;
       const locationEn=position==='No'?'':position;
       // Separate full sentences: existing Excel/PDF bilingual processors turn
       // the parenthesized English translation into a blue italic line.
       const vi='Tay vịn '+typeVi+' inox sọc nhuyễn SUS304, '+dimensionVi+', mã hiệu '+railModel+
         (locationVi?', '+locationVi:'')+'.';
       const en=typeEn.charAt(0).toUpperCase()+typeEn.slice(1)+
         ' handrail in hairline-finish stainless steel SUS304, '+dimensionEn+', model '+railModel+
         (locationEn?', '+locationEn:'')+'.';
       doc.cells.C62=vi+' ('+en+')';
     }
   }
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
     const mat=material(raw),isDoor=fn===carDoorThickness||fn===entranceThickness;
     const windowPanel=isDoor&&(glazedRows.has(viRow)||/\bDS1\b/i.test(raw));
     const thickness=mat?(windowPanel?1.5:fn(mat)):null;
     doc.materialThicknessResults.push({name,material:mat,thickness,k});
     const unverified=name==='Car wall'&&mat==='SUS-HL'&&thickness===null&&Number.isFinite(k);
     const unverifiedLam=name==='Car wall'&&mrl&&mat==='LAMINATED'&&thickness===null&&Number.isFinite(k);
     if(!Number.isFinite(thickness)&&!(unverified||unverifiedLam))continue;
     if(glazedRows.has(viRow)){
       doc.cells['C'+viRow]=doc.cells['C'+viRow].replace(/\.$/,'')+'. Tấm cửa dày '+thickness.toFixed(1)+' mm.';
       doc.cells['C'+enRow]=doc.cells['C'+enRow].replace(/\.\)$/,'')+'. Door panel thickness '+thickness.toFixed(1)+' mm.)';
       continue;
     }
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
   // v5.8.237: Optional detailed Export only — CBV/PIV/HBV button diameter.
   // Keep the measurement in the SAME Vietnamese/English sentence (no new
   // spreadsheet rows); this preserves bilingual styling, merges and row fit.
   // Target actual panel model codes rather than fixed C59/C60 etc., because
   // MR and MRL have different row layouts and MRL combines hall panels.
   const panelModel=/(?:^|\b)(?:CBV|PIV|HBV)\d*\s*[-–]?\s*[A-Z]\d+[A-Z0-9-]*/i;
   const buttonVi='đường kính nút bấm Ø33 mm';
   const buttonEn='button diameter Ø33 mm';
   function addButtonDimension(segment){
     const line=String(segment);
     if(!panelModel.test(line)||/\b(?:Not supplied|Không cung cấp)\b/i.test(line)||
       /(?:đường kính nút bấm|button diameter)\s*Ø\s*33\s*mm/i.test(line))return line;
     const trimmed=line.trim();
     // Mixed line: VI (English). Safely append to BOTH language clauses.
     const bilingual=/^([\s\S]*?)\s+\(([^()]*)\)([.!]?\s*)$/.exec(line);
     if(bilingual&&panelModel.test(bilingual[1])&&panelModel.test(bilingual[2])){
       const vi=bilingual[1].replace(/[.;,]\s*$/,'').trimEnd();
       const en=bilingual[2].replace(/[.;,]\s*$/,'').trimEnd();
       return vi+', '+buttonVi+' ('+en+', '+buttonEn+')'+(bilingual[3]||'');
     }
     // On separate bilingual lines, English is usually parenthesized.
     const wrappedEn=/^([\s]*\()([\s\S]*?)(\)\.?[\s]*)$/.exec(line);
     if(wrappedEn){
       const trailing=wrappedEn[3].includes('.')?').':')';
       return wrappedEn[1]+wrappedEn[2].replace(/[.;,]\s*$/,'')+', '+buttonEn+trailing;
     }
     const english=/^\s*(?:Hall |Car |Model |Operating |Wheelchair )/i.test(line)&&
       !/[À-ỹ]/.test(line);
     const ending=/([.;]\s*)$/.exec(line);
     return line.replace(/[.;]\s*$/,'')+', '+(english?buttonEn:buttonVi)+(ending?ending[1]: '');
   }
   for(const addr of Object.keys(doc.cells)){
     if(!/^C\d+$/.test(addr)||typeof doc.cells[addr]!=='string')continue;
     const source=doc.cells[addr];
     if(!panelModel.test(source))continue;
     doc.cells[addr]=source.split('\n').map(addButtonDimension).join('\n');
   }
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
 // Project the final, styled sheet so Form đơn, PDF and Spec Summary share the same language rules.
 async function vietnameseSheet(zip,x,lastDataColumn){
   const styles=parse(await zip.file('xl/styles.xml').async('string'));
   const fonts=all(child(styles.documentElement,'fonts'),'font');
   const xfs=all(child(styles.documentElement,'cellXfs'),'xf');
   const ss=zip.file('xl/sharedStrings.xml');
   const strings=ss?all(parse(await ss.async('string')),'si'):[];
   // Match Vietnamese vowels and tone marks. Ø (diameter) and × are engineering symbols.
   const vn=t=>/[đĐ]|[aeiouy][\u0300\u0301\u0302\u0303\u0306\u0309\u031b\u0323]/i.test(t.normalize('NFD'));
   const detailTranslation=t=>!vn(t)&&/\b(?:display floor and direction|hall indicators?|handrail|button diameter|thickness|hairline faceplate)\b/i.test(t);
   const technical=t=>/^\d+\s*V\s+(?:Single|Multi)\s*[-–]\s*\d+\s*Master Stations?$/i.test(t.trim())||/^Mitsubishi\s+Standard$/i.test(t.trim())||/^(?:[A-Z0-9][A-Z0-9./×+– -]*|\d+[\d., ×x+-]*(?:mm|kg|m\/s|m²)?)$/.test(t.trim());
   const local=t=>t.replace(/8th Floor – Tower B – Viettel Tower, 285 Cach Mang Thang 8 Street, Hoa Hung Ward, HCMC/g,'Tầng 8 – Tháp B – Tòa nhà Viettel, 285 Cách Mạng Tháng 8, phường Hòa Hưng, TP. Hồ Chí Minh').replace(/\b(\d+)\s*V\s+(Single|Multi)\s*[–-]\s*(\d+)\s*Master Stations?\b/gi,(_,v,k,n)=>v+'V '+(k.toLowerCase()==='single'?'Single':'Multi')+' - '+n+' Master station').replace(/(\d+)\s*Lines\b/gi,'$1 đường').replace(/\(Main specifications\)/ig,'').trim();
   const ranges=new Map(all(x,'mergeCell').map(el=>{const [a,b]=el.getAttribute('ref').split(':');return [a,b]}));
   const widths=new Map();
   for(const col of all(x,'col'))for(let i=+col.getAttribute('min');i<=Math.min(lastDataColumn,+col.getAttribute('max'));i++)widths.set(i,+col.getAttribute('width'));
   const richText=node=>all(node,'t').map(t=>t.textContent).join('');
   const borderTable=child(styles.documentElement,'borders'),xfTable=child(styles.documentElement,'cellXfs');
   const borders=all(borderTable,'border'),edgeCache=new Map(),styleCache=new Map();
   // In bilingual rows the upper cell has the top edge and the lower cell has the bottom edge.
   function carryBottom(upper,lower){
     const original=+upper.getAttribute('s')||0,base=xfs[original];
     const source=xfs[+lower.getAttribute('s')||0];
     const sourceId=+source.getAttribute('borderId')||0,baseId=+base.getAttribute('borderId')||0;
     const edge=child(borders[sourceId],'bottom');
     if(!edge?.getAttribute('style')||baseId===sourceId)return;
     const styleKey=original+':'+sourceId;
     if(!styleCache.has(styleKey)){
       const edgeKey=baseId+':'+sourceId;
       if(!edgeCache.has(edgeKey)){
         const border=borders[baseId].cloneNode(true),old=child(border,'bottom');
         if(old)border.replaceChild(edge.cloneNode(true),old);
         else border.insertBefore(edge.cloneNode(true),Array.from(border.children).find(e=>['diagonal','vertical','horizontal'].includes(e.localName))||null);
         edgeCache.set(edgeKey,borders.length);borders.push(border);borderTable.appendChild(border);
       }
       const xf=base.cloneNode(true);xf.setAttribute('borderId',String(edgeCache.get(edgeKey)));xf.setAttribute('applyBorder','1');
       styleCache.set(styleKey,xfs.length);xfs.push(xf);xfTable.appendChild(xf);
     }
     upper.setAttribute('s',String(styleCache.get(styleKey)));
   }
   const removedRows=new Set();
   let previousCells=new Map();
   const sheetRows=all(x,'row');
   const featureCodes=new Map();
   const englishFeatureRows=new Set();
   function featureTitle(row){
     if(!row||row.getAttribute('hidden')==='1')return '';
     const label=Array.from(row.children).find(c=>c.getAttribute('r')==='A'+row.getAttribute('r'));
     if(!label)return '';
     const node=label.getAttribute('t')==='s'?strings[+child(label,'v')?.textContent]:child(label,'is');
     return node?richText(node).trim().split(/\r?\n/)[0].trim():'';
   }
   function identifier(title){
     const prefix=/^\(([A-Z][A-Za-z0-9 /-]*)\)(?:\s|$)|^([A-Z][A-Z0-9-]*)\s+[-–—]\s+/.exec(title);
     if(prefix)return prefix[1]||prefix[2];
     // These two legacy factory labels use the full feature name in the code column.
     if(/^Multi\s*-\s*Beam Door Sensor$/i.test(title))return 'MBS';
     if(/^Non-service Temporary Release for Hall Call\s*[-–]\s*Card Reader Type/i.test(title))return 'NSCR-H';
     return '';
   }
   let inFeatureSection=false;
   for(let i=0;i<sheetRows.length-1;i++){
     const first=sheetRows[i],second=sheetRows[i+1],a=featureTitle(first);
     if(/^(?:V|VI)\.\s/.test(a))inFeatureSection=true;
     if(!inFeatureSection||!a||/^[-–]|^(?:V|VI)\./.test(a))continue;
     const b=featureTitle(second);
     if(!b||/^[-–]|^(?:V|VI)\./.test(b)||+second.getAttribute('r')!==+first.getAttribute('r')+1)continue;
     const firstCode=identifier(a),secondCode=identifier(b);
     const viFirst=vn(a)&&!vn(b)&&secondCode;
     const enFirst=!vn(a)&&vn(b)&&firstCode;
     if(!viFirst&&!enFirst)continue;
     featureCodes.set(+(viFirst?first:second).getAttribute('r'),'('+(viFirst?secondCode:firstCode)+')');
     englishFeatureRows.add(+(viFirst?second:first).getAttribute('r'));
     i++;
   }
   let inFeatures=false;
   for(const row of sheetRows){
     if(row.getAttribute('hidden')==='1')continue;
     const number=+row.getAttribute('r');if(number<2)continue;
     const cells=Array.from(row.children).filter(c=>c.localName==='c'&&/^[A-Z]\d+$/.test(c.getAttribute('r'))&&c.getAttribute('r').charCodeAt(0)-64<=lastDataColumn);
     let hadText=false,remaining=false,changed=false,height=16;
     for(const c of cells){
       const addr=c.getAttribute('r'),col=addr.charCodeAt(0)-64;
       const node=c.getAttribute('t')==='s'?strings[+child(c,'v')?.textContent]:child(c,'is');
       if(!node){if(child(c,'v'))remaining=true;continue}
       const original=richText(node);if(!original.trim())continue;
       hadText=true;
       if(col===1&&/^V\.\s/.test(original))inFeatures=true;
       const font=fonts[+(xfs[+c.getAttribute('s')||0]?.getAttribute('fontId')||0)];
       const italic=!!child(font,'i')&&child(font,'i').getAttribute('val')!=='0';
       const runs=Array.from(node.children).filter(r=>r.localName==='r');
       const protectedValue=number<=10&&col>1;
       let value=original;
       if(!protectedValue){
         if(runs.length)value=runs.map(run=>{
           const t=richText(run),props=child(run,'rPr');
           if(!t.trim())return t;
           if(child(props,'i')&&!vn(t)&&!/\/\d+\s*thang\b/i.test(t)&&!technical(t))return t.replace(/[^\n]/g,'');
           return t;
         }).join('');
         if(italic&&!vn(original)&&!technical(original)&&!/\/\d+\s*thang\b/i.test(original))value='';
         if(inFeatures&&/^\([^)]*\)\s+[A-Za-z]/.test(original)&&!vn(original))value='';
         if(col>=3&&vn(original)){
           // Detailed descriptions can mix both languages within one rich-text run.
           value=value.replace(/\(([^()\n]+)\)/g,(part,inside)=>detailTranslation(inside)?'':part)
             .split(/\r?\n/).filter(line=>!detailTranslation(line)).join('\n');
         }
         value=local(value).split(/\r?\n/).map(line=>line.trim()).filter(line=>line&&!/^Only apply for\b/i.test(line)).join('\n');
         if(vn(value))value=value.replace(/\s*\(([^()]*[a-z][^()]*)\)/gi,(match,inside)=>!vn(inside)&&/\b(?:Emergency exit|Basic specifications|Main specifications|Provided on|Not supplied|Control system)\b/i.test(inside)?'':match);
       }
       if(col===1&&featureCodes.has(number)&&vn(value)){
         const lines=value.split('\n'),code=featureCodes.get(number);
         if(!lines[0].includes(code))lines[0]=lines[0].trimEnd()+' '+code;
         value=lines.join('\n');
       }
       if(englishFeatureRows.has(number))value='';
       if(value!==original){
         const note=value.match(/\n(Chỉ áp dụng cho thang [^\n]+)$/);
         updateCell(x,c,value);
         if(note){
           const inline=child(c,'is');inline.removeChild(child(inline,'t'));
           for(const [part,red] of [[value.slice(0,note.index),false],[note[0],true]]){
             const run=x.createElementNS(n,'r'),props=x.createElementNS(n,'rPr'),t=x.createElementNS(n,'t');
             const name=x.createElementNS(n,'rFont');name.setAttribute('val','Arial');props.appendChild(name);
             if(red){props.appendChild(x.createElementNS(n,'b'));const color=x.createElementNS(n,'color');color.setAttribute('rgb','FFCC0000');props.appendChild(color)}
             run.appendChild(props);t.setAttribute('xml:space','preserve');t.textContent=part;run.appendChild(t);inline.appendChild(run);
           }
         }
         changed=true;
       }
       if(!value)continue;
       remaining=true;
       const end=(ranges.get(addr)||addr).charCodeAt(0)-64;
       let width=0;for(let k=col;k<=end;k++)width+=widths.get(k)||20;
       const size=+(child(font,'sz')?.getAttribute('val')||11);
       const max=Math.max(6,Math.floor((width-2)*10/size));
       let lines=0;
       for(const para of value.split('\n')){
         let used=0;lines++;
         for(const word of para.split(/\s+/)){
           if(used&&used+word.length+1>max){lines++;used=0}
           used+=word.length+(used?1:0);
           while(used>max){lines++;used-=max}
         }
       }
       height=Math.max(height,Math.ceil(lines*size*1.42+4));
     }
     if(hadText&&!remaining){
       removedRows.add(number);
       for(const lower of cells){
         const upper=previousCells.get(lower.getAttribute('r')[0]);
         if(upper)carryBottom(upper,lower);
       }
       row.setAttribute('hidden','1');row.setAttribute('ht','0');row.setAttribute('customHeight','1');
     }else{
       previousCells=new Map(cells.map(c=>[c.getAttribute('r')[0],c]));
       if(changed){row.setAttribute('ht',String(Math.min(409,height)));row.setAttribute('customHeight','1')}
     }
   }
   const merges=child(x.documentElement,'mergeCells');
   if(merges){
     for(const merge of all(merges,'mergeCell')){
       const [start,end]=merge.getAttribute('ref').split(':');
       const first=+start.match(/\d+$/)[0],last=+end.match(/\d+$/)[0];
       if(first===last||removedRows.has(first)||!removedRows.has(last))continue;
       let visible=last;
       while(visible>first&&removedRows.has(visible))visible--;
       const finalCell=end.replace(/\d+$/,String(visible));
       if(start===finalCell)merges.removeChild(merge);
       else merge.setAttribute('ref',start+':'+finalCell);
     }
     merges.setAttribute('count',String(all(merges,'mergeCell').length));
   }
   if(styleCache.size){
     borderTable.setAttribute('count',String(borders.length));xfTable.setAttribute('count',String(xfs.length));
     zip.file('xl/styles.xml',serial(styles));
   }
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
   // v5.8.209: Two-line merged size labels must not be clipped in Excel.
   const dimensionLabelRows=model==='MR'?[37,38,39,40]:[36,37,38];
   for(const rn of dimensionLabelRows){
     if(!/\nRộng \(W\)/.test(String(doc.cells['A'+rn]||'')))continue;
     const row=all(x,'row').find(el=>Number(el.getAttribute('r'))===rn);
     if(row){row.setAttribute('ht',String(Math.max(33,Number(row.getAttribute('ht'))||15)));
       row.setAttribute('customHeight','1');}
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
       if(part.english){
         prop.appendChild(x.createElementNS(n,'i'));
         const ink=x.createElementNS(n,'color');ink.setAttribute('rgb','FF245A86');prop.appendChild(ink);
       }
       rNode.appendChild(prop);
       const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');
       t.textContent=part.text;rNode.appendChild(t);is.appendChild(rNode);
     }
   }
   // A hall-call-box quantity can share a cell with the Vietnamese description.
   // Color/italicize only the quantity line, leaving the Vietnamese line black.
   for(const c of all(x,'c')){
     if(!/^C\d+$/.test(c.getAttribute('r')||'')||c.getAttribute('t')!=='inlineStr')continue;
     const textRuns=all(c,'t').map(t=>t.textContent||'').join('');
     if(!/(?:^|\n)\s*\d+\s*hộp\s*\//i.test(textRuns))continue;
     const is=child(c,'is');if(!is)continue;
     while(is.firstChild)is.removeChild(is.firstChild);
     for(const part of textRuns.split(/(\n)/)){
       const isQty=/^\s*\d+\s*hộp\s*\//i.test(part);
       const rNode=x.createElementNS(n,'r'),pr=x.createElementNS(n,'rPr');
       const family=x.createElementNS(n,'rFont');family.setAttribute('val','Arial');pr.appendChild(family);
       if(isQty){pr.appendChild(x.createElementNS(n,'i'));const ink=x.createElementNS(n,'color');ink.setAttribute('rgb','FF245A86');pr.appendChild(ink);}
       rNode.appendChild(pr);
       const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');t.textContent=part;rNode.appendChild(t);is.appendChild(rNode);
     }
   }
   // v5.8.119: split single-row bilingual descriptions into TWO Excel rows.
   // Feature descriptions already encoded as two source rows remain untouched.
   // Preserve the worksheet's original outside borders; the only new internal
   // NO internal separator is drawn between Vietnamese and English rows.
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
       for(const existing of Array.from(font.children).filter(e=>e.localName==='color'))font.removeChild(existing);
       const mutedBlue=sd.createElementNS(n,'color');mutedBlue.setAttribute('rgb','FF245A86');font.appendChild(mutedBlue);
       const nf=all(fonts,'font').length;fonts.appendChild(font);fonts.setAttribute('count',String(nf+1));
       xf.setAttribute('fontId',String(nf));xf.setAttribute('applyFont','1');
       const baseBorder=all(borders,'border')[Number(original.getAttribute('borderId')||0)]||all(borders,'border')[0];
       const b=baseBorder.cloneNode(true);
       let top=Array.from(b.children).find(e=>e.localName==='top');
       if(!top){top=sd.createElementNS(n,'top');b.appendChild(top);}
       top.removeAttribute('style');while(top.firstChild)top.removeChild(top.firstChild);
       const nb=all(borders,'border').length;borders.appendChild(b);borders.setAttribute('count',String(nb+1));
       xf.setAttribute('borderId',String(nb));xf.setAttribute('applyBorder','1');
       const ns=all(xfs,'xf').length;xfs.appendChild(xf);xfs.setAttribute('count',String(ns+1));
       styleCache.set(baseId,ns);return ns;
     }
     const hallQuantityStyleCache=new Map();
     function hallQuantityStyle(baseId){
       if(hallQuantityStyleCache.has(baseId))return hallQuantityStyleCache.get(baseId);
       const original=all(xfs,'xf')[baseId]||all(xfs,'xf')[0],xf=original.cloneNode(true);
       const originalFont=all(fonts,'font')[Number(original.getAttribute('fontId')||0)]||all(fonts,'font')[0];
       const font=originalFont.cloneNode(true);
       if(!Array.from(font.children).some(e=>e.localName==='i'))font.appendChild(sd.createElementNS(n,'i'));
       for(const existing of Array.from(font.children).filter(e=>e.localName==='color'))font.removeChild(existing);
       const ink=sd.createElementNS(n,'color');ink.setAttribute('rgb','FF245A86');font.appendChild(ink);
       const fontId=all(fonts,'font').length;fonts.appendChild(font);fonts.setAttribute('count',String(fontId+1));
       xf.setAttribute('fontId',String(fontId));xf.setAttribute('applyFont','1');
       const id=all(xfs,'xf').length;xfs.appendChild(xf);xfs.setAttribute('count',String(id+1));
       hallQuantityStyleCache.set(baseId,id);return id;
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
       if(!before||!/(?:[À-ỹ]|\bthang\b|\bhop\b)/i.test(before)||!en||after)return null;
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
         qtyCell.setAttribute('s',String(hallQuantityStyle(baseStyle)));
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
     // v5.8.210: English translations are already on a dedicated row and
     // styled italic/navy. Remove only their OUTER translation parentheses;
     // do not touch technical parentheses inside Vietnamese or model codes.
     for(const [rn,row] of afterSplit){
       const previous=afterSplit.get(rn-1);
       if(!previous||row.getAttribute('hidden')==='1'||previous.getAttribute('hidden')==='1')continue;
       const cell=getCell(row,'C'),above=getCell(previous,'C');
       if(!cell||!above)continue;
       const value=content(cell).trim(),vietnamese=content(above).trim();
       if(!/^\([A-Za-z0-9][\s\S]*\)$/.test(value)||/[À-ỹ]/.test(value)||!/[À-ỹ]/.test(vietnamese))continue;
       if(content(getCell(row,'A')||x.createElementNS(n,'c')).trim()||
          content(getCell(row,'B')||x.createElementNS(n,'c')).trim())continue;
       // English explanations may be short: (No), (1 set), (8 persons).
       updateCell(x,cell,value.slice(1,-1).trim());
       cell.setAttribute('s',String(englishStyle(Number(cell.getAttribute('s')||0))));
     }
     // v5.8.123: final perimeter repair for BOTH source formats.
     // Separate VI/EN rows have NO separator BETWEEN the rows,
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
       const aLower=content(getCell(lowerRow,'A')||x.createElementNS(n,'c')).trim();
       const bLower=content(getCell(lowerRow,'B')||x.createElementNS(n,'c')).trim();
       // Feature rows carry both labels separately: Vietnamese above, English below.
       // Other description rows have no label on their English row.
       if(!aLower&&!bLower)return true;
       const aUpper=content(getCell(upperRow,'A')||x.createElementNS(n,'c')).trim();
       return !bLower && /[À-ỹ]/.test(aUpper) && !/[À-ỹ]/.test(aLower) &&
         /^\s*\(?[A-Za-z]/.test(aLower);
     }
     for(const [rn,row] of afterSplit){
       if(!hasBilingualPair(rn))continue;
       const bottomRow=afterSplit.get(rn+1),vi=getCell(row,'C'),en=getCell(bottomRow,'C');
       vi.setAttribute('s',String(perimeterStyle(Number(vi.getAttribute('s')||0),
         {left:'table',right:'table',bottom:'none'})));
       en.setAttribute('s',String(perimeterStyle(Number(en.getAttribute('s')||0),
         {left:'table',right:'table',top:'none',bottom:'table'})));
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
         {left:'table',right:'table',top:'none',bottom:'none'})));
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
       ['Độ chính xác dừng tầng (mm)','Trong phạm vi ±10 mm']
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
   // v5.8.211: FINAL presentation pass. This runs after ALL row insertions,
   // source-model style normalization, and optional detailed rows. Earlier
   // passes could replace styles and erase English blue/italic properties.
   // Working from the final OOXML is shared by MR and MRL Version02, and
   // the PDF print layout reads these very same styles/rich text runs.
   {
     const stylePath='xl/styles.xml',sd=parse(await zip.file(stylePath).async('string'));
     const fonts=all(sd,'fonts')[0],xfs=all(sd,'cellXfs')[0];
     if(!fonts||!xfs)throw Error('Export bilingual final style missing');
     const originalStyleCount=all(xfs,'xf').length;
     const fontCache=new Map(),alignmentCache=new Map();
     const read=c=>all(c,'t').map(t=>t.textContent||'').join('');
     const textCell=(row,col)=>Array.from(row.children).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'))||null;
     const rows=all(x,'row').filter(row=>row.getAttribute('hidden')!=='1');
     const rowMap=new Map(rows.map(row=>[+row.getAttribute('r'),row]));
     const hasVi=v=>/[À-ỹ]/.test(String(v||''));
     const isEnglishLine=(str)=>{
       const v=String(str||'').trim().replace(/^\(|\)$/g,'').trim();
       if(!v||hasVi(v)||!/[A-Za-z]/.test(v))return false;
       // Exclude isolated engineering codes/units and bare dimensions.
       if(/^(?:[A-Z][A-Z0-9-]{0,13}|\d+(?:[.,]\d+)?\s*(?:mm|kg|m|m\/s))$/.test(v))return false;
       return true;
     };
     function normalizeEnglish(v){
       const raw=String(v||'').trim();
       // Only remove outer TRANSLATION parentheses. Keep parentheses inside
       // mathematical or engineering notations, e.g. (1.0 m/s) in a sentence.
       return /^\([\s\S]+\)$/.test(raw)?raw.slice(1,-1).trim():raw;
     }
     function languageStyle(baseId){
       const key=Number(baseId)||0;
       if(fontCache.has(key))return fontCache.get(key);
       const base=all(xfs,'xf')[key]||all(xfs,'xf')[0],xf=base.cloneNode(true);
       const font= (all(fonts,'font')[+(base.getAttribute('fontId')||0)]||all(fonts,'font')[0]).cloneNode(true);
       for(const old of Array.from(font.children).filter(c=>c.localName==='color'))old.remove();
       if(!Array.from(font.children).some(c=>c.localName==='i'))font.appendChild(sd.createElementNS(n,'i'));
       const color=sd.createElementNS(n,'color');color.setAttribute('rgb','FF245A86');font.appendChild(color);
       const fontId=all(fonts,'font').length;fonts.appendChild(font);
       xf.setAttribute('fontId',String(fontId));xf.setAttribute('applyFont','1');
       const id=all(xfs,'xf').length;xfs.appendChild(xf);fontCache.set(key,id);return id;
     }
     function middleStyle(baseId,wrap){
       const key=String(baseId||0)+':'+(wrap?'wrap':'middle');
       if(alignmentCache.has(key))return alignmentCache.get(key);
       const xf=(all(xfs,'xf')[Number(baseId)||0]||all(xfs,'xf')[0]).cloneNode(true);
       let a=Array.from(xf.children).find(e=>e.localName==='alignment');
       if(!a){a=sd.createElementNS(n,'alignment');const before=Array.from(xf.children).find(e=>['protection','extLst'].includes(e.localName));xf.insertBefore(a,before||null);}
       a.setAttribute('vertical','center');
       if(wrap)a.setAttribute('wrapText','1');
       xf.setAttribute('applyAlignment','1');
       const id=all(xfs,'xf').length;xfs.appendChild(xf);alignmentCache.set(key,id);return id;
     }
     function asRichCell(cell,parts){
       // Write Excel inline string runs with explicit run-level properties.
       // Plain Vietnamese remains black; English is navy and italic.
       for(const old of Array.from(cell.children).filter(e=>['is','v','f'].includes(e.localName)))old.remove();
       cell.setAttribute('t','inlineStr');
       const inline=x.createElementNS(n,'is');
       for(const part of parts){
         if(!part.text)continue;
         const r=x.createElementNS(n,'r'),pr=x.createElementNS(n,'rPr');
         const face=x.createElementNS(n,'rFont');face.setAttribute('val','Arial');pr.appendChild(face);
         if(part.en){
           pr.appendChild(x.createElementNS(n,'i'));
           const color=x.createElementNS(n,'color');color.setAttribute('rgb','FF245A86');pr.appendChild(color);
         }
         r.appendChild(pr);
         const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');t.textContent=part.text;r.appendChild(t);
         inline.appendChild(r);
       }
       cell.appendChild(inline);
     }
     function makeMixedParts(value){
       // Parenthetical translations are split first. A second pass then
       // styles already line-separated translations without parentheses.
       const original=String(value||'');
       const prelim=bilingualRuns(original);
       let flattened=prelim.map(p=>p.text).join('');
       const lines=flattened.split('\n');
       const result=[];
       let precedingVietnamese=false;
       for(let i=0;i<lines.length;i++){
         const line=lines[i];
         const en=(precedingVietnamese && isEnglishLine(line)) ||
                  // Description-only English after a Vietnamese header can be
                  // styled as a separated worksheet row (handled below).
                  false;
         if(en)result.push({text:normalizeEnglish(line),en:true});
         else result.push({text:line,en:false});
         if(hasVi(line))precedingVietnamese=true;
         if(i<lines.length-1)result.push({text:'\n',en:false});
       }
       // If existing bilingualRuns recognized an English substring in a
       // single physical cell, carry its language tag even for short English.
       if(prelim.some(p=>p.english) && !result.some(p=>p.en)){
         const rebuilt=[];
         for(const p of prelim){
           if(p.english)rebuilt.push({text:normalizeEnglish(p.text),en:true});
           else rebuilt.push({text:p.text,en:false});
         }
         return rebuilt;
       }
       return result;
     }
     for(const row of rows){
       const rn=+row.getAttribute('r');
       const c=textCell(row,'C');
       if(!c)continue;
       // Middle align all column C data cells, preserving original horizontal
       // alignment and wrapping. Applies to both model workbooks.
       c.setAttribute('s',String(middleStyle(c.getAttribute('s'),false)));
       const value=read(c).trim();
       if(!value)continue;
       const preceding=rowMap.get(rn-1),following=rowMap.get(rn+1);
       const neighbor=preceding||null;
       const previousC=neighbor?textCell(neighbor,'C'):null;
       const nextC=following?textCell(following,'C'):null;
       const noOwnLabel=!(textCell(row,'A')&&read(textCell(row,'A')).trim()) && !(textCell(row,'B')&&read(textCell(row,'B')).trim());
       const adjacentVi= (previousC&&hasVi(read(previousC))) || (nextC&&hasVi(read(nextC)));
       const qtyLine=/^\s*\d+\s*hộp\s*\//i.test(value);
       // An entire English translation row must have its own navy italic
       // cell XF. Short translations, such as Thailand, No or 1 set, count.
       // Source Features sometimes store English first (in a labelled
       // row) and Vietnamese second. Style English in those pairs too.
       const currentA=textCell(row,'A'),followingA=following&&textCell(following,'A');
       const englishFirstFeature=!!(nextC&&followingA&&currentA&&
         hasVi(read(nextC))&&hasVi(read(followingA))&&!hasVi(read(currentA))&&
         isEnglishLine(value)&&!hasVi(value));
       const standaloneEn=(noOwnLabel&&adjacentVi&&isEnglishLine(value)&&!hasVi(value) &&
         !/^\s*(?:I|II|III|IV|V|VI)\./.test(value))||englishFirstFeature;
       if(standaloneEn || qtyLine){
         if(standaloneEn){
           const corrected=normalizeEnglish(value);
           if(value!==corrected)updateCell(x,c,corrected);
         }
         c.setAttribute('s',String(languageStyle(c.getAttribute('s'))));
         continue;
       }
       if(!hasVi(value))continue;
       const mixed=makeMixedParts(value);
       if(mixed.some(p=>p.en))asRichCell(c,mixed);
     }
     // In both MR and MRL the dimension labels are merged A:B. They need
     // wrapText and enough row height; values in C retain 'mm' and are middle.
     for(const row of rows){
       const a=textCell(row,'A');if(!a)continue;
       const label=read(a);
       if(!/Rộng\s*\(W\)/i.test(label)||!/\n/.test(label))continue;
       for(const col of ['A','B']){
         const cell=textCell(row,col);
         if(cell)cell.setAttribute('s',String(middleStyle(cell.getAttribute('s'),true)));
       }
       const c=textCell(row,'C');
       if(c){
         const v=read(c).trim();
         if(/^\d[\d.,]*\s*[x×]\s*\d[\d.,]*(?:\s*[x×]\s*\d[\d.,]*)?\s*$/i.test(v))updateCell(x,c,v+' mm');
         c.setAttribute('s',String(middleStyle(c.getAttribute('s'),true)));
       }
       row.setAttribute('ht',String(Math.max(43,Number(row.getAttribute('ht'))||0)));
       row.setAttribute('customHeight','1');
     }
     fonts.setAttribute('count',String(all(fonts,'font').length));
     xfs.setAttribute('count',String(all(xfs,'xf').length));
     zip.file(stylePath,serial(sd));
   }
   // v5.8.212: final source-aware export normalizer (MR + MRL Version 02).
   // Critically, source feature cells use shared-string OOXML ('t=s'), so the
   // earlier text-only pass silently skipped most factory Feature rows.
   // Work on actual values + final layout, preserving every source feature code.
   {
     const sd=parse(await zip.file('xl/styles.xml').async('string'));
     const fonts=all(sd,'fonts')[0],xfs=all(sd,'cellXfs')[0];
     const rows=all(x,'row');
     const rowMap=new Map(rows.map(row=>[+row.getAttribute('r'),row]));
     const ssFile=zip.file('xl/sharedStrings.xml');
     const strings=ssFile?all(parse(await ssFile.async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent||'').join('')):[];
     const get=(row,col)=>Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'))||null;
     function value(c){if(!c)return '';
       if(c.getAttribute('t')==='s')return strings[+(all(c,'v')[0]?.textContent||0)]||'';
       if(c.getAttribute('t')==='inlineStr')return all(c,'t').map(t=>t.textContent||'').join('');
       return all(c,'v')[0]?.textContent||'';
     }
     const hasVi=t=>/[À-ỹ]/.test(String(t||''));
     const fontStyles=new Map();
     function textStyle(id,english,wrap=false){
       const key=String(id||0)+':'+(english?'en':'vi')+':'+wrap;
       if(fontStyles.has(key))return fontStyles.get(key);
       const original=all(xfs,'xf')[+id||0]||all(xfs,'xf')[0];
       const xf=original.cloneNode(true);
       const fontId=+(original.getAttribute('fontId')||0);
       const font=(all(fonts,'font')[fontId]||all(fonts,'font')[0]).cloneNode(true);
       for(const old of Array.from(font.children).filter(c=>['color','i'].includes(c.localName)))old.remove();
       if(english){font.appendChild(sd.createElementNS(n,'i'));
         const col=sd.createElementNS(n,'color');col.setAttribute('rgb','FF245A86');font.appendChild(col);
       } else {
         const col=sd.createElementNS(n,'color');col.setAttribute('rgb','FF000000');font.appendChild(col);
       }
       const newFontId=all(fonts,'font').length;fonts.appendChild(font);
       xf.setAttribute('fontId',String(newFontId));xf.setAttribute('applyFont','1');
       if(wrap){let a=Array.from(xf.children).find(el=>el.localName==='alignment');
         if(!a){a=sd.createElementNS(n,'alignment');const tail=Array.from(xf.children).find(el=>['protection','extLst'].includes(el.localName));xf.insertBefore(a,tail||null);}
         a.setAttribute('vertical','center');a.setAttribute('wrapText','1');xf.setAttribute('applyAlignment','1');
       }
       const styleId=all(xfs,'xf').length;xfs.appendChild(xf);fontStyles.set(key,styleId);return styleId;
     }
     const apply=(cell,text,isEnglish,wrap=false)=>{if(!cell)return;
       updateCell(x,cell,text);
       cell.setAttribute('s',String(textStyle(cell.getAttribute('s'),isEnglish,wrap)));
     };
     const sourceFeatures=new Set();
     let featurePairs=0;
     // The real factory templates store each Feature in two successive rows
     // with identical D code: EN above VN. Flip TEXT only, not row positions,
     // formula codes, hidden state, merges or border structure.
     for(const row of rows){
       const rn=+row.getAttribute('r');if(sourceFeatures.has(rn))continue;
       const next=rowMap.get(rn+1);if(!next)continue;
       const d1=value(get(row,'D')).trim(),d2=value(get(next,'D')).trim();
       const a1=value(get(row,'A')).trim(),a2=value(get(next,'A')).trim();
       const c1=value(get(row,'C')).trim(),c2=value(get(next,'C')).trim();
       if(!d1||d1!==d2||!a1||!a2||!c1||!c2)continue;
       if(hasVi(a1)||!hasVi(a2)||hasVi(c1)||!hasVi(c2))continue;
       // Does not match general spec fields or factory section headers.
       // Works for source features and appended ACC / NSCR-H alike.
       for(const col of ['A','C']){
         const first=get(row,col),second=get(next,col);
         const upper=value(first).trim(),lower=value(second).trim();
         apply(first,lower,false,true);
         apply(second,upper.replace(/^\(([^()]*\S)\)$/,'$1'),true,true);
       }
       for(const rr of [row,next]){
         const size=Math.max(value(get(rr,'A')).length/32,value(get(rr,'C')).length/80);
         rr.setAttribute('ht',String(Math.max(26,Math.min(116,Math.ceil(size)*15+12))));
         rr.setAttribute('customHeight','1');
       }
       sourceFeatures.add(rn);sourceFeatures.add(rn+1);featurePairs++;
     }
     // Ensure 'Giới hạn chịu lửa cửa tầng' spans A:B *and* both bilingual
     // description rows. The source MR workbook merges only A vertically;
     // source MRL may have A:B per row. Remove only these local overlaps.
     const mergeRoot=all(x,'mergeCells')[0];
     const mergeSet=()=>mergeRoot?Array.from(mergeRoot.children).filter(el=>el.localName==='mergeCell'):[];
     const rangeMatch=t=>/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(t||'');
     let fireFixed=0;
     for(const row of rows){
       const rn=+row.getAttribute('r'),a=get(row,'A'),txt=value(a);
       if(!/(?:Giới hạn chịu lửa cửa tầng|Cửa chống cháy\s*\(Fireproof door\))/i.test(txt))continue;
       const below=rowMap.get(rn+1);
       const hasSecond=!!below && !value(get(below,'A')).trim() && !value(get(below,'B')).trim() &&
         (!!value(get(below,'C')).trim() || mergeSet().some(m=>{const r=rangeMatch(m.getAttribute('ref'));return r&&+r[2]===rn&&+r[4]>rn;}));
       const endRow=hasSecond?rn+1:rn;
       // Do not overwrite another item. Both B values must be empty.
       if(value(get(row,'B')).trim()||(hasSecond&&value(get(below,'B')).trim()))continue;
       for(const m of mergeSet()){
         const q=rangeMatch(m.getAttribute('ref'));if(!q)continue;
         const c1=q[1].charCodeAt(0),c2=q[3].charCodeAt(0);
         if(c1<=66&&c2>=65&&+q[2]<=endRow&&+q[4]>=rn)m.remove();
       }
       if(mergeRoot){const item=x.createElementNS(n,'mergeCell');
         item.setAttribute('ref','A'+rn+':B'+endRow);mergeRoot.appendChild(item);
       }
       // Bilingual label with a line break; the EN line is blue italic.
       if(a){
         const parts=[
           {text:'Giới hạn chịu lửa cửa tầng',english:false},
           {text:'\n',english:false},
           {text:'Entrance Fire resistance / Fire rating',english:true}
         ];
         for(const old of Array.from(a.children).filter(e=>['is','v','f'].includes(e.localName)))old.remove();
         a.setAttribute('t','inlineStr');const inline=x.createElementNS(n,'is');
         for(const part of parts){const r=x.createElementNS(n,'r'),pr=x.createElementNS(n,'rPr');
           if(part.english){pr.appendChild(x.createElementNS(n,'i'));
             const color=x.createElementNS(n,'color');color.setAttribute('rgb','FF245A86');pr.appendChild(color);
           }
           const name=x.createElementNS(n,'rFont');name.setAttribute('val','Arial');pr.appendChild(name);
           r.appendChild(pr);const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');
           t.textContent=part.text;r.appendChild(t);inline.appendChild(r);
         }a.appendChild(inline);
         a.setAttribute('s',String(textStyle(a.getAttribute('s'),false,true)));
       }
       for(const rr of [row,...(hasSecond?[below]:[])]){
         rr.setAttribute('ht',String(Math.max(34,Number(rr.getAttribute('ht')||0))));
         rr.setAttribute('customHeight','1');
       }
       fireFixed++;
     }
     // The source hall button group uses: A label merged vertically across
     // all floors; B floor label merged across VI + EN + quantity; C remains
     // three separate rows. Preserve those merges, and normalize the C
     // quantity line in dark-blue italic, without inventing merged data cells.
     const hallQuantity=/^\s*\d+\s*hộp\s*\/\s*\d+\s*thang\s*\/\s*1\s*tầng/i;
     let quantityFixed=0;
     for(const row of rows){const c=get(row,'C'),str=value(c).trim();
       if(!hallQuantity.test(str))continue;
       c.setAttribute('s',String(textStyle(c.getAttribute('s'),true,true)));
       row.setAttribute('ht',String(Math.max(25,+row.getAttribute('ht')||0)));
       row.setAttribute('customHeight','1');quantityFixed++;
     }
     // Keep the short COP finish note, but style this SUPPLEMENTARY line
     // italic navy (user explicitly requested it). Do not duplicate text.
     let copFinishFixed=0;
     for(const row of rows){const c=get(row,'C');if(!c)continue;
       const v=value(c);if(!/Bề mặt inox sọc nhuyễn\s*SUS(?:430|304)/i.test(v))continue;
       const parts=v.split(/(Bề mặt inox sọc nhuyễn\s*SUS(?:430|304)\.?)/i);
       if(parts.length<3)continue;
       for(const old of Array.from(c.children).filter(e=>['v','is','f'].includes(e.localName)))old.remove();
       c.setAttribute('t','inlineStr');const inline=x.createElementNS(n,'is');
       for(const part of parts){if(!part)continue;const marked=/^Bề mặt inox sọc nhuyễn\s*SUS(?:430|304)/i.test(part);
         const run=x.createElementNS(n,'r'),pr=x.createElementNS(n,'rPr');
         const font=x.createElementNS(n,'rFont');font.setAttribute('val','Arial');pr.appendChild(font);
         if(marked){pr.appendChild(x.createElementNS(n,'i'));
           const ink=x.createElementNS(n,'color');ink.setAttribute('rgb','FF245A86');pr.appendChild(ink);
         }run.appendChild(pr);const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');
         t.textContent=part;run.appendChild(t);inline.appendChild(run);
       }c.appendChild(inline);copFinishFixed++;
       c.setAttribute('s',String(textStyle(c.getAttribute('s'),false,true)));
     }
     if(mergeRoot)mergeRoot.setAttribute('count',String(mergeSet().length));
     fonts.setAttribute('count',String(all(fonts,'font').length));
     xfs.setAttribute('count',String(all(xfs,'xf').length));
     zip.file('xl/styles.xml',serial(sd));
     // Audit for testing, not an extra exported file or end-user dependency.
     if(typeof window!=='undefined')window.__SPEC_EXPORT_212_AUDIT__={model,featurePairs,fireFixed,quantityFixed,copFinishFixed};
   }
   // v5.8.213: Last-pass workbook-only correction for both factory models.
   // This MUST follow v5.8.211/212 language/feature formatting to preserve
   // PDF/XLSX parity and avoid later style writers wiping the rich text.
   {
     const stylesPath='xl/styles.xml';
     const sd=parse(await zip.file(stylesPath).async('string'));
     const fonts=all(sd,'fonts')[0],xfs=all(sd,'cellXfs')[0];
     const ssPart=zip.file('xl/sharedStrings.xml');
     const shared=ssPart?all(parse(await ssPart.async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent||'').join('')):[];
     const dataRows=all(x,'row').filter(node=>node.getAttribute('hidden')!=='1');
     const merges=all(x,'mergeCells')[0];
     const wrapStyles=new Map();
     const getCell=(row,col)=>Array.from(row.children).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'))||null;
     const cellValue=c=>{if(!c)return '';
       if(c.getAttribute('t')==='s')return shared[Number(all(c,'v')[0]?.textContent||0)]||'';
       if(c.getAttribute('t')==='inlineStr')return all(c,'t').map(t=>t.textContent||'').join('');
       return all(c,'v')[0]?.textContent||'';};
     const wrapStyle=old=>{
       const key=String(old||0);
       if(wrapStyles.has(key))return wrapStyles.get(key);
       const base=all(xfs,'xf')[Number(key)]||all(xfs,'xf')[0];
       const xf=base.cloneNode(true);
       let alignment=Array.from(xf.children).find(n=>n.localName==='alignment');
       if(!alignment){alignment=sd.createElementNS(n,'alignment');
         const after=Array.from(xf.children).find(node=>['protection','extLst'].includes(node.localName));
         xf.insertBefore(alignment,after||null);}
       alignment.setAttribute('vertical','center');
       alignment.setAttribute('wrapText','1');
       xf.setAttribute('applyAlignment','1');
       const id=all(xfs,'xf').length;xfs.appendChild(xf);
       wrapStyles.set(key,id);return id;
     };
     const richBilingual=(cell,vi,en)=>{
       // The English portion occupies a NEW LINE, with no translation brackets.
       // Italic and navy are explicit in the run, not dependent on cell style.
       if(!cell)return;
       for(const node of Array.from(cell.children).filter(node=>['v','f','is'].includes(node.localName)))node.remove();
       cell.setAttribute('t','inlineStr');
       const inline=x.createElementNS(n,'is');
       for(const part of [{text:vi,en:false},{text:'\n',en:false},{text:en,en:true}]){
         const run=x.createElementNS(n,'r'),style=x.createElementNS(n,'rPr');
         const font=x.createElementNS(n,'rFont');font.setAttribute('val','Arial');style.appendChild(font);
         const color=x.createElementNS(n,'color');color.setAttribute('rgb',part.en?'FF245A86':'FF000000');style.appendChild(color);
         if(part.en)style.appendChild(x.createElementNS(n,'i'));
         run.appendChild(style);
         const txt=x.createElementNS(n,'t');txt.setAttribute('xml:space','preserve');
         // XML text is a literal newline (not slash+n).
         txt.textContent=part.text;
         run.appendChild(txt);inline.appendChild(run);
       }
       cell.appendChild(inline);
       cell.setAttribute('s',String(wrapStyle(cell.getAttribute('s'))));
     };
     const dimensionLabels={
       'Car internal size':['Kích thước phòng thang','Car internal size','R × S × C','W × D × H'],
       'Door size':['Kích thước cửa','Door size','R × C','W × H'],
       'Hoistway size':['Kích thước hố thang','Hoistway size','R × S','W × D'],
       'Machine room size':['Kích thước phòng máy','Machine room size','R × S × C','W × D × H']
     };
     // Restrict the edit to Section II. Landing/other feature labels may
     // also contain the words "Door size" elsewhere in the source workbook.
     const sectionRow=heading=>dataRows.find(row=>heading.test(cellValue(getCell(row,'A'))));
     const dimStart=Number(sectionRow(/^\s*II\.\s*(?:Các kích thước|Dimensional)/i)?.getAttribute('r'))||0;
     const dimEnd=Number(sectionRow(/^\s*III\.\s*(?:Thiết kế|Car design)/i)?.getAttribute('r'))||99999;
     let dimensionsFixed=0;const seenDimensions=[];
     for(const row of dataRows){
       const num=Number(row.getAttribute('r'));
       if(num<=dimStart||num>=dimEnd)continue;
       const a=getCell(row,'A'),b=getCell(row,'B'),c=getCell(row,'C');
       if(!a||!b||!c)continue;
       const label=cellValue(a).trim();
       const match=Object.entries(dimensionLabels).find(([english])=>label.includes(english));
       if(!match)continue;
       const [vi,en,viAxes,enAxes]=match[1],rn=Number(row.getAttribute('r'));
       // Unmerge only the dimension row's A:B range, preserving all other
       // source merges including feature headings and the travel/Pit/OH rows.
       for(const merge of Array.from(merges?.children||[]).filter(node=>node.localName==='mergeCell')){
         const m=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(merge.getAttribute('ref')||'');
         if(!m)continue;
         const col=letters=>[...letters].reduce((sum,letter)=>sum*26+letter.charCodeAt(0)-64,0);
         if(+m[2]<=rn && +m[4]>=rn && col(m[1])<=2 && col(m[3])>=1)merge.remove();
       }
       richBilingual(a,vi,en);
       richBilingual(b,viAxes,enAxes);
       c.setAttribute('s',String(wrapStyle(c.getAttribute('s'))));
       const content=cellValue(c).trim();
       if(/^\d[\d.,]*\s*[x×]\s*\d[\d.,]*(?:\s*[x×]\s*\d[\d.,]*)?$/.test(content))updateCell(x,c,content+' mm');
       // Two bilingual lines fit naturally at 35 pt in NEXIEZ-MR;
       // the former 49 pt creates excessive empty space in rows 37-40.
       // Find these rows by their label, not unstable worksheet row numbers.
       row.setAttribute('ht','35');
       row.setAttribute('customHeight','1');
       dimensionsFixed++;seenDimensions.push(rn+':'+en);
     }
     if(dimensionsFixed!==4)throw Error('Dimension rows (A/B unmerge) expected 4, found '+dimensionsFixed+' in '+model+'; rows '+seenDimensions.join(',')+'; section '+dimStart+'..'+dimEnd);
     if(merges)merges.setAttribute('count',String(Array.from(merges.children).filter(el=>el.localName==='mergeCell').length));
     // Header is company text only. The old form code/version/date are removed.
     for(const row of dataRows.filter(row=>+row.getAttribute('r')<=2)){
       const rn=+row.getAttribute('r');
       if(rn===1){
         const a=getCell(row,'A');
         if(model==='MR'&&a)updateCell(x,a,'MITSUBISHI ELECTRIC VIETNAM CO., LTD.');
         const c=getCell(row,'C');if(c)updateCell(x,c,'');
       }
       if(model==='MR'&&rn===2){
         const c=getCell(row,'C');if(c)updateCell(x,c,'');
       }
     }
     // v5.8.214: NEXIEZ-MR export polish based on the FINAL XLSX cells.
     // Separate technical codes (Mitsubishi Standard) from translations;
     // restore the blue/italic English COP model and hall indicator lines,
     // keeping the SUS304 surface statement as Vietnamese, upright, black.
     let plainStandards=0,copModelFixed=0,hallSurfacesFixed=0,hallQuantitiesFixed=0;
     if(model==='MR'){
       const normalStyles=new Map();
       function plainBlackStyle(old){
         const key=String(old||0);
         if(normalStyles.has(key))return normalStyles.get(key);
         const base=all(xfs,'xf')[Number(key)]||all(xfs,'xf')[0];
         const xf=base.cloneNode(true);
         const fontSource=all(fonts,'font')[Number(base.getAttribute('fontId')||0)]||all(fonts,'font')[0];
         const font=fontSource.cloneNode(true);
         for(const child of Array.from(font.children).filter(el=>['i','color'].includes(el.localName)))child.remove();
         const color=sd.createElementNS(n,'color');color.setAttribute('rgb','FF000000');font.appendChild(color);
         xf.setAttribute('fontId',String(all(fonts,'font').length));
         fonts.appendChild(font);xf.setAttribute('applyFont','1');
         const id=all(xfs,'xf').length;xfs.appendChild(xf);normalStyles.set(key,id);return id;
       }
       function styledRuns(cell,parts){
         for(const child of Array.from(cell.children).filter(el=>['v','f','is'].includes(el.localName)))child.remove();
         cell.setAttribute('t','inlineStr');
         const inline=x.createElementNS(n,'is');
         for(const part of parts){
           if(!part.text)continue;
           const run=x.createElementNS(n,'r'),pr=x.createElementNS(n,'rPr');
           const name=x.createElementNS(n,'rFont');name.setAttribute('val','Arial');pr.appendChild(name);
           if(part.italic)pr.appendChild(x.createElementNS(n,'i'));
           const color=x.createElementNS(n,'color');color.setAttribute('rgb',part.color||'FF000000');pr.appendChild(color);
           run.appendChild(pr);
           const text=x.createElementNS(n,'t');text.setAttribute('xml:space','preserve');text.textContent=part.text;run.appendChild(text);
           inline.appendChild(run);
         }
         cell.appendChild(inline);
         cell.setAttribute('s',String(wrapStyle(plainBlackStyle(cell.getAttribute('s')))));
       }
       const regular=(text,color='FF000000')=>({text,color,italic:false});
       const english=text=>({text,color:'FF245A86',italic:true});
       function vietnameseHallRuns(line){
         // Vietnamese material finish has the same black upright style as all Vietnamese text.
         return [regular(String(line))];
       }
       for(const row of dataRows){
         const c=getCell(row,'C');if(!c)continue;
         const value=cellValue(c).trim();if(!value)continue;
         // Engineering standards are independent values, not English rows.
         if(/^Mitsubishi Standard$/i.test(value)){
           c.setAttribute('s',String(plainBlackStyle(c.getAttribute('s'))));plainStandards++;
           continue;
         }
         // COP English model line (e.g. 'Model CBV1-N1010E ...') can be
         // combined in a factory cell with Vietnamese and a short finish note.
         const label=cellValue(getCell(row,'A'));
         if(/Bảng điều khiển\s*\(Car operation panel\)/i.test(label)&&/\bModel\s+[A-Z0-9-]+/i.test(value)){
           const match=/\(\s*(Model\s+[A-Z0-9-]+[^()\n]*?)\s*\)/i.exec(value);
           const chunks=[];
           if(match){
             const vi=value.slice(0,match.index).trim();
             const suffix=value.slice(match.index+match[0].length).replace(/^\s*\.?\s*/,'').trim();
             if(vi)chunks.push(regular(vi),regular('\n'));
             chunks.push(english(match[1].trim().replace(/\.$/,'')));
             if(suffix){for(const ln of suffix.split(/\n/).map(v=>v.trim()).filter(Boolean)){
               chunks.push(regular('\n'));
               // Existing COP note is a distinct secondary note, keep its
               // earlier blue/italic convention for consistency.
               chunks.push(/^Bề mặt inox sọc nhuyễn\s*SUS(?:430|304)/i.test(ln)?english(ln):regular(ln));
             }}
           }else{
             const lines=value.split(/\r?\n/).map(v=>v.trim()).filter(Boolean);
             for(const [i,ln] of lines.entries()){
               if(i)chunks.push(regular('\n'));
               chunks.push(/^Model\s+/i.test(ln)?english(ln.replace(/^\((.*)\)$/,'$1')):regular(ln));
             }
           }
           if(chunks.length){styledRuns(c,chunks);copModelFixed++;}
         }
         // Hall display description: the Vietnamese SUS304 finish is a
         // physical/material specification, NOT an English translation.
         if(/Hiển thị\s*(?:&|và).*nút\s*(?:gọi|bấm)\s*tầng/i.test(value) &&
            /Hall indicators\s*&\s*buttons/i.test(value)){
           const parts=[];
           const lines=value.split(/\r?\n/).map(v=>v.trim()).filter(Boolean);
           for(const [i,line] of lines.entries()){
             if(i)parts.push(regular('\n'));
             if(/^\(?\s*Hall indicators\s*&\s*buttons\b/i.test(line)){
               const en=line.replace(/^\(([\s\S]*)\)\s*\.?$/,'$1').trim();
               parts.push(english(en));
             }else if(/^\d+\s*hộp\s*\//i.test(line)){
               parts.push(english(line));hallQuantitiesFixed++;
             }else parts.push(...vietnameseHallRuns(line));
           }
           styledRuns(c,parts);hallSurfacesFixed++;
           continue;
         }
         // Hall quantity may be in a dedicated C row rather than appended
         // as a third line. It is still navy + italic.
         if(/^\d+\s*hộp\s*\/\s*\d+\s*thang\s*\/\s*\d+\s*tầng/i.test(value)){
           styledRuns(c,[english(value)]);hallQuantitiesFixed++;
         }
       }
     }
     if(typeof window!=='undefined')window.__SPEC_EXPORT_214_AUDIT__={model,plainStandards,copModelFixed,hallSurfacesFixed,hallQuantitiesFixed};
     xfs.setAttribute('count',String(all(xfs,'xf').length));
     fonts.setAttribute('count',String(all(fonts,'font').length));
     zip.file(stylesPath,serial(sd));
     if(typeof window!=='undefined')window.__SPEC_EXPORT_213_AUDIT__={model,dimensionsFixed,headerUpdated:true};
   }
   // v5.8.216 - feature-pair perimeter normalizer, LAST workbook edit.
   // Source Feature lines are A:B merged independently for VI and EN, with
   // a single C description. Old fixes styled only A and C. Excel also
   // inherits the B-cell edge of each A:B merge, which created misplaced
   // horizontal lines in the actually selected/hidden feature output.
   // Normalize the six A/B/C cells AS ONE two-row item (both models).
   {
     const sp='xl/styles.xml',sd=parse(await zip.file(sp).async('string'));
     const fonts=all(sd,'fonts')[0],borders=all(sd,'borders')[0],xfs=all(sd,'cellXfs')[0];
     const data=all(x,'sheetData')[0],mergeRoot=all(x,'mergeCells')[0];
     const rs=Array.from(data.children).filter(el=>el.localName==='row');
     const index=new Map(rs.map(row=>[+row.getAttribute('r'),row]));
     const stringsPart=zip.file('xl/sharedStrings.xml');
     const shared=stringsPart?all(parse(await stringsPart.async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent||'').join('')):[];
     const get=(row,c)=>Array.from(row?.children||[]).find(e=>e.localName==='c'&&e.getAttribute('r')===c+row.getAttribute('r'))||null;
     const value=c=>!c?'':c.getAttribute('t')==='s'?(shared[+(all(c,'v')[0]?.textContent||0)]||''):
       c.getAttribute('t')==='inlineStr'?all(c,'t').map(t=>t.textContent||'').join(''):(all(c,'v')[0]?.textContent||'');
     const vi=t=>/[À-ỹ]/.test(String(t||''));
     const styleMemo=new Map();
     // Clone fonts, borders and cell styles without modifying the factory
     // styles used by non-Feature rows or either technical calculation.
     function pairStyle(oldStyle,english,column,topRow){
       const cacheKey=[oldStyle,english,column,topRow].join(':');
       if(styleMemo.has(cacheKey))return styleMemo.get(cacheKey);
       const original=all(xfs,'xf')[+oldStyle||0]||all(xfs,'xf')[0];
       const xf=original.cloneNode(true);
       const font=(all(fonts,'font')[+(original.getAttribute('fontId')||0)]||all(fonts,'font')[0]).cloneNode(true);
       for(const child of Array.from(font.children).filter(el=>['i','color'].includes(el.localName)))child.remove();
       if(english)font.appendChild(sd.createElementNS(n,'i'));
       const ink=sd.createElementNS(n,'color');ink.setAttribute('rgb',english?'FF245A86':'FF000000');font.appendChild(ink);
       xf.setAttribute('fontId',String(all(fonts,'font').length));fonts.appendChild(font);
       xf.setAttribute('applyFont','1');
       const border=(all(borders,'border')[+(original.getAttribute('borderId')||0)]||all(borders,'border')[0]).cloneNode(true);
       // Only one perimeter per two-row Feature. No rule between VI and EN.
       const rules={left:column==='A'||column==='C',right:column==='B'||column==='C',
         top:topRow,bottom:!topRow};
       for(const sideName of ['left','right','top','bottom']){
         let side=Array.from(border.children).find(el=>el.localName===sideName);
         if(!side){side=sd.createElementNS(n,sideName);
           const order=['left','right','top','bottom','diagonal','vertical','horizontal','start','end'];
           const following=Array.from(border.children).find(el=>order.indexOf(el.localName)>order.indexOf(sideName));
           border.insertBefore(side,following||null);
         }
         while(side.firstChild)side.removeChild(side.firstChild);
         if(rules[sideName]){
           side.setAttribute('style','thin');
           const color=sd.createElementNS(n,'color');color.setAttribute('rgb','FF000000');side.appendChild(color);
         }else side.removeAttribute('style');
       }
       xf.setAttribute('borderId',String(all(borders,'border').length));borders.appendChild(border);
       xf.setAttribute('applyBorder','1');
       let align=Array.from(xf.children).find(el=>el.localName==='alignment');
       if(!align){align=sd.createElementNS(n,'alignment');
         const next=Array.from(xf.children).find(el=>['protection','extLst'].includes(el.localName));
         xf.insertBefore(align,next||null);
       }
       align.setAttribute('vertical','center');align.setAttribute('wrapText','1');xf.setAttribute('applyAlignment','1');
       const sid=all(xfs,'xf').length;xfs.appendChild(xf);styleMemo.set(cacheKey,sid);return sid;
     }
     const merges=()=>Array.from(mergeRoot?.children||[]).filter(el=>el.localName==='mergeCell');
     const overlappingMerge=(rn)=>merges().filter(el=>{
       const m=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(el.getAttribute('ref')||'');
       if(!m)return false;
       const start=[...m[1]].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0);
       const end=[...m[3]].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0);
       return start<=2&&end>=1&&+m[2]<=rn&&+m[4]>=rn;
     });
     function ensureCell(row,c){
       let cell=get(row,c);if(cell)return cell;
       cell=x.createElementNS(n,'c');cell.setAttribute('r',c+row.getAttribute('r'));
       const pos=c.charCodeAt(0);
       const next=Array.from(row.children).find(el=>el.localName==='c'&&(el.getAttribute('r')||'').charCodeAt(0)>pos);
       row.insertBefore(cell,next||null);return cell;
     }
     let featurePairs=0,featureStyles=0,hiddenPairs=0;
     const processed=new Set();
     for(const row of rs){
       const rn=+row.getAttribute('r'),next=index.get(rn+1);
       if(!next||processed.has(rn)||processed.has(rn+1))continue;
       const d1=value(get(row,'D')).trim(),d2=value(get(next,'D')).trim();
       const a1=value(get(row,'A')).trim(),a2=value(get(next,'A')).trim();
       const c1=value(get(row,'C')).trim(),c2=value(get(next,'C')).trim();
       if(!d1||d1!==d2||!a1||!a2||!c1||!c2)continue;
       if(vi(a1)===vi(a2))continue;
       // The pair is recognized from the SAME source Feature ID D,
       // not from two adjacent visible rows after filtering.
       const hidden1=row.getAttribute('hidden')==='1';
       const hidden2=next.getAttribute('hidden')==='1';
       if(hidden1!==hidden2)throw Error('Export Feature has a half-hidden bilingual pair: '+d1+' rows '+rn+'-'+(rn+1));
       processed.add(rn);processed.add(rn+1);
       if(hidden1){hiddenPairs++;continue;}
       if(!vi(a1)&&vi(a2)){
         // Legacy English-first template. Swap the text on BOTH columns.
         // The feature code in D stays untouched.
         for(const col of ['A','C']){
           const upper=get(row,col),lower=get(next,col);
           const first=value(upper),second=value(lower);
           updateCell(x,upper,second);updateCell(x,lower,first);
         }
       }
       const expected=['A'+rn+':B'+rn,'A'+(rn+1)+':B'+(rn+1)];
       for(const rr of [rn,rn+1]){
         const found=overlappingMerge(rr);
         if(found.some(el=>!expected.includes(el.getAttribute('ref'))))
           throw Error('Export Feature unexpected merge at '+d1+' row '+rr+': '+found.map(el=>el.getAttribute('ref')).join(', '));
         if(!found.length){
           const item=x.createElementNS(n,'mergeCell');item.setAttribute('ref','A'+rr+':B'+rr);mergeRoot.appendChild(item);
         }
       }
       for(const [theRow,english,topRow] of [[row,false,true],[next,true,false]]){
         for(const col of ['A','B','C']){
           const cell=ensureCell(theRow,col);
           const old=cell.getAttribute('s')||0;
           cell.setAttribute('s',String(pairStyle(old,english,col,topRow)));featureStyles++;
         }
         const a=get(theRow,'A'),c=get(theRow,'C');
         if(english){
           for(const cell of [a,c]){
             const str=value(cell),clean=str.replace(/^\s*\(([A-Za-z][^()]*)\)\s*$/,'$1');
             if(clean!==str)updateCell(x,cell,clean);
           }
         }
         const longest=Math.max(value(a).length/37,value(c).length/73);
         const height=Math.max(24,Math.min(124,Math.ceil(Math.max(1,longest))*15+9));
         theRow.setAttribute('ht',String(height));theRow.setAttribute('customHeight','1');
       }
       featurePairs++;
     }
     // Factory MRL Version 02 also contains *two physical rows* for the
     // English translation of car-ceiling / landing-entrance material. Those
     // entries have no D feature code, so the Feature-only font fixer above
     // would not touch them (e.g. C48 CL2 English remained dark grey).
     const translationStyles=new Map();let translationRows=0;
     function translationStyle(baseId){
       const key=String(baseId||0);if(translationStyles.has(key))return translationStyles.get(key);
       const source=all(xfs,'xf')[+key]||all(xfs,'xf')[0],xf=source.cloneNode(true);
       const font=(all(fonts,'font')[+(source.getAttribute('fontId')||0)]||all(fonts,'font')[0]).cloneNode(true);
       for(const e of Array.from(font.children).filter(el=>['i','color'].includes(el.localName)))e.remove();
       font.appendChild(sd.createElementNS(n,'i'));
       const color=sd.createElementNS(n,'color');color.setAttribute('rgb','FF245A86');font.appendChild(color);
       xf.setAttribute('fontId',String(all(fonts,'font').length));fonts.appendChild(font);
       xf.setAttribute('applyFont','1');
       let align=Array.from(xf.children).find(el=>el.localName==='alignment');
       if(!align){align=sd.createElementNS(n,'alignment');
         const after=Array.from(xf.children).find(el=>['protection','extLst'].includes(el.localName));
         xf.insertBefore(align,after||null);
       }
       align.setAttribute('wrapText','1');align.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');
       const id=all(xfs,'xf').length;xfs.appendChild(xf);translationStyles.set(key,id);return id;
     }
     for(const row of rs){
       const rn=+row.getAttribute('r'),above=index.get(rn-1);
       if(!above||processed.has(rn)||row.getAttribute('hidden')==='1'||above.getAttribute('hidden')==='1')continue;
       const c=get(row,'C'),prior=get(above,'C');if(!c||!prior)continue;
       const en=value(c).trim(),vietnamese=value(prior).trim();
       const label=value(get(above,'A')).trim();
       if(!en||!vietnamese||!vi(vietnamese)||!label||!/^\(?[A-Za-z]/.test(en)||vi(en))continue;
       if(value(get(row,'A')).trim()||value(get(row,'B')).trim())continue;
       // Translate rows belonging to the same vertically merged label only;
       // prevent repainting a technical item as English (Mitsubishi Standard).
       const coupled=overlappingMerge(rn).some(m=>{
         const a=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m.getAttribute('ref')||'');
         return a && +a[2]===rn-1 && +a[4]>=rn;
       });
       if(!coupled)continue;
       const stripped=en.replace(/^\s*\(([\s\S]*)\)\s*\.?\s*$/,'$1').trim();
       if(stripped!==en)updateCell(x,c,stripped);
       c.setAttribute('s',String(translationStyle(c.getAttribute('s'))));translationRows++;
     }
     // Inline two-language factory cells (e.g. 'Không cung cấp (Not
     // supplied)') also occur widely in MRL Version 02. Preserve one cell,
     // add a literal line break and rich-text blue italic translation.
     const mixedStyles=new Map();let mixedTranslations=0;
     function mixedWrapStyle(baseId){
       const id=String(baseId||0);if(mixedStyles.has(id))return mixedStyles.get(id);
       const xf=(all(xfs,'xf')[+id]||all(xfs,'xf')[0]).cloneNode(true);
       let a=Array.from(xf.children).find(e=>e.localName==='alignment');
       if(!a){a=sd.createElementNS(n,'alignment');const after=Array.from(xf.children).find(e=>['protection','extLst'].includes(e.localName));xf.insertBefore(a,after||null);}
       a.setAttribute('wrapText','1');a.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');
       const key=all(xfs,'xf').length;xfs.appendChild(xf);mixedStyles.set(id,key);return key;
     }
     for(const row of rs){
       if(row.getAttribute('hidden')==='1')continue;
       const rn=+row.getAttribute('r');if(rn<11||rn>340)continue;
       const c=get(row,'C');if(!c||all(c,'r').length)continue; // existing rich-text formatting wins
       const original=value(c).trim();if(!vi(original))continue;
       const m=/^([\s\S]*?)\s*\(([^()]*[a-z][^()]*)\)\s*\.?\s*$/.exec(original);
       if(!m||!vi(m[1])||!m[2].trim())continue;
       const vietnamese=m[1].trim(),english=m[2].trim();
       // Exclude unit annotations/codes: a parenthetical with lowercase
       // explanatory words is language; a pure code such as SUS-HL is not.
       if(!/[a-z]{2}/.test(english))continue;
       for(const old of Array.from(c.children).filter(e=>['v','is','f'].includes(e.localName)))old.remove();
       c.setAttribute('t','inlineStr');const is=x.createElementNS(n,'is');
       for(const entry of [{text:vietnamese,english:false},{text:'\n',english:false},{text:english,english:true}]){
         const run=x.createElementNS(n,'r'),properties=x.createElementNS(n,'rPr');
         const font=x.createElementNS(n,'rFont');font.setAttribute('val','Arial');properties.appendChild(font);
         if(entry.english)properties.appendChild(x.createElementNS(n,'i'));
         const color=x.createElementNS(n,'color');color.setAttribute('rgb',entry.english?'FF245A86':'FF000000');properties.appendChild(color);
         run.appendChild(properties);const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');
         t.textContent=entry.text;run.appendChild(t);is.appendChild(run);
       }
       c.appendChild(is);c.setAttribute('s',String(mixedWrapStyle(c.getAttribute('s'))));
       // Keeping both languages in one merged/labelled item avoids inserts
       // that would shift border and merge references downstream.
       const estimated=Math.ceil(vietnamese.length/76)+Math.ceil(english.length/76);
       row.setAttribute('ht',String(Math.max(+row.getAttribute('ht')||0,Math.min(98,estimated*15+7))));
       row.setAttribute('customHeight','1');mixedTranslations++;
     }
     if(mergeRoot)mergeRoot.setAttribute('count',String(merges().length));
     fonts.setAttribute('count',String(all(fonts,'font').length));
     borders.setAttribute('count',String(all(borders,'border').length));
     xfs.setAttribute('count',String(all(xfs,'xf').length));
     // Assertions are part of the Export: malformed borders cannot reach users.
     for(const rn of processed){
       const row=index.get(rn);if(row?.getAttribute('hidden')==='1')continue;
       for(const col of ['A','B','C']){
         const cell=get(row,col);if(!cell)throw Error('Feature border missing cell '+col+rn);
         const style=all(xfs,'xf')[+cell.getAttribute('s')||0];
         const border=all(borders,'border')[+style.getAttribute('borderId')||0];
         const bottom=Array.from(border.children).find(el=>el.localName==='bottom');
         const top=Array.from(border.children).find(el=>el.localName==='top');
         const isLower=processed.has(rn-1) && value(get(index.get(rn-1),'D'))===value(get(row,'D'));
         if((top?.getAttribute('style')||'')!==(isLower?'':'thin') ||
           (bottom?.getAttribute('style')||'')!==(isLower?'thin':''))
           throw Error('Feature border invariant broken: '+col+rn);
         if(isLower){
           const font=all(fonts,'font')[+style.getAttribute('fontId')||0];
           if(!Array.from(font.children).some(el=>el.localName==='i') ||
             !Array.from(font.children).some(el=>el.localName==='color'&&el.getAttribute('rgb')==='FF245A86'))
             throw Error('Feature English italic/navy invariant broken: '+col+rn);
         }
       }
     }
     zip.file(sp,serial(sd));
     if(typeof window!=='undefined')window.__SPEC_EXPORT_216_AUDIT__={model,featurePairs,featureStyles,hiddenPairs,translationRows,mixedTranslations};
   }
   // v5.8.217: final MRL presentation invariant, AFTER every factory
   // conversion, feature-border and bilingual post-processor. Do not use
   // source-XLSX styles to guess whether a surface phrase is English:
   // "bề mặt inox sọc nhuyễn SUS304" is Vietnamese and stays upright black.
   // English hall-panel description AND requested quantity are blue italic.
   // This normalizer edits actual final OOXML runs, so Excel and PDF agree.
   if(model==='MRL'){
     const sp='xl/styles.xml',sd=parse(await zip.file(sp).async('string'));
     const fontRoot=all(sd,'fonts')[0],xfRoot=all(sd,'cellXfs')[0];
     const ssPart=zip.file('xl/sharedStrings.xml');
     const shared=ssPart?all(parse(await ssPart.async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent||'').join('')):[];
     const visibleRows=all(x,'row').filter(row=>row.getAttribute('hidden')!=='1');
     const styleCache=new Map();
     const findCell=(row,col)=>Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'))||null;
     const getText=c=>!c?'':c.getAttribute('t')==='s'
       ?shared[Number(all(c,'v')[0]?.textContent||0)]||''
       :c.getAttribute('t')==='inlineStr'?all(c,'t').map(t=>t.textContent||'').join('')
       :all(c,'v')[0]?.textContent||'';
     function formattedStyle(baseId,english){
       const key=String(baseId||0)+':'+(english?'en':'vi');
       if(styleCache.has(key))return styleCache.get(key);
       const original=all(xfRoot,'xf')[Number(baseId)||0]||all(xfRoot,'xf')[0];
       const xf=original.cloneNode(true);
       const fontSource=all(fontRoot,'font')[Number(original.getAttribute('fontId')||0)]||all(fontRoot,'font')[0];
       const font=fontSource.cloneNode(true);
       for(const node of Array.from(font.children).filter(el=>['i','color'].includes(el.localName)))node.remove();
       if(english)font.appendChild(sd.createElementNS(n,'i'));
       const ink=sd.createElementNS(n,'color');ink.setAttribute('rgb',english?'FF245A86':'FF000000');font.appendChild(ink);
       xf.setAttribute('fontId',String(all(fontRoot,'font').length));xf.setAttribute('applyFont','1');fontRoot.appendChild(font);
       let alignment=Array.from(xf.children).find(c=>c.localName==='alignment');
       if(!alignment){alignment=sd.createElementNS(n,'alignment');
         const next=Array.from(xf.children).find(el=>['protection','extLst'].includes(el.localName));
         xf.insertBefore(alignment,next||null);
       }
       alignment.setAttribute('wrapText','1');alignment.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');
       const id=all(xfRoot,'xf').length;xfRoot.appendChild(xf);styleCache.set(key,id);return id;
     }
     function rewrite(cell,segments){
       const base=Number(cell.getAttribute('s')||0);
       for(const el of Array.from(cell.children).filter(c=>['v','f','is'].includes(c.localName)))el.remove();
       cell.setAttribute('t','inlineStr');const inline=x.createElementNS(n,'is');
       for(const part of segments){
         if(!part.text)continue;
         const run=x.createElementNS(n,'r'),pr=x.createElementNS(n,'rPr');
         const face=x.createElementNS(n,'rFont');face.setAttribute('val','Arial');pr.appendChild(face);
         if(part.english)pr.appendChild(x.createElementNS(n,'i'));
         const color=x.createElementNS(n,'color');color.setAttribute('rgb',part.english?'FF245A86':'FF000000');pr.appendChild(color);
         run.appendChild(pr);
         const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');t.textContent=part.text;run.appendChild(t);inline.appendChild(run);
       }
       cell.appendChild(inline);
       const onlyEnglish=segments.every(part=>part.english||!part.text.trim());
       cell.setAttribute('s',String(formattedStyle(base,onlyEnglish)));
     }
     function styledLines(source){
       const parts=[];
       const lines=String(source).split(/\r?\n/).map(v=>v.trim()).filter(Boolean);
       for(let i=0;i<lines.length;i++){
         const original=lines[i];
         const quantity=/^\d+\s*hộp\s*\/\s*\d+\s*thang\s*\//i.test(original);
         const englishHall=/^\(?\s*Hall indicators\s*&\s*buttons\b/i.test(original);
         const englishModel=/^\(?\s*Model\s+[A-Z0-9-]+/i.test(original)&&!/[À-ỹ]/.test(original);
         const english=quantity||englishHall||englishModel;
         const line=english&&!quantity?original.replace(/^\s*\((.+)\)\s*\.?$/,'$1').trim():original;
         if(i)parts.push({text:'\n',english:false});
         parts.push({text:line,english});
       }
       return parts;
     }
     let thaiFixed=0,doorFixed=0,hallFixed=0,finishFixed=0;
     for(const row of visibleRows){
       const rn=+row.getAttribute('r'),c=findCell(row,'C');if(!c)continue;
       let source=getText(c),trimmed=source.trim();if(!trimmed)continue;
       // Source evaluator may return "Thailan" (no trailing d) even when
       // the original cached workbook has the correct "Thailand".
       if(rn<35&&/Thái Lan\s*\(?Thailan(?:d)?\)?/i.test(trimmed)){
         rewrite(c,[{text:'Thái Lan\n',english:false},{text:'Thailand',english:true}]);thaiFixed++;continue;
       }
       const label=getText(findCell(row,'A'));
       if(rn<40&&/Loại cửa\s*\(Door type\)/i.test(label)&&/^(?:CO|2S)(?:\s*-.*)?$/i.test(trimmed)){
         const d=trimmed.slice(0,2).toUpperCase(),desc=d==='CO'
           ?['CO - Hai cánh mở ra hai bên từ trung tâm','2-panel center opening doors']
           :['2S - Hai cánh mở về một phía','2-panel side sliding doors'];
         rewrite(c,[{text:desc[0]+'\n',english:false},{text:desc[1],english:true}]);doorFixed++;continue;
       }
       // The hall indicator body may be one cell (VI+EN+quantity), or
       // three source rows; both must get the same explicit font/runs.
       if(/^(?:Hiển thị\s*&\s*nút gọi tầng|\(?Hall indicators\s*&\s*buttons|\d+\s*hộp\s*\/)/i.test(trimmed)){
         const segments=styledLines(source);
         // Some workbook cells contain VI+EN on the same line; rely on
         // explicit text boundaries, not the legacy rich-run styling.
         if(segments.length) {rewrite(c,segments);hallFixed++;}
         continue;
       }
       // COP material is Vietnamese regardless of the uppercase "Bề mặt".
       if(/Bề mặt inox sọc nhuyễn\s*SUS(?:430|304)/i.test(source)){
         const segments=styledLines(source);
         if(segments.length){rewrite(c,segments);finishFixed++;}
       }
     }
     fontRoot.setAttribute('count',String(all(fontRoot,'font').length));
     xfRoot.setAttribute('count',String(all(xfRoot,'xf').length));
     zip.file(sp,serial(sd));
     if(typeof window!=='undefined')window.__SPEC_EXPORT_217_AUDIT__={model,thaiFixed,doorFixed,hallFixed,finishFixed};
   }
   // v5.8.226: compact I-IV and align MR Excel masthead with MRL, Feature unchanged.
   // (a) Feature-group headings in the V section are distinctly paler than I-V.
   // (b) Each visible Feature row is tall enough for *all* wrapped A:B and C
   //     content, without shifting any merged references, borders, or rows.
   {
     const stylePart='xl/styles.xml';const sd=parse(await zip.file(stylePart).async('string'));
     const fills=all(sd,'fills')[0],xfs=all(sd,'cellXfs')[0];
     const rows=all(x,'row'),rowById=new Map(rows.map(r=>[+r.getAttribute('r'),r]));
     const ssFile=zip.file('xl/sharedStrings.xml');
     const shared=ssFile?all(parse(await ssFile.async('string')),'si').map(item=>all(item,'t').map(t=>t.textContent||'').join('')):[];
     const get=(row,col)=>Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'))||null;
     const value=c=>!c?'':c.getAttribute('t')==='s'?(shared[+(all(c,'v')[0]?.textContent||0)]||''):
       c.getAttribute('t')==='inlineStr'?all(c,'t').map(t=>t.textContent||'').join(''):(all(c,'v')[0]?.textContent||'');
     const lightFill=sd.createElementNS(n,'fill'),pattern=sd.createElementNS(n,'patternFill');
     pattern.setAttribute('patternType','solid');
     const fg=sd.createElementNS(n,'fgColor');fg.setAttribute('rgb','FFEDF7FC');
     const bg=sd.createElementNS(n,'bgColor');bg.setAttribute('indexed','64');
     pattern.append(fg,bg);lightFill.append(pattern);
     const lightId=all(fills,'fill').length;fills.appendChild(lightFill);
     const fillCache=new Map();
     function paleStyle(base){
       const key=String(base||0);if(fillCache.has(key))return fillCache.get(key);
       const old=all(xfs,'xf')[+key]||all(xfs,'xf')[0],xf=old.cloneNode(true);
       xf.setAttribute('fillId',String(lightId));xf.setAttribute('applyFill','1');
       const id=all(xfs,'xf').length;xfs.appendChild(xf);fillCache.set(key,id);return id;
     }
     // v5.8.220: compact Feature rows to the *actual* wrapped line count.
     // Previous fixed (16.4 pt * lines + 10 pt) over-allocated every bilingual
     // row, particularly English rows with one or two lines. Source Excel uses
     // A+B as the label and C as the description; all use Arial 11 pt here.
     const colWidths={AB:20.5546875+25.5546875,C:70.5546875};
     const excelPx=width=>Math.floor(width*7+5);
     const cellPx={AB:excelPx(colWidths.AB)-12,C:excelPx(colWidths.C)-12};
     const canvas=(typeof document!=='undefined')?document.createElement('canvas'):null;
     const ctx=canvas?.getContext('2d')||null;
     // The same font family and size as the outgoing Feature cell styles.
     function wrappedLines(source,pixels,english){
       const body=String(source||'').replace(/\r/g,'');if(!body.trim())return 1;
       if(ctx)ctx.font=(english?'italic ':'')+'11pt Arial';
       const measure=word=>ctx?ctx.measureText(word).width:
         Array.from(word).reduce((sum,c)=>sum+(/[MW@%]/.test(c)?9.0:/[il1'.,:;]/.test(c)?3.5:7.0),0);
       const space=measure(' ');
       let count=0;
       for(const segment of body.split('\n')){
         if(!segment.trim()){count++;continue;}
         let used=0;count++;
         for(const word of segment.trim().split(/\s+/)){
           const size=measure(word);
           if(used>0&&used+space+size>pixels){count++;used=0;}
           if(size>pixels){
             // Unbroken technical tokens/URLs may need multiple physical lines.
             if(used){count++;used=0;}
             count+=Math.ceil(size/pixels)-1;used=size%pixels||pixels;
           }else used=used?used+space+size:size;
         }
       }
       return count;
     }
     // MR Excel uses its own margin to prevent bottom strokes touching borders.
     // PDF does not use these OOXML heights; MRL sizing stays unchanged.
     const compactFeatureHeight=lines=>model==='MR'
       ?Math.max(24,Math.min(360,Math.ceil(lines*15.2+7)))
       :Math.max(18,Math.min(360,Math.ceil(lines*14.1+3.5)));
     let featureRowsFitted=0,groupsShaded=0;
     // Never use source row indices here. Earlier bilingual splits insert rows,
     // so the final five group headings move (MR: 85/111/147/213/255;
     // MRL: 91/117/153/220/262 as of this build). Match final cell text.
     for(const row of rows){
       const rn=+row.getAttribute('r');if(row.getAttribute('hidden')==='1')continue;
       const caption=value(get(row,'A')).trim();
       if(model==='MRL'&&/^(?:Hành trình\s*\(Travel\)|Chiều sâu Pit\s*\(Pit depth\)|Đỉnh hố thang\s*\(Overhead)/i.test(caption)){
         row.setAttribute('ht','20');row.setAttribute('customHeight','1');
       }
       if(/^[-–]\s*Tính năng\b/i.test(caption)&&!value(get(row,'B')).trim()&&!value(get(row,'C')).trim()){
         for(const col of ['A','B','C']){
           let cell=get(row,col);
           if(!cell){cell=x.createElementNS(n,'c');cell.setAttribute('r',col+rn);
             const next=Array.from(row.children).find(e=>e.localName==='c'&&(e.getAttribute('r')||'').charCodeAt(0)>col.charCodeAt(0));
             row.insertBefore(cell,next||null);
           }
           cell.setAttribute('s',String(paleStyle(cell.getAttribute('s'))));
         }
         row.setAttribute('ht','21');row.setAttribute('customHeight','1');groupsShaded++;
       }
       // Only actual Feature lines: D has the same identifier on a bilingual
       // pair of consecutive source rows. Do not adjust unrelated data rows.
       const code=value(get(row,'D')).trim(),before=rowById.get(rn-1),after=rowById.get(rn+1);
       const beforeSame=before&&value(get(before,'D')).trim()===code;
       const afterSame=after&&value(get(after,'D')).trim()===code;
       if(!code||!(beforeSame||afterSame)||!value(get(row,'A')).trim())continue;
       const a=value(get(row,'A')),c=value(get(row,'C'));
       // Measure A:B and C independently. Use a small safety margin, but no
       // fixed 10-point excess: one-line and two-line Feature rows can now shrink.
       const english=/^\s*\(?[A-Za-z]/.test(a)&&!/[À-ỹ]/.test(a);
       const lines=Math.max(wrappedLines(a,cellPx.AB,english),wrappedLines(c,cellPx.C,english));
       const needed=compactFeatureHeight(lines);
       row.setAttribute('ht',String(needed));row.setAttribute('customHeight','1');featureRowsFitted++;
     }
     // v5.8.224: refit ALL visible customer-facing Excel rows (not only Features) with safer padding.
     // Source workbook basic/spec/car-design rows sometimes retain 13.8pt despite
     // two bilingual lines; Excel clips the lower line even though PDF auto-wraps.
     // Honor original column widths, merged label geometry, and font size, and
     // enlarge only when the content cannot fit. Never alter C# calculation cells.
     const mergeRoot=all(x,'mergeCells')[0];
     const mergeNodes=mergeRoot?all(mergeRoot,'mergeCell'):[];
     const mergeByStart=new Map();
     const occupied=new Set();
     const cellIndex=code=>Array.from(code).reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
     for(const m of mergeNodes){
       const ref=String(m.getAttribute('ref')||'');
       const match=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
       if(!match)continue;
       const c1=cellIndex(match[1]),c2=cellIndex(match[3]),r1=+match[2],r2=+match[4];
       mergeByStart.set(match[1]+r1,{cols:c2-c1+1,rows:r2-r1+1});
       for(let rn=r1;rn<=r2;rn++)for(let col=c1;col<=c2;col++)
         if(col<=3)occupied.add(rn+':'+col);
     }
     let normalRowsFitted=0,headingsMerged=0;
     const colWidthsPx=[0,excelPx(20.5546875),excelPx(25.5546875),excelPx(70.5546875)];
     const fontRoot=all(sd,'fonts')[0],fontRecords=fontRoot?all(fontRoot,'font'):[];
     const existingXfs=all(xfs,'xf');
     const romanHeading=/^\s*[IVXLCDM]+\.\s+/i;
     const originNote=/Toàn bộ thiết bị thang máy được nhập khẩu|All main components are imported/i;
     for(const row of rows){
       const rn=+row.getAttribute('r');
       if(rn<=5||rn>310||row.getAttribute('hidden')==='1')continue;
       const a=value(get(row,'A')).trim(),b=value(get(row,'B')).trim(),c=value(get(row,'C')).trim();
       if(originNote.test(a))continue; // source bilingual origin paragraph is already laid out
       const roman=romanHeading.test(a)&&!b&&!c;
       if(roman&&mergeRoot&&!occupied.has(rn+':1')&&!occupied.has(rn+':2')&&!occupied.has(rn+':3')){
         const node=x.createElementNS(n,'mergeCell');node.setAttribute('ref','A'+rn+':C'+rn);
         mergeRoot.appendChild(node);mergeByStart.set('A'+rn,{cols:3,rows:1});
         for(let cc=1;cc<=3;cc++)occupied.add(rn+':'+cc);
         headingsMerged++;
       }
       if(roman){
         if((+row.getAttribute('ht')||0)<20){row.setAttribute('ht','20');row.setAttribute('customHeight','1');normalRowsFitted++;}
         continue;
       }
       const code=value(get(row,'D')).trim(),prev=rowById.get(rn-1),next=rowById.get(rn+1);
       const isFeature=code&&((prev&&value(get(prev,'D')).trim()===code)||(next&&value(get(next,'D')).trim()===code));
       if(isFeature)continue; // compact measurements already applied above
       const group=/^[-–]\s*Tính năng\b/i.test(a)&&!b&&!c;
       if(group)continue; // preset 21pt, blue fill
       let required=0,anyText=false;
       for(const [colName,colIndex] of [['A',1],['B',2],['C',3]]){
         const cell=get(row,colName),t=value(cell);
         if(!t.trim())continue;
         anyText=true;
         const merge=mergeByStart.get(colName+rn)||{cols:1,rows:1};
         if(merge.rows>1)continue; // rowspan is distributed over multiple rows
         const endCol=Math.min(3,colIndex+merge.cols-1);
         let px=0;for(let i=colIndex;i<=endCol;i++)px+=colWidthsPx[i];
         px=Math.max(50,px-12);
         const style=existingXfs[+(cell.getAttribute('s')||0)];
         const font=fontRecords[+(style?.getAttribute('fontId')||0)];
         const sz=font?all(font,'sz')[0]:null;
         const fontSize=sz?Math.max(8,Math.min(13,Number(sz.getAttribute('val'))||11)):11;
         const italic=Boolean(font&&all(font,'i').length);
         // MR needs more room in native Excel (not PDF), including rich text
         // and short entrance descriptions; do not resize MRL rows here.
         const lines=wrappedLines(t,px,italic);
         const lineSpacing=model==='MR'?1.52:1.38;
         const breathing=model==='MR'?8:6;
         required=Math.max(required,Math.ceil(lines*(fontSize*lineSpacing)+breathing));
       }
       if(!anyText)continue;
       required=Math.max(model==='MR'?25:20,Math.min(360,required));
       const original=Number(row.getAttribute('ht')||13.8);
       if(required>original+0.75){row.setAttribute('ht',String(required));row.setAttribute('customHeight','1');normalRowsFitted++;}
     }
     // Some labels span multiple physical rows (e.g. Entrance design blocks).
     // The merged parent cell may fit in total, but if each child row is too
     // tight the table looks visually stuck together. Guarantee enough TOTAL
     // height for the merged text and distribute any deficit across the rows.
     for(const [startKey,merge] of mergeByStart.entries()){
       const m=/^([A-Z]+)(\d+)$/.exec(startKey);if(!m)continue;
       const colName=m[1],startRn=+m[2];
       if(startRn<=5||startRn>310||merge.rows<=1)continue;
       if(!['A','B','C'].includes(colName))continue;
       const row=rowById.get(startRn),cell=get(row,colName),t=value(cell).trim();
       if(!t)continue;
       const colIndex={A:1,B:2,C:3}[colName];
       const endCol=Math.min(3,colIndex+merge.cols-1);
       let px=0;for(let i=colIndex;i<=endCol;i++)px+=colWidthsPx[i];
       px=Math.max(50,px-12);
       const style=existingXfs[+(cell?.getAttribute('s')||0)];
       const font=fontRecords[+(style?.getAttribute('fontId')||0)];
       const sz=font?all(font,'sz')[0]:null;
       const fontSize=sz?Math.max(8,Math.min(13,Number(sz.getAttribute('val'))||11)):11;
       const italic=Boolean(font&&all(font,'i').length);
       const lines=wrappedLines(t,px,italic);
       const requiredTotal=model==='MR'
         ?Math.max(merge.rows*25,Math.ceil(lines*(fontSize*1.52)+10))
         :Math.max(merge.rows*20,Math.ceil(lines*(fontSize*1.38)+8));
       let currentTotal=0;const blockRows=[];
       for(let rr=startRn;rr<startRn+merge.rows;rr++){
         const r=rowById.get(rr); if(!r||r.getAttribute('hidden')==='1')continue;
         const current=Math.max(0,Number(r.getAttribute('ht')||13.8));
         currentTotal+=current;blockRows.push([r,current]);
       }
       if(!blockRows.length||requiredTotal<=currentTotal+0.75)continue;
       const deficit=requiredTotal-currentTotal;
       const share=deficit/blockRows.length;
       for(const [r,current] of blockRows){
         const next=Math.max(model==='MR'?25:20,Math.ceil((current+share)*2)/2);
         if(next>current+0.24){r.setAttribute('ht',String(next));r.setAttribute('customHeight','1');normalRowsFitted++;}
       }
     }
     // On MR, make sure the bilingual origin has full text and wraps. Some
     // versions of Excel clip 'Thailand' visually when inherited wrapText=off.
     if(model==='MR'){
       const originCell=get(rowById.get(12),'C');
       if(originCell){
         const originValue=value(originCell);
         if(/Thailan(?!d)/i.test(originValue)){
           // Only overwrite when upstream evaluator still returns the typo.
           updateCell(x,originCell,originValue.replace(/Thailan(?!d)/gi,'Thailand'));
         }
         const baseId=+(originCell.getAttribute('s')||0);
         const base=all(xfs,'xf')[baseId]||all(xfs,'xf')[0];
         const xf=base.cloneNode(true);
         let align=Array.from(xf.children).find(el=>el.localName==='alignment');
         if(!align){align=sd.createElementNS(n,'alignment');xf.appendChild(align);}
         align.setAttribute('wrapText','1');align.setAttribute('vertical','center');
         xf.setAttribute('applyAlignment','1');
         const newStyleId=all(xfs,'xf').length;xfs.appendChild(xf);
         originCell.setAttribute('s',String(newStyleId));
         const row=rowById.get(12);
         row.setAttribute('ht',String(Math.max(43,Number(row.getAttribute('ht')||0))));
         row.setAttribute('customHeight','1');
       }
     }
     // v5.8.226: compact ONLY sections I-IV in XLSX. Feature group V stays
     // completely unchanged, including previous feature row heights/merges.
     // The earlier 1.52x+8pt margin made nearly every two-line cell 42pt high.
     // Preserve enough leading for Excel 11pt font, rich text and rowspans.
     const introRow=rows.find(rr=>/^I\.\s*Đặc tính kỹ thuật cơ bản/i.test(value(get(rr,'A')).trim()));
     const featuresRow=rows.find(rr=>/Các tính năng\s*\(Features\)/i.test(value(get(rr,'A')).trim()));
     const introId=+(introRow?.getAttribute('r')||0);
     const featuresId=+(featuresRow?.getAttribute('r')||0);
     let compactedIIV=0;
     if(introId && featuresId && featuresId>introId){
       const rowsInSection=new Set();
       for(const rr of rows){
         const rn=+rr.getAttribute('r');
         if(rn<=introId||rn>=featuresId||rr.getAttribute('hidden')==='1')continue;
         const original=Number(rr.getAttribute('ht')||13.8);
         // Zero-height continuation rows are intentionally hidden by the
         // source entrance-design layout. Never open those rows again.
         if(rr.hasAttribute('ht')&&original<1)continue;
         const label=value(get(rr,'A')).trim();
         if(/^\s*(?:I|II|III|IV)\.\s/.test(label)||/^[-–]\s*Tính năng/.test(label))continue;
         let required=20,found=false;
         for(const [colName,colIndex] of [['A',1],['B',2],['C',3]]){
           const cell=get(rr,colName),v=value(cell);
           if(!v.trim())continue;
           found=true;
           const merge=mergeByStart.get(colName+rn)||{cols:1,rows:1};
           if(merge.rows>1)continue;
           let px=0;for(let ci=colIndex;ci<=Math.min(3,colIndex+merge.cols-1);ci++)px+=colWidthsPx[ci];
           const style=existingXfs[+(cell.getAttribute('s')||0)];
           const font=fontRecords[+(style?.getAttribute('fontId')||0)];
           const fsize=Number(all(font,'sz')[0]?.getAttribute('val'))||11;
           const ital=Boolean(font&&all(font,'i').length);
           const lines=wrappedLines(v,Math.max(50,px-15),ital);
           // Line leading ~14.7pt for 11pt Arial, plus vertical padding 4pt.
           required=Math.max(required,Math.ceil(lines*fsize*1.34+4));
         }
         if(!found)continue;
         required=Math.min(280,required);
         // Do not increase already tall rows in this compacting pass;
         // preserve extra capacity for unexpected Excel font substitutions.
         if(original>required+1){rr.setAttribute('ht',String(required));rr.setAttribute('customHeight','1');compactedIIV++;}
         rowsInSection.add(rn);
       }
       // Give vertically-merged A/B labels enough *total* height, without
       // blindly expanding every continuation row to the old 25pt minimum.
       for(const [startKey,merge] of mergeByStart.entries()){
         if(merge.rows<=1)continue;
         const match=/^([A-C])(\d+)$/.exec(startKey);if(!match)continue;
         const colName=match[1],rn=+match[2];
         if(rn<=introId||rn>=featuresId)continue;
         const cell=get(rowById.get(rn),colName),v=value(cell);if(!v.trim())continue;
         const colIndex={A:1,B:2,C:3}[colName];
         let px=0;for(let ci=colIndex;ci<=Math.min(3,colIndex+merge.cols-1);ci++)px+=colWidthsPx[ci];
         const style=existingXfs[+(cell.getAttribute('s')||0)];
         const font=fontRecords[+(style?.getAttribute('fontId')||0)];
         const fsize=Number(all(font,'sz')[0]?.getAttribute('val'))||11;
         const ital=Boolean(font&&all(font,'i').length);
         const target=Math.max(20,Math.ceil(wrappedLines(v,Math.max(50,px-15),ital)*fsize*1.34+5));
         const block=Array.from({length:merge.rows},(_,i)=>rowById.get(rn+i)).filter(Boolean);
         const total=block.reduce((sum,rr)=>sum+Math.max(0,Number(rr.getAttribute('ht')||13.8)),0);
         if(total+1<target){
           const visible=block.find(rr=>Number(rr.getAttribute('ht')||13.8)>0);
           if(visible){visible.setAttribute('ht',String(Math.ceil(Number(visible.getAttribute('ht')||13.8)+target-total)));visible.setAttribute('customHeight','1');}
         }
       }
     }
     // v5.8.229 — MR AutoFit retained. MRL uses a balanced, final-fit pass below.
     // Excel does not reliably AutoFit complex wrapped rich text with vertically
     // merged labels. Compute heights from the real column widths and the
     // longest visible cell instead. Never shrink rows here; preserve Feature
     // section V, source-hidden rows, and company header/logo for both models.
     let autoFitIIV=0;
     const autoFitRows=[];
     if(model==='MR'&&introId&&featuresId){
       for(const rr of rows){
         const rn=+rr.getAttribute('r');
         if(rn<=introId||rn>=featuresId||rr.getAttribute('hidden')==='1')continue;
         if(rr.getAttribute('ht')!=null&&Number(rr.getAttribute('ht'))<1)continue;
         const label=value(get(rr,'A')).trim();
         if(/^\s*(?:I|II|III|IV)\.\s/.test(label))continue;
         let target=20,measuredLines=1,hasText=false;
         for(const [col,colIndex] of [['A',1],['B',2],['C',3]]){
           const cell=get(rr,col),v=value(cell);
           if(!v.trim())continue;
           const merge=mergeByStart.get(col+rn)||{cols:1,rows:1};
           if(merge.rows>1)continue; // accounted for as a block previously
           hasText=true;
           let px=0;
           for(let ci=colIndex;ci<=Math.min(3,colIndex+merge.cols-1);ci++)px+=colWidthsPx[ci];
           // Reserve room for native Excel rendering, cell padding and borders.
           px=Math.max(65,px-30);
           const style=existingXfs[+(cell.getAttribute('s')||0)];
           const font=fontRecords[+(style?.getAttribute('fontId')||0)];
           const fontSize=Number(all(font,'sz')[0]?.getAttribute('val'))||11;
           const italic=Boolean(font&&all(font,'i').length);
           const lines=wrappedLines(v,px,italic);
           measuredLines=Math.max(measuredLines,lines);
           const long=lines>=3;
           const needed=Math.ceil(lines*fontSize*(long?1.52:1.35)+(long?12:5));
           target=Math.max(target,needed);
         }
         if(!hasText)continue;
         target=Math.min(360,target);
         const previous=Number(rr.getAttribute('ht')||13.8);
         if(target>previous+1){
           rr.setAttribute('ht',String(target));
           rr.setAttribute('customHeight','1');
           autoFitIIV++;
           autoFitRows.push({row:rn,lines:measuredLines,before:previous,after:target});
         }
       }
     }
     // v5.8.229: MRL I-IV *balanced* row fitting. v228 only enlarged rows,
     // so heights pre-allocated by old passes (e.g. 34pt for 1-line labels,
     // 79pt for the control system) were never reduced. Refit from content.
     // Keep original 0pt continuation rows, 1:V headings, feature rows, merge
     // ranges, layout for MR and all PDF-specific formatting untouched.
     let mrlRowsBalanced=0;
     const mrlFitSamples=[];
     if(model==='MRL'&&introId&&featuresId&&featuresId>introId){
       const compactLineHeight=fontSize=>fontSize*1.34;
       const originalLineHeight=fontSize=>fontSize*1.34;
       // If a cell spans several rows, ensure the sum of visible row heights
       // can accommodate its content, but never re-open any 0-height row.
       const rowspanNeeds=[];
       for(const rr of rows){
         const rn=+rr.getAttribute('r');
         if(rn<=introId||rn>=featuresId||rr.getAttribute('hidden')==='1')continue;
         const current=Number(rr.getAttribute('ht')??13.8);
         if(current===0)continue;
         const label=value(get(rr,'A')).trim();
         if(/^\s*(?:I|II|III|IV)\.\s/.test(label))continue;
         let necessary=18.5,hasContent=false;
         const info=[];
         for(const [col,colIndex] of [['A',1],['B',2],['C',3]]){
           const cell=get(rr,col),body=value(cell);
           if(!body.trim())continue;
           const merge=mergeByStart.get(col+rn)||{cols:1,rows:1};
           let px=0;
           for(let i=colIndex;i<=Math.min(3,colIndex+merge.cols-1);i++)px+=colWidthsPx[i];
           px=Math.max(45,px-13);
           const style=existingXfs[+(cell.getAttribute('s')||0)];
           const font=fontRecords[+(style?.getAttribute('fontId')||0)];
           const fsize=Number(all(font,'sz')[0]?.getAttribute('val'))||11;
           const italic=Boolean(font&&all(font,'i').length);
           const lines=wrappedLines(body,px,italic);
           const height=Math.ceil(lines*compactLineHeight(fsize)+(lines>=3?5:3));
           if(merge.rows>1){
             rowspanNeeds.push({rn,rows:merge.rows,height,cell:col+rn});
             continue;
           }
           hasContent=true;
           necessary=Math.max(necessary,height);
           info.push({col,lines});
         }
         if(!hasContent)continue;
         necessary=Math.max(19,Math.min(250,necessary));
         if([13,17,19,26,30,35,41,71,79,82,85].includes(rn))mrlFitSamples.push({row:rn,from:current,to:necessary,info,detail:'probe'});
         // Only adjust when difference is meaningful: avoid jitter from half-
         // point rounding. This fits content in both directions (not only up).
         if(Math.abs(necessary-current)>1){
           rr.setAttribute('ht',String(necessary));
           rr.setAttribute('customHeight','1');
           mrlRowsBalanced++;
           if(mrlFitSamples.length<32||[13,19,26,35,71,79,82,85].includes(rn))
             mrlFitSamples.push({row:rn,from:current,to:necessary,info});
         }
       }
       // A vertical merge often extends into intentionally collapsed rows.
       // Put the deficit on an already-visible row in that merged block.
       for(const item of rowspanNeeds){
         const span=[];
         for(let rn=item.rn;rn<item.rn+item.rows;rn++){
           const row=rowById.get(rn);
           if(!row||row.getAttribute('hidden')==='1')continue;
           span.push(row);
         }
         const actual=span.reduce((sum,row)=>sum+Math.max(0,Number(row.getAttribute('ht')??13.8)),0);
         if(actual+1>=item.height)continue;
         const visible=span.find(row=>Number(row.getAttribute('ht')??13.8)>0);
         if(!visible)continue;
         visible.setAttribute('ht',String(Math.ceil(Number(visible.getAttribute('ht')??13.8)+item.height-actual+2)));
         visible.setAttribute('customHeight','1');
         mrlRowsBalanced++;
       }
     }
     // Company and address are free-standing text in the factory template.
     // Merge the top masthead cells before the logo image (row 3 onward), so
     // the MRL address/MEVN company name does not spill or truncate in Excel.
     let mastheadMerged=0;
     for(const rn of [1,2]){
       const row=rowById.get(rn);if(!row||!value(get(row,'A')).trim())continue;
       if(value(get(row,'B')).trim()||value(get(row,'C')).trim())continue;
       if(!mergeRoot||[1,2,3].some(c=>occupied.has(rn+':'+c)))continue;
       const m=x.createElementNS(n,'mergeCell');m.setAttribute('ref','A'+rn+':C'+rn);
       mergeRoot.appendChild(m);
       for(let col=1;col<=3;col++)occupied.add(rn+':'+col);
       if((+row.getAttribute('ht')||0)<20){row.setAttribute('ht','20');row.setAttribute('customHeight','1');}
       mastheadMerged++;
     }
     if(mergeRoot)mergeRoot.setAttribute('count',String(all(mergeRoot,'mergeCell').length));
     fills.setAttribute('count',String(all(fills,'fill').length));
     xfs.setAttribute('count',String(all(xfs,'xf').length));
     zip.file(stylePart,serial(sd));
     if(typeof window!=='undefined')window.__SPEC_EXPORT_229_AUDIT__={model,groupsShaded,featureRowsFitted,normalRowsFitted,headingsMerged,mastheadMerged,compactedIIV,autoFitIIV,autoFitRows,mrlRowsBalanced,mrlFitSamples};
   }
   // v5.8.226: align MR Excel masthead exactly with the existing MRL V02
   // layout: row1 company name, row2 address, logo over rows 3-5, title row6.
   // This is an XLSX-only change, after feature calculations and row fitting.
   // The PDF layout engine remains intact; technical content is untouched.
   if(model==='MR'){
     const data=all(x,'sheetData')[0];
     const rows=all(data,'row');
     const getCell=(row,col)=>Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'));
     const first=rows.find(row=>row.getAttribute('r')==='1');
     const second=rows.find(row=>row.getAttribute('r')==='2');
     if(!first||!second)throw Error('MR masthead: missing first 2 rows');
     let a1=getCell(first,'A'),a2=getCell(second,'A');
     if(!a1){a1=x.createElementNS(n,'c');a1.setAttribute('r','A1');first.insertBefore(a1,first.firstChild);}
     if(!a2){a2=x.createElementNS(n,'c');a2.setAttribute('r','A2');second.insertBefore(a2,second.firstChild);}
     updateCell(x,a1,'MITSUBISHI ELECTRIC VIETNAM');
     updateCell(x,a2,'8th Floor – Tower B – Viettel Tower, 285 Cach Mang Thang 8 Street, Hoa Hung Ward, HCMC');
     second.setAttribute('ht','20');second.setAttribute('customHeight','1');
     // MRL merges both its masthead lines across A:C; otherwise the address
     // clips in native Excel even when the OOXML cell has the complete text.
     const mastheadMerges=all(x,'mergeCells')[0];
     if(!mastheadMerges)throw Error('MR masthead: mergeCells missing');
     if(!all(mastheadMerges,'mergeCell').some(m=>m.getAttribute('ref')==='A2:C2')){
       const merge=x.createElementNS(n,'mergeCell');merge.setAttribute('ref','A2:C2');
       mastheadMerges.appendChild(merge);
       mastheadMerges.setAttribute('count',String(all(mastheadMerges,'mergeCell').length));
     }
     // Match the MRL address typography instead of inheriting a bold header.
     const st=parse(await zip.file('xl/styles.xml').async('string'));
     const styles=all(st,'cellXfs')[0],fonts=all(st,'fonts')[0];
     const xf=(all(styles,'xf')[+(a1.getAttribute('s')||0)]||all(styles,'xf')[0]).cloneNode(true);
     const font=(all(fonts,'font')[+(xf.getAttribute('fontId')||0)]||all(fonts,'font')[0]).cloneNode(true);
     for(const old of Array.from(font.children).filter(e=>['b','sz'].includes(e.localName)))old.remove();
     const bold=st.createElementNS(n,'b'),size=st.createElementNS(n,'sz');size.setAttribute('val','9');
     font.appendChild(bold);font.appendChild(size);
     const fontId=all(fonts,'font').length;fonts.appendChild(font);fonts.setAttribute('count',String(fontId+1));
     xf.setAttribute('fontId',String(fontId));xf.setAttribute('applyFont','1');
     let align=Array.from(xf.children).find(el=>el.localName==='alignment');
     if(!align){align=st.createElementNS(n,'alignment');xf.appendChild(align);}
     align.setAttribute('wrapText','0');align.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');
     const xfId=all(styles,'xf').length;styles.appendChild(xf);styles.setAttribute('count',String(xfId+1));
     a2.setAttribute('s',String(xfId));zip.file('xl/styles.xml',serial(st));
     // One extra row before the main title (MR title at 5 -> MRL title at 6).
     // No formula cells remain on SpecFormExport, but preserve every merge and
     // print range as rows are shifted so exported Excel opens without repair.
     for(const row of rows){
       const old=+row.getAttribute('r');if(old<5)continue;
       row.setAttribute('r',String(old+1));
       for(const cell of Array.from(row.children).filter(c=>c.localName==='c')){
         const ref=cell.getAttribute('r');if(ref)cell.setAttribute('r',ref.replace(/\d+$/,String(old+1)));
       }
     }
     for(const merge of all(x,'mergeCell')){
       const ref=merge.getAttribute('ref');if(!ref)continue;
       merge.setAttribute('ref',ref.replace(/([A-Z]+)(\d+)/g,(all,col,rn)=>col+(Number(rn)>=5?Number(rn)+1:rn)));
     }
     // Correct MR factory source's repeated 'IV. Features' heading without
     // changing its Feature content, styles, merges or row heights.
     for(const row of rows){
       const cell=getCell(row,'A');
       const txt=cell?all(cell,'t').map(v=>v.textContent||'').join(''):'';
       if(/^IV\.\s*Các tính năng\s*\(Features\)/i.test(txt))updateCell(x,cell,txt.replace(/^IV\./,'V.'));
     }
     // The A5:C5 title moved down by one; reserve the new fifth row for the
     // MRL-style image anchor. The original MR heading line stays unchanged.
     for(const dn of all(workbook,'definedName')){
       if(dn.getAttribute('name')==='_xlnm.Print_Area')
         dn.textContent=String(dn.textContent||'').replace(/(\$[A-Z]+\$)(\d+)/g,(v,col,rn)=>col+(Number(rn)>=5?Number(rn)+1:rn));
     }
     for(const breakNode of all(x,'brk')){
       const id=Number(breakNode.getAttribute('id'));
       if(id>=5)breakNode.setAttribute('id',String(id+1));
     }
     // Source MRL uses this same PNG (also used in the PDF fallback); export
     // it as a *real Excel drawing*, not a text approximation or PDF-only logo.
     const logoBase64=window.SPEC_MITSUBISHI_LOGO_BASE64;
     if(!logoBase64||!/^iVBOR/.test(logoBase64))throw Error('MR masthead: MRL logo unavailable');
     const drawPath='xl/drawings/specHeaderLogo.xml',picPath='xl/media/specHeaderLogo.png';
     const drawRid='rIdSpecHeaderLogo',pictureRid='rIdSpecHeaderImage';
     const imageXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
       <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <xdr:twoCellAnchor editAs="oneCell">
          <xdr:from><xdr:col>0</xdr:col><xdr:colOff>34290</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from>
          <xdr:to><xdr:col>0</xdr:col><xdr:colOff>1123950</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>151444</xdr:rowOff></xdr:to>
          <xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Mitsubishi Electric logo"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
           <xdr:blipFill><a:blip r:embed="${pictureRid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
           <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>
        </xdr:twoCellAnchor>
       </xdr:wsDr>`;
     const relsImage=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${pkg}"><Relationship Id="${pictureRid}" Type="${r}/image" Target="/xl/media/specHeaderLogo.png"/></Relationships>`;
     const relsSheet=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${pkg}"><Relationship Id="${drawRid}" Type="${r}/drawing" Target="/xl/drawings/specHeaderLogo.xml"/></Relationships>`;
     zip.file(picPath,logoBase64,{base64:true});zip.file(drawPath,imageXml);
     zip.file('xl/drawings/_rels/specHeaderLogo.xml.rels',relsImage);
     zip.file('xl/worksheets/_rels/sheet2.xml.rels',relsSheet);
     let drawing=all(x,'drawing')[0];
     if(!drawing){drawing=x.createElementNS(n,'drawing');x.documentElement.appendChild(drawing);}
     drawing.setAttributeNS(r,'r:id',drawRid);
     const ctPath='[Content_Types].xml',types=parse(await zip.file(ctPath).async('string'));
     const typeNS='http://schemas.openxmlformats.org/package/2006/content-types';
     const png=Array.from(types.documentElement.children).some(t=>t.getAttribute('Extension')==='png');
     if(!png){const node=types.createElementNS(typeNS,'Default');node.setAttribute('Extension','png');node.setAttribute('ContentType','image/png');types.documentElement.appendChild(node);}
     const node=types.createElementNS(typeNS,'Override');node.setAttribute('PartName','/xl/drawings/specHeaderLogo.xml');
     node.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.drawing+xml');types.documentElement.appendChild(node);
     zip.file(ctPath,serial(types));
     if(typeof window!=='undefined')window.__SPEC_EXPORT_226_MR_HEADER__={company:true,address:true,logo:true,titleRow:6};
   }
   // v5.8.233: Use the same bilingual inline rich-text convention in A/B
   // as in C: Vietnamese first, literal line break, English without its
   // translation parentheses, italic navy #245A86. Never split worksheet
   // rows: preserved merges, borders, feature codes and quantity cells.
   {
     const stylesPart='xl/styles.xml';
     const styleDoc=parse(await zip.file(stylesPart).async('string'));
     const xfs=all(styleDoc,'cellXfs')[0],fonts=all(styleDoc,'fonts')[0];
     const rows=all(x,'sheetData')[0] ? all(all(x,'sheetData')[0],'row') : [];
     const rowByNum=new Map(rows.map(row=>[Number(row.getAttribute('r')),row]));
     const mergeNodes=all(x,'mergeCell');
     const mainDataTop=8; // do not rewrite company address, logo or main banner
     const sharedPart=zip.file('xl/sharedStrings.xml');
     const shared=sharedPart?all(parse(await sharedPart.async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent||'').join('')):[];
     const getCell=(row,col)=>Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'))||null;
     const cellText=cell=>!cell?'':cell.getAttribute('t')==='s'?(shared[Number(all(cell,'v')[0]?.textContent||0)]||''):
       cell.getAttribute('t')==='inlineStr'?all(cell,'t').map(t=>t.textContent||'').join(''):(all(cell,'v')[0]?.textContent||'');
     const hasVietnamese=v=>/[À-ỹ]/.test(v);
     const codeOnly=/^(?:CAR|CWT|MR|MRL|COP|HPI|PIH|BA|BMS|P\d{1,2}D?|W|D|H|R|S)$/i;
     // Only a parenthesized ENGLISH translation at the end of a Vietnamese
     // label is converted. Internal notes (lần đầu), abbreviations (CAR)
     // and function codes such as (OEPS-A) remain untouched.
     function bilingualLabel(source){
       const original=String(source||'').trim();
       if(!original||/^\(/.test(original)||!hasVietnamese(original))return null;
       let vi='',en='',suffix='';
       const inline=/^([\s\S]+?)\s*\(\s*([A-Za-z][^()\n\r]*)\s*\)\s*([:;.]*)$/.exec(original);
       if(inline){vi=inline[1].trim();en=inline[2].trim();suffix=inline[3]||'';}
       else {
         const separated=/^([\s\S]+?)\s*\n\s*\(?\s*([A-Za-z][^()\r\n]*)\s*\)?\s*([:;.]*)$/.exec(original);
         if(!separated)return null;
         vi=separated[1].trim();en=separated[2].trim();suffix=separated[3]||'';
       }
       if(!hasVietnamese(vi)||hasVietnamese(en)||!/[a-z]/i.test(en)||codeOnly.test(en)||!en.trim())return null;
       // Main heading has a trailing colon in its source; retain it at the
       // English end rather than losing punctuation.
       return {vi,en:en+suffix};
     }
     function rPrFromBase(base,isEnglish){
       const baseFont=all(fonts,'font')[Number(base.getAttribute('fontId')||0)]||all(fonts,'font')[0];
       const pr=x.createElementNS(n,'rPr');
       const name=x.createElementNS(n,'rFont');name.setAttribute('val','Arial');pr.appendChild(name);
       const size=all(baseFont,'sz')[0];if(size){const el=x.createElementNS(n,'sz');el.setAttribute('val',size.getAttribute('val'));pr.appendChild(el);}
       if(all(baseFont,'b').length)pr.appendChild(x.createElementNS(n,'b'));
       if(isEnglish){
         pr.appendChild(x.createElementNS(n,'i'));
         const ink=x.createElementNS(n,'color');ink.setAttribute('rgb','FF245A86');pr.appendChild(ink);
       }else{
         const color=all(baseFont,'color')[0];
         if(color)pr.appendChild(x.importNode(color,true));
       }
       return pr;
     }
     const wrapCache=new Map();
     function wrappedStyle(old){
       const key=String(old||0);if(wrapCache.has(key))return wrapCache.get(key);
       const xf=(all(xfs,'xf')[Number(key)]||all(xfs,'xf')[0]).cloneNode(true);
       let align=Array.from(xf.children).find(el=>el.localName==='alignment');
       if(!align){align=styleDoc.createElementNS(n,'alignment');xf.appendChild(align);}
       align.setAttribute('wrapText','1');align.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');
       const styleIndex=all(xfs,'xf').length;xfs.appendChild(xf);wrapCache.set(key,styleIndex);return styleIndex;
     }
     function addRun(inline,base,text,english){
       if(!text)return;
       const run=x.createElementNS(n,'r');run.appendChild(rPrFromBase(base,english));
       const el=x.createElementNS(n,'t');el.setAttribute('xml:space','preserve');el.textContent=text;
       run.appendChild(el);inline.appendChild(run);
     }
     const measuredCanvas=(typeof document!=='undefined')?document.createElement('canvas'):null;
     const ctx=measuredCanvas?.getContext('2d')||null;
     function visibleLines(str,width,fontSize,italic){
       const body=String(str||'').trim();if(!body)return 1;
       if(ctx)ctx.font=(italic?'italic ':'')+fontSize+'pt Arial';
       const measure=word=>ctx?ctx.measureText(word).width:word.length*fontSize*0.67;
       const space=measure(' ');let count=0;
       for(const part of body.split('\n')){
         count++;let used=0;
         for(const word of part.split(/\s+/).filter(Boolean)){
           const size=measure(word);
           if(used>0&&used+space+size>width){count++;used=0;}
           if(size>width){if(used>0){count++;used=0;}const wraps=Math.ceil(size/width);count+=wraps-1;used=size%width||width;}
           else used=used?used+space+size:size;
         }
       }
       return count;
     }
     const colWidth={A:20.5546875,B:25.5546875,C:70.5546875};
     const colPixel=col=>Math.floor((colWidth[col]||20)*7+5);
     const mergeInfos=[];
     for(const node of mergeNodes){
       const match=/^([A-C])(\d+):([A-C])(\d+)$/.exec(node.getAttribute('ref')||'');
       if(match)mergeInfos.push({col:match[1],r1:+match[2],endCol:match[3],r2:+match[4]});
     }
     const indexes={A:0,B:1,C:2},letters=['A','B','C'];
     let converted=0,notChangedCodes=0,heightAdjusted=0;
     const samples=[];
     for(const row of rows){
       const rn=Number(row.getAttribute('r'));
       if(rn<mainDataTop||rn>310||row.getAttribute('hidden')==='1')continue;
       for(const col of ['A','B']){
         const cell=getCell(row,col);if(!cell)continue;
         const value=cellText(cell);
         const normalized=bilingualLabel(value);
         if(!normalized){if(/\((?:CAR|CWT)\)/.test(value))notChangedCodes++;continue;}
         // Already-correct bilingual rich strings (e.g. dimensional A/B)
         // should remain byte-for-byte equivalent; no need to reserialize.
         const inlineExisting=all(cell,'is')[0];
         if(inlineExisting&&all(inlineExisting,'r').length>1&&value.includes('\n')){
           const enRuns=all(inlineExisting,'r').filter(run=>
             all(run,'i').length>0&&all(run,'color').some(c=>/245A86$/i.test(c.getAttribute('rgb')||'')));
           if(enRuns.length)continue;
         }
         const oldStyle=Number(cell.getAttribute('s')||0);
         const xf=all(xfs,'xf')[oldStyle]||all(xfs,'xf')[0];
         for(const child of Array.from(cell.children).filter(node=>['v','f','is'].includes(node.localName)))child.remove();
         cell.setAttribute('t','inlineStr');
         const inline=x.createElementNS(n,'is');
         addRun(inline,xf,normalized.vi,false);
         addRun(inline,xf,'\n',false);
         addRun(inline,xf,normalized.en,true);
         cell.appendChild(inline);cell.setAttribute('s',String(wrappedStyle(oldStyle)));
         converted++;
         if(samples.length<15)samples.push({cell:col+rn,vi:normalized.vi,en:normalized.en});
         const cols=mergeInfos.find(m=>m.col===col&&m.r1===rn)||{col,endCol:col,r1:rn,r2:rn};
         const span=letters.slice(indexes[col],indexes[cols.endCol]+1);
         const pixels=Math.max(60,span.reduce((sum,c)=>sum+colPixel(c),0)-17);
         const font=all(fonts,'font')[Number(xf.getAttribute('fontId')||0)]||all(fonts,'font')[0];
         const fontSize=Number(all(font,'sz')[0]?.getAttribute('val')||11);
         const lines=visibleLines(normalized.vi,pixels,fontSize,false)+visibleLines(normalized.en,pixels,fontSize,true);
         const needed=Math.max(20,Math.ceil(lines*fontSize*1.35+5));
         const block=[];for(let r=rn;r<=cols.r2;r++){const item=rowByNum.get(r);if(item&&item.getAttribute('hidden')!=='1')block.push(item);}
         if(!block.length)continue;
         const sum=block.reduce((total,item)=>total+Math.max(0,Number(item.getAttribute('ht')??13.8)),0);
         // For merged vertical blocks, preserve zero-height continuation rows.
         if(sum+0.75<needed){
           const target=block.find(item=>Number(item.getAttribute('ht')??13.8)>0);if(!target)continue;
           const previous=Number(target.getAttribute('ht')??13.8);
           target.setAttribute('ht',String(Math.min(360,Math.ceil(previous+needed-sum+1))));
           target.setAttribute('customHeight','1');heightAdjusted++;
         }
       }
     }
     xfs.setAttribute('count',String(all(xfs,'xf').length));
     zip.file(stylesPart,serial(styleDoc));
     if(typeof window!=='undefined')window.__SPEC_EXPORT_233_AUDIT__={model,converted,notChangedCodes,heightAdjusted,samples};
   }
   // v5.8.238 — final Excel presentation pass. It intentionally operates
   // AFTER all bilingual rich-text conversions so row fit is based on what
   // the customer will see. Does not change printHtml / PDF or Feature V.
   {
     const stylePath='xl/styles.xml';
     const styleDoc=parse(await zip.file(stylePath).async('string'));
     const styleXfs=all(all(styleDoc,'cellXfs')[0],'xf');
     const styleFonts=all(all(styleDoc,'fonts')[0],'font');
     const targetRows=all(x,'row');
     const byRow=new Map(targetRows.map(rr=>[Number(rr.getAttribute('r')),rr]));
     const sharedPart=zip.file('xl/sharedStrings.xml');
     const shared=sharedPart?all(parse(await sharedPart.async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent||'').join('')):[];
     const cell=(row,col)=>Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'))||null;
     const textOf=c=>!c?'':c.getAttribute('t')==='s'?(shared[Number(all(c,'v')[0]?.textContent||0)]||''):
       c.getAttribute('t')==='inlineStr'?all(c,'t').map(t=>t.textContent||'').join(''):(all(c,'v')[0]?.textContent||'');
     let entranceFixed=0,rowsReduced=0;
     if(model==='MRL'){
       // The English material + thickness is present on a collapsed source
       // continuation row. The prior visible combined cell was malformed:
       // "(Stainless..., dày 1.2 mm.". Read the validated English source,
       // strip its parentheses, and render with separate rich-text fonts.
       for(const row of targetRows){
         const rn=Number(row.getAttribute('r'));
         if(row.getAttribute('hidden')==='1')continue;
         const a=textOf(cell(row,'A'));
         if(!/^(?:Khung cửa tầng|Cửa tầng)(?:\s|\n|$)/i.test(a.trim()))continue;
         const c=cell(row,'C'),raw=textOf(c),following=byRow.get(rn+1);
         const nextC=following?.getAttribute('hidden')==='1'?cell(following,'C'):null;
         const followingEN=textOf(nextC).trim();
         if(!c||!/SUS\s*430J1L|Stainless steel|E-102/i.test(raw))continue;
         const candidates=raw.split(/\r?\n/),viPart=(candidates[0]||'').trim();
         let enPart=followingEN&&/(?:Stainless steel|E-102|thickness)/i.test(followingEN)
           ?followingEN : candidates.slice(1).join(' ').trim();
         if(!enPart)continue;
         enPart=enPart.replace(/^\s*\(+/,'').replace(/\)+\s*\.?\s*$/,'')
           .replace(/\bdày\s*(\d+(?:[.,]\d+)?\s*mm)/gi,'thickness $1')
           .replace(/[.\s]+$/,'').trim();
         if(!enPart||/[À-ỹ]/.test(enPart))continue;
         const th=/\bthickness\s+(\d+(?:[.,]\d+)?\s*mm)/i.exec(enPart);
         let vi=viPart;
         if(th&&!/\bdày\s*\d/i.test(vi))
           vi=vi.replace(/[.;]\s*$/,'').trimEnd()+', dày '+th[1]+'.';
         for(const child of Array.from(c.children).filter(el=>['v','f','is'].includes(el.localName)))child.remove();
         c.setAttribute('t','inlineStr');
         const inline=x.createElementNS(n,'is');
         const base=styleXfs[Number(c.getAttribute('s')||0)]||styleXfs[0];
         const font=styleFonts[Number(base?.getAttribute('fontId')||0)]||styleFonts[0];
         const fontSize=all(font,'sz')[0]?.getAttribute('val')||'11';
         const append=(body,isEN)=>{
           const run=x.createElementNS(n,'r'),pr=x.createElementNS(n,'rPr');
           const face=x.createElementNS(n,'rFont');face.setAttribute('val','Arial');pr.appendChild(face);
           const size=x.createElementNS(n,'sz');size.setAttribute('val',fontSize);pr.appendChild(size);
           if(isEN)pr.appendChild(x.createElementNS(n,'i'));
           const color=x.createElementNS(n,'color');color.setAttribute('rgb',isEN?'FF245A86':'FF000000');pr.appendChild(color);
           run.appendChild(pr);
           const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');t.textContent=body;run.appendChild(t);
           inline.appendChild(run);
         };
         append(vi+'\n',false);append(enPart+'.',true);c.appendChild(inline);
         entranceFixed++;
       }
     }
     // Re-evaluate the short spec/dimension rows in Excel. Previous source
     // templates set 36/63/79 pt even for short bilingual values. Derive
     // physical lines independently per cell and shrink, NEVER add rows or
     // touch 0pt / hidden continuation rows. Feature and print remain intact.
     const painter=(typeof document!=='undefined')?document.createElement('canvas').getContext('2d'):null;
     function physicalLines(source,pixels){
       const paragraphs=String(source||'').replace(/\r/g,'').split('\n');
       if(!painter)return paragraphs.reduce((sum,p)=>sum+Math.max(1,Math.ceil(p.length/(pixels/7.5))),0);
       painter.font='11pt Arial';const space=painter.measureText(' ').width;
       let total=0;
       for(const p of paragraphs){
         if(!p.trim()){total++;continue;}
         let used=0;total++;
         for(const word of p.trim().split(/\s+/)){
           const w=painter.measureText(word).width;
           if(used>0&&used+space+w>pixels){total++;used=0;}
           if(w>pixels){
             const n=Math.ceil(w/pixels);
             if(used>0){total++;used=0;}
             total+=n-1;used=w%pixels||pixels;
           }else used=used?used+space+w:w;
         }
       }
       return Math.max(1,total);
     }
     // Final physical width of Excel columns A/B/C from factory workbook.
     const widths={A:20.5546875,B:25.5546875,C:70.5546875};
     const pixelsFor=col=>Math.max(55,Math.floor(widths[col]*7+5)-16);
     // Preserve existing vertical merge geometry when fitting; a spanning label
     // is already displayed across multiple rows and must not be counted twice.
     const mergedStarts=new Map();
     for(const m of all(x,'mergeCell')){
       const match=/^([A-C])(\d+):([A-C])(\d+)$/.exec(m.getAttribute('ref')||'');
       if(match)mergedStarts.set(match[1]+match[2],{start:match[1],end:match[3],endRow:+match[4]});
     }
     const columnOrder=['A','B','C'];
     const fittingRow=(row)=>{
       const rn=Number(row.getAttribute('r'));
       if(row.getAttribute('hidden')==='1'||Number(row.getAttribute('ht')??13.8)===0)return false;
       if(model==='MRL')return rn>=12&&rn<=54;
       return rn>=51&&rn<=57;
     };
     for(const row of targetRows){
       if(!fittingRow(row))continue;
       const rn=Number(row.getAttribute('r'));
       const a=textOf(cell(row,'A'));
       if(/^\s*(?:I|II|III|IV|V)\.\s/.test(a))continue;
       // Do not compress extra motor/ride-comfort/noise tables or long prose:
       // they contain complex mixed styles and long wrapped sentences.
       if(model==='MRL'&&rn>=36&&rn<=44)continue;
       const previous=Number(row.getAttribute('ht')??13.8);
       if(previous<28)continue;
       let lineCount=1,any=false;
       for(const col of columnOrder){
         const c=cell(row,col),body=textOf(c).trim();if(!body)continue;
         const m=mergedStarts.get(col+rn);
         if(m&&m.endRow>rn)continue;
         any=true;
         const covered=m?columnOrder.slice(columnOrder.indexOf(col),columnOrder.indexOf(m.end)+1):[col];
         const width=covered.reduce((sum,v)=>sum+pixelsFor(v),0);
         lineCount=Math.max(lineCount,physicalLines(body,width));
       }
       if(!any)continue;
       // 11pt Arial, 14.8pt line step, 4pt total breathing room; min 20pt.
       const required=Math.min(280,Math.max(20,Math.ceil(lineCount*14.0+3)));
       if(previous>required+1){
         row.setAttribute('ht',String(required));row.setAttribute('customHeight','1');rowsReduced++;
       }
     }
     if(typeof window!=='undefined')window.__SPEC_EXPORT_238_AUDIT__={model,entranceFixed,rowsReduced};
   }
   // v5.8.239 — FIX excess whitespace across sections I–IV (XLSX only).
   // Use final visible content, do not count vertically merged labels twice,
   // preserve hidden 0pt rows and preserve V. Features exactly as generated.
   {
     const sd=parse(await zip.file('xl/styles.xml').async('string'));
     const xfs=all(all(sd,'cellXfs')[0],'xf'),fonts=all(all(sd,'fonts')[0],'font');
     const sheetRows=all(x,'row');
     const siPart=zip.file('xl/sharedStrings.xml');
     const si=siPart?all(parse(await siPart.async('string')),'si').map(el=>all(el,'t').map(t=>t.textContent||'').join('')):[];
     const at=(row,col)=>Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'))||null;
     const textOf=c=>!c?'':c.getAttribute('t')==='s'?si[+(all(c,'v')[0]?.textContent||0)]||'':c.getAttribute('t')==='inlineStr'?all(c,'t').map(t=>t.textContent||'').join(''):all(c,'v')[0]?.textContent||'';
     const cells=['A','B','C'],widths={A:20.5546875,B:25.5546875,C:70.5546875};
     const mergeStarts=new Map();
     for(const m of all(x,'mergeCell')){
       const p=/^([A-C])(\d+):([A-C])(\d+)$/.exec(m.getAttribute('ref')||'');
       if(p)mergeStarts.set(p[1]+p[2],{first:p[1],last:p[3],endRow:+p[4]});
     }
     const ctx=(typeof document!=='undefined')?document.createElement('canvas').getContext('2d'):null;
     const countLines=(source,pixels,fontSize,slack)=>{
       const parts=String(source||'').replace(/\r/g,'').split('\n');
       if(!ctx)return parts.reduce((sum,line)=>sum+Math.max(1,Math.ceil(line.length/(pixels/7))),0);
       ctx.font=fontSize+'pt Arial';let total=0;
       for(const part of parts){
         let used=0;total++;
         if(!part.trim())continue;
         for(const word of part.trim().split(/\s+/)){
           const width=ctx.measureText(word).width;
           const space=ctx.measureText(' ').width;
           if(used>0&&used+space+width>pixels*slack){total++;used=0;}
           if(width>pixels*slack){
             if(used){total++;used=0;}
             total+=Math.ceil(width/(pixels*slack))-1;
             used=width%(pixels*slack)||pixels*slack;
           }else used=used?used+space+width:width;
         }
       }
       return total;
     };
     const sectionV=sheetRows.find(row=>/^V\.\s/.test(textOf(at(row,'A')).trim()))?.getAttribute('r');
     const beforeV=sectionV?Number(sectionV):999;
     let rowsTuned=0;const samples=[];
     for(const row of sheetRows){
       const rn=+row.getAttribute('r');
       if(rn<11||rn>=beforeV||rn>310||row.getAttribute('hidden')==='1')continue;
       const old=Number(row.getAttribute('ht')??13.8);
       if(old===0)continue;
       const a=textOf(at(row,'A')).trim();
       if(/^\s*(?:I|II|III|IV|V)\.\s/.test(a))continue;
       if(![...cells].some(col=>textOf(at(row,col)).trim()))continue;
       let neededLines=1;
       for(const col of cells){
         const cell=at(row,col),text=textOf(cell).trim();if(!text)continue;
         const merge=mergeStarts.get(col+rn);
         if(merge&&merge.endRow>rn)continue; // height is shared across merged block
         const start=cells.indexOf(col),end=merge?cells.indexOf(merge.last):start;
         const pixelWidth=cells.slice(start,end+1).reduce((sum,k)=>sum+Math.floor(widths[k]*7+5),0)-10;
         const xf=xfs[Number(cell.getAttribute('s')||0)]||xfs[0];
         const font=fonts[Number(xf.getAttribute('fontId')||0)]||fonts[0];
         const size=Math.max(9,Math.min(12,Number(all(font,'sz')[0]?.getAttribute('val')||11)));
         // Excel's text metrics differ slightly from browser canvas. Apply
         // small font-width slack to short A/B labels to avoid false wraps.
         const slack=col==='C'?1.02:1.17;
         neededLines=Math.max(neededLines,countLines(text,pixelWidth,size,slack));
       }
       const target=Math.max(20,Math.min(330,Math.ceil(neededLines*14.7+3)));
       // Only compact oversized rows, do not undo intentional custom Feature heights.
       if(old>target+1.5){
         row.setAttribute('ht',String(target));row.setAttribute('customHeight','1');
         rowsTuned++;
         if(samples.length<25)samples.push([rn,old,target,neededLines]);
       }
     }
     // v5.8.240 — Excel-only: prevent Hall indicators text from overlapping
     // the following floor block. Earlier compact passes reduced tall C cells
     // to 77pt even when Vietnamese, English, and quantity occupy >=5 visual
     // lines. Always perform this FINAL fit after every I-IV shrink pass.
     // No row insertion, border change, or PDF formatting is involved.
     let mrHallRowsFitted=0;
     const mrHallFitSamples=[];
     if(model==='MR'){
       const headingIV=sheetRows.find(rr=>/^IV\.\s/.test(textOf(at(rr,'A')).trim()));
       const afterIV=headingIV?Number(headingIV.getAttribute('r')):0;
       const availablePx=Math.floor(widths.C*7+5)-24;
       for(const row of sheetRows){
         const rn=Number(row.getAttribute('r'));
         if(rn<=afterIV||rn>=beforeV||row.getAttribute('hidden')==='1')continue;
         const old=Number(row.getAttribute('ht')??13.8);
         if(old<1)continue; // keep collapsed continuation rows collapsed
         const c=at(row,'C'),desc=textOf(c).trim();
         if(!/Hiển thị\s*&\s*nút gọi tầng|Hall indicators\s*&\s*buttons/i.test(desc))continue;
         // Arial 11pt, wrapping in the ACTUAL width of column C. Conservative
         // 6% width reserve protects native Excel layout vs browser canvas.
         const lines=countLines(desc,availablePx,11,0.94);
         const fitted=Math.min(320,Math.max(24,Math.ceil(lines*17.2+10)));
         if(fitted>old+0.5){
           row.setAttribute('ht',String(fitted));
           row.setAttribute('customHeight','1');
           mrHallRowsFitted++;
         }
         mrHallFitSamples.push({row:rn,lines,from:old,to:Math.max(old,fitted)});
       }
     }
     // MR row 32 is a technical CWT rail designation, not a bilingual
     // translation. Reuse the plain black, non-italic styles of row 31.
     let plainRail=false;
     if(model==='MR'){
       const row31=sheetRows.find(r=>+r.getAttribute('r')===31);
       const row32=sheetRows.find(r=>+r.getAttribute('r')===32);
       if(row31&&row32){
         for(const col of ['B','C']){
           const source=at(row31,col),target=at(row32,col);
           if(source&&target){target.setAttribute('s',source.getAttribute('s')||'0');plainRail=true;}
         }
       }
     }
     if(typeof window!=='undefined')window.__SPEC_EXPORT_240_AUDIT__={model,rowsTuned,samples,plainRail,mrHallRowsFitted,mrHallFitSamples};
   }
   // v5.8.241 — Excel-only: fix detailed MRL CBV bilingual text in column C.
   // Detail mode used to append `(Model CBV..., button diameter Ø33 mm)`
   // to the Vietnamese sentence as a black, non-italic run.  Rewrite only
   // those malformed cells as genuine in-cell bilingual rich text; do not
   // insert rows, alter merges, change standard exports, or modify PDF.
   {
     let mrlCopBilingualFixed=0;
     const correctedRows=[];
     if(model==='MRL'){
       const sd=parse(await zip.file('xl/styles.xml').async('string'));
       const xfs=all(all(sd,'cellXfs')[0],'xf');
       const fonts=all(all(sd,'fonts')[0],'font');
       const tableRows=all(x,'row');
       const sharedPart=zip.file('xl/sharedStrings.xml');
       const shared=sharedPart?all(parse(await sharedPart.async('string')),'si').map(si=>all(si,'t').map(t=>t.textContent||'').join('')):[];
       const cAt=(row,col)=>Array.from(row?.children||[]).find(c=>c.localName==='c'&&c.getAttribute('r')===col+row.getAttribute('r'))||null;
       const txt=c=>!c?'':c.getAttribute('t')==='s'?(shared[+(all(c,'v')[0]?.textContent||0)]||''):
         c.getAttribute('t')==='inlineStr'?all(c,'t').map(t=>t.textContent||'').join(''):(all(c,'v')[0]?.textContent||'');
       const ctx=typeof document!=='undefined'?document.createElement('canvas').getContext('2d'):null;
       function physicalLines(source,pixels){
         const paragraphs=String(source||'').replace(/\r/g,'').split('\n');
         if(!ctx)return paragraphs.reduce((v,t)=>v+Math.max(1,Math.ceil(t.length/(pixels/7))),0);
         ctx.font='11pt Arial';const space=ctx.measureText(' ').width;
         let lines=0;
         for(const paragraph of paragraphs){
           lines++;let used=0;
           for(const word of paragraph.trim().split(/\s+/)){
             if(!word)continue;
             const px=ctx.measureText(word).width;
             if(used>0&&used+space+px>pixels){lines++;used=0;}
             if(px>pixels){
               if(used){lines++;used=0;}
               lines+=Math.ceil(px/pixels)-1;used=px%pixels||pixels;
             }else used=used?used+space+px:px;
           }
         }
         return lines;
       }
       for(const row of tableRows){
         if(row.getAttribute('hidden')==='1')continue;
         const c=cAt(row,'C'),source=txt(c);
         if(!c||!/\bCBV1?-/i.test(source)||!/đường kính nút bấm\s*Ø33\s*mm/i.test(source))continue;
         const [firstLine,...remaining]=source.replace(/\r/g,'').split('\n');
         const parts=/^(.+?)\s+\((Model\s+CBV[^\n]+?)\)\.?\s*$/i.exec(firstLine.trim());
         if(!parts)continue;
         const vi=parts[1].trim();
         const en=parts[2].trim().replace(/[.\s]+$/,'')+'.';
         const tail=remaining.join('\n').trim();
         const base=xfs[Number(c.getAttribute('s')||0)]||xfs[0];
         const font=fonts[Number(base?.getAttribute('fontId')||0)]||fonts[0];
         const size=all(font,'sz')[0]?.getAttribute('val')||'11';
         const inline=x.createElementNS(n,'is');
         function appendRun(content,isEnglish){
           if(!content)return;
           const run=x.createElementNS(n,'r'),pr=x.createElementNS(n,'rPr');
           const family=x.createElementNS(n,'rFont');family.setAttribute('val','Arial');pr.appendChild(family);
           const sz=x.createElementNS(n,'sz');sz.setAttribute('val',size);pr.appendChild(sz);
           if(isEnglish)pr.appendChild(x.createElementNS(n,'i'));
           const color=x.createElementNS(n,'color');color.setAttribute('rgb',isEnglish?'FF245A86':'FF000000');pr.appendChild(color);
           const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');t.textContent=content;
           run.append(pr,t);inline.appendChild(run);
         }
         appendRun(vi+'\n',false);
         appendRun(en+(tail?'\n':''),true);
         appendRun(tail,false);
         for(const child of Array.from(c.children).filter(el=>['v','f','is'].includes(el.localName)))child.remove();
         c.setAttribute('t','inlineStr');c.appendChild(inline);
         // The corrected English is on its own line.  Fit the actual
         // displayed C text after all earlier compact-row passes.
         const width=Math.floor(70.5546875*7+5)-24;
         const lineCount=physicalLines([vi,en,tail].filter(Boolean).join('\n'),width);
         const current=Number(row.getAttribute('ht')??13.8);
         const fitted=Math.max(24,Math.ceil(lineCount*17.2+9));
         if(current<fitted){row.setAttribute('ht',String(fitted));row.setAttribute('customHeight','1');}
         correctedRows.push({row:Number(row.getAttribute('r')),lines:lineCount,from:current,to:Math.max(current,fitted)});
         mrlCopBilingualFixed++;
       }
     }
     if(typeof window!=='undefined')window.__SPEC_EXPORT_241_AUDIT__={model,mrlCopBilingualFixed,correctedRows};
   }
   // v5.8.243: In-cell Rich Text for the two Control descriptions, applied
   // after ALL legacy normalization and row-height shrink passes. This keeps
   // English blue italic on its own line in native Excel and in the PDF print.
   {
     const sd=parse(await zip.file('xl/styles.xml').async('string'));
     const xfs=all(all(sd,'cellXfs')[0],'xf');
     const fonts=all(all(sd,'fonts')[0],'font');
     const tableRows=all(x,'row');
     const ctx=(typeof document!=='undefined')?document.createElement('canvas').getContext('2d'):null;
     const pixelWidth=Math.floor(70.5546875*7+5)-22;
     const visualLines=(body,fontSize)=>{
       if(ctx)ctx.font=fontSize+'pt Arial';
       const measure=word=>ctx?ctx.measureText(word).width:word.length*fontSize*.64;
       const space=measure(' ');
       let lines=0;
       for(const phrase of String(body).split('\n')){
         lines++;let used=0;
         for(const word of phrase.trim().split(/\s+/)){
           if(!word)continue;
           const w=measure(word);
           if(used&&used+space+w>pixelWidth){lines++;used=0;}
           if(w>pixelWidth){if(used){lines++;used=0;}
             lines+=Math.ceil(w/pixelWidth)-1;used=w%pixelWidth||pixelWidth;
           }else used=used?used+space+w:w;
         }
       }
       return lines;
     };
     const fitted=[];
     for(const address of (model==='MRL'?['C26','C35']:['C35'])){
       // Look up the ORIGINAL cell; final physical C35 is shifted +2.
       const rawAddress=address==='C26'?'C25':address==='C35'?(model==='MR'?'C33':'C32'):address;
       const source=String(doc.cells[rawAddress]||'');
       if(!source.includes('\n'))continue;
       const chunks=source.split('\n');
       const groupTriplet=address==='C26'&&chunks.length>=3;
       const en=groupTriplet?chunks.slice(2).join('\n'):chunks.slice(1).join('\n');
       if(!chunks[0]||!en)continue;
       const contentRuns=groupTriplet?
         [[chunks[0]+'\n',false],[chunks[1]+'\n',false],[en,true]]:
         [[chunks[0]+'\n',false],[en,true]];
       const cell=all(x,'c').find(c=>c.getAttribute('r')===address);
       if(!cell)throw Error('Control bilingual export cell missing: '+address);
       const row=tableRows.find(r=>+r.getAttribute('r')===Number(address.slice(1)));
       if(!row)throw Error('Control bilingual export row missing: '+address);
       const oldStyle=xfs[Number(cell.getAttribute('s')||0)]||xfs[0];
       const font=fonts[Number(oldStyle?.getAttribute('fontId')||0)]||fonts[0];
       const size=Number(all(font,'sz')[0]?.getAttribute('val')||11);
       for(const child of Array.from(cell.children).filter(e=>['v','is','f'].includes(e.localName)))child.remove();
       cell.setAttribute('t','inlineStr');
       const inline=x.createElementNS(n,'is');
       for(const [content,english] of contentRuns){
         const run=x.createElementNS(n,'r'),pr=x.createElementNS(n,'rPr');
         const family=x.createElementNS(n,'rFont');family.setAttribute('val','Arial');pr.appendChild(family);
         const sz=x.createElementNS(n,'sz');sz.setAttribute('val',String(size));pr.appendChild(sz);
         if(english)pr.appendChild(x.createElementNS(n,'i'));
         const color=x.createElementNS(n,'color');color.setAttribute('rgb',english?'FF245A86':'FF000000');pr.appendChild(color);
         const t=x.createElementNS(n,'t');t.setAttribute('xml:space','preserve');t.textContent=content;
         run.append(pr,t);inline.appendChild(run);
       }
       cell.appendChild(inline);
       const newXf=oldStyle.cloneNode(true);
       let alignment=Array.from(newXf.children).find(e=>e.localName==='alignment');
       if(!alignment){alignment=sd.createElementNS(n,'alignment');newXf.appendChild(alignment);}
       alignment.setAttribute('wrapText','1');alignment.setAttribute('vertical','center');
       newXf.setAttribute('applyAlignment','1');
       const stylesRoot=all(sd,'cellXfs')[0];const id=all(stylesRoot,'xf').length;
       stylesRoot.appendChild(newXf);stylesRoot.setAttribute('count',String(id+1));
       cell.setAttribute('s',String(id));
       // Measure actual display lines, then reserve a small safety margin.
       const lines=visualLines(source,size);
       const desired=Math.min(310,Math.ceil(lines*size*1.40+7));
       const previous=Number(row.getAttribute('ht')||13.8);
       if(desired>previous){row.setAttribute('ht',String(desired));row.setAttribute('customHeight','1');}
       fitted.push({address,lines,height:Math.max(previous,desired)});
     }
     zip.file('xl/styles.xml',serial(sd));
     if(typeof window!=='undefined')window.__SPEC_EXPORT_243_AUDIT__={model,controlCells:fitted};
   }
   // v5.8.244: The standard export has an unused source spacer at physical
   // row 37, immediately before section II. Collapse only this empty row;
   // detailed Export uses row 37 for actual technical data and is untouched.
   if(!doc.includeTechDetails){
     const sheetData=all(x,'sheetData')[0];
     let spacer=all(x,'row').find(row=>row.getAttribute('r')==='37');
     if(!spacer){
       spacer=x.createElementNS(n,'row');spacer.setAttribute('r','37');
       const later=Array.from(sheetData.children).find(row=>row.localName==='row'&&Number(row.getAttribute('r'))>37);
       sheetData.insertBefore(spacer,later||null);
     }
     // Non-empty content here would indicate a new shift bug; fail visibly
     // rather than hiding customer-facing text beneath a section heading.
     if(Array.from(spacer.children).some(c=>c.localName==='c'&&(all(c,'t').some(t=>t.textContent.trim())||all(c,'v').some(v=>v.textContent.trim()))))
       throw Error('Unexpected content in pre-section II spacer row 37');
     spacer.setAttribute('ht','0');spacer.setAttribute('customHeight','1');spacer.setAttribute('hidden','1');
   }
   // Section II is merged A:C: its right-hand cells must remain empty.
   // Stale control-system text in C38 previously expanded the entire heading.
   // A position-independent check makes the guard safe when optional rows move.
   const iiHeading=all(x,'row').find(row=>Array.from(row.children).some(c=>c.localName==='c'&&/^A\d+$/.test(c.getAttribute('r')||'')&&
     /^II\.\s*Các kích thước liên quan/.test(all(c,'t').map(t=>t.textContent||'').join(''))));
   if(iiHeading){
     const rn=iiHeading.getAttribute('r');
     const c=Array.from(iiHeading.children).find(cell=>cell.localName==='c'&&cell.getAttribute('r')==='C'+rn);
     if(c&&(all(c,'t').some(t=>t.textContent.trim())||all(c,'v').some(v=>v.textContent.trim())))
       throw Error('Unexpected content in section II merged C'+rn);
   }
   if(doc.onlyVi)await vietnameseSheet(zip,x,3);
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
   const re=/\((?:[A-Za-z0-9][^()\r\n]*|[A-Za-z0-9][^()\r\n]*\([^()]*\)[^()\r\n]*)\)/g;
   const out=[];let last=0,m;
   const englishPhrase=(part,preceding)=>{
     const inside=part.slice(1,-1).trim();
     // Only classify explanatory text after an actual Vietnamese description.
     // Covers short translations (No), capacities (8 persons), quantity
     // (1 set), and speed (60 mpm - 1 m/s), which the old whitelist missed.
     if(!/(?:[À-ỹ]|\bthang\b|\bhop\b)/i.test(preceding)||/[À-ỹ]/.test(inside))return false;
     if(!/[A-Za-z]/.test(inside)||inside.length<2)return false;
     if(/^\d{2,5}\s*(?:kg|mm|m)$/i.test(inside)||/^[A-Z][A-Z0-9-]{1,7}$/.test(inside))return false;
     return true;
   };
   while((m=re.exec(raw))){
     const segment=m[0];
     if(!englishPhrase(segment,raw.slice(0,m.index)))continue;
     const prefix=raw.slice(last,m.index);
     if(prefix)out.push({text:prefix,english:false});
     const earlier=raw.slice(0,m.index);
     const currentLine=earlier.slice(earlier.lastIndexOf('\n')+1);
     // Newline if English text follows Vietnamese on the same physical line.
     if(/[À-ỹ]/.test(currentLine)&&currentLine.trim() && !/\n$/.test(prefix))out.push({text:'\n',english:false});
     out.push({text:segment.slice(1,-1).trim(),english:true});last=m.index+segment.length;
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
   @page{size:A4 portrait;margin:12mm 10mm 13mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;font-size:9pt;color:#1b1b1b;margin:0;padding:0}header{margin:0 0 12px;border-bottom:2px solid #b5161c;padding:7px 0 10px}header strong{font-size:13pt;color:#10345c}header small{display:block;margin-top:3px}table{border-collapse:collapse;width:100%;table-layout:fixed}td{border:1px solid #777;padding:4px 5px;vertical-align:top;white-space:pre-line;overflow-wrap:anywhere;line-height:1.33}td:nth-child(1){width:31%;font-weight:600}td:nth-child(2){width:21%}td:nth-child(3){width:48%}tr.section td{font-weight:700;background:#e9edf2;color:#132e4c;break-after:avoid}tr.subsection td{font-weight:700;background:#f3f4f5;break-after:avoid}tr.feature td:first-child{font-weight:500}tr{break-inside:avoid}.spec-english,em{color:#245A86;font-style:italic}.toolbar{position:sticky;top:0;padding:12px;background:#f7f8fa;border-bottom:1px solid #aaa;margin-bottom:12px;font-family:Arial,sans-serif}.toolbar button{padding:9px 16px;background:#ad1119;color:white;border:0;border-radius:5px;cursor:pointer}.toolbar small{margin-left:12px;color:#555}footer{color:#666;font-size:8pt;padding-top:8px} @media print{.toolbar{display:none}header{margin:0 0 10px}footer{display:none}} 
   </style></head><body><div class="toolbar"><button onclick="window.print()">In / Save as A4 PDF</button><small>Chọn máy in “Save as PDF” trong hộp thoại / Select Save as PDF.</small></div><header><strong>MITSUBISHI ELECTRIC VIETNAM — SPECIFICATION</strong><small>${escape(title)} | ${escape(v('B8'))} — ${escape(v('B9'))}</small></header><table><tbody>${rows.join('')}</tbody></table><footer>Source: SpecFormExport | A4 portrait | Excel Database</footer></body></html>`;
 }
 window.SPEC_EXPORT={vietnameseSheet,exportModel,xlsxFromModel,printHtml,attachMRMaterialThickness};
})();
