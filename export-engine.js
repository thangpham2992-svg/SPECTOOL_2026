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
   const cells={},hiddenRows=new Set(),overrides={},sections={},unresolved=[];
   const output=db.sheets.SpecFormExport;
   for(const [addr,entry] of Object.entries(output)){
     if(entry&&typeof entry==='object'&&Object.hasOwn(entry,'f')){
       const result=eng.safe('SpecFormExport',addr);
       if(result.ok)cells[addr]=result.value;
       else {cells[addr]='';unresolved.push({addr,formula:entry.f,error:result.error})}
     }else cells[addr]=entry?.v??entry??'';
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
       // C# clears helper columns D/E and highlights nonstandard features.
       for(const er of featureRows){cells['D'+er]='';cells['E'+er]='';}
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
   // Per-component material override: special glass door assemblies are SUS304,
   // independent of the global SUS Standard selection. Never promote Jamb,
   // Transom or Front return to SUS304 unless the global SUS304 option is chosen.
   // The source formulas already handle DSJ1 and Large window; include DS1 here.
   if(model==='MRL'){
     const read=addr=>text(eng.safe('SpecForm',addr).value);
     for(const [input,exportRows] of [['D46',[49,50]],['E73',[72,73]],['E74',[74,75]]]){
       const selected=read(input);
       if(window.SPEC_NY40_RULES?.isGlassDoor(selected) && /^SUS-HL/i.test(selected)){
         for(const row of exportRows){const key='C'+row;
           cells[key]=text(cells[key]).replace(/SUS430J1L|SUS430/gi,'SUS304');
         }
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
   // v4.2: Remove the legacy decorative-weight exclusion note from any export.
   for(const addr of Object.keys(cells))if(typeof cells[addr]==='string')cells[addr]=cells[addr]
     .replace(/\s*[-–]?\s*Trừ\s*:\s*Handrail\s*&\s*mirror\s*nhập khẩu\s*;\s*sàn\s*\(By local\)/gi,'');
   // MR source has two historical MBSS #N/A lookups on feature rows that are removed
   // when MBSS does not apply; only unresolved visible output cells block export.
   const fatal=unresolved.filter(e=>!hiddenRows.has(cellRow(e.addr))&&/^[ABC]/.test(e.addr));
   if(fatal.length)throw Error('Công thức Export chưa tính được / Cannot evaluate '+fatal.slice(0,5).map(x=>x.addr+' '+x.error).join(' | '));
   return {model,cells,hiddenRows,removedFeatures,includedFeatures,sections,hiddenHall,hallLanternExportRows,totalFormulaUnresolved:unresolved.length,fatal};
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
     addRailRow(31,'Đối trọng (CWT)',false);
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
   return {zip,freezeCount,hiddenRows:doc.hiddenRows.size,filename:doc.model};
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
     rows.push('<tr class="'+rowClass+'"><td>'+escape(a)+'</td><td>'+escape(b)+'</td><td>'+escape(c)+'</td></tr>');
   }
   let title=model==='MR'?'NEXIEZ-MR ALPHA':'NEXIEZ-MRL Version02 ALPHA';
   return `<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>SPEC ${title} - A4 PDF</title><style>
   @page{size:A4 portrait;margin:12mm 10mm 13mm}*{box-sizing:border-box}body{font-family:Arial,"Times New Roman",sans-serif;font-size:9pt;color:#1b1b1b;margin:0;padding:0}header{margin:0 0 12px;border-bottom:2px solid #b5161c;padding:7px 0 10px}header strong{font-size:13pt;color:#10345c}header small{display:block;margin-top:3px}table{border-collapse:collapse;width:100%;table-layout:fixed}td{border:1px solid #777;padding:4px 5px;vertical-align:top;white-space:pre-line;overflow-wrap:anywhere;line-height:1.33}td:nth-child(1){width:31%;font-weight:600}td:nth-child(2){width:21%}td:nth-child(3){width:48%}tr.section td{font-weight:700;background:#e9edf2;color:#132e4c;break-after:avoid}tr.subsection td{font-weight:700;background:#f3f4f5;break-after:avoid}tr.feature td:first-child{font-weight:500}tr{break-inside:avoid}.toolbar{position:sticky;top:0;padding:12px;background:#f7f8fa;border-bottom:1px solid #aaa;margin-bottom:12px;font-family:Arial,sans-serif}.toolbar button{padding:9px 16px;background:#ad1119;color:white;border:0;border-radius:5px;cursor:pointer}.toolbar small{margin-left:12px;color:#555}footer{color:#666;font-size:8pt;padding-top:8px} @media print{.toolbar{display:none}header{margin:0 0 10px}footer{display:none}} 
   </style></head><body><div class="toolbar"><button onclick="window.print()">In / Save as A4 PDF</button><small>Chọn máy in “Save as PDF” trong hộp thoại / Select Save as PDF.</small></div><header><strong>MITSUBISHI ELECTRIC VIETNAM — SPECIFICATION</strong><small>${escape(title)} | ${escape(v('B8'))} — ${escape(v('B9'))}</small></header><table><tbody>${rows.join('')}</tbody></table><footer>Source: SpecFormExport | A4 portrait | Excel Database</footer></body></html>`;
 }
 window.SPEC_EXPORT={exportModel,xlsxFromModel,printHtml};
})();
