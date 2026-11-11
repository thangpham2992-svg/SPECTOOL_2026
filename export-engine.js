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
 const parse=s=>new DOMParser().parseFromString(s,'application/xml');
 const serial=d=>new XMLSerializer().serializeToString(d);
 const all=(node,tag)=>Array.from(node.getElementsByTagNameNS(n,tag));
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
   // Supplemental features present in Manual Version02 but absent from Version01 SpecForm table.
   // They are emitted in a separate clearly identified appendix; not silently discarded.
   if(model==='MRL' && window.SPEC_V52_FEATURES){
     const v52=window.SPEC_V52_FEATURES;
     const op=text(eng.safe('SpecForm','D22').value);
     const extras=v52.virtual(model,db);
     let pos=312,started=false;
     for(const f of extras){
       const st=v52.status(model,f.code,op,db);
       const enabled=Object.hasOwn(changes,f.cell)?on(changes[f.cell]):st==='S';
       if(st==='X'||!enabled)continue;
       if(!started){cells['A'+pos]='VI. Tính năng bổ sung theo Manual Version02 / Additional features';pos++;started=true;}
       cells['A'+pos]='('+f.code+') '+f.en;
       cells['C'+pos]='Feature listed in NEXIEZ-MRL Alpha Package Version02 manual. Application: '+(st==='S'?'Standard':'Available option')+'.';
       pos++;
       cells['A'+pos]=f.vi;
       cells['C'+pos]='Tính năng theo Manual NEXIEZ-MRL Version02. Trạng thái: '+(st==='S'?'Tiêu chuẩn.':'Tùy chọn.');
       includedFeatures.push({group:f.table.replace('TblFeature_',''),code:f.code,row:null,option:st,exportRows:[pos-1,pos]});
       pos++;
     }
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
   const workbook=parse(await zip.file('xl/workbook.xml').async('string'));
   const rels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
   const sheetnodes=all(workbook,'sheet');
   const sheet=sheetnodes.find(x=>x.getAttribute('name')==='SpecFormExport');
   if(!sheet)throw Error('Không có SpecFormExport / Missing sheet');
   const exportRid=sheet.getAttributeNS(r,'id')||sheet.getAttribute('r:id');
   const relationships=Array.from(rels.documentElement.children);
   const ref=relationships.find(x=>x.getAttribute('Id')===exportRid);
   if(!ref)throw Error('No workbook relationship for Export sheet');
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
     .filter(([addr,value])=>/^A\d+$/.test(addr)&&cellRow(addr)<=300&&!doc.hiddenRows.has(cellRow(addr))&&sectionTitle.test(text(value)))
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
