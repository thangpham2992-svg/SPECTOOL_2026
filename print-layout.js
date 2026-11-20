'use strict';
/* Print from exported XLSX OOXML rather than resynthesizing a three-column table.
   Uses the workbook's SpecFormExport cell styles, column proportions, merged ranges,
   row heights, and export-engine-hidden feature rows. Runs fully offline in browser. */
(()=>{
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 // Exact Mitsubishi Electric logo embedded in the MRL Version02 source workbook.
 // Embed it in the print document to keep PDF self-contained/offline (MR fallback).
 const FALLBACK_LOGO='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAQQAAABzCAYAAAB+WHzHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAIdUAACHVAQSctJ0AAEBMSURBVHhe7d0LvG1T9QfwSlGhECGlQgk9kFQiIVGiJI+iJCFRSQm9UK48k6QiFFei6CWJniryiAoliqRUSsQ/FNX8z+9adxzLumuf19773n01x+czP+ecveeajzHH+I3HnHOdB6VChQoVmkUFEAoVKjRGBRAKFSo0RgUQChUqNEYFEAoVKjRGBRAKFSo0RgUQChUqNEYFEAoVKjRGDxxA+O9/6/Kf/9TF74UKFZoSPTAAAQDcfXdKt9+e0s031+Wuu2Z9WahQocnSAwMQ/v3vlP72t5SuvTalSy9N6ZJLUrrllllfFipUaLL0wACEf/4zpcsuS+mzn03psMNSOuaYlK65ZtaXhQoVmizN24AgT8A7uO22lE47LaU3vCGljTdOabvtUjr//JTuuacOJwoVKjQpmrcBgbL/5S8pXXRRSvvum9IznpHSssumtOaaKR17bB1C3HnnrMqFChWaiOZtQOAdXHVVrfybbJLSAgvkGeUpLbVUSrvvntLXvpbSn/9cdh0KFZokzduAcO+9KZ19dkrbbJPSCiukNN98NSAstFBKa6yR0l57pXTFFfWOQwkdChWakOZdQKDkN92U0lFHpbTKKiktuGANBgpgWGSRlF7ykpTOPDOlG29M6V//mvVgoUKFetG8BwhxAOmPf0zp3HNT2m23lJZcMqWHPOQ+QFAe+tCUVl45pX32Sem88+ozCiVsKFRoXJr3AEGYwDtw3uDAA1Naf/06RGiCQZQllqi9hI98JKUbbqi3J0voUKhQT8paM4+RE4lChc9/vk4kPv7xtTfQBQiSjBKMO+yQ0oUX1oeVJCILFSrUSVlr5jGyzShUePe7U1pxxZQe8YjZgSDKgx+c0vzzp7T22ikdcURKF19cjjQXKjQOZa2Zh4i772zBjBkpbbBBSosuWit9Fxg0i7MJm26a0okn3ucllHxCoUKzUdaWeYQo8R13pPTDH9bbjEIF1r8LANpFjmH55ettyMsvT+nWW0voUKhQB2VtmUdIQvC66+r7CquuOjnPoFke9rCUNtsspVNPrb0MyclChQrdj7KmzCPEqp9ySkqve11Kyywzu8JPVACIbciddkrpnHPq5GShQoXuR1lTRpy49v/3f/UR5T32qBOJzUNIUykOK62+ekpHHlkfVpJgLLmEQoXGKGvJiJNQ4Re/SOnkk1PaaKM6H9A+hDTZImxwiOmNb0zpS1+qzyYUQChUaIyylow48Q5cUtp119o76FL0qRRnE573vJT23julCy6ojzSXw0qFClWUNWREieVWbBMeemh9tdk2Y5eST6XwLhZbLKV11qkTjI40e2/CXKL/5jkq/86hUbv8JwOV4vsu8rnvu55V4vl2na7PplI8G+NutnPvvffOVrqe6Xq2XadN8bl6zbY9025nMqVrPNF+s3TVif6a82yOSelqvz3Odp0u6nomSvvzXm1MhbKGjChl5lb3Fb773drFBwZc/i4ln2pxsvHJT07pve9N6cc/rl+/NpfIIt50003p29/+dvrqV786Vr6WvaIf57Fdd9112UnKXlIH3ZOBzPfnn39++sY3vjH27De/+c104YUXpquuuipdf/316ZJLLklnnXVW+vrXv56+853vpF/+8pfp6quvTuedd979+pyoaPdHP/pRjrRuyJHcP9Ntt92WfvOb36Tvf//76YwzzkgnnXRSOu6449InPvGJ9OlPfzrNnDmzeu4HP/hBNY477rijGjPh/Vf2zDx79tlnV3XM94orrshL8beq7TbhwR/+8IeqLe0qX8phn7n8+te/no1/vQoeqPvTn/40i9cf05133lmNx7j0jW9R1+9//vOf0913310pe4zh4osvrr43BvP95Cc/mT71qU+l448/vuIDfhhTzJfC+olv5557bvWs9bj88svTX/7yl3TXXXf1VOZrrrmmqusZ/Lduv/vd79Jll11WraXv9PeLHFbfMoDXBmbtGFEiFI4bf+hDKa211vTzBr0KgHFYyW3JLJhziwgjwds1h0SveMUrxsqrX/3q9MEPfrASBELYJs8R0K985SvpHe94R9puu+3Gnt1hhx3S/vvvXykohTv44IPTlltumbbZZpv01re+NZ1yyinZOTo17bzzzvfrc6Ky/fbbp/e85z2VYP71r39NV155ZTr99NPTHnvskTbccMO06qqrpuWWWy4ts8wy6SlPeUp6znOek175ylemd77znenMM8+slOT27JEBAz8J87bbblu1bb6Ui4IDmiZRFsqLT/pfY401qrJpXj9zwYO3vOUts423q2y11VZp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';
 const OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
 // Safari, Chrome and other browsers on iPhone/iPad all use the iOS WebKit printing engine.
 // Keep the original print stylesheet completely unchanged on PC and Android.
 const isIOS=(()=>{const ua=navigator.userAgent||'';return /iPad|iPhone|iPod/i.test(ua)||(/Macintosh/i.test(ua)&&navigator.maxTouchPoints>1)})();
 const q=(node,n)=>Array.from(node?.childNodes||[]).find(x=>x.localName===n)||null;
 const qa=(node,n)=>Array.from(node?.getElementsByTagNameNS(ns,n)||[]);
 const attr=(node,n,defaultValue='')=>node?.getAttribute(n)??defaultValue;
 const esc=v=>String(v??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
 const colNo=s=>[...String(s)].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
 const colAddr=n=>{let s='';while(n>0){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26)}return s};
 const addrParts=a=>{const m=/^([A-Z]+)(\d+)$/.exec(a||'');return m?{col:colNo(m[1]),row:+m[2]}:null};
 const parse=str=>new DOMParser().parseFromString(String(str).replace(/^\uFEFF/, '').replace(/^\s+(?=<\?xml\b)/, ''),'application/xml');
 function color(node,defaultColor){if(!node)return defaultColor;const rgb=attr(node,'rgb');if(/^[0-9a-f]{8}$/i.test(rgb))return '#'+rgb.slice(2);if(/^[0-9a-f]{6}$/i.test(rgb))return '#'+rgb;
   const theme=attr(node,'theme');if(theme!=='')return ({0:'#ffffff',1:'#1b2635',2:'#f6f8fa',3:'#24364a',4:'#22587d',5:'#b1222a'})[Number(theme)]||defaultColor;
   return defaultColor;
 }
 function styleBook(root){const s={};const fonts=q(root,'fonts'),fills=q(root,'fills'),borders=q(root,'borders'),cellxfs=q(root,'cellXfs');
   s.fonts=Array.from(fonts?.children||[]).map(f=>({size:Number(attr(q(f,'sz'),'val',9))||9,bold:!!q(f,'b'),italic:!!q(f,'i'),color:color(q(f,'color'),'#1e2a36'),name:attr(q(f,'name'),'val','Arial')}));
   s.fills=Array.from(fills?.children||[]).map(f=>color(q(q(f,'patternFill'),'fgColor'),'transparent'));
   s.borders=Array.from(borders?.children||[]).map(b=>{const sides={};for(const n of ['left','right','top','bottom']){const side=q(b,n);const st=attr(side,'style');if(st&&st!=='none')sides[n]=(st==='medium'?1.2:0.45)+'px solid '+color(q(side,'color'),'#a9b4c0');}return sides});
   s.xfs=Array.from(cellxfs?.children||[]).map(x=>({font:s.fonts[Number(attr(x,'fontId',0))]||{},fill:s.fills[Number(attr(x,'fillId',0))],border:s.borders[Number(attr(x,'borderId',0))]||{},align:q(x,'alignment')}));return s;
 }
 function formatStyle(index,book,uniformBorders=false){const x=book.xfs[index]||{};const f=x.font||{};const rules=[];
   if(f.bold)rules.push('font-weight:700');if(f.italic)rules.push('font-style:italic');
   rules.push('font-size:'+Math.max(7.2,Math.min(13,Number(f.size||9)))+'pt');
   if(f.color)rules.push('color:'+f.color);if(x.fill&&x.fill!=='transparent'&&x.fill!=='#000000')rules.push('background:'+x.fill);
   if(x.align){const a=attr(x.align,'horizontal');const v=attr(x.align,'vertical');if(['center','right','left','justify'].includes(a))rules.push('text-align:'+a);if(v==='center')rules.push('vertical-align:middle');if(attr(x.align,'wrapText')==='1')rules.push('overflow-wrap:break-word');}
   if(x.fill&&/^#[0-9a-f]{6}$/i.test(x.fill)){const v=x.fill.slice(1);const rr=parseInt(v.slice(0,2),16),gg=parseInt(v.slice(2,4),16),bb=parseInt(v.slice(4,6),16);if((rr*299+gg*587+bb*114)/1000<125)rules.push('color:#ffffff');}
   if(!uniformBorders){const b=x.border||{};for(const key of ['left','right','top','bottom'])if(b[key])rules.push('border-'+key+':'+b[key]);}
   return rules.join(';');
 }
 function readText(c,strings){if(!c)return '';const type=attr(c,'t');if(type==='inlineStr')return qa(q(c,'is'),'t').map(t=>t.textContent).join('');let v=q(c,'v')?.textContent||'';if(type==='s')return strings[Number(v)]||'';if(type==='b')return v==='1'?'TRUE':'FALSE';return v}
 async function fromXlsx(bytes,model){
   const zip=await JSZip.loadAsync(bytes);
   const wb=parse(await zip.file('xl/workbook.xml').async('string'));
   const rels=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
   const sheet=qa(wb,'sheet').find(x=>attr(x,'name')==='Specification')||qa(wb,'sheet')[0];
   if(!sheet)throw Error('Export workbook has no Specification sheet');
   const rid=sheet.getAttributeNS(OFFICE,'id')||attr(sheet,'r:id');
   const rel=Array.from(rels.documentElement.children).find(x=>attr(x,'Id')===rid);
   if(!rel)throw Error('Missing workbook sheet relationship');
   let target=attr(rel,'Target').replace(/^\//,'');if(!target.startsWith('xl/'))target='xl/'+target;
   const data=parse(await zip.file(target).async('string'));
   const styles=styleBook(parse(await zip.file('xl/styles.xml').async('string')).documentElement);
   // Prevent legacy brown/tan fill from appearing in the PDF export.
   for(const xf of styles.xfs){if(['#d2b48c','#c19a6b','#deb887','#a52a2a'].includes(String(xf.fill||'').toLowerCase()))xf.fill='transparent';}
   const strings=zip.file('xl/sharedStrings.xml')?qa(parse(await zip.file('xl/sharedStrings.xml').async('string')),'si').map(si=>qa(si,'t').map(t=>t.textContent).join('')):[];
   const sheetData=q(data.documentElement,'sheetData'),rows=Array.from(sheetData?.children||[]).filter(x=>x.localName==='row');
   const visibleRows=new Map(),cells=new Map();let largest=0;
   // Rows above 300 are usually technical helpers. Only show the deliberate
   // additional-feature appendix (VI.) and its bilingual entries.
   // Include the genuine source/origin and local-material notes even when
   // they fall AFTER the 300th worksheet row. MR puts this note at A300;
   // MRL Version02 uses A305. The previous 300-row cutoff silently lost it
   // in PDF even though the same value was preserved in XLSX.
   const isOriginDisclaimer=t => /Toàn bộ thiết bị thang máy được nhập khẩu|All main components are imported|Some auxiliary parts of the lifts/i.test(String(t||''));
   const isExtraRow=(rn,row)=>{
     if(rn<=300)return true;
     const c=Array.from(row.children).find(el=>el.localName==='c'&&attr(el,'r')==='A'+rn);
     const t=readText(c,strings).trim();
     if(isOriginDisclaimer(t))return true;
     if(rn<312||rn>475)return false;
     return /^VI\.\s|^\([A-Z0-9-]+\)\s|^Tính năng theo Manual|^Hạn chế|^Cáp động/i.test(t)||Boolean(c&&t&&rn>=313);
   };
   for(const row of rows){let rn=Number(attr(row,'r'));if(!rn||!isExtraRow(rn,row)||attr(row,'hidden')==='1')continue;
     visibleRows.set(rn,row);
     for(const c of Array.from(row.children).filter(x=>x.localName==='c')){const a=attr(c,'r'),p=addrParts(a);if(!p||p.col>3)continue;
       const v=readText(c,strings);cells.set(a,{value:v,style:Number(attr(c,'s',0))});if(v.trim())largest=Math.max(largest,rn);
     }
   }
   if(!largest)throw Error('Export SpecFormExport is empty');
   // The PDF must never silently omit source footnotes included in Excel.
   // Validate against ALL non-hidden source notes, not a fixed row index.
   const disclaimerRows=rows.filter(row=>attr(row,'hidden')!=='1' && Array.from(row.children).some(el=>el.localName==='c' && /^A\d+$/.test(attr(el,'r')) && isOriginDisclaimer(readText(el,strings))));
   for(const row of disclaimerRows){
     const rn=Number(attr(row,'r'));
     if(!visibleRows.has(rn))throw Error('PDF export lost source footnote at Excel row '+rn);
   }
   const end=largest;
   // Merge map comes from original Excel print sheet; supports vertical and horizontal merges.
   const merges=new Map(),covered=new Set();
   for(const m of qa(data,'mergeCell')){const ref=attr(m,'ref'),sp=ref.split(':');if(sp.length!==2)continue;const a=addrParts(sp[0]),b=addrParts(sp[1]);if(!a||!b||a.col>3||b.col>3||a.row>end)continue;
     const validRows=[];for(let r=a.row;r<=b.row;r++)if(visibleRows.has(r))validRows.push(r);
     if(!validRows.length||!visibleRows.has(a.row))continue;
     const key=colAddr(a.col)+a.row;merges.set(key,{colspan:b.col-a.col+1,rowspan:validRows.length});
     for(const vr of validRows)for(let col=a.col;col<=b.col;col++){const k=colAddr(col)+vr;if(k!==key)covered.add(k)}
   }
   // The workbook has heading texts in A but often does not define A:C merged cells
   // for individual feature groups. Merge these visual rows in the PDF/print view only.
   // Keep the exported XLSX's native merge structure and all regular data rows intact.
   const groupTitle=/^\s*[-–]\s*Tính năng\b/i;
   for(const [rn] of visibleRows){
     const a=String(cells.get('A'+rn)?.value||'');
     const b=String(cells.get('B'+rn)?.value||'');
     const c=String(cells.get('C'+rn)?.value||'');
     if(!groupTitle.test(a)||b.trim()||c.trim())continue;
     if(merges.has('A'+rn))continue;
     merges.set('A'+rn,{colspan:3,rowspan:1});
     covered.add('B'+rn);covered.add('C'+rn);
   }
   // v3.8: merge Roman-numbered specification section headings for the print view.
   // The blue background is guaranteed by the same export style used for XLSX.
   const romanSection=/^\s*[IVXLCDM]+\.\s+/i;
   const blueRomanRows=new Set();
   for(const [rn] of visibleRows){
     const a=String(cells.get('A'+rn)?.value||'');
     const b=String(cells.get('B'+rn)?.value||'');
     const c=String(cells.get('C'+rn)?.value||'');
     if(!romanSection.test(a))continue;
     blueRomanRows.add(rn);
     if(b.trim()||c.trim()||merges.has('A'+rn)||covered.has('A'+rn))continue;
     merges.set('A'+rn,{colspan:3,rowspan:1});covered.add('B'+rn);covered.add('C'+rn);
   }
   // Improve print legibility of the two original header lines without altering XLSX cells.
   // The company/address text is in A1/A2 and B1/B2 are blank, so span these on paper.
   for(const headrow of [1,2]){if(visibleRows.has(headrow)&&!String(cells.get('B'+headrow)?.value||'').trim()&&!merges.has('A'+headrow)){merges.set('A'+headrow,{colspan:2,rowspan:1});covered.add('B'+headrow)}}
   const cols=q(data.documentElement,'cols');const widths=[20.5546875,25.5546875,70.5546875];
   for(const entry of Array.from(cols?.children||[])){const min=Number(attr(entry,'min')),max=Number(attr(entry,'max')),width=Number(attr(entry,'width'));for(let i=1;i<=3;i++)if(i>=min&&i<=max&&width>0)widths[i-1]=width}
   const sum=widths.reduce((a,b)=>a+b,0);const colgroup=widths.map(w=>'<col style="width:'+(w/sum*100).toFixed(5)+'%">').join('');
   const first=Array.from(cells.entries()).find(([a])=>a==='B8')?.[1]?.value||'';
   const project=esc(first);const modelName=model==='MR'?'NEXIEZ-MR':'NEXIEZ-MRL Version2';
   // The XLSX source stores corporate identity in cells A1/A2 and the logo as
   // an Excel drawing (MRL) rather than a worksheet cell. The prior PDF
   // renderer silently discarded that drawing and left the title cramped.
   const company=String(cells.get('A1')?.value||'').trim();
   const companyAddress=String(cells.get('A2')?.value||'').trim();
   // Document ID and revision intentionally omitted from PDF masthead (both models).
   const titleRow=[5,6,7].find(r=>/ĐẶC TÍNH KỸ THUẬT|Main specifications/i.test(String(cells.get('A'+r)?.value||'')))||5;
   const reportTitle=String(cells.get('A'+titleRow)?.value||'ĐẶC TÍNH KỸ THUẬT (Main specifications)').trim();
   // Recover the original drawing media when it exists; otherwise use the
   // same official mark from the MRL Excel template for the MR PDF as well.
   const logoEntry=zip.file('xl/media/image.png');
   const logoSrc=logoEntry?'data:image/png;base64,'+await logoEntry.async('base64'):FALLBACK_LOGO;
   const masthead='<header class="spec-masthead">'+
     '<div class="spec-masthead-top"><div class="spec-company">'+esc(company)+
     (companyAddress?'<div class="spec-address">'+esc(companyAddress)+'</div>':'')+
     '</div></div>'+
     '<div class="spec-logo-row"><img class="spec-logo" alt="Mitsubishi Electric" src="'+logoSrc+'"></div>'+
     '<h1 class="spec-report-title">'+esc(reportTitle)+'</h1></header>';
   // Detect a bilingual pair from the real exported XLSX border styles across
   // ALL sections. Section III may have a vertically merged label on the left,
   // unlike Section V features where both lines carry their own labels.
   // Section I/II/IV rows that already contain both languages in the SAME cell
   // have no internal border; they retain a very light item separator.
   // One bilingual-row detector shared by MR and MRL Version02.
   // It uses cell CONTENT rather than template border styles: template borders
   // differ between models, and are not a reliable sign of bilingual pairing.
   const bilingualUpper=new Set(),bilingualLower=new Set();
   const hasVietnamese=s=>/[À-ỹ]/.test(String(s||''));
   const isEnglish=s=>/^\s*\(?[A-Za-z]/.test(String(s||''))&&!hasVietnamese(s);
   // Feature names may be plain English (Multi - Beam) or parenthesized with spaces (SDE - One Side).
   for(let rn=1;rn<end;rn++){
     if(!visibleRows.has(rn)||!visibleRows.has(rn+1))continue;
     const upper=String(cells.get('C'+rn)?.value||'').trim();
     const lower=String(cells.get('C'+(rn+1))?.value||'').trim();
     const label=String(cells.get('A'+rn)?.value||'').trim();
     const nextLabel=String(cells.get('A'+(rn+1))?.value||'').trim();
     if(!upper||!lower||romanSection.test(label)||romanSection.test(nextLabel))continue;
     // A Vietnam-first pair normally has an English parenthetical line below.
     // A feature pair follows the source's English-first / Vietnamese-second convention.
     const vnThenEn=hasVietnamese(upper)&&isEnglish(lower)&&/^\(/.test(lower)
       &&!nextLabel&&!String(cells.get('B'+(rn+1))?.value||'').trim();
     const enThenVn=isEnglish(upper)&&hasVietnamese(lower)
       &&isEnglish(label)&&hasVietnamese(nextLabel);
     if(vnThenEn||enThenVn){bilingualUpper.add(rn);bilingualLower.add(rn+1)}
   }
   // Prelude ends at the first real specification section (different row in
   // MR and MRL). Do not strip borders from Manufacturer or Basic Spec rows.
   const firstSectionRow=blueRomanRows.size?Math.min(...blueRomanRows):11;
   const printRows=[],printedCells=new Set();
   for(let rowNum=1;rowNum<=end;rowNum++){
     const row=visibleRows.get(rowNum);if(!row||rowNum<=titleRow)continue;
     const rowCells=[];const isBlank=[1,2,3].every(i=>!String(cells.get(colAddr(i)+rowNum)?.value||'').trim());
     const rowHeight=Number(attr(row,'ht',attr(q(data.documentElement,'sheetFormatPr'),'defaultRowHeight','13.8')))||13.8;
     for(let col=1;col<=3;col++){
       const key=colAddr(col)+rowNum;if(covered.has(key))continue;
       printedCells.add(key);
       const val=cells.get(key)||{value:'',style:0};const merge=merges.get(key)||{colspan:1,rowspan:1};
       const extra=merge.colspan>1?' colspan="'+merge.colspan+'"':'';const vextra=merge.rowspan>1?' rowspan="'+merge.rowspan+'"':'';
       // PDF grid must never inherit irregular workbook border colors/weights.
       // Font, fill and alignment still come from the original XLSX style.
       let style=formatStyle(val.style,styles,true);
       // Explicit inline background prevents theme tint and browser print defaults
       // from suppressing the original pale-blue I/II/III/IV section headers.
       if(blueRomanRows.has(rowNum))style+=';background-color:#d9edf7;color:#173c56;font-weight:700';
       rowCells.push('<td'+extra+vextra+' data-cell="'+key+'" style="'+style+'">'+esc(val.value).replace(/\r\n|\r|\n/g,'<br>')+'</td>');
     }
     // Excel's explicit row heights are respected as minimums; text-heavy wrapped cells can expand.
     const customHeight=attr(row,'ht')!=='';
     const minHeight=Math.max(isBlank?2:0,Math.min(72,rowHeight*1.07));
     const rowValue=[1,2,3].map(i=>String(cells.get(colAddr(i)+rowNum)?.value||'')).join(' ');
     const cl=(rowNum<firstSectionRow?' excel-prelude':'')+(isBlank?' excel-space':'')+(/^(?:I|II|III|IV|V|VI)\.\s|CÁC TÍNH NĂNG|CÁC TÍNH NĂNG|Các tính năng/i.test(rowValue.trim())?' excel-heading':'')+(/Các tính năng \(Features\)/i.test(rowValue)?' excel-features-title':'');
     const pairClass=bilingualUpper.has(rowNum)?' excel-bilingual-upper':bilingualLower.has(rowNum)?' excel-bilingual-lower':'';
     const originNote=isOriginDisclaimer(String(cells.get('A'+rowNum)?.value||''))?' excel-origin-note':'';
     printRows.push('<tr class="excel-row'+cl+pairClass+originNote+'" data-source-row="'+rowNum+'" style="height:'+minHeight.toFixed(1)+'px">'+rowCells.join('')+'</tr>');
   }
   // Parity guard: any visible, non-empty customer-facing source cell in
   // columns A:C must be represented in the print table. Excel may merge away
   // covered cells, but the PDF must never silently skip an unmerged value.
   for(const [rowNum] of visibleRows){
     if(rowNum<=titleRow||rowNum>end)continue;
     for(let col=1;col<=3;col++){
       const key=colAddr(col)+rowNum;
       if(!String(cells.get(key)?.value||'').trim()||covered.has(key))continue;
       if(!printedCells.has(key))throw Error('PDF Export missing Excel cell '+key);
     }
   }
   const iosPrintCSS=isIOS?`
   /* iOS: preserve exactly the same vector grid colors as desktop PDF. */
   @media print {
     html,body,.paper,table.excel,table.excel td { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
   }
   `:'';
   const css=`
   /* Zero printer page margins suppress Chromium's unrelated date/url/title header.
      Print-safe margins are provided by repeating .paper padding on every page. */
   @page { size: A4 portrait; margin: 0; }
   html,body{margin:0;padding:0;color:#212a35;font-family:Arial,'Times New Roman',sans-serif;font-size:8.4pt;background:white}
   *{box-sizing:border-box} .print-toolbar{display:flex;align-items:center;justify-content:space-between;gap:14px;padding:12px 18px;background:#f4f6f8;border-bottom:1px solid #dce1e7;font-family:Arial,sans-serif}
   .print-toolbar .title{font-size:13px;font-weight:700;color:#1d2d41}.print-toolbar .sub{font-size:11px;color:#718096;margin-top:3px}
   .print-toolbar button{border:none;border-radius:6px;background:#b9232c;color:#fff;padding:10px 18px;font-size:12px;font-weight:700;cursor:pointer}
   .paper{margin:14px auto;max-width:920px;padding:5mm 4mm;box-shadow:0 5px 22px #14233718;background:white}
   .spec-masthead{display:block;margin-bottom:10px;break-inside:avoid;page-break-inside:avoid;font-family:Arial,sans-serif;color:#20344c}
   .spec-masthead-top{display:flex;align-items:flex-start;justify-content:space-between;gap:12px}
   .spec-company{font-size:10pt;font-weight:700;line-height:1.15;max-width:78%}
   .spec-address{font-size:7.5pt;font-weight:400;line-height:1.26;margin-top:3px;color:#4b5968}
   .spec-logo-row{height:57px;display:flex;align-items:center;justify-content:flex-start;margin:3px 0 6px}
   .spec-logo{display:block;width:126px;height:auto;max-height:56px;object-fit:contain}
   .spec-report-title{font-size:13.5pt;line-height:1.25;text-align:center;font-weight:700;margin:5px 0 9px;color:#203f65}
   
   table.excel{border-collapse:collapse;border-spacing:0;table-layout:fixed;width:100%;font-size:8.2pt;line-height:1.18}
   table.excel td{padding:3px 4px;vertical-align:middle;white-space:pre-wrap;overflow-wrap:break-word;word-break:normal;line-height:1.23;min-width:0}
   /* Unified PDF border system: MR and MRL use exactly the same vector stroke.
      Source workbook edges are deliberately NOT emitted as inline borders.
      0.75pt = 1 CSS pixel; all normal sides have the same width/color. */
   table.excel tr[data-source-row]:not(.excel-prelude):not(.excel-space) td{
     border:0.75pt solid #596571;
   }
   /* The inner rule of a true Vietnamese/English pair is lighter, not thicker.
      Apply to both adjacent cells to avoid border-collapse color conflicts.
      Rowspanned labels keep one uninterrupted outline. */
   /* Compact the source/origin paragraph, preserving every character while
      avoiding a lone almost-empty final A4 page (especially MRL's row 305). */
   table.excel tr.excel-origin-note td{font-size:7.45pt!important;line-height:1.12;padding:2px 4px!important;vertical-align:top}
   table.excel tr.excel-bilingual-upper td:not([rowspan]){border-bottom:0.5pt solid #c6ced6!important}
   table.excel tr.excel-bilingual-lower td:not([rowspan]){border-top:0.5pt solid #c6ced6!important}
   table.excel tr.excel-bilingual-upper{break-after:avoid;page-break-after:avoid}
   table.excel tr.excel-bilingual-lower{break-before:avoid;page-break-before:avoid}
   /* The decorative prelude/blank spacer rows are outside the grid. */
   table.excel tr.excel-prelude td,table.excel tr.excel-space td{border:0}
   /* MRL entrance is a single logical device/floor section: start it on a new
      page to prevent a top-floor row being orphaned from main/other floors. */
   main.mrl-export tr[data-source-row="69"]{break-before:page;page-break-before:always}
   tr.excel-row{break-inside:avoid;page-break-inside:avoid}
   tr.excel-space td{padding:0 2px} tr.excel-heading{break-after:avoid;page-break-after:avoid} tr.excel-features-title{break-before:auto;page-break-before:auto}
   @media print{html,body{background:#fff;print-color-adjust:exact;-webkit-print-color-adjust:exact} .print-toolbar{display:none!important}.paper{box-shadow:none;margin:0;width:100%;max-width:100%;padding:13mm 11mm 14mm;-webkit-box-decoration-break:clone;box-decoration-break:clone;zoom:1}table.excel{width:100%;} }
   `;
   return '<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>'+esc(modelName)+' | A4 Portrait Specification</title><style>'+css+iosPrintCSS+'</style></head><body>'+
   '<div class="print-toolbar"><div><div class="title">Bản in Spec / Print-ready Specification</div><div class="sub">'+esc(modelName)+' · '+project+' · A4 dọc · Tắt “Headers and footers” trong hộp thoại in nếu được hỏi</div></div><button onclick="window.print()">In / Save as PDF</button></div>'+
   '<main class="paper '+(model==='MRL'?'mrl-export':'')+'">'+masthead+'<table class="excel"><colgroup>'+colgroup+'</colgroup><tbody>'+printRows.join('')+'</tbody></table></main></body></html>';
 }
 window.SPEC_PRINT_LAYOUT={fromXlsx};
})();
