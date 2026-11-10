'use strict';
/* Print from exported XLSX OOXML rather than resynthesizing a three-column table.
   Uses the workbook's SpecFormExport cell styles, column proportions, merged ranges,
   row heights, and export-engine-hidden feature rows. Runs fully offline in browser. */
(()=>{
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 const OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
 const q=(node,n)=>Array.from(node?.childNodes||[]).find(x=>x.localName===n)||null;
 const qa=(node,n)=>Array.from(node?.getElementsByTagNameNS(ns,n)||[]);
 const attr=(node,n,defaultValue='')=>node?.getAttribute(n)??defaultValue;
 const esc=v=>String(v??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
 const colNo=s=>[...String(s)].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
 const colAddr=n=>{let s='';while(n>0){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26)}return s};
 const addrParts=a=>{const m=/^([A-Z]+)(\d+)$/.exec(a||'');return m?{col:colNo(m[1]),row:+m[2]}:null};
 const parse=str=>new DOMParser().parseFromString(str,'application/xml');
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
 function formatStyle(index,book){const x=book.xfs[index]||{};const f=x.font||{};const rules=[];
   if(f.bold)rules.push('font-weight:700');if(f.italic)rules.push('font-style:italic');
   rules.push('font-size:'+Math.max(7.2,Math.min(13,Number(f.size||9)))+'pt');
   if(f.color)rules.push('color:'+f.color);if(x.fill&&x.fill!=='transparent'&&x.fill!=='#000000')rules.push('background:'+x.fill);
   if(x.align){const a=attr(x.align,'horizontal');const v=attr(x.align,'vertical');if(['center','right','left','justify'].includes(a))rules.push('text-align:'+a);if(v==='center')rules.push('vertical-align:middle');if(attr(x.align,'wrapText')==='1')rules.push('overflow-wrap:break-word');}
   if(x.fill&&/^#[0-9a-f]{6}$/i.test(x.fill)){const v=x.fill.slice(1);const rr=parseInt(v.slice(0,2),16),gg=parseInt(v.slice(2,4),16),bb=parseInt(v.slice(4,6),16);if((rr*299+gg*587+bb*114)/1000<125)rules.push('color:#ffffff');}
   const b=x.border||{};for(const key of ['left','right','top','bottom'])if(b[key])rules.push('border-'+key+':'+b[key]);
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
   for(const row of rows){let rn=Number(attr(row,'r'));if(!rn||rn>300||attr(row,'hidden')==='1')continue;
     visibleRows.set(rn,row);
     for(const c of Array.from(row.children).filter(x=>x.localName==='c')){const a=attr(c,'r'),p=addrParts(a);if(!p||p.col>3)continue;
       const v=readText(c,strings);cells.set(a,{value:v,style:Number(attr(c,'s',0))});if(v.trim())largest=Math.max(largest,rn);
     }
   }
   if(!largest)throw Error('Export SpecFormExport is empty');
   const end=Math.min(largest,300);
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
   const project=esc(first);const modelName=model==='MR'?'NEXIEZ-MR ALPHA':'NEXIEZ-MRL Version02 ALPHA';
   const printRows=[];
   for(let rowNum=1;rowNum<=end;rowNum++){
     const row=visibleRows.get(rowNum);if(!row)continue;
     const rowCells=[];const isBlank=[1,2,3].every(i=>!String(cells.get(colAddr(i)+rowNum)?.value||'').trim());
     const rowHeight=Number(attr(row,'ht',attr(q(data.documentElement,'sheetFormatPr'),'defaultRowHeight','13.8')))||13.8;
     for(let col=1;col<=3;col++){
       const key=colAddr(col)+rowNum;if(covered.has(key))continue;
       const val=cells.get(key)||{value:'',style:0};const merge=merges.get(key)||{colspan:1,rowspan:1};
       const extra=merge.colspan>1?' colspan="'+merge.colspan+'"':'';const vextra=merge.rowspan>1?' rowspan="'+merge.rowspan+'"':'';
       let style=formatStyle(val.style,styles);
       // Explicit inline background prevents theme tint and browser print defaults
       // from suppressing the original pale-blue I/II/III/IV section headers.
       if(blueRomanRows.has(rowNum))style+=';background-color:#d9edf7;color:#173c56;font-weight:700';
       rowCells.push('<td'+extra+vextra+' data-cell="'+key+'" style="'+style+'">'+esc(val.value).replace(/\r\n|\r|\n/g,'<br>')+'</td>');
     }
     // Excel's explicit row heights are respected as minimums; text-heavy wrapped cells can expand.
     const customHeight=attr(row,'ht')!=='';
     const minHeight=Math.max(isBlank?2:0,Math.min(72,rowHeight*1.07));
     const rowValue=[1,2,3].map(i=>String(cells.get(colAddr(i)+rowNum)?.value||'')).join(' ');
     const cl=(isBlank?' excel-space':'')+(/^(?:I|II|III|IV|V|VI)\.\s|CÁC TÍNH NĂNG|CÁC TÍNH NĂNG|Các tính năng/i.test(rowValue.trim())?' excel-heading':'')+(/Các tính năng \(Features\)/i.test(rowValue)?' excel-features-title':'');
     printRows.push('<tr class="excel-row'+cl+'" data-source-row="'+rowNum+'" style="height:'+minHeight.toFixed(1)+'px">'+rowCells.join('')+'</tr>');
   }
   const css=`
   @page { size: A4 portrait; margin: 10mm 9mm 11mm 9mm; }
   html,body{margin:0;padding:0;color:#212a35;font-family:Arial,'Times New Roman',sans-serif;font-size:8.4pt;background:white}
   *{box-sizing:border-box} .print-toolbar{display:flex;align-items:center;justify-content:space-between;gap:14px;padding:12px 18px;background:#f4f6f8;border-bottom:1px solid #dce1e7;font-family:Arial,sans-serif}
   .print-toolbar .title{font-size:13px;font-weight:700;color:#1d2d41}.print-toolbar .sub{font-size:11px;color:#718096;margin-top:3px}
   .print-toolbar button{border:none;border-radius:6px;background:#b9232c;color:#fff;padding:10px 18px;font-size:12px;font-weight:700;cursor:pointer}
   .paper{margin:14px auto;max-width:920px;padding:5mm 4mm;box-shadow:0 5px 22px #14233718;background:white}
   table.excel{border-collapse:collapse;border-spacing:0;table-layout:fixed;width:100%;font-size:8.2pt;line-height:1.18}
   table.excel td{padding:2px 3px;vertical-align:middle;white-space:pre-wrap;overflow-wrap:break-word;word-break:normal;line-height:1.2;min-width:0}
   tr.excel-row{break-inside:avoid;page-break-inside:avoid}
   tr.excel-space td{padding:0 2px} tr.excel-heading{break-after:avoid;page-break-after:avoid} tr.excel-features-title{break-before:page;page-break-before:always}
   @media print{html,body{background:#fff;print-color-adjust:exact;-webkit-print-color-adjust:exact} .print-toolbar{display:none!important}.paper{box-shadow:none;padding:0;margin:0;width:100%;max-width:100%}table.excel{width:100%;} .paper{zoom:.96;width:104.17%;} }
   `;
   return '<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>'+esc(modelName)+' | A4 Portrait Specification</title><style>'+css+'</style></head><body>'+
   '<div class="print-toolbar"><div><div class="title">Bản in Spec / Print-ready Specification</div><div class="sub">'+esc(modelName)+' · '+project+' · A4 dọc / A4 Portrait · Dựa trên SpecFormExport</div></div><button onclick="window.print()">In / Save as PDF</button></div>'+
   '<main class="paper"><table class="excel"><colgroup>'+colgroup+'</colgroup><tbody>'+printRows.join('')+'</tbody></table></main></body></html>';
 }
 window.SPEC_PRINT_LAYOUT={fromXlsx};
})();
