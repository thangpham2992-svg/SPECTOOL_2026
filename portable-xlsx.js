/* Cross-platform, format-preserving OOXML export. Keeps source cell styles,
   mergeCells, row/column geometry and original workbook style tables as XML text.
   No DOMParser / XMLSerializer in XLSX generation (iOS parity). */
(()=>{'use strict';
 const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
 const attr=(tag,key)=>{const k=key.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');const m=tag.match(new RegExp('(?:^|\\s)'+k+'\\s*=\\s*(["\\\'])(.*?)\\1','i'));return m?m[2]:''};
 const set=(tag,k,v)=>{const open=tag.match(/^<[^>]*>/)?.[0]||tag;const suffix=tag.slice(open.length);const name=k.replace(':','\:');const regex=new RegExp('('+name+'\\s*=\\s*)(["\\\']).*?\\2','i');const revised=regex.test(open)?open.replace(regex,(_m,part)=>part+'"'+esc(v)+'"'):open.replace(/\s*\/?\s*>$/,m=>' '+k+'="'+esc(v)+'"'+m);return revised+suffix};
 const rm=(tag,k)=>tag.replace(new RegExp('\\s'+k.replace(':','\\:')+'\\s*=\\s*(["\\\']).*?\\1','gi'),'');
 const findTags=(xml,name)=>[...xml.matchAll(new RegExp('<(?:[\\w.-]+:)?'+name+'\\b[^>]*\\/?>','g'))].map(x=>x[0]);
 const tagBlock=(xml,name)=>{let m=xml.match(new RegExp('<(?:[\\w.-]+:)?'+name+'\\b[^>]*>[\\s\\S]*?<\\/?(?:[\\w.-]+:)?'+name+'>','i'));return m?.[0]||''};
 const rowN=a=>Number(/\d+/.exec(a)?.[0]||0);
 function textCell(raw,addr,value,prefix){
   const open=raw?.match(/^<[^>]+>/)?.[0] || `<${prefix}c r="${addr}">`;
   let t=rm(rm(open.replace(/\s*\/?>$/,'>'),'t'),'cm');t=set(t,'t',typeof value==='number'?'n':typeof value==='boolean'?'b':'inlineStr');
   if(value==='')return t.replace(/>$/, '/>');
   const isValue=typeof value==='number' ? `<${prefix}v>${value}</${prefix}v>` :
     typeof value==='boolean' ? `<${prefix}v>${value?'1':'0'}</${prefix}v>` :
     `<${prefix}is><${prefix}t xml:space="preserve">${esc(value)}</${prefix}t></${prefix}is>`;
   return t+isValue+`</${prefix}c>`;
 }
 async function create(zip,doc){
   const wbPath='xl/workbook.xml',relPath='xl/_rels/workbook.xml.rels',stylesPath='xl/styles.xml';
   let workbook=await zip.file(wbPath).async('string'),rels=await zip.file(relPath).async('string');
   let styles=await zip.file(stylesPath).async('string');
   const sheetTags=findTags(workbook,'sheet'),sheet=sheetTags.find(t=>attr(t,'name')==='SpecFormExport');
   if(!sheet)throw Error('SpecFormExport sheet not found in source workbook');
   const rid=attr(sheet,'r:id')||attr(sheet,'id');
   const relTags=findTags(rels,'Relationship');
   const relationship=relTags.find(t=>attr(t,'Id')===rid);
   if(!relationship)throw Error('Export sheet workbook relationship missing: '+rid);
   let target=attr(relationship,'Target');
   let sheetPath=target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'').replace(/^xl\//,'');
   const sourceFile=zip.file(sheetPath);
   if(!sourceFile)throw Error('Source export worksheet missing: '+sheetPath);
   const original=await sourceFile.async('string');
   const pref=(original.match(/<(\w+:)?worksheet\b/)?.[1]||'');
   const rowtag=pref+'row',celltag=pref+'c';
   const mergeMatch=original.match(new RegExp('<'+pref+'mergeCells\\b[^>]*>[\\s\\S]*?<\\/'+pref+'mergeCells>'));
   const mergeCount=(mergeMatch?.[0].match(new RegExp('<'+pref+'mergeCell\\b','g'))||[]).length;
   let xml=original;
   const pending=new Map();
   for(const [a,val] of Object.entries(doc.cells||{})){
      if(!/^[ABC]\d+$/.test(a)||val===null||val===undefined)continue;
      const n=rowN(a);if(!pending.has(n))pending.set(n,[]);pending.get(n).push([a,val]);
   }
   const blue=/^\s*[IVXLCDM]+\.\s+/i;
   const headingRows=[...new Set(Object.entries(doc.cells||{}).filter(([a,v])=>/^A\d+$/.test(a)&&rowN(a)<=300&&!doc.hiddenRows?.has(rowN(a))&&blue.test(String(v||''))).map(([a])=>rowN(a)))];
   const headingSet=new Set(headingRows);
   // Existing styles stay intact. Clone each referenced XF with a new blue fill.
   const baseStyles=[...styles.matchAll(new RegExp('<'+pref+'xf\\b[^>]*\\/?>','g'))];
   const xfElement=(styles.match(/<([\w.-]+:)?cellXfs\b/)?.[1]||'x:')+'xf';
   const stylePrefix=xfElement.slice(0,-2);
   const xfBlock=styles.match(new RegExp('<'+stylePrefix+'cellXfs\\b([^>]*)>([\\s\\S]*?)<\\/'+stylePrefix+'cellXfs>'));
   const fillsBlock=styles.match(new RegExp('<'+stylePrefix+'fills\\b([^>]*)>([\\s\\S]*?)<\\/'+stylePrefix+'fills>'));
   if(!xfBlock||!fillsBlock)throw Error('Source XLSX missing cell styles/fills');
   const baseXfs=[...xfBlock[2].matchAll(new RegExp('<'+xfElement+'\\b[^>]*\\/>|<'+xfElement+'\\b[^>]*>[\\s\\S]*?<\\/'+xfElement+'>','g'))].map(m=>m[0]);
   let fillId=Number(attr(fillsBlock[0].split('>')[0]+'>','count')||0);
   const blueFill=`<${stylePrefix}fill><${stylePrefix}patternFill patternType="solid"><${stylePrefix}fgColor rgb="FFD9EDF7"/><${stylePrefix}bgColor indexed="64"/></${stylePrefix}patternFill></${stylePrefix}fill>`;
   const xfCache=new Map(),addXf=[];
   function blueStyle(n){n=Number(n)||0;if(xfCache.has(n))return xfCache.get(n);let t=baseXfs[n]||baseXfs[0];if(!t)throw Error('Base XF missing');t=set(t,'fillId',fillId);t=set(t,'applyFill','1');let id=baseXfs.length+addXf.length;addXf.push(t);xfCache.set(n,id);return id}
   const processed=new Set();
   const rowRe=new RegExp('<'+rowtag+'\\b([^>]*)>([\\s\\S]*?)<\\/'+rowtag+'>','g');
   xml=xml.replace(rowRe,(whole,rawAttrs,body)=>{
      const number=Number(attr('<'+rowtag+rawAttrs+'>','r'));
      processed.add(number);
      const vals=new Map(pending.get(number)||[]);
      let rowOpen='<'+rowtag+rawAttrs+'>';
      // Row 28 is 13.8pt in source: a modest 18pt for the two adjoining rows.
      if(number===29||number===30){rowOpen=set(rowOpen,'ht','18');rowOpen=set(rowOpen,'customHeight','1')}
      if(doc.hiddenRows?.has(number)){rowOpen=set(rowOpen,'hidden','1');rowOpen=set(rowOpen,'ht','0');rowOpen=set(rowOpen,'customHeight','1')}
      if(doc.hallLanternExportRows?.includes(number)){
        const desc=String(doc.cells['C'+number]||'');const lines=desc.split(/\r?\n/).reduce((a,s)=>a+Math.max(1,Math.ceil(s.length/69)),0);
        rowOpen=set(rowOpen,'ht',Math.min(185,Math.max(24,lines*16+14)));rowOpen=set(rowOpen,'customHeight','1');
      }
      const cellRe=new RegExp('<'+celltag+'\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/'+celltag+'>)','g');
      const found=new Set();
      body=body.replace(cellRe,c=>{
         const a=attr(c.slice(0,c.indexOf('>')+1),'r');if(!a)return c;found.add(a);
         let next=vals.has(a)?textCell(c,a,vals.get(a),pref):c;
         if(headingSet.has(number)&&/^[ABC]\d+$/.test(a)){
           let op=next.match(/^<[^>]+>/)?.[0];if(op){let id=blueStyle(attr(op,'s'));next=set(op,'s',id)+next.slice(op.length)}
         }
         return next;
      });
      const extras=[];
      for(const [a,val] of vals){if(!found.has(a)&&val!==''){
        let cell=textCell(null,a,val,pref);
        if(headingSet.has(number)){let op=cell.match(/^<[^>]+>/)[0];cell=set(op,'s',blueStyle(0))+cell.slice(op.length)}
        extras.push(cell)
      }}
      if(headingSet.has(number))for(const col of ['A','B','C']){
        let a=col+number;if(!found.has(a)&&!vals.has(a)){
          let cell=textCell(null,a,'',pref),op=cell.match(/^<[^>]+>/)[0];extras.push(set(op,'s',blueStyle(0))+cell.slice(op.length))
        }
      }
      if(extras.length){
        body+=extras.join('');
        const tags=[...body.matchAll(cellRe)].map(m=>m[0]);
        if(tags.join('')===body)body=tags.sort((a,b)=>{
          const col=t=>attr(t.slice(0,t.indexOf('>')+1),'r').replace(/\d+/g,'');
          const index=t=>[...col(t)].reduce((v,ch)=>v*26+ch.charCodeAt(0)-64,0);
          return index(a)-index(b);
        }).join('');
      }
      return rowOpen+body+'</'+rowtag+'>';
   });
   const newRows=[];
   for(const [num,rawArr] of [...pending.entries()].sort((a,b)=>a[0]-b[0]))if(!processed.has(num)){
     const arr=rawArr.filter(([,v])=>v!=='');if(!arr.length)continue;
     let cells=arr.map(([a,v])=>textCell(null,a,v,pref)).join('');
     newRows.push('<'+rowtag+' r="'+num+'">'+cells+'</'+rowtag+'>');
   }
   if(newRows.length)xml=xml.replace('</'+pref+'sheetData>',newRows.join('')+'</'+pref+'sheetData>');
   if(headingRows.length){
     styles=styles.replace(fillsBlock[0],'<'+stylePrefix+'fills'+fillsBlock[1].replace(/count="\d+"/,'count="'+(fillId+1)+'"')+'>'+fillsBlock[2]+blueFill+'</'+stylePrefix+'fills>');
     styles=styles.replace(xfBlock[0],'<'+stylePrefix+'cellXfs'+xfBlock[1].replace(/count="\d+"/,'count="'+(baseXfs.length+addXf.length)+'"')+'>'+xfBlock[2]+addXf.join('')+'</'+stylePrefix+'cellXfs>');
   }
   // Workbook markup stays byte-for-byte intact except for removing inactive sheet references.
   for(const t of sheetTags)if(t!==sheet)workbook=workbook.replace(t,'');
   workbook=workbook.replace(sheet, set(sheet,'name','Specification'));
   const defPref=workbook.match(/<(\w+:)?workbook\b/)?.[1]||'';
   const printArea=`<${defPref}definedNames><${defPref}definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Specification&apos;!$A$1:$C$${Math.max(1,...Object.keys(doc.cells||{}).filter(a=>/^[ABC]\d+$/.test(a)).map(rowN))}</${defPref}definedName></${defPref}definedNames>`;
   if(new RegExp('<'+defPref+'definedNames\\b').test(workbook))workbook=workbook.replace(new RegExp('<'+defPref+'definedNames\\b[^>]*>[\\s\\S]*?<\\/'+defPref+'definedNames>'),printArea);
   else workbook=workbook.replace('</'+defPref+'sheets>','</'+defPref+'sheets>'+printArea);
   workbook=workbook.replace(new RegExp('<'+defPref+'calcPr\\b[^>]*\\/?>','g'),'');
   for(const rt of relTags)if(rt!==relationship&&(/\/worksheet$|\/calcChain$/.test(attr(rt,'Type'))))rels=rels.replace(rt,'');
   const ctPath='[Content_Types].xml';let ct=await zip.file(ctPath).async('string');
   ct=ct.replace(/<Override\b[^>]*\/>/g,tag=>{
     const path=attr(tag,'PartName');return(path==='/xl/calcChain.xml'||(/^\/xl\/worksheets\/sheet\d+\.xml$/.test(path)&&path!=='/'+sheetPath))?'':tag;
   });
   zip.remove('xl/calcChain.xml');
   // Do not strip mergeCells or base style blocks: raw source markup is retained.
   if(mergeCount){const after=(xml.match(new RegExp('<'+pref+'mergeCell\\b','g'))||[]).length;
     if(after!==mergeCount)throw Error('Merge preservation failed: '+after+'/'+mergeCount);
   }
   zip.file(sheetPath,xml);zip.file(stylesPath,styles);zip.file(wbPath,workbook);zip.file(relPath,rels);zip.file(ctPath,ct);
   // Other worksheet parts are intentionally left unreferenced inside ZIP, not exposed in Excel.
   return {zip,freezeCount:Object.keys(doc.cells||{}).length,hiddenRows:doc.hiddenRows?.size||0,filename:doc.model,portable:true,mergeCount};
 }
 window.SPEC_PORTABLE_XLSX={create};
})();
