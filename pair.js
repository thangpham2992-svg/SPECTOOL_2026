const pairTemplate='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';
const edgeStyles={"company": [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25], "version": [26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41], "title": [42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57], "project": [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25], "entry": [10, 11, 12, 13, 14, 15, 16, 17, 58, 59, 60, 61, 62, 63, 64, 65], "header": [66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 5], "section": [81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 6], "label": [96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 7], "value": [111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 8], "en": [126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 9], "feature_value": [142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 141], "feature_en": [158, 159, 160, 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 157]};
const bilingualStyles={"96,4": 173, "96,8": 174, "96,12": 175, "97,4": 176, "97,8": 177, "97,12": 178, "98,4": 179, "98,8": 180, "98,12": 181, "99,4": 182, "99,8": 183, "99,12": 184, "100,8": 185, "101,8": 186, "102,8": 187, "103,8": 188, "104,4": 189, "105,4": 190, "106,4": 191, "107,4": 192, "111,4": 193, "111,8": 194, "111,12": 195, "112,4": 196, "112,8": 197, "112,12": 198, "113,4": 199, "113,8": 200, "113,12": 201, "114,4": 202, "114,8": 203, "114,12": 204, "115,8": 205, "116,8": 206, "117,8": 207, "118,8": 208, "119,4": 209, "120,4": 210, "121,4": 211, "122,4": 212, "126,4": 213, "126,8": 214, "126,12": 215, "127,4": 216, "127,8": 217, "127,12": 218, "128,4": 219, "128,8": 220, "128,12": 221, "129,4": 222, "129,8": 223, "129,12": 224, "130,8": 225, "131,8": 226, "132,8": 227, "133,8": 228, "134,4": 229, "135,4": 230, "136,4": 231, "137,4": 232, "158,4": 233, "158,8": 234, "158,12": 235, "159,4": 236, "159,8": 237, "159,12": 238, "160,4": 239, "160,8": 240, "160,12": 241, "161,4": 242, "161,8": 243, "161,12": 244, "162,8": 245, "163,8": 246, "164,8": 247, "165,8": 248, "166,4": 249, "167,4": 250, "168,4": 251, "169,4": 252, "142,4": 253, "142,8": 254, "142,12": 255, "143,4": 256, "143,8": 257, "143,12": 258, "144,4": 259, "144,8": 260, "144,12": 261, "145,4": 262, "145,8": 263, "145,12": 264, "146,8": 265, "147,8": 266, "148,8": 267, "149,8": 268, "150,4": 269, "151,4": 270, "152,4": 271, "153,4": 272};
const formStyles={"company": 1, "version": 2, "title": 3, "project": 1, "entry": 4, "header": 5, "section": 6, "label": 7, "value": 8, "en": 9, "feature_value": 141, "feature_en": 157};
const xmlText=v=>String(v??'').replace(/[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu,'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'})[c]);
function alignedRows(cars){
 const order=[],rows=new Map();
 cars.forEach((car,i)=>{
  let cursor=0;
  for(const r of car.rows){
   if(!rows.has(r.key)){order.splice(cursor,0,r.key);rows.set(r.key,{...r,itpTypes:cars.map(()=>null),parts:cars.map(()=>[]),values:Array(cars.length).fill(null)})}
   cursor=order.indexOf(r.key)+1;rows.get(r.key).itpTypes[i]=r.itpType??null;rows.get(r.key).values[i]=r.value;rows.get(r.key).parts[i]=r.englishParts||[];
  }
 });
 return order.map(key=>rows.get(key));
}
const hasVN=text=>/[đĐ]|\p{Script=Latin}\p{M}/u.test(text.normalize('NFD'));
function englishText(text){
 return text.trim().replace(/^\(\s*((?:[^()]|\([^()]*\))*)\s*\)([.:]?)$/,'$1$2').trim();
}
function labelParts(text){
 const blocks=text.trim().match(/^([^]*?)\s*\(((?:[^()]|\([^()]*\))+)\)\s*([.:]?)$/);
 if(blocks&&(hasVN(blocks[1])||/\bthang\b/i.test(blocks[1]))&&!hasVN(blocks[2])&&/[a-z]/i.test(blocks[2]))return [blocks[1].trim(),blocks[2].trim()+blocks[3]];
 const slash=text.match(/^([^]*?)\s+\/\s+([^]+)$/);
 if(slash&&hasVN(slash[1])&&!hasVN(slash[2]))return [slash[1],slash[2]];
 return null;
}
function displayRuns(value,english,mixed,hints){
 const runs=[];
 function add(text,en){
  if(!text.trim())return;
  runs.push({text:(runs.length?'\n':'')+(en?englishText(text):text.trim()),en});
 }
 if(english){add(String(value),true);return runs}
 if(!mixed){return [{text:String(value),en:false}]}
 const source=String(value);
 let rest=source;
 function plain(chunk){
  for(const part of chunk.split(/\r?\n/)){
   const pair=labelParts(part);
   if(pair){add(pair[0],false);add(pair[1],true)}
   else add(part,runs.length>0&&hasVN(source)&&!hasVN(part)&&/[a-z]{2}/i.test(part)&&!/^Mitsubishi Standard$/i.test(part.trim()));
  }
 }
 for(const hint of hints||[]){
  const i=rest.indexOf(hint);if(i<0)continue;
  const prefix=rest.slice(0,i),suffix=rest.slice(i+hint.length);
  const wrapped=prefix.endsWith('(')&&suffix.startsWith(')');
  plain(wrapped?prefix.slice(0,-1):prefix);add(hint,true);rest=wrapped?suffix.slice(1):suffix;
 }
 plain(rest);return runs;
}
function sameControlGroup(a,b){return a?.groupId!=null&&a.groupId===b?.groupId}
function groupSheet(template,cars,featureNotes,showItpTypes){
 const last=String.fromCharCode(66+cars.length),width=108/cars.length,merges=[],lines=[],carry=new Map();
 function line(values,kind='value'){
  const n=lines.length+1;const headingArea=n<=8;
  let height=headingArea?([17,16,2,2,47,3,27,27][n-1]):21,cells=[];
  for(const [col,value,span=1,style=kind,depth=1,edges=15,englishLabel=false,hints=[],bold=false,scopeNote=''] of values){
   const letter=String.fromCharCode(65+col);
   if(span>1||depth>1)merges.push('<mergeCell ref="'+letter+n+':'+String.fromCharCode(64+col+span)+(n+depth-1)+'"/>');
   const capacity=headingArea?Array.from({length:span},(_,offset)=>[25,15][col+offset]??width).reduce((a,b)=>a+b,0)-2:col<2?(span>1?42:24):width*span-2;
   const english=englishLabel||style==='en'||style==='feature_en';
   const mixed=['label','title','section','project','value','feature_value'].includes(style)||(style==='header'&&col===0);
   const dims=style==='label'&&col===1&&/^R[^\n]*\nW[^\n]*$/.test(value);
   const runs=dims?String(value).split('\n').map((text,i)=>({text:(i?'\n':'')+text,en:i===1})):displayRuns(value,english,mixed,hints);
   if(scopeNote)runs.push({text:'\n'+scopeNote,scope:true});
   const t=text=>'<t xml:space="preserve">'+xmlText(text)+'</t>';
   const englishSize=style==='title'?'<sz val="16"/>':'';
   const content=bold?'<r><rPr><rFont val="Arial"/><b/></rPr>'+t(String(value))+'</r>':runs.some(run=>run.en||run.scope)?runs.map(run=>'<r>'+(run.scope?'<rPr><rFont val="Arial"/><b/><color rgb="FFCC0000"/></rPr>':run.en?'<rPr><rFont val="Arial"/><i/><color rgb="FF245A86"/>'+englishSize+'</rPr>':'<rPr><rFont val="Arial"/></rPr>')+t(run.text)+'</r>').join(''):t(runs.map(run=>run.text).join(''));
   height=Math.max(height,(headingArea?13.5:16)*runs.map(run=>run.text).join('').split('\n').reduce((sum,t)=>sum+Math.max(1,Math.ceil(t.length/capacity)),0));
   for(let y=0;y<depth;y++)for(let x=0;x<span;x++){
    const mask=(x===0?1:0)|(x===span-1?2:0)|(y===0?4:0)|(y===depth-1?8:0);
    const anchor=x===0&&y===0;
    const baseStyle=anchor&&edges===15?formStyles[style]:edgeStyles[style][(anchor?15:mask)&edges];
    const faint=mask & ~edges & 12;
    const id=faint?bilingualStyles[[baseStyle,faint]]:baseStyle;
    const addr=String.fromCharCode(65+col+x)+(n+y);
    const cell='<c r="'+addr+'" s="'+id+'"'+(anchor?' t="inlineStr"><is>'+content+'</is></c>':'/>');
    if(!y)cells[col+x]=cell;
    else{if(!carry.has(n+y))carry.set(n+y,[]);carry.get(n+y)[col+x]=cell}
   }
  }
  for(const [col,cell] of (carry.get(n)||[]).entries())if(cell)cells[col]=cell;
  carry.delete(n);
  lines.push('<row r="'+n+'" ht="'+Math.min(height,409)+'" customHeight="1">'+cells.join('')+'</row>');
 }
 const rows=alignedRows(cars),total=2+cars.length;
 const quantityRow=rows.find(row=>row.quantity&&row.a);
 const liftCounts=cars.map((car,i)=>parseInt(quantityRow?.values[i],10)||1);
 const preamble=rows.filter(r=>r.preamble);
 const meta=cars[0].rows.find(r=>r.meta)?.meta||{};
 line([[0,meta.company||'',total-1,'company'],[total-1,meta.doc||'',1,'version']]);
 line([[0,meta.address||'',total-1,'company'],[total-1,meta.revision||'',1,'version']]);
 line([]);line([]);
 line([[0,'ĐẶC TÍNH KỸ THUẬT (Main specifications)',total,'title']]);line([]);
 for(const [label,pattern] of [['Công trình (Project)',/project|công trình/i],['Hạng mục (Item)',/item|hạng mục/i]]){
  const field=preamble.find(r=>pattern.test(r.a));line([[0,label,1,'project'],[1,field?.values[0]||'',total-1,'entry']]);
 }
 line([[0,'Tên thang (Elev name)',2],...cars.map((c,i)=>[2+i,c.name])],'header');
 const vietnamese=text=>/[đĐ]|\p{Script=Latin}\p{M}/u.test(String(text??'').normalize('NFD'));
 function splitTranslation(value){
  if(!value)return null;
  const match=String(value).match(/^([\s\S]+?)\s+(\([^()]+\)[.!]?)$/);
  if(!match||!vietnamese(match[1])||vietnamese(match[2])||!/[A-Za-z]/.test(match[2]))return null;
  return [match[1].trim(),match[2]];
 }
 const body=rows.filter(r=>!r.preamble).flatMap(row=>{
  const r={...row,quantity:row.quantity||/^\s*(?:Số lượng(?: thang)?|Quantity)\s*(?:\(|$)/iu.test(row.a.split(/\r?\n/)[0])};
  if(r.quantity&&!r.a&&r.values.some(v=>/\bsets?\b/i.test(v||'')))r.lower=true;
  if(r.heading||r.upper||r.lower)return [r];
  const pairs=r.values.map(splitTranslation);
  if(!pairs.some(Boolean))return [r];
  return [{...r,upper:true,values:r.values.map((v,i)=>pairs[i]?.[0]??v)},
          {...r,a:'',b:'',lower:true,values:r.values.map((v,i)=>pairs[i]?.[1]??(v==null?null:''))}];
 });
 function bilingualBreak(i){
  const upper=body[i],lower=body[i+1];
  if(!upper||!lower||upper.heading||lower.heading)return false;
  if(lower.lower&&(upper.upper||(upper.quantity&&lower.quantity)))return true;
  const texts=lower.values.filter(Boolean);
  return !lower.a&&!lower.b&&texts.length>0&&texts.every(v=>/^\s*\(/.test(v));
 }
 const separateLabels=new Set(['Tải trọng','Loại cửa','Tốc độ','Số điểm dừng','Tầng phục vụ','Tầng không phục vụ','Cửa thoát hiểm','Loại mở cửa','Vị trí phòng máy','Vị trí đối trọng','An toàn đối trọng','Cấu trúc hố thang','Thanh dẫn hướng']);
 let labelEnd=-1,inFeatures=false,basicSection=false,separateLiftCells=false;
 for(const [index,r] of body.entries()){
  if(r.heading)basicSection=/^\s*I\./i.test(r.a);
  if(r.a)separateLiftCells=/Ngưỡng cửa phòng thang|\bCar sill\b/i.test(r.a)||basicSection&&separateLabels.has(r.a.split(/[\n(]/)[0].trim());
  const edges=15^(bilingualBreak(index-1)?4:0)^(bilingualBreak(index)?8:0);
  const featureHeading=/^\s*[-–—_]\s*Tính năng(?:\s|$)/iu.test(r.a)&&!r.b&&r.values.every(v=>!v);
  if(r.heading)inFeatures=/tính năng|features?/iu.test(r.a);
  if(featureHeading)inFeatures=true;
  if(r.heading){line([[0,r.a.trim(),total]],'section');continue}
  if(featureHeading||(r.fullWidth&&!r.b&&r.values.every(v=>!v))){line([[0,r.a.trim(),total,'label']]);continue}
  const groupControl=basicSection&&/group control operation|điều khiển (vận hành )?nhóm/i.test(r.key);
  if(groupControl)r.values=r.values.map((v,i)=>v!=null&&cars[i].groupId!=null?v+' · '+cars.filter(c=>sameControlGroup(c,cars[i])).map(c=>c.name).join(' + '):v);
  const dimension=/Car internal size|Door size|Hoistway size|Machine room size/i.test(r.a);
  const perLift=/Hoistway size|Machine room size/i.test(r.a);
  const hall=/Hall indicators & buttons|Hiển thị và nút bấm/i.test(r.key);
  const hallCount=/\d+\s*hộp\s*\/\s*\d+\s*thang\s*\/\s*1\s*tầng/gi;
  // Match the PIV families in the single form. Only replace quantity notes for known models.
  function hallPattern(value){
   let capacity=0;
   return String(value).split('\n').map(line=>{
    const model=line.match(/\b(PIV\d*|HBV\d*)[- ]?([AC]\d{3,4})/i);
    if(model){
     const family=model[2].toUpperCase();
     capacity=!/^PIV/i.test(model[1])?0:
      ['A1020','A1040','A1090','C720','C740','C776','C790'].includes(family)?2:
      ['A1010','A1030','A1080','C710','C730','C766','C780'].includes(family)?1:0;
    }
    return capacity?line.replace(hallCount,'{PIV:'+capacity+'}'):line;
   }).join('\n');
  }
  const scope=inFeatures?featureNotes.get(r.key):null;
  const labelsEnglish=r.lower||/^\([A-Z0-9 -]+\)/.test(r.a)&&!hasVN(r.a);
  const scopeNote=scope?(labelsEnglish?'Only apply for ':'Chỉ áp dụng cho thang ')+scope.join(', '):'';
  const common=scope?[...new Set(r.values.filter(v=>v!=null&&v.trim()))]:[];
  const compare=common.length===1?cars.map(()=>common[0]):hall?r.values.map(v=>v==null?v:hallPattern(v)):r.values;
  const values=[];
  if(index>labelEnd){
   let end=index;while(r.a&&end+1<body.length&&!body[end+1].a&&!body[end+1].heading)end++;
   const split=body.slice(index,end+1).some(item=>item.b);
   labelEnd=end;
   values.push([0,r.a,split?1:2,'label',end-index+1,end===index?edges:15,Boolean(r.lower&&r.a&&!hasVN(r.a)),[],false,scopeNote]);
  }
  if(r.b){
   let end=index;
   while(end<labelEnd&&!body[end+1].b)end++;
   values.push([1,r.b,1,'label',end-index+1,end===index?edges:15]);
  }
  for(let i=0;i<cars.length;){let j=i+1;while(!separateLiftCells&&!dimension&&!r.quantity&&j<cars.length&&compare[i]!=null&&compare[i]===compare[j]&&(!groupControl||sameControlGroup(cars[i],cars[j])))j++;
   let value=compare[i]??'—';
   if(perLift&&liftCounts[i]>=2&&/mm\s*$/i.test(value))value=value.trimEnd()+'/1 thang';
   if(hall){
    const lifts=liftCounts.slice(i,j).reduce((a,b)=>a+b,0);
    value=value.replace(/\{PIV:([12])\}/g,(_,capacity)=>Math.ceil(lifts/Number(capacity))+' hộp / '+lifts+' thang / 1 tầng');
   }
   const emphasizeIdentity=(/Manufacturer/i.test(r.a)&&value.trim()==='MITSUBISHI ELECTRIC')||
      (/Model/i.test(r.a)&&/^NEXIEZ-MRL?\b/i.test(value.trim()));
   values.push([2+i,value,j-i,(inFeatures?'feature_':'')+(r.lower||/^\s*\(/.test(r.values[i]||'')?'en':'value'),1,edges,false,r.parts[i],emphasizeIdentity]);i=j;
  }
  line(values);
  if(showItpTypes&&r.itpTypes.some(type=>type!==null)&&!body[index+1]?.itpTypes.some(type=>type!==null)){
   const config=[[0,'Cấu hình liên lạc nội bộ (ITP configuration)',2,'label']];
   for(let i=0;i<cars.length;){
    const type=r.itpTypes[i];let span=1;
    while(type&&sameControlGroup(cars[i],cars[i+span])&&type===r.itpTypes[i+span])span++;
    config.push([2+i,type||'—',span,'feature_value']);
    i+=span;
   }
   line(config);
  }
 }
 return template.replace(/<dimension[^>]*\/>/,'<dimension ref="A1:'+last+lines.length+'"/>').replace(/<cols>.*?<\/cols>/s,'<cols><col min="1" max="1" width="25" customWidth="1"/><col min="2" max="2" width="15" customWidth="1"/><col min="3" max="'+(2+cars.length)+'" width="'+width+'" customWidth="1"/></cols>').replace(/<sheetData>.*?<\/sheetData>/s,'<sheetData>'+lines.join('')+'</sheetData><mergeCells count="'+merges.length+'">'+merges.join('')+'</mergeCells>').replace(/<pageMargins[^>]*\/>/,'<pageMargins left=".394" right=".394" top=".45" bottom=".45" header=".15" footer=".15"/>');
}
async function buildPairBook(cars){
 const zip=await JSZip.loadAsync(pairTemplate,{base64:true}),read=p=>zip.file(p).async('string');
 const template=await read('xl/worksheets/sheet1.xml'),wb='xl/workbook.xml',rels='xl/_rels/workbook.xml.rels',types='[Content_Types].xml';
 let relationships=await read(rels),content=await read(types),sheets='',titles='',start=0;
 const featureNotes=new Map(),itpTypes=new Set();
 for(const row of alignedRows(cars)){
  for(const type of row.itpTypes)if(type!==null)itpTypes.add(type);
  const names=cars.filter((car,index)=>Boolean(row.values[index]?.trim())).map(car=>car.name);
  if(names.length&&names.length<cars.length)featureNotes.set(row.key,names);
 }
 const showItpTypes=itpTypes.size>1;
 const count=Math.ceil(cars.length/4);
 for(let i=0;start<cars.length;i++){
  let size=Math.min(4,Math.ceil((cars.length-start)/Math.max(1,count-i)));
  if(sameControlGroup(cars[start+size-1],cars[start+size])){
   let end=start+size;
   while(sameControlGroup(cars[end-1],cars[end]))end++;
   if(end-start<=4)size=end-start;
   else while(sameControlGroup(cars[start+size-1],cars[start+size]))size--;
  }
  const group=cars.slice(start,start+size),n=i+1,id=i?'rIdCar'+n:'rId1',name='Spec '+(start+1)+'-'+(start+size),path='xl/worksheets/sheet'+n+'.xml';
  zip.file(path,groupSheet(template,group,featureNotes,showItpTypes));
  sheets+='<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="'+name+'" sheetId="'+n+'" r:id="'+id+'"/>';
  titles+='<definedName name="_xlnm.Print_Titles" localSheetId="'+i+'">'+name.replace(/^(.+)$/,"'$1'")+'!$9:$9</definedName>';
  if(i){relationships=relationships.replace('</Relationships>','<Relationship Id="'+id+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/'+path+'"/></Relationships>');
   content=content.replace('</Types>','<Override PartName="/'+path+'" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')}
  start+=size;
 }
 zip.file(wb,(await read(wb)).replace(/<sheets>.*?<\/sheets>/s,'<sheets>'+sheets+'</sheets>').replace(/<definedNames>.*?<\/definedNames>/s,'<definedNames>'+titles+'</definedNames>'));
 zip.file(rels,relationships);zip.file(types,content);return zip.generateAsync({type:'blob'});
}
