'use strict';
const el=id=>document.getElementById(id), cars=[], pending=new Map(), openGroups=new Set();
const sections={P:'Công trình',I:'I. Thông số cơ bản',II:'II. Kích thước',III:'III. Thiết kế cabin',IV:'IV. Thiết kế cửa tầng',V:'V. Tính năng'};
let section='I',detail=false,token=0,version=0,cache={},editing=null,created=0;
let leadTimes=new Map();
const featureDisclosure=new Map();
async function readLeadTime(car){
 try{const r=await request(car,'leadTime');return Number.isFinite(r?.days)&&r.days>0?r:null;}catch(e){return null;}
}
// Keep factory ETD per car intact: a control bank ships at the slowest member's ETD.
function effectiveLeadTime(car){
 if(!car.group)return leadTimes.get(car)||null;
 const members=car.group.filter(c=>cars.includes(c));
 const valid=members.filter(c=>Number.isFinite(leadTimes.get(c)?.days)&&leadTimes.get(c).days>0);
 if(!valid.length||valid.length!==members.length)return {days:null,reasons:['Chưa tính được ETD của tất cả thang trong nhóm.']};
 const days=Math.max(...valid.map(c=>leadTimes.get(c).days));
 const drivers=valid.filter(c=>leadTimes.get(c).days===days);
 const reasons=['ETD chung nhóm '+displayMode(car.group.mode)+' · '+groupLabel(car)+': lấy theo '+drivers.map(c=>columnLabel(cars.indexOf(c))+' ('+c.name+')').join(', ')+' ('+days+' days).'];
 for(const c of drivers)for(const reason of leadTimes.get(c).reasons||[])if(!reasons.includes(reason))reasons.push(reason);
 return {days,reasons};
}
function leadTimeCell(car,result){
 const cell=make('span',result?.days?result.days+' days':'—');cell.className='summary-etd-value'+(result?.days>35?' extended':'');
 if(result?.reasons?.length){cell.title=result.reasons.join('\n');cell.setAttribute('aria-label',cell.textContent+' · '+result.reasons.join('; '));}
 return cell;
}
function appendLeadTimeResults(body){
 const table=make('table');table.className='summary-etd-popup';
 for(const [i,car] of cars.entries()){const row=table.insertRow();row.insertCell().textContent=columnLabel(i)+' · '+car.name;const r=effectiveLeadTime(car);row.insertCell().append(leadTimeCell(car,r));const reasons=row.insertCell();reasons.textContent=r?(r.reasons||[]).join('; ')||(r.days?'Tiêu chuẩn':'Chưa tính được ETD'):'Chưa tính được ETD';}
 body.append(make('h3','ETD / Lead time'),table);
}
function make(tag,text){const n=document.createElement(tag);if(text!=null)n.textContent=text;return n}
function status(text){el('statusText').textContent=text;el('desktop-check-floating-status').textContent=text}
function request(car,action,payload){
 if(action==='set')clearErrorMarks();
 return car.ready.then(()=>new Promise((resolve,reject)=>{
  const t=++token,timer=setTimeout(()=>{pending.delete(t);reject(Error(car.name+' · Hết thời gian chờ'))},90000);
  pending.set(t,{source:car.frame.contentWindow,done:d=>{clearTimeout(timer);d.error?reject(Error(car.name+' · '+d.error)):resolve(d.result)}});
  car.frame.contentWindow.postMessage({channel:'multi-spec',token:t,action,payload},'*');
 }));
}
window.addEventListener('message',e=>{const d=e.data,p=pending.get(d?.token);if(d?.channel==='multi-result'&&p?.source===e.source){pending.delete(d.token);p.done(d)}});
window.addEventListener('message',async e=>{
 if(e.data?.channel!=='summary-operation')return;
 const car=cars.find(c=>c.frame.contentWindow===e.source);
 if(!car?.group)return;
 const main=document.querySelector('main');main.inert=true;
 try{
  if(parseInt(e.data.value,10)<car.group.length)throw Error('Hệ điều khiển cần phù hợp với số cột đã ghép.');
  await selectGroupMode(car,e.data.value);
 }catch(err){
  await request(car,'set',{key:controlKey,value:car.group.mode});status(err.message);
 }finally{main.inert=false;await refresh()}
});
function createFrame(name,frameModel){
 const frame=make('iframe');frame.title=name;frame.hidden=true;
 const ready=new Promise(resolve=>frame.addEventListener('load',resolve,{once:true}));
 frame.src='spec.html?v=R96&model='+frameModel;document.querySelector('main').append(frame);
 return {name,frame,ready};
}
function add(){
 clearErrorMarks();
 let name;do{name='P'+(++created)}while(cars.some(c=>c.name===name));const car=createFrame(name,model);cars.push(car);refreshColumnSelector(cars.length-1);
}
function current(){return cars[el('car').selectedIndex]}
function columnLabel(index){return 'Cột '+(index+1)}
function refreshColumnSelector(selected=el('car').selectedIndex){
 const select=el('car');
 select.replaceChildren(...cars.map((_,i)=>new Option(columnLabel(i),String(i))));
 select.selectedIndex=cars.length?Math.max(0,Math.min(selected,cars.length-1)):-1;
}
function show(){
 cars.forEach(c=>c.frame.hidden=!detail||c!==current());el('overview').hidden=detail;el('detailQuick').textContent=detail?'← Spec Summary':'Form đơn cột đang chọn';
 if(!detail)refresh();
}
async function refresh(viewState){
 const revision=++version;
 try{
 const target=[...cars];
 const [entries,etds]=await Promise.all([Promise.all(Object.keys(sections).map(async id=>[id,await Promise.all(target.map(c=>request(c,'fields',{section:id,group:el('featureGroup').value})))])),Promise.all(target.map(readLeadTime))]);
 if(revision!==version)return;cache=Object.fromEntries(entries);leadTimes=new Map(target.map((c,i)=>[c,etds[i]]));
 for(const bank of new Set(cars.map(c=>c.group).filter(Boolean))){
   if(internalBanks(bank,bank.mode))for(const member of bank)member.group=null;
 }
 draw(viewState);
 }catch(e){if(revision===version)status(e.message)}
}
let model,isMRL,controlKey,speedKey,stopsKey,floorsKey;
function setSummaryModel(value){
 model=value;isMRL=model==='MRL';controlKey=isMRL?'D22':'D21';speedKey=isMRL?'D17':'D15';
 stopsKey=isMRL?'D18':'D16';floorsKey=isMRL?'D19':'D17';
 const name=isMRL?'NEXIEZ-MRL Version2':'NEXIEZ-MR';
 el('modelBadge').textContent=name;el('summaryModel').value=model;window.name='spec-summary-'+model;updateModeLinks(model);document.title='Spec Summary · '+name;
}
setSummaryModel(new URLSearchParams(location.search).get('model')==='MRL'?'MRL':'MR');
const MAX_STOP_DIFFERENCE=1;
function titleInput(car){
 const input=make('input');input.value=car.name;input.title='Sửa mã thang';input.setAttribute('aria-label','Mã thang '+car.name);
 input.onfocus=()=>input.select();input.onkeydown=e=>{if(e.key==='Escape')input.value=car.name;if(['Escape','Enter'].includes(e.key))input.blur()};
 input.onchange=()=>{const name=input.value.trim();
  if(!name){status('Vui lòng nhập tên thang.');input.value=car.name;return}
  clearErrorMarks();car.name=name;car.frame.title=name;draw();
 };return input;
}
function columnCount(car){
 return Number(cache.I?.[cars.indexOf(car)]?.find(f=>f.key==='D11')?.value);
}
function internalBanks(members,mode){
 const size=parseInt(mode,10);
 return size>1&&members.every(c=>columnCount(c)>0&&columnCount(c)%size===0);
}
function internalLabel(car,size,count){
 const names=car.name.split(/[,;]/).map(v=>v.trim()).filter(Boolean);
 if(names.length===count){
  const groups=[];
  for(let i=0;i<names.length;i+=size)groups.push(names.slice(i,i+size).join(' + '));
  return groups.join(' · ');
 }
 return car.name+' · '+(count/size)+' nhóm × '+size+' thang';
}
function groupLabel(car){return car.group.map(c=>c.name).join(' + ')}
function groupMode(size){return size<=2?size+'C-2BC':size+'C-SAI-2200C'}
function displayMode(value){return String(value).replace(/^(\dC)-SAI-2200C$/,'$1-AI2200C')}
// A control bank has one rated speed. Use the intersection of the live options,
// not a fixed speed table; different capacities can have different permitted speeds.
function commonSpeedOptions(fields){
 const first=fields[0];
 if(!first||fields.some(f=>!f))return [];
 return (first.options||[]).filter(([v])=>v&&fields.every(f=>f.locked?f.value===v:f.options?.some(o=>o[0]===v)));
}
async function readBankSpeeds(members){
 return (await Promise.all(members.map(c=>request(c,'fields',{section:'I'})))).map(list=>list.find(f=>f.key===speedKey));
}
async function writeBankSpeed(members,value){
 const fields=await readBankSpeeds(members);
 if(!commonSpeedOptions(fields).some(o=>o[0]===value))throw Error('Tốc độ '+value+' m/min không khả dụng cho tất cả thang trong nhóm.');
 for(const [i,car] of members.entries())if(fields[i].value!==value)await request(car,'set',{section:'I',key:speedKey,value});
 await assertBankSpeed(members,value);
}
async function assertBankSpeed(members,value){
 const fields=await readBankSpeeds(members);
 if(fields.some(f=>f?.value!==value)||!commonSpeedOptions(fields).some(o=>o[0]===value))throw Error('Tốc độ chung chưa đồng bộ hoặc không còn hợp lệ.');
}
async function setBankSpeed(car,value){
 const bank=car.group;
 if(!bank)return request(car,'set',{section:'I',key:speedKey,value});
 await bankUpdate(car,bank,()=>writeBankSpeed(bank,value),{features:false});
}
window.addEventListener('message',async e=>{
 if(e.data?.channel!=='summary-speed')return;
 const car=cars.find(c=>c.frame.contentWindow===e.source);if(!car)return;
 const main=document.querySelector('main');main.inert=true;
 try{await setBankSpeed(car,e.data.value);status('Đã cập nhật tốc độ '+(car.group?groupLabel(car):car.name)+'.')}
 catch(err){status(err.message)}finally{await refresh();main.inert=false}
});
async function writeGroupMode(members,value){
 const fields=await Promise.all(members.map(c=>request(c,'fields',{section:'I'})));
 const prior=fields.map(list=>list.find(f=>f.key===controlKey));
 if(prior.some(f=>!f||f.locked||!f.options?.some(o=>o[0]===value)))throw Error('Kiểu điều khiển chưa khả dụng cho tất cả thang trong nhóm.');
 const results=await Promise.allSettled(members.map(c=>request(c,'set',{key:controlKey,value})));
 if(results.some(r=>r.status==='rejected')){
  await Promise.all(members.map((c,i)=>request(c,'set',{key:controlKey,value:prior[i].value})));
  throw Error('Chưa cập nhật được hệ điều khiển chung.');
 }
}
async function sharedFields(car){
 return (await request(car,'fields',{section:'V',group:'all'})).filter(f=>f.sharedCode);
}
async function copyGroupFeatures(primary,members){
 const fields=await sharedFields(primary);
 for(const car of members.filter(c=>c!==primary)){
  const target=new Map((await sharedFields(car)).map(f=>[f.key,f]));
  for(const f of fields){
   if(f.sharedCode==='ITP type')continue;
   const peer=target.get(f.key);
   if(peer?.value===f.value)continue;
   if(!peer||peer.locked)throw Error(car.name+' · '+f.sharedCode+' chưa phù hợp cấu hình.');
   await request(car,'set',{section:'V',key:f.key,value:f.value,featureGroup:f.featureTable});
  }
  const type=fields.find(f=>f.sharedCode==='ITP type');
  if(type){
   const peer=(await sharedFields(car)).find(f=>f.key===type.key);
   if(peer?.value!==type.value)await request(car,'set',{section:'V',key:type.key,value:type.value,featureGroup:type.featureTable});
  }
  const actual=new Map((await sharedFields(car)).map(f=>[f.key,f.value]));
  for(const f of fields)if(actual.get(f.key)!==f.value)throw Error(car.name+' · '+f.sharedCode+' chưa đồng bộ.');
 }
}
async function readPivs(car){
 return (await request(car,'fields',{section:'IV'})).filter(f=>f.pivKey===f.key);
}
async function syncPivs(primary,members,before,joining){
 const source=await readPivs(primary),peer=members.find(c=>c!==primary);
 const other=await readPivs(peer);
 for(const f of source){
  const g=other.find(g=>g.key===f.key);
  if(!g||f.value===g.value)continue;
  if(joining&&!f.doublePiv&&g.doublePiv){
   await request(primary,'set',{key:f.key,value:g.value,section:'IV',mirrorPiv:true});
  }else if(f.doublePiv||g.doublePiv&&before?.some(v=>v.key===f.key&&v.doublePiv)){
   await request(peer,'set',{key:f.key,value:f.value,section:'IV',mirrorPiv:true});
  }
 }
}
async function bankUpdate(primary,members,apply,options={}){
 const saved=await Promise.all(members.map(c=>request(c,'snapshot')));
 try{
  const before=options.piv?await readPivs(primary):null;
  await apply();
  if(options.features!==false&&!isMRL)await copyGroupFeatures(primary,members);
  if(options.piv)await syncPivs(primary,members,before,options.joining);
 }catch(err){await Promise.all(members.map((c,i)=>request(c,'restore',saved[i])));throw err}
}
async function editFeature(car,f,payload){
 if(payload.key===speedKey){
  await setBankSpeed(car,payload.value);
 }else if(car.group?.mode==='2C-2BC'&&f.pivKey===payload.key){
  await bankUpdate(car,car.group,()=>request(car,'set',payload),{piv:true,features:false});
 }else if(!isMRL&&car.group&&f.sharedCode){
  await bankUpdate(car,car.group,()=>request(car,'set',payload));
 }else await request(car,'set',payload);
}
window.addEventListener('message',async e=>{
 if(e.data?.channel!=='summary-piv')return;
 const car=cars.find(c=>c.frame.contentWindow===e.source);if(!car)return;
 const main=document.querySelector('main');main.inert=true;
 try{await editFeature(car,{pivKey:e.data.payload.key},e.data.payload)}
 catch(err){status(err.message)}finally{main.inert=false;await refresh()}
});
window.addEventListener('message',async e=>{
 if(e.data?.channel!=='summary-feature'||isMRL)return;
 const car=cars.find(c=>c.frame.contentWindow===e.source);if(!car)return;
 const main=document.querySelector('main');main.inert=true;
 try{await editFeature(car,{sharedCode:true},e.data.payload)}
 catch(err){status(err.message)}finally{await refresh();main.inert=false}
});
async function selectGroupMode(car,value){
 const speed=(await readBankSpeeds([car]))[0]?.value;
 await bankUpdate(car,car.group,async()=>{await writeBankSpeed(car.group,speed);await writeGroupMode(car.group,value);await assertBankSpeed(car.group,speed)},{piv:value==='2C-2BC',joining:true});
 car.group.mode=value;
}
async function changeGroup(car,other){
 const members=other?[...(car.group||[car]),other]:car.group;
 const size=other?members.length:1;
 if(size>(isMRL?4:2)||other?.group)return false;
 const main=document.querySelector('main');main.inert=true;
 try{
  let value=groupMode(size);
  if(other){
   const fields=await request(car,'fields',{section:'I'});
   const currentMode=fields.find(f=>f.key===controlKey)?.value;
   if(parseInt(currentMode,10)>=size)value=currentMode;
  }
  const within=other&&internalBanks(members,value);
  if(other&&!within){
   const speed=(await readBankSpeeds([car]))[0]?.value;
   await bankUpdate(car,members,async()=>{await writeBankSpeed(members,speed);await writeGroupMode(members,value);await assertBankSpeed(members,speed)},{piv:value==='2C-2BC',joining:true});
  }
  else await writeGroupMode(members,value);
  const selected=current(),ordered=cars.filter(c=>members.includes(c)),first=cars.indexOf(ordered[0]);
  if(other&&!within){
   const rest=cars.filter(c=>!members.includes(c));
   rest.splice(first,0,...ordered);cars.splice(0,cars.length,...rest);
  }
  ordered.mode=value;
  for(const c of members)c.group=other&&!within?ordered:null;
  refreshColumnSelector(cars.indexOf(selected));
  status(within?'Đã đặt '+displayMode(value)+' riêng trong từng cột.':other?'Đã ghép '+groupLabel(car)+' · '+displayMode(value):'Đã tách nhóm · 1C-2BC');return true;
 }catch(e){status(e.message);return false}finally{main.inert=false;await refresh()}
}
function renderGroup(cell,car,view){
 if(!car.group){
  const fields=cache.I?.[cars.indexOf(car)]||[];
  const size=parseInt(fields.find(f=>f.key===controlKey)?.value,10),count=columnCount(car);
  if(size>1&&count>0&&count%size===0)cell.append(make('strong',internalLabel(car,size,count)));
 }
 if(car.group){
  cell.colSpan=view.filter(c=>c.group===car.group).length;
  cell.append(make('strong',displayMode(car.group.mode)+' · '+groupLabel(car)));
  const split=make('button','Tách nhóm');split.onclick=()=>changeGroup(car,null);cell.append(split);
 }
 if((car.group?.length||1)>=(isMRL?4:2))return;
 const select=make('select');select.setAttribute('aria-label','Group với thang khác cho '+car.name);
 select.add(new Option('Group với…',''));
 cars.forEach((other,i)=>{if(other!==car&&!other.group)select.add(new Option(other.name,String(i)))});
 select.onchange=()=>{if(select.value!=='')changeGroup(car,cars[Number(select.value)])};cell.append(select);
}
// All input columns share the same geometry; export pagination is independent.
function updateSummaryFit(){
 const matrix=document.querySelector('.matrix');
 if(!matrix||!matrix.clientWidth)return;
 const label=Math.min(300,Math.max(220,Math.round(matrix.clientWidth*.24)));
 const minimum=label+cars.length*150;
 matrix.style.setProperty('--summary-label-width',label+'px');
 matrix.classList.toggle('wide-matrix',minimum>matrix.clientWidth);
}
function summaryTableColumns(table,count){
 const cols=make('colgroup');
 for(let i=0;i<=count;i++){const col=make('col');if(!i)col.className='summary-label-col';cols.append(col)}
 table.append(cols);return table;
}
window.addEventListener('resize',updateSummaryFit);
// Rebuilding the entire matrix temporarily collapses the page. Preserve the
// document AND matrix scroll, using the edited field as a stable anchor.
function rememberSummaryView(input){
 const matrix=document.querySelector('.matrix');
 const panel=input?.closest('section[id^="part-"]');
 const anchor=panel&&input.dataset.fieldKey?{id:panel.id,key:input.dataset.fieldKey,car:input.dataset.carIndex,y:input.getBoundingClientRect().top}:null;
 if(panel){section=panel.id.slice(5);el('featureGroup').hidden=section!=='V';document.querySelectorAll('nav.sections button').forEach(b=>b.classList.toggle('active',b.dataset.section===section));}
 return {x:window.scrollX,y:window.scrollY,left:matrix.scrollLeft,top:matrix.scrollTop,anchor};
}
function restoreSummaryView(state,matrix){
 window.scrollTo({left:state.x,top:state.y,behavior:'instant'});
 matrix.scrollLeft=state.left;matrix.scrollTop=state.top;
 const a=state.anchor;if(!a)return;
 const panel=el(a.id),selector='[data-field-key="'+CSS.escape(a.key)+'"][data-car-index="'+CSS.escape(a.car)+'"]';
 let node=panel?.querySelector('select'+selector+',input'+selector)||panel?.querySelector(selector);
 if(!node||!node.getClientRects().length)return;
 const delta=node.getBoundingClientRect().top-a.y;
 if(matrix.classList.contains('wide-matrix'))matrix.scrollTop+=delta;
 else window.scrollBy({top:delta,left:0,behavior:'instant'});
}
function draw(viewState){
 const state=viewState||rememberSummaryView();
 if(!el('carErrors').hidden)renderErrors();
 const from=0,to=cars.length,view=cars,matrix=document.querySelector('.matrix');
 matrix.replaceChildren();
 matrix.style.setProperty('--car-min-total',150*view.length+'px');
 updateSummaryFit();
 for(const id of Object.keys(sections)){
  if(!cache[id])continue;
  const panel=id==='P'?el('sharedProject'):make('section');
  if(id==='P')panel.replaceChildren();else{panel.id='part-'+id;panel.append(make('h2',sections[id]));matrix.append(panel)}
  if(id==='P'){
   for(const f of cache.P[0]||[]){
    const label=make('label',f.title),input=make('input');input.value=project.get(f.key)??f.value;
    project.set(f.key,input.value);input.onchange=async()=>{project.set(f.key,input.value);
     try{await syncProject();status('Đã cập nhật thông tin công trình.')}catch(e){status(e.message)}};
    label.append(input);panel.append(label);
   }
  }else drawSection(id,cache[id],panel,from,to,view);
 }
 if(!el('carErrors').hidden)markFailedFields();
 restoreSummaryView(state,matrix);
}
const project=new Map([['C8','Thang máy']]);
async function syncProject(targets=cars){
 await Promise.all(targets.map(async c=>{for(const [key,value] of project)await request(c,'set',{key,value,section:'P'})}));
}
function inactiveValue(item,section){
 if(!item||item.locked)return true;
 if(section==='V'&&item.value==='0'&&item.options?.some(([value,label])=>value==='0'&&label==='Không'))return true;
 const label=item.options?.find(([v])=>v===item.value)?.[1];
 const empty=/^(?:|[-—–]+|n\/?a|none|no|not applicable|not applied|not provided|not supplied|không|không có|không áp dụng|không cung cấp|chưa chọn|— chọn)$/i;
 return [item.value,label].some(v=>v!=null&&empty.test(String(v).trim().split(/\s+\/\s+|\s*[(]/)[0].trim()));
}
// Display-only labels: never change source keys, options or exported descriptions.
function summaryFieldTitle(f,section){
 const title=String(f.title||'');
 if(section==='V'&&/^FE-C30(?:\s*·|\s*$)/i.test(title))return 'FE-C30 · Vận hành khẩn cấp cho lực lượng chữa cháy QC06';
 const dimensionCells=isMRL?['D28','D29','D30','D31','D32','D33','D34','D35','D36','D37']:['D29','D30','D31','D32','D33','D34','D35','D36','D37','D38','D39','D40','D41'];
 const dimension=section==='II'&&(dimensionCells.includes(f.sourceCell||f.key)||/^(?:AA|BB|HL|JJ|HH|AH|BH|AM|BM|MH|TR|Travel|OH|PD)\s*[-–—]/i.test(title));
 if(dimension&&!/\bmm\b/i.test(title))return title+' (mm)';
 return title;
}
function drawSection(section,cache,matrix,from,to,view){
 const seen=new Map();for(const list of cache)for(const f of list)seen.set(f.key,f);
 const buckets=new Map();
 for(const f of seen.values()){
  const row=cache.map(list=>list.find(x=>x.key===f.key));
  if(!row.slice(from,to).some(Boolean))continue;
  if(el('diff').checked&&new Set(row.map((x,i)=>f.key===controlKey&&cars[i].group?groupLabel(cars[i]):x?.value)).size<2)continue;
  if(!summaryFieldTitle(f,section).toLocaleLowerCase().includes(el('search').value.trim().toLocaleLowerCase()))continue;
  const group=['II','III','IV','V'].includes(section)?f.group:sections[section];
  if(!buckets.has(group))buckets.set(group,[]);buckets.get(group).push({f,row});
 }
 const summaryTable=summaryTableColumns(make('table'),view.length),header=summaryTable.insertRow();header.append(make('th','Tên thang'));
 for(const car of view){const th=make('th');th.append(section==='I'?titleInput(car):car.name);
  const index=cars.indexOf(car);
  if(section==='I'&&index>0){const copy=make('button','⧉');copy.type='button';copy.className='copy-column';copy.title='Sao chép cấu hình '+cars[index-1].name+' sang '+car.name;copy.setAttribute('aria-label',copy.title);copy.onclick=()=>copyColumn(cars[index-1],car);th.append(copy)}
  header.append(th)}
 summaryTable.className='name-header';matrix.append(summaryTable);
 if(section==='I'){const groupTable=summaryTableColumns(make('table'),view.length);
 matrix.append(groupTable);const groupRow=groupTable.insertRow();const groupTitle=groupRow.insertCell();groupTitle.textContent='Nhóm điều khiển';const hint=make('small',isMRL?'1C: từng thang độc lập. 2C / 3C / 4C: đúng 2 / 3 / 4 thang trong một nhóm.':'1C: từng thang độc lập. 2C: đúng 2 thang trong một nhóm.');hint.textContent+=' Quantity là tổng số thang trong cột. Ví dụ P1,P3 · 2 thang · 2C là một nhóm riêng.';groupTitle.title=hint.textContent;hint.textContent=isMRL?'1C: đơn · 2C/3C/4C: nhóm theo số thang':'1C: đơn · 2C: nhóm 2 thang';hint.style.display='block';groupTitle.append(hint);
 view.forEach((c,i)=>{
  if(c.group&&view[i-1]?.group===c.group)return;
  renderGroup(groupRow.insertCell(),c,view);
 });
 const etdRow=groupTable.insertRow();etdRow.className='summary-etd-row';etdRow.insertCell().textContent='ETD / Lead time';
 for(let i=0;i<view.length;){
  const car=view[i],cell=etdRow.insertCell();let span=1;
  if(car.group)while(i+span<view.length&&view[i+span].group===car.group)span++;
  cell.colSpan=span;cell.append(leadTimeCell(car,effectiveLeadTime(car)));i+=span;
 }
 }
 if(!buckets.size){matrix.append(make('p','Không có hạng mục phù hợp.'));return}
 for(const [group,rows] of buckets){
  const details=make('section');
  const collapsible=section==='V';
  const provided=item=>item&&((item.value==='1'&&item.options?.some(([v,label])=>v==='0'&&label==='Không'))||!inactiveValue(item,section));
  const inactive=collapsible&&rows.every(({row})=>row.slice(from,to).every(item=>!provided(item)));
  let featureToggle=null;
  if(collapsible){
   details.className='feature-fold';details.dataset.featureGroup=group;
   const heading=make('h3');const toggle=make('button');toggle.type='button';toggle.className='feature-toggle';
   const caret=make('span','▾');caret.className='fold-caret';caret.setAttribute('aria-hidden','true');
   const title=make('span',group);title.className='fold-title';
   const count=rows.filter(({row})=>row.slice(from,to).some(provided)).length;
   const meta=make('small',count+' / '+rows.length+' tính năng');
   toggle.append(caret,title,meta);heading.append(toggle);details.append(heading);featureToggle=toggle;
  }else if(group!==sections[section]&&(group!==rows[0].f.title||rows.length>1)){
   const heading=make('h3',group);
   if(section==='V'&&rows.every(({row})=>row.slice(from,to).every(item=>inactiveValue(item,section))))heading.classList.add('inactive-group');
   details.append(heading);
  }
  const table=summaryTableColumns(make('table'),view.length);
  for(const {f,row} of rows){
   const tr=table.insertRow();
   if(!isMRL&&section==='II'&&['D36','D37','D38'].includes(f.sourceCell||f.key))tr.classList.add('mr-room-dimension');
   if(section==='IV'&&/cửa chống cháy|fireproof\s*door/i.test(f.title))tr.classList.add('fireproof-door-row');
   const displayTitle=summaryFieldTitle(f,section);
   const label=tr.insertCell();label.textContent=displayTitle;
   const source=row[from];
   if(source&&!source.locked&&view.length>1&&f.key!==controlKey){const quick=make('button','»');quick.type='button';quick.className='quick-copy';quick.title='Copy '+view[0].name+' sang tất cả các cột bên phải';quick.setAttribute('aria-label',quick.title+' · '+displayTitle);quick.onclick=()=>copyRow(f.key,section,view);label.append(quick)}
   row.slice(from,to).forEach((item,j)=>{const car=view[j],partner=f.key===controlKey&&car.group;

    const sharedSpeed=f.key===speedKey&&car.group;
    if(sharedSpeed&&view[j-1]?.group===car.group)return;
    const sharedQty=item?.pivKey&&item.key!==item.pivKey&&item.doublePiv&&car.group?.mode==='2C-2BC'&&car.group.every(c=>row[cars.indexOf(c)]?.doublePiv);
    if(sharedQty&&view[j-1]?.group===car.group)return;
    const cell=tr.insertCell();cell.dataset.fieldKey=f.key;cell.dataset.carIndex=String(from+j);

    if(sharedSpeed){
     cell.colSpan=view.filter(c=>c.group===car.group).length;
     const members=car.group.map(c=>row[cars.indexOf(c)]),options=commonSpeedOptions(members);
     const speed=members[0]?.value,consistent=members.every(f=>f?.value===speed)&&options.some(o=>o[0]===speed);
     const input=make('select');input.dataset.fieldKey=f.key;input.dataset.carIndex=String(from+j);
     input.setAttribute('aria-label',groupLabel(car)+' · Tốc độ chung');
     if(!consistent){const opt=new Option(options.length?'— Chọn tốc độ chung —':'Không có tốc độ chung phù hợp','');opt.disabled=true;input.add(opt)}
     for(const [value,label] of options)input.add(new Option(label,value));
     input.value=consistent?speed:'';input.disabled=!options.length;
     if(!consistent)input.style.borderColor='#b42318';
     input.onchange=async()=>{const state=rememberSummaryView(input),main=document.querySelector('main');main.inert=true;
      try{await setBankSpeed(car,input.value);status('Đã cập nhật tốc độ chung · '+groupLabel(car))}
      catch(e){status(e.message)}finally{await refresh(state);main.inert=false}
     };cell.append(input);return;
    }
    if(sharedQty){cell.colSpan=view.filter(c=>c.group===car.group).length;cell.textContent='1 hộp / 2 thang / tầng · '+groupLabel(car);return}
    if(!item){cell.textContent='—';cell.classList.add('empty-slot');return}
    const input=make(item.options?'select':'input');
    if(f.key==='D11'){input.type='number';input.min='1';input.step='1'}
    if(item.options){for(const [value,label] of item.options){if(partner&&parseInt(value,10)<partner.length)continue;input.add(new Option(f.key===controlKey?displayMode(value):label,value));}
     if(!item.options.some(o=>o[0]===item.value))input.add(new Option(item.value,item.value));}
    input.dataset.fieldKey=item.key;input.dataset.carIndex=String(from+j);input.value=item.value;input.disabled=item.locked;input.classList.toggle('inactive-value',inactiveValue(item,section));if(section==='V'&&!item.locked&&!inactiveValue(item,section))input.classList.add('feature-provided');input.setAttribute('aria-label',car.name+' · '+displayTitle);
    input.onchange=async()=>{const state=rememberSummaryView(input),main=document.querySelector('main');main.inert=true;
     try{if(partner)await selectGroupMode(car,input.value);else await editFeature(car,item,{key:f.key,value:input.value,section,featureGroup:item.featureTable||el('featureGroup').value});status(car.name+' ✓')}
     catch(e){status(e.message)}finally{await refresh(state);main.inert=false}
    };cell.append(input);
   });
  }
  details.append(table);
  if(featureToggle){
   const key=group;
   const setOpen=open=>{table.hidden=!open;featureToggle.setAttribute('aria-expanded',String(open));featureToggle.querySelector('.fold-caret').textContent=open?'▾':'▸';};
   setOpen(featureDisclosure.get(key)??!inactive);
   featureToggle.onclick=()=>{const open=table.hidden;featureDisclosure.set(key,open);setOpen(open);};
  }
  matrix.append(details);
 }
}
async function copyTransaction(members,action){
 const main=document.querySelector('main');main.inert=true;
 let saved;
 try{
  saved=await Promise.all(members.map(c=>request(c,'draft')));
  const message=await action();await syncProject(members);status(message+' · Chưa kiểm tra');clearErrorMarks();
 }catch(e){
  if(saved){const results=await Promise.allSettled(members.map((c,i)=>request(c,'openDraft',saved[i])));
   status(e.message+(results.some(r=>r.status==='rejected')?' · Khôi phục chưa đầy đủ, kiểm tra lại.':' · Đã hoàn tác.'));}
  else status(e.message);
 }finally{await refresh();main.inert=false}
}
async function copyRow(key,section,view){
 const [source,...targets]=view,affected=[...new Set([source,...targets].flatMap(c=>c.group||[c]))];
 await copyTransaction(affected,async()=>{
  const group=el('featureGroup').value;
  const read=async c=>(await request(c,'fields',{section,group})).find(f=>f.key===key);
  const f=await read(source);if(!f||f.locked)throw Error('Không thể copy trường nguồn.');
  const done=new Set(),skipped=[];let count=0;
  for(const car of targets){
   const peer=await read(car);
   if(!peer||peer.locked||key==='D11'&&car.group||peer.options&&!peer.options.some(o=>o[0]===f.value)){skipped.push(car.name);continue}
   const scope=car.group&&(key===speedKey||(!isMRL&&f.sharedCode))?car.group:car;
   if(done.has(scope))continue;done.add(scope);
   if(peer.value!==f.value)await editFeature(car,peer,{key,value:f.value,section,featureGroup:peer.featureTable||group});
   const actual=await read(car);if(actual?.value!==f.value)throw Error(car.name+' · Giá trị copy không được giữ nguyên.');count++;
  }
  return 'Copy '+source.name+' · '+count+' cột/nhóm'+(skipped.length?' · Bỏ qua không tương thích: '+skipped.join(', '):'');
 });
}
async function copyColumn(source,target){
 if(!confirm('Copy cấu hình '+source.name+' sang '+target.name+'? Tên thang và nhóm điều khiển hiện tại được giữ nguyên; thông số dùng chung trong nhóm sẽ được đồng bộ.'))return;
 const affected=[...new Set([...(target.group||[target])])];
 await copyTransaction(affected,async()=>{
  const [draft,fields]=await Promise.all([request(source,'draft'),request(target,'fields',{section:'I'})]);
  const operation=fields.find(f=>f.key===controlKey)?.value;
  // Bank membership belongs to the Summary, not the copied draft.
  draft.values[controlKey]=operation;
  if(target.group){draft.values.D11=fields.find(f=>f.key==='D11').value;}
  await request(target,'openDraft',draft);
  const actual=(await request(target,'fields',{section:'I'})).find(f=>f.key===controlKey)?.value;
  if(actual!==operation)throw Error(target.name+' · Không giữ được hệ điều khiển hiện tại.');
  if(target.group){
   const speed=(await readBankSpeeds([target]))[0]?.value;
   await writeBankSpeed(target.group,speed);
   if(!isMRL)await copyGroupFeatures(target,target.group);
   if(target.group.mode==='2C-2BC')await syncPivs(target,target.group,null,true);
   await assertBankSpeed(target.group,speed);
  }
  return 'Đã copy '+source.name+' → '+target.name;
 });
}
function edit(row,index,section){
 editing={row,index,section};const f=row[index];el('caption').textContent=f.title;
 const input=make(f.options?'select':'input');if(f.options)for(const [v,t] of f.options)input.add(new Option(t,v));
 if(f.key==='D11'){input.type='number';input.min='1';input.step='1'}
 input.value=f.value;el('entry').replaceChildren(input);el('picks').replaceChildren();
 cars.forEach((c,i)=>{const label=make('label'),box=make('input');box.type='checkbox';box.checked=i===index;label.append(box,c.name);el('picks').append(label)});
 function eligible(){[...el('picks').querySelectorAll('input')].forEach((box,i)=>{const x=row[i];box.disabled=([controlKey,speedKey].includes(f.key)&&!!cars[i].group)||!x||x.locked||!!x.options&&!x.options.some(o=>o[0]===input.value);if(box.disabled)box.checked=false})}
 input.oninput=eligible;eligible();el('editor').showModal();
}
el('cancel').onclick=()=>el('editor').close();
el('apply').onclick=async()=>{
 const value=el('entry').firstChild.value,key=editing.row[editing.index].key,selected=[...el('picks').querySelectorAll('input')].flatMap((x,i)=>x.checked&&!x.disabled?[cars[i]]:[]);
 if(!selected.length)return;
 el('editor').close();const f=editing.row[editing.index],main=document.querySelector('main');
 main.inert=true;
 try{
  const done=new Set();
  for(const car of selected){
   const scope=f.sharedCode&&car.group?car.group:car;if(done.has(scope))continue;
   await editFeature(car,f,{key,value,section:editing.section,featureGroup:f.featureTable||el('featureGroup').value});done.add(scope);
  }
  status('Đã áp dụng lựa chọn.');
 }catch(err){status(err.message)}finally{await refresh();main.inert=false}

};
for(const [id,label] of Object.entries(sections)){const btn=make('button',label);btn.dataset.section=id;btn.classList.toggle('active',id===section);btn.onclick=()=>{section=id;el('search').value='';el('featureGroup').hidden=id!=='V';document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('active',b===btn));el('part-'+id)?.scrollIntoView({block:'start'})};document.querySelector('nav').append(btn)}
for(const [key,name] of Object.entries({Emergency:'Khẩn cấp',Operation:'Vận hành',Door:'Cửa',Signal:'Tín hiệu',Traffic:'Giao thông'}))el('featureGroup').add(new Option(name,'TblFeature_'+key));
el('featureGroup').insertBefore(new Option('Tất cả nhóm tính năng','all'),el('featureGroup').firstChild);el('featureGroup').value='all';el('featureGroup').onchange=refresh;el('search').oninput=draw;el('diff').onchange=draw;
el('car').onchange=()=>{show()};
el('detailQuick').onclick=()=>{detail=!detail;show()};
el('summaryModel').onchange=e=>{const next=e.target.value;e.target.value=model;if(next!==model)openSpecMode('summary',next)};
el('addCar').onclick=async()=>{add();await syncProject([current()]);show()};
el('renameCar').onclick=()=>{const c=current(),name=prompt('Mã thang',c.name)?.trim();if(!name)return;clearErrorMarks();c.name=name;c.frame.title=name;show()};
el('deleteCar').onclick=async()=>{const i=el('car').selectedIndex;if(cars.length===1||!confirm('Xóa '+columnLabel(i)+' ('+current().name+')?'))return;if(cars[i].group&&!await changeGroup(cars[i],null))return;clearErrorMarks();cars[i].frame.remove();cars.splice(i,1);refreshColumnSelector(i);show()};
let lastFailures=[];
function locate(car,issue){
 const index=cars.indexOf(car);
 for(const id of Object.keys(sections)){
  const field=cache[id]?.[index]?.find(f=>f.key===issue.cell||Boolean(issue.cell)&&f.sourceCell===issue.cell);
  if(field)return {id,field};
 }
}
async function jumpToIssue(car,issue){
 const index=cars.indexOf(car);if(index<0)return;
 el('car').selectedIndex=index;el('search').value='';el('diff').checked=false;
 if(el('featureGroup').value!=='all'){el('featureGroup').value='all';await refresh()}
 const target=locate(car,issue);if(!target){status('Lỗi này chưa có địa chỉ ô nhập trong Spec Summary.');return}
 section=target.id;el('featureGroup').hidden=section!=='V';draw();
 const tab=[...document.querySelectorAll('nav.sections button')].find(b=>b.textContent===sections[section]);
 document.querySelectorAll('nav.sections button').forEach(b=>b.classList.toggle('active',b===tab));
 const scope=el('part-'+section);
 const selector='[data-car-index="'+index+'"][data-field-key="'+CSS.escape(target.field.key)+'"]';
 let cell=scope.querySelector('input'+selector+',select'+selector)||scope.querySelector(selector);
 if(!cell&&car.group)cell=scope.querySelector('[data-car-index="'+cars.indexOf(cars.find(c=>c.group===car.group))+'"][data-field-key="'+CSS.escape(target.field.key)+'"]');
 if(!cell)return;
 openFeatureForError(cell);
 cell.scrollIntoView({block:'center',inline:'nearest'});cell.focus({preventScroll:true});
 for(const node of [cell,tab].filter(Boolean)){node.classList.remove('error-pulse');void node.offsetWidth;node.classList.add('error-pulse')}
}
function clearErrorMarks(){
 el('carErrors').hidden=true;
 document.querySelectorAll('.has-spec-error,.error-pulse').forEach(node=>node.classList.remove('has-spec-error','error-pulse'));
}
function openFeatureForError(node){
 const fold=node.closest('.feature-fold');const toggle=fold?.querySelector('.feature-toggle');
 if(toggle?.getAttribute('aria-expanded')==='false')toggle.click();
}
function markFailedFields(){
 const failedSections=new Set();
 for(const {car,errors} of lastFailures){
  const index=cars.indexOf(car);
  for(const issue of errors){
   const target=issue.cell&&locate(car,issue);if(!target)continue;
   failedSections.add(target.id);
   const panel=el('part-'+target.id);
   const key=CSS.escape(target.field.key);
   const selector='[data-car-index="'+index+'"][data-field-key="'+key+'"]';
   panel?.querySelectorAll(selector).forEach(node=>{node.classList.add('has-spec-error');openFeatureForError(node)});
  }
 }
 document.querySelectorAll('nav.sections button').forEach(tab=>{
  tab.classList.toggle('has-spec-error',[...failedSections].some(id=>sections[id]===tab.textContent));
 });
}
function renderErrors(){
 const panel=el('carErrors'),grid=make('div');grid.className='error-grid';
 const visible=cars;grid.style.gridTemplateColumns='240px repeat('+visible.length+',minmax(150px,1fr))';grid.style.minWidth=(240+150*visible.length)+'px';
 grid.append(make('strong','Kết quả kiểm tra · '+lastFailures.length+'/'+cars.length+' thang có lỗi'));
 for(const car of visible){
  const column=make('div'),errors=lastFailures.find(f=>f.car===car)?.errors||[];
  column.append(make('strong',car.name+' · '+(errors.length?errors.length+' lỗi':'Đạt')));
  for(const e of errors){const issue=typeof e==='string'?{message:e,cell:''}:e;
   const button=make(issue.cell?'button':'p',issue.message);if(issue.cell)button.onclick=()=>jumpToIssue(car,issue);column.append(button)}
  grid.append(column);
 }
 panel.replaceChildren(grid);
}
async function reportErrors(failures){
 if(el('featureGroup').value!=='all'){el('featureGroup').value='all';await refresh()}
 lastFailures=failures;detail=false;cars.forEach(c=>c.frame.hidden=true);el('overview').hidden=false;
 const panel=el('carErrors');panel.hidden=false;draw();panel.scrollIntoView({block:'start'});status('Cần xử lý lỗi của '+failures.map(f=>f.car.name).join(', '));
}
function topServiceFloor(list){
 const last=String(list??'').normalize('NFC').trim().split(/[~,;\n]+/).map(s=>s.trim()).filter(Boolean).pop()||'';
 // Numeric ranges used by the form include 1-6 and B2-12. Preserve names like PH-A.
 const range=last.match(/^.+?\s+(?:[-–—]|đến|to)\s+(.+)$/iu)||last.match(/^(?:[\p{L}]*-?\d+[\p{L}]*|G|M)[-–—]([\p{L}]*-?\d+[\p{L}]*|[\p{L}]+)$/iu);
 return (range?.[1]||last).trim();
}
function validatePair(a,b,fieldsA,fieldsB,maxStopDifference){
 const issues=[];
 function values(key){
  return [fieldsA,fieldsB].map(fields=>String(fields.find(f=>f.key===key||f.sourceCell===key)?.value??'').trim());
 }
 function mismatch(cell,label,values,rule){
  issues.push({cell,message:label+' · '+a.name+' = '+(values[0]||'chưa nhập')+', '+b.name+' = '+(values[1]||'chưa nhập')+'. '+rule});
 }
 const speeds=values(speedKey);
 if(speeds.some(v=>!v)||Number(speeds[0])!==Number(speeds[1]))mismatch(speedKey,'Tốc độ',speeds,'Hai thang cùng nhóm phải có cùng tốc độ.');
 const stops=values(stopsKey);
 if(stops.some(v=>!v||!Number.isInteger(Number(v)))||Math.abs(Number(stops[0])-Number(stops[1]))>maxStopDifference)
  mismatch(stopsKey,'Số điểm dừng',stops,'Hai thang cùng nhóm được chênh tối đa '+maxStopDifference+' điểm dừng.');
 const tops=values(floorsKey).map(topServiceFloor);
 if(tops.some(v=>!v)||tops[0].toUpperCase()!==tops[1].toUpperCase())mismatch(floorsKey,'Tầng trên cùng',tops,'Tên tầng trên cùng của hai thang phải giống nhau.');
 return issues;
}
function checkPopup(title,failures=[],canExport=false){
 const dialog=el('checkExportDialog'),body=el('checkExportBody');
 el('checkExportTitle').textContent=title;body.replaceChildren();
 for(const {car,errors} of failures){
  body.append(make('strong',car.name));
  for(const error of errors){
   const issue=typeof error==='string'?{message:error}:error;
   const node=make(issue.cell?'button':'p',issue.message);
   if(issue.cell)node.onclick=()=>{dialog.close();jumpToIssue(car,issue)};
   body.append(node);
  }
 }
 if(canExport)body.append(make('p',cars.map(c=>c.name).join(' · ')+' · '+(el('only-vi').checked?'Vietnamese Only':'Song ngữ')+(el('export-tech-details').checked?' · Có trích xuất chi tiết':'')+' · Chọn khổ giấy và định dạng xuất'));
 if(canExport||failures.some(f=>cars.includes(f.car)))appendLeadTimeResults(body);
 for(const id of ['confirmExport','confirmPdf','exportOrientation'])el(id).hidden=!canExport;
 dialog.returnValue='';if(!dialog.open)dialog.showModal();
 return new Promise(resolve=>dialog.addEventListener('close',()=>resolve(canExport?dialog.returnValue:''),{once:true}));
}
async function run(format){
 let preview;
 const main=document.querySelector('main'),book=[],failures=[],groupFields=new Map();main.inert=true;clearErrorMarks();
 const dialog=el('checkExportDialog');
 el('checkExportTitle').textContent='Đang kiểm tra Spec';
 el('checkExportBody').replaceChildren();
 for(const id of ['confirmExport','confirmPdf','exportOrientation'])el(id).hidden=true;
 el('confirmPdf').onclick=e=>{preview=window.open('','_blank');if(!preview){e.preventDefault();el('checkExportBody').append(make('p','Hãy cho phép mở tab để xem bản in PDF.'));}else{preview.document.title='Spec Summary';preview.document.body.textContent='Đang tạo bản in…'}};
 dialog.returnValue='';dialog.showModal();
 try{
  for(const car of cars){
   if(!dialog.open)return;
   const progress='Đang kiểm tra '+car.name+'…';
   el('checkExportBody').textContent=progress;status(progress);let errors=[];leadTimes.set(car,null);
   try{
    await syncProject([car]);leadTimes.set(car,await readLeadTime(car));
    errors=await request(car,'validate');
    {const fields=await request(car,'fields',{section:'I'});groupFields.set(car,fields);
     const field=fields.find(x=>x.key===controlKey);
     const count=Number(fields.find(f=>f.key==='D11')?.value),required=parseInt(field?.value,10);
     if(car.group&&field?.value!==car.group.mode)errors.push({cell:controlKey,message:'Hệ điều khiển chưa đồng bộ với nhóm '+groupLabel(car)+'.'});
     if(!Number.isSafeInteger(count)||count<1)errors.push({cell:'D11',message:'Nhập số lượng thang là số nguyên dương.'});
     else if(!car.group&&required>1&&count%required!==0)errors.push({cell:'D11',message:car.name+' có '+count+' thang. Một nhóm '+displayMode(field.value)+' yêu cầu đúng '+required+' thang. Có thể cấu hình nhiều nhóm đủ số thang trong cùng cột.'});}
   }catch(e){errors.push(e.message)}
   if(errors.length)failures.push({car,errors});
  }
  for(const bank of new Set(cars.map(c=>c.group).filter(g=>g?.mode==='2C-2BC'))){
   const [a,b]=await Promise.all(bank.map(readPivs));
   for(const f of a){
    const g=b.find(field=>field.key===f.key);
    if(!g||!(f.doublePiv||g.doublePiv)||f.value===g.value)continue;
    for(const car of bank){
     const issue={cell:f.key,message:groupLabel(car)+' dùng chung một hộp PIV đôi mỗi tầng. Chọn cùng mã hộp.'};
     const failed=failures.find(x=>x.car===car);if(failed)failed.errors.push(issue);else failures.push({car,errors:[issue]});
    }
   }
  }
  if(!isMRL){
   const banks=new Set(cars.map(c=>c.group).filter(Boolean));
   for(const bank of banks){
    const lists=await Promise.all(bank.map(sharedFields));
    const keys=new Map(lists.flat().map(f=>[f.key,f.sharedCode]));
    for(const [key,code] of keys){
     if(new Set(lists.map(list=>list.find(f=>f.key===key)?.value)).size<=1)continue;
     bank.forEach((car,i)=>{
      const cell=lists[i].some(f=>f.key===key)?key:lists[i].find(f=>f.sharedCode==='ITP')?.key||controlKey;
      const issue={cell,message:code+' chưa đồng nhất trong nhóm '+groupLabel(car)+'. Chọn lại tại mục V.'};
      const failed=failures.find(f=>f.car===car);
      if(failed)failed.errors.push(issue);else failures.push({car,errors:[issue]});
     });
    }
   }
  }
  for(const [car,fields] of groupFields){
   if(car.group&&car.group.every(c=>groupFields.has(c))){
    const total=car.group.reduce((n,c)=>n+Number(groupFields.get(c).find(f=>f.key==='D11')?.value||1),0);
    if(total!==parseInt(car.group.mode,10)){
     const issue={cell:'D11',message:groupLabel(car)+' có '+total+' thang. '+displayMode(car.group.mode)+' yêu cầu đúng '+parseInt(car.group.mode,10)+' thang.'};
     const failed=failures.find(f=>f.car===car);
     if(failed)failed.errors.push(issue);else failures.push({car,errors:[issue]});
    }
   }
   for(const mate of car.group||[]){
   if(!groupFields.has(mate)||cars.indexOf(car)>=cars.indexOf(mate))continue;
   const issues=validatePair(car,mate,fields,groupFields.get(mate),MAX_STOP_DIFFERENCE);
   if(issues.length)for(const member of [car,mate]){
    const failed=failures.find(f=>f.car===member);
    if(failed)failed.errors.push(...issues);
    else failures.push({car:member,errors:[...issues]});
   }
   }
  }
  if(!dialog.open)return;
  draw();
  if(failures.length){if(preview&&!preview.closed)preview.close();await reportErrors(failures);main.inert=false;await checkPopup('Check Spec · Cần sửa lỗi',failures);return}
  status('Check Spec đạt · '+cars.map(c=>c.name).join(', '));
  const choice=await checkPopup('Check Spec đạt · '+cars.length+' cột thang',[],true);
  if(!['xlsx','pdf'].includes(choice))return;
  format=choice;const orientation=el('exportOrientation').value==='landscape'?'landscape':'portrait';const paper=orientation==='portrait'?'A4 dọc':'A4 ngang';

  for(const car of cars){
   try{book.push({bankLabel:!car.group&&internalBanks([car],groupFields.get(car)?.find(f=>f.key===controlKey)?.value)?internalLabel(car,parseInt(groupFields.get(car).find(f=>f.key===controlKey).value,10),columnCount(car)):'',onlyVi:el('only-vi').checked,orientation,model,operation:groupFields.get(car)?.find(f=>f.key===controlKey)?.value,name:car.name,groupId:car.group?cars.indexOf(car.group[0]):null,rows:await request(car,'pairRows',{includeTechDetails:el('export-tech-details').checked})})}
   catch(e){failures.push({car,errors:[e.message]})}
  }
  if(failures.length){if(preview&&!preview.closed)preview.close();await reportErrors(failures);main.inert=false;await checkPopup('Check Spec · Cần sửa lỗi',failures);return}
  if(format==='pdf'){
   const markup=await request(cars[0],'summaryPrint',book);
   if(!preview||preview.closed)throw Error('Tab bản in đã đóng.');
   preview.document.open();preview.document.write(markup);preview.document.close();preview.focus();
   status('Đã mở bản in '+paper+'. Chọn In / Save as PDF để lưu.');
  }else{const blob=await request(cars[0],'pairBook',book),url=URL.createObjectURL(blob),a=make('a');a.download=(isMRL?'Spec_Summary_NEXIEZ_MRL_Version2.xlsx':'Multi_Spec_NEXIEZ_MR.xlsx');a.href=url;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);status('Đã xuất Excel '+paper+'.')}

 }catch(e){if(preview&&!preview.closed)preview.close();status(e.message);main.inert=false;await checkPopup('Không thể hoàn tất kiểm tra hoặc xuất file',[{car:{name:''},errors:[e.message]}])}finally{main.inert=false;if(el('checkExportTitle').textContent==='Đang kiểm tra Spec')status('Đã đóng kiểm tra Spec.')}
}
for(const id of ['check','desktop-check-floating','m-check'])el(id).onclick=()=>run('check');
for(const id of ['exportBook','desktop-export-xlsx','m-xlsx'])el(id).onclick=()=>run('xlsx');
for(const id of ['exportPdf','desktop-export-pdf','m-pdf'])el(id).onclick=()=>run('pdf');
document.querySelectorAll('main button:not(.active)').forEach(b=>b.classList.add('btn','btn-secondary'));
for(let i=0;i<4;i++)add();el('car').selectedIndex=0;show();

function summaryDraftGroups(){
 return [...new Set(cars.map(c=>c.group).filter(Boolean))].map(bank=>({
  mode:bank.mode,members:bank.map(c=>cars.indexOf(c))
 }));
}
function validateSummaryDraft(d){
 const invalid=()=>{throw Error('File chưa đúng định dạng bản nhập Spec Summary.')};
 if(!d||d.format!=='SpecWebTool.summary'||d.version!==1||!['MR','MRL'].includes(d.model)||
  !Array.isArray(d.cars)||!d.cars.length||!Array.isArray(d.groups)||!Array.isArray(d.project)||
  !d.project.every(x=>Array.isArray(x)&&['C7','C8'].includes(x[0])&&typeof x[1]==='string')||
  !d.cars.every(c=>c&&typeof c.name==='string'&&c.name.trim()&&c.input?.model===d.model)||
  typeof d.onlyVi!=='boolean'||typeof d.details!=='boolean')invalid();
 const used=new Set();
 for(const g of d.groups){
  if(!g||!['2C-2BC',...(d.model==='MRL'?['3C-SAI-2200C','4C-SAI-2200C']:[])].includes(g.mode)||
     !Array.isArray(g.members)||g.members.length<2||g.members.length>parseInt(g.mode))invalid();
  for(const i of g.members){if(!Number.isInteger(i)||!d.cars[i]||used.has(i))invalid();used.add(i)}
 }
}
async function saveSummaryDraft(){
 const inputs=await Promise.all(cars.map(c=>request(c,'draft')));
 const d={format:'SpecWebTool.summary',version:1,model,project:[...project],onlyVi:el('only-vi').checked,
  details:el('export-tech-details').checked,groups:summaryDraftGroups(),selected:cars.indexOf(current()),
  cars:cars.map((c,i)=>({name:c.name,input:inputs[i]}))};
 const name=await request(cars[0],'draftName',project.get('C7')||'Cong trinh');
 const url=URL.createObjectURL(new Blob([JSON.stringify(d,null,2)],{type:'application/json'}));
 const a=make('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();
 setTimeout(()=>URL.revokeObjectURL(url),60000);el('draft-feedback').textContent='Đã tạo bản nhập '+name;
}
async function openSummaryDraft(file){
 const d=JSON.parse((await file.text()).replace(/^\uFEFF/,''));validateSummaryDraft(d);
 const next=[];
 try{
  for(const c of d.cars){
   const car=createFrame(c.name,d.model);next.push(car);await request(car,'openDraft',c.input);
   for(const [key,value] of d.project)await request(car,'set',{section:'P',key,value});
  }
  for(const g of d.groups){const bank=g.members.map(i=>next[i]);bank.mode=g.mode;for(const car of bank)car.group=bank}
  const nextCache={};
  for(const id of Object.keys(sections))nextCache[id]=await Promise.all(next.map(c=>request(c,'fields',{section:id,group:'all'})));
  ++version;const old=cars.splice(0,cars.length,...next);for(const car of old)car.frame.remove();
  setSummaryModel(d.model);cache=nextCache;created=cars.length;
 }catch(e){for(const c of next)if(!cars.includes(c))c.frame.remove();throw e}
 project.clear();for(const [k,v] of d.project)project.set(k,v);
 const selected=Number.isInteger(d.selected)&&cars[d.selected]?d.selected:0;
 refreshColumnSelector(selected);section='I';detail=false;
 el('detailQuick').textContent='Form đơn cột đang chọn';el('overview').hidden=false;el('featureGroup').value='all';el('featureGroup').hidden=true;
 el('diff').checked=false;el('search').value='';el('only-vi').checked=d.onlyVi;el('export-tech-details').checked=d.details;
 document.querySelectorAll('nav.sections button').forEach(btn=>btn.classList.toggle('active',btn.textContent===sections.I));
 lastFailures=[];clearErrorMarks();draw();status('Chưa kiểm tra');
 el('draft-feedback').textContent='Đã mở '+file.name+' · '+cars.length+' cột thang.';
}
async function handleSummaryFile(action){
 const main=document.querySelector('main');main.inert=true;
 try{await action()}catch(e){el('draft-feedback').textContent='Chưa hoàn tất. '+e.message}finally{main.inert=false}
}
el('save-summary-draft').onclick=()=>handleSummaryFile(saveSummaryDraft);
el('open-summary-draft').onclick=()=>el('summary-draft-file').click();
el('summary-draft-file').onchange=e=>{const f=e.target.files[0];e.target.value='';if(f)handleSummaryFile(()=>openSummaryDraft(f))};
