'use strict';
const el=id=>document.getElementById(id), cars=[], pending=new Map(), openGroups=new Set();
const sections={P:'Công trình',I:'I. Thông số cơ bản',II:'II. Kích thước',III:'III. Thiết kế cabin',IV:'IV. Thiết kế cửa tầng',V:'V. Tính năng'};
let section='I',page=0,detail=false,token=0,version=0,cache={},editing=null,created=0;
function make(tag,text){const n=document.createElement(tag);if(text!=null)n.textContent=text;return n}
function status(text){el('statusText').textContent=text}
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
function add(){
 clearErrorMarks();
 let name;do{name='P'+(++created)}while(cars.some(c=>c.name===name));const frame=make('iframe');frame.title=name;frame.hidden=true;
 const ready=new Promise(resolve=>frame.addEventListener('load',resolve,{once:true}));frame.src='spec.html?model='+model;
 cars.push({name,frame,ready});document.querySelector('main').append(frame);el('car').add(new Option(name));el('car').selectedIndex=cars.length-1;
}
function current(){return cars[el('car').selectedIndex]}
function show(){
 cars.forEach(c=>c.frame.hidden=!detail||c!==current());el('overview').hidden=detail;el('detail').textContent=detail?'Quay lại bảng nhiều thang':'Form đơn';
 if(!detail)refresh();
}
async function refresh(){
 const revision=++version;
 try{
 const snapshot=Object.fromEntries(await Promise.all(Object.keys(sections).map(async id=>[id,await Promise.all(cars.map(c=>request(c,'fields',{section:id,group:el('featureGroup').value})))])));
 if(revision!==version)return;cache=snapshot;
 for(const bank of new Set(cars.map(c=>c.group).filter(Boolean))){
   if(internalBanks(bank,bank.mode))for(const member of bank)member.group=null;
 }
 draw();
 }catch(e){if(revision===version)status(e.message)}
}
const isMRL=new URLSearchParams(location.search).get('model')==='MRL';
const model=isMRL?'MRL':'MR';
const modelName=isMRL?'NEXIEZ-MRL Version2':'NEXIEZ-MR';
el('modelBadge').textContent=modelName;
el('summaryModel').textContent=modelName+' ALPHA';
document.title='Spec Summary · '+modelName;
const controlKey=isMRL?'D22':'D21';
const speedKey=isMRL?'D17':'D15', stopsKey=isMRL?'D18':'D16', floorsKey=isMRL?'D19':'D17';
const MAX_STOP_DIFFERENCE=1;
function titleInput(car){
 const input=make('input');input.value=car.name;input.title='Sửa mã thang';input.setAttribute('aria-label','Mã thang '+car.name);
 input.onfocus=()=>input.select();input.onkeydown=e=>{if(e.key==='Escape')input.value=car.name;if(['Escape','Enter'].includes(e.key))input.blur()};
 input.onchange=()=>{const name=input.value.trim();
  if(!name){status('Vui lòng nhập tên thang.');input.value=car.name;return}
  clearErrorMarks();car.name=name;car.frame.title=name;el('car').options[cars.indexOf(car)].text=name;draw();
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
 if(car.group?.mode==='2C-2BC'&&f.pivKey===payload.key){
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
 await bankUpdate(car,car.group,()=>writeGroupMode(car.group,value),{piv:value==='2C-2BC',joining:true});
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
  if(other&&!within)await bankUpdate(car,members,()=>writeGroupMode(members,value),{piv:value==='2C-2BC',joining:true});
  else await writeGroupMode(members,value);
  const selected=current(),ordered=cars.filter(c=>members.includes(c)),first=cars.indexOf(ordered[0]);
  if(other&&!within){
   const rest=cars.filter(c=>!members.includes(c));
   rest.splice(first,0,...ordered);cars.splice(0,cars.length,...rest);
  }
  ordered.mode=value;
  for(const c of members)c.group=other&&!within?ordered:null;
  el('car').replaceChildren(...cars.map(c=>new Option(c.name)));el('car').selectedIndex=cars.indexOf(selected);
  page=Math.floor(first/4);
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
function draw(){
 if(!el('carErrors').hidden)renderErrors();
 const from=page*4,to=Math.min(from+4,cars.length),view=cars.slice(from,to),matrix=document.querySelector('.matrix');matrix.replaceChildren();
 el('prev').disabled=page===0;el('next').disabled=to===cars.length;el('range').textContent=view.map(c=>c.name).join(' · ');
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
function drawSection(section,cache,matrix,from,to,view){
 const seen=new Map();for(const list of cache)for(const f of list)seen.set(f.key,f);
 const buckets=new Map();
 for(const f of seen.values()){
  const row=cache.map(list=>list.find(x=>x.key===f.key));
  if(!row.slice(from,to).some(Boolean))continue;
  if(el('diff').checked&&new Set(row.map((x,i)=>f.key===controlKey&&cars[i].group?groupLabel(cars[i]):x?.value)).size<2)continue;
  if(!f.title.toLocaleLowerCase().includes(el('search').value.trim().toLocaleLowerCase()))continue;
  const group=['II','III','IV','V'].includes(section)?f.group:sections[section];
  if(!buckets.has(group))buckets.set(group,[]);buckets.get(group).push({f,row});
 }
 const summaryTable=make('table'),header=summaryTable.insertRow();header.append(make('th',section==='I'?'Tên thang':'Hạng mục'));
 for(const car of view){const th=make('th');th.append(section==='I'?titleInput(car):car.name);header.append(th)}
 if(section==='I'){const groupRow=summaryTable.insertRow();const groupTitle=groupRow.insertCell();groupTitle.textContent='Nhóm điều khiển';const hint=make('small',isMRL?'1C: từng thang độc lập. 2C / 3C / 4C: đúng 2 / 3 / 4 thang trong một nhóm.':'1C: từng thang độc lập. 2C: đúng 2 thang trong một nhóm.');hint.textContent+=' Quantity là tổng số thang trong cột. Ví dụ P1,P3 · 2 thang · 2C là một nhóm riêng.';hint.style.display='block';groupTitle.append(hint);
 view.forEach((c,i)=>{
  if(c.group&&view[i-1]?.group===c.group)return;
  renderGroup(groupRow.insertCell(),c,view);
 });
 }matrix.append(summaryTable);
 if(!buckets.size){matrix.append(make('p','Không có hạng mục phù hợp.'));return}
 for(const [group,rows] of buckets){
  const details=make('section');
  if(group!==sections[section]&&(group!==rows[0].f.title||rows.length>1)){
   const heading=make('h3',group);
   if(section==='V'&&rows.every(({row})=>row.slice(from,to).every(item=>inactiveValue(item,section))))heading.classList.add('inactive-group');
   details.append(heading);
  }
  const table=make('table');
  for(const {f,row} of rows){
   const tr=table.insertRow();const label=tr.insertCell();label.textContent=f.title;
   const first=row.findIndex((x,i)=>x&&!x.locked&&!(f.key===controlKey&&cars[i].group));
   if(first>=0){const bulk=make('button','Áp dụng cho…');bulk.className='bulk-edit';bulk.onclick=()=>edit(row,first,section);label.append(bulk)}
   row.slice(from,to).forEach((item,j)=>{const car=view[j],partner=f.key===controlKey&&car.group;

    const sharedQty=item?.pivKey&&item.key!==item.pivKey&&item.doublePiv&&car.group?.mode==='2C-2BC'&&car.group.every(c=>row[cars.indexOf(c)]?.doublePiv);
    if(sharedQty&&view[j-1]?.group===car.group)return;
    const cell=tr.insertCell();cell.dataset.fieldKey=f.key;cell.dataset.carIndex=String(from+j);

    if(sharedQty){cell.colSpan=view.filter(c=>c.group===car.group).length;cell.textContent='1 hộp / 2 thang / tầng · '+groupLabel(car);return}
    if(!item){cell.textContent='—';cell.classList.add('empty-slot');return}
    const input=make(item.options?'select':'input');
    if(f.key==='D11'){input.type='number';input.min='1';input.step='1'}
    if(item.options){for(const [value,label] of item.options){if(partner&&parseInt(value,10)<partner.length)continue;input.add(new Option(f.key===controlKey?displayMode(value):label,value));}
     if(!item.options.some(o=>o[0]===item.value))input.add(new Option(item.value,item.value));}
    input.dataset.fieldKey=item.key;input.dataset.carIndex=String(from+j);input.value=item.value;input.disabled=item.locked;input.classList.toggle('inactive-value',inactiveValue(item,section));input.setAttribute('aria-label',car.name+' · '+f.title);
    input.onchange=async()=>{const main=document.querySelector('main');main.inert=true;
     try{if(partner)await selectGroupMode(car,input.value);else await editFeature(car,item,{key:f.key,value:input.value,section,featureGroup:item.featureTable||el('featureGroup').value});status(car.name+' ✓')}
     catch(e){status(e.message)}finally{await refresh();main.inert=false}
    };cell.append(input);
   });
  }
  details.append(table);matrix.append(details);
 }
}
function edit(row,index,section){
 editing={row,index,section};const f=row[index];el('caption').textContent=f.title;
 const input=make(f.options?'select':'input');if(f.options)for(const [v,t] of f.options)input.add(new Option(t,v));
 if(f.key==='D11'){input.type='number';input.min='1';input.step='1'}
 input.value=f.value;el('entry').replaceChildren(input);el('picks').replaceChildren();
 cars.forEach((c,i)=>{const label=make('label'),box=make('input');box.type='checkbox';box.checked=i===index;label.append(box,c.name);el('picks').append(label)});
 function eligible(){[...el('picks').querySelectorAll('input')].forEach((box,i)=>{const x=row[i];box.disabled=(f.key===controlKey&&!!cars[i].group)||!x||x.locked||!!x.options&&!x.options.some(o=>o[0]===input.value);if(box.disabled)box.checked=false})}
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
for(const [id,label] of Object.entries(sections)){const btn=make('button',label);btn.classList.toggle('active',id===section);btn.onclick=()=>{section=id;el('search').value='';el('featureGroup').hidden=id!=='V';document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('active',b===btn));el('part-'+id)?.scrollIntoView({block:'start'})};document.querySelector('nav').append(btn)}
for(const [key,name] of Object.entries({Emergency:'Khẩn cấp',Operation:'Vận hành',Door:'Cửa',Signal:'Tín hiệu',Traffic:'Giao thông'}))el('featureGroup').add(new Option(name,'TblFeature_'+key));
el('featureGroup').insertBefore(new Option('Tất cả nhóm tính năng','all'),el('featureGroup').firstChild);el('featureGroup').value='all';el('featureGroup').onchange=refresh;el('search').oninput=draw;el('diff').onchange=draw;
el('prev').onclick=()=>{page--;draw()};el('next').onclick=()=>{page++;draw()};
el('car').onchange=()=>{page=Math.floor(el('car').selectedIndex/4);show()};
el('detail').onclick=()=>{detail=!detail;show()};
el('addCar').onclick=async()=>{add();await syncProject([current()]);page=Math.floor((cars.length-1)/4);show()};
el('renameCar').onclick=()=>{const c=current(),name=prompt('Mã thang',c.name)?.trim();if(!name)return;clearErrorMarks();c.name=name;c.frame.title=name;el('car').selectedOptions[0].text=name;show()};
el('deleteCar').onclick=async()=>{if(cars.length===1||!confirm('Xóa '+current().name+'?'))return;const i=el('car').selectedIndex;if(cars[i].group&&!await changeGroup(cars[i],null))return;clearErrorMarks();cars[i].frame.remove();cars.splice(i,1);el('car').remove(i);page=Math.floor(el('car').selectedIndex/4);show()};
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
 page=Math.floor(index/4);el('car').selectedIndex=index;el('search').value='';el('diff').checked=false;
 if(el('featureGroup').value!=='all'){el('featureGroup').value='all';await refresh()}
 const target=locate(car,issue);if(!target){status('Lỗi này chưa có địa chỉ ô nhập trong Spec Summary.');return}
 section=target.id;el('featureGroup').hidden=section!=='V';draw();
 const tab=[...document.querySelectorAll('nav.sections button')].find(b=>b.textContent===sections[section]);
 document.querySelectorAll('nav.sections button').forEach(b=>b.classList.toggle('active',b===tab));
 const scope=el('part-'+section);
 const selector='[data-car-index="'+index+'"][data-field-key="'+CSS.escape(target.field.key)+'"]';
 let cell=scope.querySelector('input'+selector+',select'+selector)||scope.querySelector(selector);
 if(!cell&&car.group)cell=scope.querySelector('[data-car-index="'+cars.indexOf(cars.slice(page*4,page*4+4).find(c=>c.group===car.group))+'"][data-field-key="'+CSS.escape(target.field.key)+'"]');
 if(!cell)return;
 cell.scrollIntoView({block:'center',inline:'nearest'});cell.focus({preventScroll:true});
 for(const node of [cell,tab].filter(Boolean)){node.classList.remove('error-pulse');void node.offsetWidth;node.classList.add('error-pulse')}
}
function clearErrorMarks(){
 el('carErrors').hidden=true;
 document.querySelectorAll('.has-spec-error,.error-pulse').forEach(node=>node.classList.remove('has-spec-error','error-pulse'));
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
   panel?.querySelectorAll(selector).forEach(node=>node.classList.add('has-spec-error'));
  }
 }
 document.querySelectorAll('nav.sections button').forEach(tab=>{
  tab.classList.toggle('has-spec-error',[...failedSections].some(id=>sections[id]===tab.textContent));
 });
}
function renderErrors(){
 const panel=el('carErrors'),grid=make('div');grid.className='error-grid';
 const visible=cars.slice(page*4,page*4+4);grid.style.gridTemplateColumns='28% repeat('+visible.length+',minmax(0,1fr))';
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
 lastFailures=failures;page=Math.floor(cars.indexOf(failures[0].car)/4);detail=false;cars.forEach(c=>c.frame.hidden=true);el('overview').hidden=false;el('detail').textContent='Form đơn';
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
 if(canExport)body.append(make('p',cars.map(c=>c.name).join(' · ')+' · '+(el('only-vi').checked?'Vietnamese Only':'Song ngữ')+(el('export-tech-details').checked?' · Có trích xuất chi tiết':'')+' · A4 ngang'));
 el('confirmExport').hidden=!canExport;
 dialog.returnValue='';dialog.showModal();
 return new Promise(resolve=>dialog.addEventListener('close',()=>resolve(dialog.returnValue==='export'),{once:true}));
}
async function run(exporting){
 const main=document.querySelector('main'),book=[],failures=[],groupFields=new Map();main.inert=true;clearErrorMarks();
 try{
  for(const car of cars){
   status('Đang kiểm tra '+car.name+'…');let errors=[];
   try{
    await syncProject([car]);errors=await request(car,'validate');
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
  if(failures.length){await reportErrors(failures);main.inert=false;await checkPopup('Check Spec · Cần sửa lỗi',failures);return}
  status('Check Spec đạt · '+cars.map(c=>c.name).join(', '));
  if(!await checkPopup('Check Spec đạt · '+cars.length+' cột thang',[],true))return;
  exporting=true;
  for(const car of cars){
   try{book.push({bankLabel:!car.group&&internalBanks([car],groupFields.get(car)?.find(f=>f.key===controlKey)?.value)?internalLabel(car,parseInt(groupFields.get(car).find(f=>f.key===controlKey).value,10),columnCount(car)):'',onlyVi:el('only-vi').checked,model,operation:groupFields.get(car)?.find(f=>f.key===controlKey)?.value,name:car.name,groupId:car.group?cars.indexOf(car.group[0]):null,rows:await request(car,'pairRows',{includeTechDetails:el('export-tech-details').checked})})}
   catch(e){failures.push({car,errors:[e.message]})}
  }
  if(failures.length){await reportErrors(failures);main.inert=false;await checkPopup('Check Spec · Cần sửa lỗi',failures);return}
  if(exporting){const blob=await request(cars[0],'pairBook',book),url=URL.createObjectURL(blob),a=make('a');a.download=(isMRL?'Spec_Summary_NEXIEZ_MRL_Version2.xlsx':'Multi_Spec_NEXIEZ_MR.xlsx');a.href=url;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000)}
  status(exporting?'Đã xuất Excel A4 ngang.':'Check Spec PASS · '+cars.map(c=>c.name).join(', '));
 }catch(e){status(e.message);main.inert=false;await checkPopup('Không thể xuất Excel',[{car:{name:''},errors:[e.message]}])}finally{main.inert=false}
}
el('check').onclick=()=>run(false);el('exportBook').onclick=()=>run(true);
document.querySelectorAll('main button:not(.active)').forEach(b=>b.classList.add('btn','btn-secondary'));
for(let i=0;i<4;i++)add();el('car').selectedIndex=0;show();
