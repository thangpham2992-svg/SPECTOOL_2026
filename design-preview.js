/* R105 — read-only cabin preview, original G11L1 artwork. */
(function(){
'use strict';
const library=()=>window.R105_CAR_LIBRARY;
let loading;
function loadLibrary(){
 if(library())return Promise.resolve();
 if(!loading)loading=new Promise((resolve,reject)=>{
  const s=document.createElement('script');s.src='design-car-assets.js?v=R105';s.onload=()=>library()?resolve():reject(Error('Invalid cabin library'));s.onerror=()=>{s.remove();reject(Error('Không tải được thư viện ảnh cabin.'))};document.head.append(s);
 }).catch(e=>{loading=null;throw e});
 return loading;
}
const norm=v=>String(v??'').trim();
const none=v=>/^(?:No|None|Not supplied|Without)$/i.test(v);
function resolveSpec(v,lib){
 const db=lib.db,images=lib.images,missing=[],layers=[];
 const val=k=>norm(v[k]);
 const fail=(name,value)=>missing.push({name,value:value||'Chưa nhập / Blank'});
 const door=val('D15'),key=door==='CO'?'png_co':/^2S(?:L)?$/.test(door)?'png_2s':null;
 if(!key)fail('Loại cửa / Door type',door);
 if(val('D14')&&!none(val('D14')))fail('Cabin quan sát / Observation',val('D14'));
 if(val('D16')!=='1D-1G')fail('Hướng mở cửa / Entrance',val('D16'));
 function add(file,name,value){if(file&&Object.hasOwn(images,file))layers.push(file);else fail(name,value);}
 function entry(group,text){return db[group]?.find(r=>r.text===text)}
 function layer(group,text,name,value=text){const row=entry(group,text);if(!row||!key||!row[key])fail(name,value);else add(row[key],name,value);}
 function finish(cell,paint,pattern){const f=val(cell);if(f==='SUS-HL')return 'Hairline stainless-steel';if(f==='SUS-M')return 'Mirror stainless-steel';if(f==='SUS-Shiny vibration')return 'Shiny vibration';if(f==='Painted')return val(paint);if(/^(?:Laminated steel sheet|Pattern-printed steel sheet PPS)$/.test(f))return val(pattern);return f;}
 function surface(group,cell,paint,pattern,name){layer(group,finish(cell,paint,pattern),name,[val(cell),val(cell)==='Painted'?val(paint):/Laminated|Pattern-printed/.test(val(cell))?val(pattern):''].filter(Boolean).join(' · '));}
 add('car_00_base','Nền cabin','Base');
 // The floor is deliberately independent of the selected specification.
 layer('CarFlooring','PR801','Sàn minh họa / Reference floor');
 surface('CarTransompanel','D49','D50','D51','Đố cửa trên / Transom');
 if(/DSJ1|\bDS1\b|Large\s*window|\bwindow\b|glass/i.test(val('D46')))fail('Cửa kính / Glass door',val('D46'));else surface('CarDoors','D46','D47','D48','Cửa cabin / Car door');
 surface('CarFrontreturnpanel','D52','D53','','Vách trước / Front return');
 surface('CarWalls','D43','D44','D45','Vách cabin / Car walls');
 const kick=val('D54');layer('CarKickplate',kick==='SUS-HL'?'Hairline stainless-steel':kick,'Viền chân vách / Kickplate',kick);
 if(!none(val('D66')))layer('CarMirrors',val('D66'),'Gương / Mirror');
 const cop=val('D61'),loc=val('D60'),button=val('D59'),copFull=[button,cop,loc].join(' · ');
 const cp=/^([CMPND]\d{3,4})E(TL)?$/.exec(cop),pos=loc==='Front'?'[For front return panel]':loc==='Side'?'[For side wall]':'';
 let copRow=cp&&pos&&/^CBV[1-6]$/.test(button)?entry('CarCop',`${pos} CBVx-${cp[1]}`):null;
 if(!copRow||!key||!copRow[key])fail('COP chính / Main COP',copFull);else add(copRow[key]+'a','COP chính / Main COP',copFull);
 const aux=val('D62');
 // The original library only provides an auxiliary image corresponding to the main COP.
 if(aux&&!none(aux)){
  const am=aux.match(/(?:^|[-\s])([CMPND]\d{3,4})E(TL)?(?:$|\s)/);
  if(!copRow||!key||!copRow[key]||!am||am[1]!==cp?.[1]||Boolean(am[2])!==Boolean(cp?.[2]))fail('COP phụ / Auxiliary COP',aux);
  else add(copRow[key]+'b','COP phụ / Auxiliary COP',aux);
 }
 if(cp?.[2]&&copRow&&key){const touch=copRow['png_t_less_'+(key==='png_co'?'co':'2s')];add(touch?touch+'_main':'','COP Touchless',cop);if(aux&&!none(aux))add(touch?touch+'_aux':'','COP phụ Touchless',aux);}
 const rail=val('D64'),position=val('D65');
 if(!none(rail)){
  const model=/^(YH-\d+[A-Z]+)(?:\s|$)/.exec(rail)?.[1]||rail;
  const places={ '1 pc on rear wall':['frontwall'],'1 pc on right side wall':['rightwall'],'1 pc on left side wall':['leftwall'],'2 pcs on both side walls':['rightwall','leftwall'],'2 pcs on rear & right side wall':['frontwall','rightwall'],'2 pcs on rear & left side wall':['frontwall','leftwall'],'3 pcs on 3 walls':['frontwall','rightwall','leftwall']};
  if(!places[position])fail('Vị trí tay vịn / Handrail location',position);else for(const p of places[position])layer('CarHandrails'+p,model,'Tay vịn / Handrail',rail+' · '+position);
 }
 const ceiling=val('D41'),color=val('D42'),cm=/^(CL[12]|DL[567])-([CW])$/.exec(ceiling);
 if(cm)layer('CarCeiling',`${cm[1]}:${color} [${cm[2]==='C'?'cool':'warm'} white]`,'Trần cabin / Ceiling',ceiling+' · '+color);else fail('Trần cabin / Ceiling',ceiling+' · '+color);
 add('mask_car','Mask','mask_car');
 return {layers,missing};
}
const cache=new Map();
function decode(name){if(!cache.has(name)){const p=new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>reject(Error('Không đọc được ảnh: '+name));img.src='data:image/webp;base64,'+library().images[name]});cache.set(name,p);p.catch(()=>cache.delete(name));if(cache.size>48)cache.delete(cache.keys().next().value);}return cache.get(name);}
function node(tag,text,css){const el=document.createElement(tag);if(text)el.textContent=text;if(css)el.style.cssText=css;return el;}
async function open(values){
 const old=document.getElementById('r105-cabin-dialog');old?.close();old?.remove();
 const dialog=node('dialog','','border:1px solid #b6c6d5;border-radius:12px;padding:0;width:min(900px,92vw);max-height:90vh;color:#17365d;background:#fff;font:14px Arial,sans-serif;box-shadow:0 12px 50px #162e4d44;');dialog.id='r105-cabin-dialog';
 const header=node('div','','display:flex;align-items:center;justify-content:space-between;padding:14px 20px;background:#e8f1fa;border-bottom:1px solid #c8d9e8;gap:12px;');header.append(node('strong','Render cabin / Phối cảnh cabin','font-size:18px;'));
 const close=node('button','Đóng / Close','padding:7px 12px;border:1px solid #a9bdce;border-radius:5px;background:white;color:#17365d;cursor:pointer;');close.type='button';close.onclick=()=>dialog.close();header.append(close);
 const body=node('div','','padding:16px 20px;');body.append(node('p','Đang tải thư viện đồ họa cabin…'));
 dialog.append(header,body);document.body.append(dialog);dialog.addEventListener('close',()=>dialog.remove(),{once:true});dialog.showModal();
 try{
  await loadLibrary();if(!dialog.isConnected)return;
  const {layers,missing}=resolveSpec(values,library());body.replaceChildren();
  body.append(node('div','NEXIEZ-MRL Version2 · '+norm(values.C8),'font-weight:bold;margin-bottom:10px;'));
  if(missing.length){body.append(node('p','Thư viện chưa đủ dữ liệu để tái tạo đồ họa. / Insufficient image library data to generate the design preview.','background:#fff4df;color:#80531b;border:1px solid #e3c48c;border-radius:6px;padding:12px;font-weight:bold;'));const ul=node('ul','','padding-left:24px;line-height:1.7;');for(const m of missing)ul.append(node('li',m.name+': '+m.value));body.append(ul);}
  else{const imgs=await Promise.all(layers.map(decode));if(!dialog.isConnected)return;const canvas=node('canvas','','display:block;width:100%;max-width:705px;height:auto;margin:0 auto;background:#fff;');canvas.width=705;canvas.height=570;const ctx=canvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,705,570);for(const img of imgs)ctx.drawImage(img,0,0,705,570);body.append(canvas);}
  body.append(node('p','Sàn minh họa PR801 (không phụ thuộc loại sàn Spec). Hình phối cảnh tham khảo, không theo kích thước AA/BB/HL thực tế; COP dùng hình CBVx theo mã cơ sở. Không thay đổi dữ liệu Spec, Check Spec hoặc Export.','font-size:12px;color:#586b7b;margin:12px 0 0;line-height:1.5;'));
 }catch(e){if(dialog.isConnected)body.replaceChildren(node('p','Không thể render cabin: '+e.message));}
}
/* R107: original hall layers, combined design appendix. No synthetic artwork. */
let hallLoading;
function loadHall(){
 if(window.R107_HALL_LIBRARY)return Promise.resolve(window.R107_HALL_LIBRARY);
 if(!hallLoading)hallLoading=new Promise((resolve,reject)=>{const s=document.createElement('script');s.src='design-hall-assets.js?v=R107';s.onload=()=>window.R107_HALL_LIBRARY?resolve(window.R107_HALL_LIBRARY):reject(Error('Invalid hall library'));s.onerror=()=>{s.remove();reject(Error('Không tải được thư viện cửa tầng.'))};document.head.append(s);}).catch(e=>{hallLoading=null;throw e});
 return hallLoading;
}
const absent=v=>!norm(v)||none(norm(v));
function hallSpec(v,lib,level='main'){
 const val=k=>norm(v[k]), main=level==='main', j=main?'71':'72',d=main?'73':'74',f=level==='top'?75:main?78:81;
 const missing=[],layers=[],key=val('D15')==='CO'?'png_co':/^2S(?:L)?$/.test(val('D15'))?'png_2s':null;
 const fail=(name,value)=>missing.push({name,value:value||'Chưa nhập / Blank'});
 function add(name,label,value){if(name&&Object.hasOwn(lib.images,name))layers.push(name);else fail(label,value);}
 function row(group,text){return lib.db[group]?.find(r=>r.text===text);}
 function layer(group,text,label,value=text){const r=row(group,text);if(r&&key)add(r[key],label,value);else fail(label,value);}
 const finish=(a,b)=>val(a)==='SUS-HL'?'Hairline stainless-steel':val(a)==='SUS-M'?'Mirror stainless-steel':val(a)==='Painted'?val(b):val(a);
 if(!key)fail('Loại cửa / Door type',val('D15'));
 add('hall_00_base','Nền / Base','Base');layer('HallWall','White 1','Building wall');layer('HallFloor','Gray','Building floor');
 const jamb=val('D'+j),jf=finish('E'+j,'F'+j),df=finish('E'+d,'F'+d);
 layer('HallJumb',jamb.replace('-','')+':'+jf,'Khung bao / Jamb',jamb+' · '+jf);
 // E312 requires a separate transom selection in the original tool, absent from Spec.
 if(jamb==='E-312')fail('Đố cửa tầng / Hall transom','E-312: chưa có lựa chọn trong Spec');
 const doors=row('HallDoors',df);
 if(doors&&key&&doors[key]){add(doors[key]+'_r','Cửa tầng / Landing door',df);add(doors[key]+'_l','Cửa tầng / Landing door',df);}else fail('Cửa tầng / Landing door',df);
 const piv=val('E'+f),hbv=val('E'+(f+1)),indicator=val('E'+(f+2)),lantern=val(level==='top'?'HL_TOP':main?'HL_MAIN':'HL_OTHER');
 let buttons=0,fixtures=0;
 function panel(value){if(absent(value))return;const m=/^(PIV|HBV)[1-4]-([AC]\d{3,4})N(TL)?$/.exec(value);const group='HallButtons'+(buttons++?'B':'A');const r=m&&row(group,m[1]+'*-'+m[2]);if(!r||!key){fail('Hộp gọi tầng / Hall button',value);return;}add(r[key],'Hộp gọi tầng / Hall button',value);if(m[3])add(r['png_t_less_'+(key==='png_co'?'co':'2s')],'Touchless',value);}
 function fixture(value){if(absent(value))return;const group='HallOtherfixtures'+(fixtures++?'B':'A');layer(group,value,'Đèn báo tầng / Hall fixture',value);}
 panel(piv);panel(hbv);fixture(indicator);fixture(lantern);
 // Positions A/B overlap if multiple devices are selected. Do not invent placement.
 if(buttons+fixtures>1)fail('Vị trí thiết bị / Fixture layout','Nhiều thiết bị: chưa có dữ liệu vị trí lắp đặt');
 return {layers,missing,details:[['Jamb',jamb+' · '+jf],['Doors',df+' · '+val('D15')],['PIV',piv],['HBV',hbv],['PIH / PID',indicator],['Hall Lantern',lantern]]};
}
async function composite(layers,lib,w,h){
 const canvas=document.createElement('canvas');canvas.width=w;canvas.height=h;const ctx=canvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);
 const images=await Promise.all(layers.map(name=>new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>reject(Error('Không đọc được ảnh: '+name));img.src='data:image/webp;base64,'+lib.images[name];})));
 for(const img of images)ctx.drawImage(img,0,0,w,h);return canvas.toDataURL('image/png');
}
async function collectDesign(v){
 const items=[];
 async function item(title,details,load,resolve,w,h){const o={title,details,missing:[],src:''};items.push(o);try{const lib=await load();const plan=resolve(lib);o.missing=plan.missing;if(plan.details)o.details=plan.details;if(!o.missing.length)o.src=await composite(plan.layers,lib,w,h);}catch(e){o.missing=[{name:'Thư viện / Library',value:e.message}];}return o;}
 const val=k=>norm(v[k]);
 const cabinDetails=[['Ceiling',val('D41')+' · '+val('D42')],['Doors',val('D46')+' · '+val('D15')],['Walls',val('D43')+' '+val('D44')+' '+val('D45')],['Transom',val('D49')],['Front return',val('D52')],['Kick plate',val('D54')],['Flooring (reference)', 'PR801'],['COP',val('D59')+'-'+val('D61')],['Handrails',val('D64')+' · '+val('D65')],['Mirror',val('D66')]];
 await item('Cabin / Phòng thang',cabinDetails,async()=>{await loadLibrary();return library();},lib=>resolveSpec(v,lib),705,570);
 // A single representative entrance per floor category; merge only identical actual selections.
 const seen=new Map();
 for(const [level,title] of [['main','Main floor / Tầng chính'],['other','Other floors / Tầng khác'],['top','Top floor / Tầng trên cùng']]){
  if(level==='other'&&Number(v.D18)===2)continue;
  const main=level==='main',j=main?'71':'72',d=main?'73':'74',f=level==='top'?75:main?78:81;
  const signature=JSON.stringify(['D'+j,'E'+j,'F'+j,'E'+d,'F'+d,'E'+f,'E'+(f+1),'E'+(f+2),level==='top'?'HL_TOP':main?'HL_MAIN':'HL_OTHER'].map(val));
  if(seen.has(signature)){seen.get(signature).title+=' · '+title;continue;}
  const o=await item(title,[],loadHall,lib=>hallSpec(v,lib,level),405,570);seen.set(signature,o);
 }
 return items;
}
const esc107=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const designCSS=`.r107-design{font:8pt Arial,sans-serif;color:#20344c;line-height:1.22}.r107-design h2{font-size:12pt;margin:2mm 0 1.5mm;color:#17365d;border-bottom:1px solid #a8bacb;padding-bottom:1mm}.r107-design h3{font-size:8.5pt;margin:0 0 1mm;color:#17365d}.r107-design .design-meta{font-size:9pt;margin-bottom:2mm;overflow-wrap:anywhere}.r107-cabin{display:grid;grid-template-columns:61% 39%;gap:2mm;align-items:start}.r107-design img{display:block;max-width:100%;object-fit:contain;margin:auto}.r107-cabin img{width:100%;height:auto;aspect-ratio:705/570;max-height:93mm}.r107-design dl{margin:0;font-size:7.2pt;line-height:1.15;overflow-wrap:anywhere}.r107-design dt{font-weight:bold;margin-top:.6mm}.r107-design dd{margin:0}.r107-halls{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:2mm}.r107-halls.one{grid-template-columns:1fr}.r107-halls.two{grid-template-columns:repeat(2,minmax(0,1fr))}.r107-halls.two .r107-hallbody{display:grid;grid-template-columns:48% 52%;gap:1.5mm;align-items:start}.r107-halls.two .r107-missing{grid-column:1/-1}.r107-halls.one .r107-hallbody{display:grid;grid-template-columns:40% 60%;align-items:start;gap:3mm}.r107-halls img{height:58mm;width:auto;aspect-ratio:405/570;max-width:100%}.r107-halls:not(.one):not(.two) dl{font-size:6.6pt}.r107-missing{border:1px solid #d8b980;background:#fff8eb;padding:2mm;color:#77541c;font-size:7.5pt;overflow-wrap:anywhere}.r107-design .design-note{font-size:7pt;color:#546476;margin:2mm 0 0}.r107-hallbody{min-width:0}.r107-design .r107-hall{min-width:0}
/* R109: A4 appendix uses the actual printable space below the masthead. */
.r107-design-page .r107-design{display:flex;flex-direction:column;box-sizing:border-box;gap:1.4mm;min-height:0;}
.r107-design-page .design-meta{flex:none;margin-bottom:0;font-size:9pt;}
.r107-design-page .design-section{display:flex;flex-direction:column;min-height:0;flex:1;}
.r107-design-page .design-section.cabin-section{flex:1.18;}
.r107-design-page .design-section h2{flex:none;margin:0 0 1.5mm;}
.r107-design-page .r107-cabin{flex:1;min-height:0;grid-template-columns:minmax(0,70fr) minmax(0,30fr);gap:3mm;align-items:stretch;}
.r107-design-page .r107-design img{height:100%;width:100%;max-width:100%;max-height:100%;min-height:0;aspect-ratio:auto;object-fit:contain;object-position:left center;margin:0;}
.r107-design-page .r107-design dl{font-size:8.5pt;line-height:1.22;align-self:start;min-width:0;}
.r107-design-page .r107-design dt{margin-top:.8mm;}
.r107-design-page .r107-halls{flex:1;min-height:0;gap:3mm;}
.r107-design-page .r107-hall{display:flex;flex-direction:column;min-height:0;}
.r107-design-page .r107-hall h3{flex:none;margin-bottom:1mm;font-size:9pt;}
.r107-design-page .r107-hallbody{flex:1;min-height:0;display:grid;grid-template-rows:minmax(0,1fr);align-items:stretch;}
.r107-design-page .r107-halls.one .r107-hallbody{grid-template-columns:minmax(0,48fr) minmax(0,52fr);gap:4mm;}
.r107-design-page .r107-halls.two .r107-hallbody{grid-template-columns:minmax(0,59fr) minmax(0,41fr);gap:2mm;}
.r107-design-page .r107-halls.two dl,.r107-design-page .r107-halls.three dl{font-size:7.6pt;}
.r107-design-page .r107-halls.three .r107-hallbody{grid-template-columns:1fr;grid-template-rows:minmax(0,1fr) auto;gap:1mm;}
.r107-design-page .r107-design .design-note{flex:none;margin:0;font-size:8pt;}
.r107-design-page .only-cabin .r107-cabin{grid-template-columns:1fr;grid-template-rows:minmax(0,1fr) auto;gap:2mm;}
.r107-design-page .only-cabin .r107-cabin img{object-position:center;}
.r107-design-page .only-cabin .r107-cabin dl{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:1mm 3mm;}
.r107-design-page .only-hall .r107-halls.one .r107-hallbody{grid-template-columns:minmax(0,62fr) minmax(0,38fr);}
.r107-design-page .only-hall .r107-halls.two .r107-hallbody{grid-template-columns:1fr;grid-template-rows:minmax(0,1fr) auto;}
.r107-design-page .only-hall .r107-halls.two img,.r107-design-page .r107-halls.three img{object-position:center;}
.r107-design-page .r107-design.compact-details dl{font-size:7.2pt;line-height:1.1;}
`;
function designMarkup(items,v,onlyVi=false){
 const e=esc107;
 const shown=x=>{const s=norm(x);return s&&!/^(?:No|None|Not supplied|Without|Không cung cấp(?:\s*\(Not supplied\))?)$/i.test(s)};
 const details=o=>'<dl>'+o.details.filter(([k,x])=>shown(x)).map(([k,x])=>'<div><dt>'+e(k)+'</dt><dd>'+e(x)+'</dd></div>').join('')+'</dl>';
 const content=o=>o.src?'<img width="'+(o===items[0]?705:405)+'" height="570" alt="'+e(o.title)+'" src="'+o.src+'">'+details(o):'';
 const cabin=items[0];
 const halls=items.slice(1).filter(o=>o.src);
 const hallBlock=halls.length?'<div class="design-section hall-section"><h2>IV. Phối cảnh cửa tầng / Landing design</h2><div class="r107-halls'+(halls.length===1?' one':halls.length===2?' two':' three')+'">'+halls.map(o=>'<div class="r107-hall"><h3>'+e(o.title)+'</h3><div class="r107-hallbody">'+content(o)+'</div></div>').join('')+'</div></div>':'';
 return '<div class="r107-design'+(cabin.src?(halls.length?'':' only-cabin'):' only-hall')+'"><div class="design-meta"><b>Công trình / Project:</b> '+e(v.C7)+' &nbsp; <b>Hạng mục / Item:</b> '+e(v.C8)+'<br><b>Model:</b> NEXIEZ-MRL Version2</div>'+
  (cabin.src?'<div class="design-section cabin-section"><h2>III. Phối cảnh cabin / Cabin design</h2><div class="r107-cabin">'+content(cabin)+'</div></div>':'')+
  hallBlock+
  '<p class="design-note">Hình ảnh mang tính chất tham khảo.</p></div>';
}
async function openCombined(v){
 const old=document.getElementById('r107-design-dialog');old?.close();old?.remove();
 const d=node('dialog','','width:min(850px,95vw);max-height:94vh;border:1px solid #aebdce;border-radius:8px;padding:12px;background:white;color:#17365d;font:14px Arial;');d.id='r107-design-dialog';const close=node('button','Đóng / Close','float:right;padding:6px 12px;cursor:pointer');close.onclick=()=>d.close();const body=node('div','Đang tải thư viện đồ họa…');d.append(close,body);document.body.append(d);d.addEventListener('close',()=>d.remove(),{once:true});d.showModal();try{const items=await collectDesign(v);if(d.isConnected)body.innerHTML='<style>'+designCSS+'</style><h2>Render Design · A4 Portrait</h2>'+designMarkup(items,v);}catch(e){body.textContent=e.message;}
}
window.R105_CABIN_PREVIEW={open:openCombined,resolveSpec,hallSpec,collectDesign,designMarkup,designCSS};

})();
