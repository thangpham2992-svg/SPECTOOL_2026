// Selection shared by MR duplex members, reviewed against Features_MR and DataTable.
const mrBankFeatures=new Set(['ITP','FER','COS','GCBK','SOHS','FSAT','MFP','EER-P','EER-S']);
function sharedMrCode(input){
 if($('model').value!=='MR')return '';
 if(input.id==='itp-type')return 'ITP type';
 const code=input.closest('.feature-line')?.querySelector('.fcode')?.textContent.trim();
 return mrBankFeatures.has(code)?code:'';
}
window.addEventListener('message',async event=>{
 const data=event.data;
 if(window.parent===window||event.source!==window.parent||data?.channel!=='multi-spec')return;
 const reply={channel:'multi-result',token:data.token};
 try{
  if(data.action==='snapshot'){
   reply.result={values:{...changes},touched:[...featureTouched]};
  }else if(data.action==='restore'){
   changes={...data.payload.values};featureTouched=new Set(data.payload.touched);
   renderAllSpecRows();renderFeatures();invalidateCheck();reply.result=true;
  }else if(data.action==='fields'){
   const result=[];
   const section=data.payload.section;
   function readControls(root,featureGroup){
    if(!root)return;
    const keys=new Set();
    for(const input of root.querySelectorAll('select,input[data-cell]')){
     if(input.matches('.mr-original-internal,[aria-hidden="true"]')||input.closest('[hidden]'))continue;
     const hall=input.closest('.hall-choice-item');
     if(input.classList.contains('hall-qty')&&!input.id){
      const source=hall?.querySelector('.field [data-cell]')?.dataset.cell;
      if(source)input.id='multi-hall-qty-'+source;
     }
     const key=input.dataset.cell||(input.id?'ui:'+input.id:'');
     if(!key||keys.has(key))continue;
     keys.add(key);
     const row=input.closest('.spec-row'),feature=input.closest('.feature-line');
     const heading=row?.querySelector('.spec-row-title')?.textContent?.trim()||'';
     const field=input.closest('.field')?.querySelector('.nm')?.textContent?.trim()||'';
     const device=input.closest('.car-design-line')?.querySelector('.car-design-label')?.textContent;
     const floor=hall?.closest('.hall-floor-card')?.querySelector('.hall-floor-title')?.textContent?.replace(/^\d+/, '').trim();
     const hallItem=hall?.querySelector('.hall-choice-label')?.textContent;
     const group=featureGroup||device||floor||input.closest('.dimension-group')?.querySelector('h3')?.textContent||heading||'Thông số';
     let title=input.getAttribute('aria-label')||heading||field||input.labels?.[0]?.textContent||key;
     if(input.id==='mr-source-choice')title='Mã P trùng cấu hình / Source code';
     if(row?.classList.contains('multi')&&field)title=heading+' · '+field;
     if(device)title=input.closest('.car-design-variants')?field:device;
     if(hallItem)title=input.classList.contains('hall-qty')?hallItem+' · Số hộp':hallItem;
     if(feature)title=feature.querySelector('.fcode')?.textContent+' · '+feature.querySelector('.fdesc')?.firstChild?.textContent;
     if(input.id==='itp-type')title='ITP · Loại liên lạc nội bộ / Interphone type';
     result.push({sharedCode:sharedMrCode(input),key,sourceCell:input.dataset.cell||input.closest('.field')?.querySelector('.mr-original-internal[data-cell]')?.dataset.cell||'',title,group,featureTable:featureGroup?activeFeatureTable:null,value:input.type==='checkbox'?(input.checked?'1':'0'):input.value,
      locked:input.disabled||input.readOnly,options:input.options?Array.from(input.options).filter(o=>!o.disabled).map(o=>[o.value,o.textContent]):input.type==='checkbox'?[['1','Có'],['0','Không']]:null});
    }
   }
   if(section==='V'){
    const previous=activeFeatureTable;
    const tables=data.payload.group==='all'?[...new Set(sourceFeatureRows().map(f=>f.table))]:[data.payload.group];
    for(const t of tables){activeFeatureTable=t;renderFeatures();readControls($('features'),$('feature-tabs').querySelector('.feature-tab-current')?.textContent||t)}
    activeFeatureTable=previous;renderFeatures();
   }else readControls(section==='P'?$('project-fields'):document.querySelector('.spec-section[data-section="'+section+'"]'));
   reply.result=result;
  }else if(data.action==='set'){
   const {key,value}=data.payload;
   const feature=data.payload.section==='V'?sourceFeatureRows().find(f=>f.cell===key):null;
   if(feature&&activeFeatureTable!==feature.table){activeFeatureTable=feature.table;renderFeatures()}
   if(data.payload.section==='V'&&!feature&&data.payload.featureGroup&&data.payload.featureGroup!=='all'){activeFeatureTable=data.payload.featureGroup;renderFeatures()}
   const input=key.startsWith('ui:')?document.getElementById(key.slice(3)):activeInput(key);
   if(!input||input.disabled||input.readOnly||input.closest('[hidden]'))throw Error('Trường này được khóa theo cấu hình.');
   if(input.options&&!Array.from(input.options).some(o=>o.value===value&&!o.disabled))throw Error('Giá trị không áp dụng cho cấu hình này.');
   if(input.type==='checkbox')input.checked=value==='1';else input.value=value;
   input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new Event('change',{bubbles:true}));reply.result=true;
  }else if(data.action==='validate'){window.__lastCheck=null;
   const messages=runCheckSpec(),issues=window.__lastCheck?.errors;
   reply.result=issues?.length?orderedCheckErrors(issues).map(e=>({message:e.message||'Vui lòng kiểm tra thông số.',cell:e.cell||''})):messages.map(message=>({message,cell:''}));
  }else   if(data.action==='pairRows'){
   if(runCheckSpec().length)throw Error('Vui lòng xử lý các mục trong Check Spec.');
   const x=await buildFinalExport(Boolean(data.payload?.includeTechDetails),false);reply.result=await window.SPEC_PRINT_LAYOUT.fromXlsx(x.data,x.model,true);
  }else if(data.action==='pairBook'){reply.result=await buildPairBook(data.payload);
  }else if(data.action==='excel'){
   if(runCheckSpec().length)throw Error('Vui lòng xử lý các mục trong Check Spec.');
   reply.result=(await buildFinalExport()).data;
  }else if(data.action==='zip'){
   const archive=new JSZip();
   for(const file of data.payload)archive.file(file.name,file.bytes);
   reply.result=await archive.generateAsync({type:'blob'});
  }else return;
 }catch(error){reply.error=error.message||String(error)}
 event.source.postMessage(reply,'*');
});

window.addEventListener('change',event=>{
 const input=event.target;
 if(window.parent!==window&&event.isTrusted&&input.dataset.cell===($('model').value==='MRL'?'D22':'D21'))
  window.parent.postMessage({channel:'summary-operation',value:input.value},'*');
},true);

window.addEventListener('change',event=>{
 const input=event.target;
 if(!event.isTrusted||window.parent===window||!sharedMrCode(input))return;
 const payload={key:input.dataset.cell,value:input.type==='checkbox'?(input.checked?'1':'0'):input.value,section:'V',featureGroup:activeFeatureTable};
 event.stopImmediatePropagation();
 renderFeatures();
 window.parent.postMessage({channel:'summary-feature',payload},'*');
},true);
