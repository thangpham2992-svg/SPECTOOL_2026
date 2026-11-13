/* NEXIEZ-MRL Alpha Package Version02 [G11L1P], 2026-05-08: manual Table 6.2-6.5. */
(function(){'use strict';
const catalog={"BMS-GW":{"group":"Emergency","statuses":"AAA"},"CSB":{"group":"Emergency","statuses":"AAA"},"EER-P":{"group":"Emergency","statuses":"AAA"},"EER-S":{"group":"Emergency","statuses":"AAA"},"FE":{"group":"Emergency","statuses":"AAA"},"FER":{"group":"Emergency","statuses":"AAA"},"MELD":{"group":"Emergency","statuses":"AAA"},"OEPS-A":{"group":"Emergency","statuses":"AAA"},"OEPS-M":{"group":"Emergency","statuses":"XXX"},"OEPS-SA":{"group":"Emergency","statuses":"AAA"},"THMF":{"group":"Emergency","statuses":"SSS"},"WP-S":{"group":"Emergency","statuses":"AAA"},"ABP":{"group":"Operation","statuses":"ASS"},"AS":{"group":"Operation","statuses":"AAA"},"BP":{"group":"Operation","statuses":"AAA"},"CCC":{"group":"Operation","statuses":"SSS"},"CFO-A":{"group":"Operation","statuses":"SSS"},"CFO-B":{"group":"Operation","statuses":"SSS"},"CLO-A":{"group":"Operation","statuses":"SSS"},"CLO-B":{"group":"Operation","statuses":"SSS"},"COS":{"group":"Operation","statuses":"XSS"},"EMS":{"group":"Operation","statuses":"AAA"},"EVRC-C":{"group":"Operation","statuses":"AAA"},"EVRC-S":{"group":"Operation","statuses":"AAA"},"EVRC-H":{"group":"Operation","statuses":"AAX"},"EXIT SWITCH":{"group":"Operation","statuses":"AAA"},"FCC-A":{"group":"Operation","statuses":"AAS"},"FCC-P":{"group":"Operation","statuses":"AAA"},"FHC-P":{"group":"Operation","statuses":"AAA"},"FSAT":{"group":"Operation","statuses":"SSS"},"GCBK":{"group":"Operation","statuses":"XSS"},"HE-B":{"group":"Operation","statuses":"AAA"},"HOS":{"group":"Operation","statuses":"AAA"},"HOS-T":{"group":"Operation","statuses":"XXX"},"IND":{"group":"Operation","statuses":"SSS"},"NS-CB":{"group":"Operation","statuses":"AAA"},"NS":{"group":"Operation","statuses":"XXX"},"NSCR-C":{"group":"Operation","statuses":"AAA"},"NS-T":{"group":"Operation","statuses":"XXX"},"NST":{"group":"Operation","statuses":"SSS"},"NXL":{"group":"Operation","statuses":"SSS"},"OLH":{"group":"Operation","statuses":"SSS"},"RCS":{"group":"Operation","statuses":"XXX"},"RET1":{"group":"Operation","statuses":"AAA"},"SCS-B":{"group":"Operation","statuses":"AAA"},"SHBT":{"group":"Operation","statuses":"XXX"},"1D-2G":{"group":"Door","statuses":"AAA"},"CLTS":{"group":"Door","statuses":"SSS"},"D69":{"group":"Door","statuses":"SSS"},"DDOP":{"group":"Door","statuses":"SSS"},"DFOP":{"group":"Door","statuses":"SSS"},"DODA":{"group":"Door","statuses":"SSS"},"DSAC":{"group":"Door","statuses":"SSS"},"DKO-TB":{"group":"Door","statuses":"AAA"},"DLD":{"group":"Door","statuses":"SSS"},"EDC":{"group":"Door","statuses":"SSS"},"EDM":{"group":"Door","statuses":"AAA"},"HMS":{"group":"Door","statuses":"XXX"},"IDKO":{"group":"Door","statuses":"SSS"},"MBS":{"group":"Door","statuses":"SSS"},"NDG":{"group":"Door","statuses":"SSS"},"RDC":{"group":"Door","statuses":"SSS"},"ROHB":{"group":"Door","statuses":"SSS"},"SDE":{"group":"Door","statuses":"AAA"},"SNCD":{"group":"Door","statuses":"AAA"},"SR-1 Beam":{"group":"Door","statuses":"SSS"},"SR-2 Beams":{"group":"Door","statuses":"AAA"},"AAN-B":{"group":"Signal","statuses":"AAS"},"AAN-G":{"group":"Signal","statuses":"AAA"},"ACB":{"group":"Signal","statuses":"AAA"},"ACS":{"group":"Signal","statuses":"XXX"},"AECC":{"group":"Signal","statuses":"AAA"},"AECH":{"group":"Signal","statuses":"AAS"},"AHC":{"group":"Signal","statuses":"AAA"},"AIL":{"group":"Signal","statuses":"XXA"},"AUTL":{"group":"Signal","statuses":"AAX"},"BPL":{"group":"Signal","statuses":"AAX"},"CBE":{"group":"Signal","statuses":"SSS"},"DAC":{"group":"Signal","statuses":"SSS"},"DAH":{"group":"Signal","statuses":"SSS"},"CID":{"group":"Signal","statuses":"XXX"},"CID-S":{"group":"Signal","statuses":"AAA"},"EMB":{"group":"Signal","statuses":"SSS"},"EXCL":{"group":"Signal","statuses":"AAX"},"EXCLC":{"group":"Signal","statuses":"XXX"},"FELC":{"group":"Signal","statuses":"AAA"},"FELH":{"group":"Signal","statuses":"XXX"},"FERC":{"group":"Signal","statuses":"AAA"},"FHL":{"group":"Signal","statuses":"AAA"},"HBE":{"group":"Signal","statuses":"AAA"},"HID":{"group":"Signal","statuses":"XXX"},"HID-S":{"group":"Signal","statuses":"AAA"},"ITP":{"group":"Signal","statuses":"AAA"},"ITV":{"group":"Signal","statuses":"AAA"},"OLHL":{"group":"Signal","statuses":"AAA"},"OSL":{"group":"Signal","statuses":"XXX"},"RESL":{"group":"Signal","statuses":"XXX"},"PIE":{"group":"Signal","statuses":"SSX"},"PIH-D":{"group":"Signal","statuses":"AAX"},"PIM-C":{"group":"Signal","statuses":"AAX"},"PIW-D":{"group":"Signal","statuses":"AAA"},"TCP":{"group":"Signal","statuses":"XXA"},"BSO":{"group":"Traffic","statuses":"XXX"},"CAT":{"group":"Traffic","statuses":"XXS"},"CFS":{"group":"Traffic","statuses":"XXA"},"CNPS":{"group":"Traffic","statuses":"XXA"},"DPS":{"group":"Traffic","statuses":"XXA"},"DRO":{"group":"Traffic","statuses":"XXS"},"EL-SC":{"group":"Traffic","statuses":"XXX"},"ESO-N":{"group":"Traffic","statuses":"XXS"},"ESO-W":{"group":"Traffic","statuses":"XXS"},"FFS":{"group":"Traffic","statuses":"AAA"},"HCL":{"group":"Traffic","statuses":"AAA"},"LTS":{"group":"Traffic","statuses":"XXA"},"MFP":{"group":"Traffic","statuses":"AAA"},"NN":{"group":"Traffic","statuses":"XXS"},"PTC":{"group":"Traffic","statuses":"XXS"},"SCPS":{"group":"Traffic","statuses":"XXA"},"SFPS":{"group":"Traffic","statuses":"XXA"},"SOHS":{"group":"Traffic","statuses":"XSS"},"STH":{"group":"Traffic","statuses":"XXA"},"TFS":{"group":"Traffic","statuses":"XXX"},"TLF":{"group":"Traffic","statuses":"XXA"},"UCPS":{"group":"Traffic","statuses":"XXA"},"UPS":{"group":"Traffic","statuses":"XXS"},"VIP-S":{"group":"Traffic","statuses":"XXX"}};
const aliases={"ITV Cable":"ITV","Multi - Beam Door Sensor":"MBS","SDE - One Side":"SDE","SDE - Both Sides":"SDE","EXIT SW":"EXIT SWITCH"};
const en={"BMS-GW":"Building Management System Gateway","THMF":"Detection of high temperature in inverter heat-sink","WP-S":"Supervisory Panel – Serial Communication","BP":"Bypass","CFO-B":"Car Fan Shut Off – Block Sign","CLO-B":"Car Light Shut Off – Block Sign","HOS-T":"Hall Out of Service – Timer Type","NS-T":"Non-service to Specific Floors – Timer Type","NST":"Not Start Operation","SHBT":"Special Schedule Operation","1D-2G":"One-door Two-gate","CLTS":"Door Close Limit Switch On Start","D69":"Door Close Torque UP Control","DDOP":"Double Door Operation","DFOP":"Door Forcefully Open Operation","EDC":"Expediting of Door-close","IDKO":"Installation Extended Door Open","SNCD":"Quick Closing System with Sensor","ACS":"Auxiliary Control Station","CBE":"Click Type Car Call Button","DAC":"Direction Arrow in Car","DAH":"Direction Arrow on Halls","CID":"Car Information Display","CID-S":"Car LCD Position Indicator","FELH":"FE Operation Signal Lamp on Halls","HBE":"Click Type Hall Call Button","HID":"Hall Information Display","HID-S":"Hall LCD Position Indicator","PIE":"Hall Position Indicator (With Call Button)","PIH-D":"Hall Position Indicator Digital","PIM-C":"Hall LCD Position Indicator with Hall Button","PIW-D":"Supervisory Panel Position Indicator","OSL":"Hall Non-service Lamp","CAT":"Car Allocation Tuning","DRO":"Dynamic Rule Set Optimizer","EL-SC":"Elevator Operation Integrated with Security Control System","ESO-W":"Energy-saving Operation – Allocation Control","NN":"Distinction of Traffic Flow with Neural Network","STH":"See-through Operation","TLF":"Learning Function"};
const vi={"BMS-GW":"Cổng kết nối BMS","THMF":"Phát hiện nhiệt độ cao tại tản nhiệt biến tần","WP-S":"Bảng giám sát – giao tiếp nối tiếp","BP":"Bỏ qua lệnh gọi","CFO-B":"Tắt quạt cabin bằng thao tác nút","CLO-B":"Tắt đèn cabin bằng thao tác nút","NST":"Chế độ không khởi động","CBE":"Nút gọi cabin kiểu Click","DAC":"Mũi tên báo chiều trong cabin","DAH":"Mũi tên báo chiều tại tầng","PIH-D":"Bộ hiển thị tầng kỹ thuật số","PIW-D":"Hiển thị tầng trên bảng giám sát","CAT":"Tối ưu phân bổ cabin","NN":"Nhận diện luồng giao thông bằng mạng nơ-ron","TLF":"Chức năng tự học"};
const complimentary=new Set(['FCC-A','FCC-P','FHC-P','SOHS','GCBK','COS','PTC','UPS','DPS']);
const canonical=code=>aliases[code]||code;
const kind=control=>{const s=String(control||'').trim();return s==='1C-2BC'?0:s==='2C-2BC'?1:/^(?:[3-9]C|SAI|AI-|AI-2200C|AI-22)/i.test(s)?2:-1;};
function status(model,code,operation,db){
 const key=canonical(code),group=kind(operation);
 if(group<0)return 'X';
 if(model==='MRL'&&catalog[key])return catalog[key].statuses[group];
 if(model==='MR'&&key==='FE-C30')return 'A';
 const t=db?.tables?.Tbl39_Feature,dt=db?.sheets?.DataTable;if(!t||!dt)return null;
 const start=Number((t.ref.split(':')[0].match(/\d+/)||[])[0]||0)+1,end=Number((t.ref.split(':')[1].match(/\d+/)||[])[0]||0);
 for(let r=start;r<=end;r++)if(String(dt['A'+r]||'')===key){
  // MR: ALPHA flag is authoritative; not a single Version02/manual flag can turn it on.
  if(model==='MR'&&String(dt['Q'+r]||'').trim().toUpperCase()!=='Y')return 'X';
  const ctl=String(operation||'').trim().toUpperCase();
  const col=(model==='MR'&&(/^(AI-22|AI22)$/.test(ctl)))?'I':group===0?'G':group===1?'H':'J';
  const v=dt[col+r];return typeof v==='object'?String(v.v||'X').trim().toUpperCase():String(v||'X').trim().toUpperCase();
 }
 return null;
}
function fixDatabase(model,db){
 if(!db?.sheets?.DataTable||!db.tables?.Tbl39_Feature)return db;
 const dt=db.sheets.DataTable,t=db.tables.Tbl39_Feature;
 const start=Number((t.ref.split(':')[0].match(/\d+/)||[])[0]||0)+1,end=Number((t.ref.split(':')[1].match(/\d+/)||[])[0]||0);
 for(let r=start;r<=end;r++){
  const code=String(dt['A'+r]||'');const k=canonical(code);
  // The MR source contains EXCLC (Exclusive Operation Signal Light - Car)
  // with A/A/A/X compatibility, but its ALPHA flag is blank in the
  // distributed embedded database. Restore only this verified MR option.
  // MRL Version02 explicitly disallows EXCLC (XXX) in its own manual.
  if(model==='MR'&&k==='EXCLC'&&
     ['G','H','I'].some(c=>String(dt[c+r]||'').toUpperCase()==='A')){
   dt['Q'+r]='Y';
  }
  if(model==='MRL'&&catalog[k]){
   const st=catalog[k].statuses;
   dt['G'+r]=st[0];dt['H'+r]=st[1];dt['I'+r]=st[2];dt['J'+r]=st[2];
   // Version01 ALPHA blank flags must not make Version02 options disappear.
   dt['Q'+r]='Y';
  }
 }
 return db;
}
function virtual(model,db){
 if(model!=='MRL')return [];
 const present=new Set();const sh=db?.sheets?.SpecForm||{};
 for(const [name,t] of Object.entries(db?.tables||{}))if(name.startsWith('TblFeature_')){
  const a=Number((t.ref.split(':')[0].match(/\d+/)||[])[0]||0)+1,b=Number((t.ref.split(':')[1].match(/\d+/)||[])[0]||0);
  for(let r=a;r<=b;r++){const k=sh['B'+r];if(k)present.add(canonical(String(k)));}
 }
 return Object.entries(catalog).filter(([code,d])=>!present.has(code)&&/[AS]/.test(d.statuses)).map(([code,d])=>({
  table:'TblFeature_'+d.group,cell:'V52_'+code.replace(/[^A-Z0-9]/gi,'_'),code,virtual:true,
  en:en[code]||code,vi:vi[code]||en[code]||code,option:d.statuses[0],suggestion:'Manual V02',apply:d.statuses[0]==='S'?1:0
 }));
}
window.SPEC_V52_FEATURES={catalog,complimentary,canonical,status,fixDatabase,virtual};
})();