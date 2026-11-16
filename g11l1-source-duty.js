'use strict';
/* Port candidate of original StandardDuty1 SelectBalanceWeight, SelectBuffer,
   SelectCarPlatform (A) and composite stage 1. Fails closed if table input unavailable. */
(function(root){
 const T=root.G11L1_TS_ALPHA_TABLES, E=root.G11L1_MACHINE_SOURCE;
 const math={h:(x,y)=>Math.sqrt(x*x+y*y)-90,h0:(x,y)=>Math.sqrt(x*x+y*y)};
 const N=v=>Number(v), c=(v,test)=>v==='-'||test(v);
 const has=(t,str)=>String(t).includes(String(str));
 const req=(r,name)=>{if(!r)throw Error(`Source selector missing: ${name}`);return r;};
 function balance(d){
 const table=(T.BalanceWeight1||[]);
 const row=table.find(r=>c(r.CapacityRange,v=>has(v,d.CapacityRange))&&c(r.Entrance,v=>has(v,d.Entrance))&&c(r['Rated speed [m/s]'],v=>has(v,d.Speed))&&c(r['Code number'],v=>has(v,d.CodeNumber))&&c(r['Lower AA [mm]'],v=>N(v)<=d.AAxx)&&c(r['Upper AA [mm]'],v=>d.AAxx<=N(v))&&c(r['Lower BB [mm]'],v=>N(v)<=d.BBxx)&&c(r['Upper BB [mm]'],v=>d.BBxx<=N(v))&&c(r['Lower TR [mm]'],v=>N(v)<=d.TRxm)&&c(r['Upper TR [mm]'],v=>d.TRxm<=N(v))&&c(r['Cwt. Location'],v=>has(v,d.CwtLocation)));
 req(row,'BalanceWeight1');d.BalanceWeightNum=N(row.Total);return row;
 }
 function buffer(d){
 const s=d.Speed,sets=d.SetsL;
 const types=sets==='Yes'?{96:'LSB10',105:'LSB10',120:'LSB16',150:'LSB18'}:{60:'LSB10',96:'LSB16',105:'LSB18',120:'OBC20M',150:'OBC25M'};
 d.CarBuffer=d.CwtBuffer=req(types[s],`Buffer ${s} SETS=${sets}`);
 if(has(d.CapacityRange,'A'))d.CarBufferNum=d.CwtBufferNum=1;
 else if(has(d.CapacityRange,'B')){d.CarBufferNum=d.BGxx<1340?1:2;d.CwtBufferNum=1;}
 else throw Error('Unrecognized CapacityRange '+d.CapacityRange);
 const max={'LSB10':3250,'LSB16':3250,'LSB18':3250,'OBC20M':5000,'OBC25M':5000};
 d.MaxCarBufferLoad=max[d.CarBuffer]*d.CarBufferNum;
 d.MaxCwtBufferLoad=max[d.CwtBuffer]*d.CwtBufferNum;
 }
 function selectA(d){
 const cols=T.CarSling1||[];
 const row=cols.find(r=>c(r['Code number'],v=>has(v,d.CodeNumber))&&c(r['Upper CAP [kg]'],v=>d.Capacity<=N(v))&&c(r['Lower TR [mm]'],v=>N(v)<=d.TRxm)&&c(r['Upper TR [mm]'],v=>d.TRxm<=N(v))&&c(r['Upper AS [mm]'],v=>d.ASxx<=N(v)));
 req(row,'CarSling1');d.UprightSize=row['Upright size'];d.UprightReinforcement=row['Upright reinforcement'];
 const old=d.CapacityRange;
 d.UprightBrace=old==="A0'"?'FlatBar':((old==="A1'"||old==="A2'")&&((d.DoorType==='CO'&&d.EExx>680)||(d.DoorType!=='CO'&&d.EExx>716)))?'ScrewBar':'FlatBar';
 const isG1=d.Entrance==='1D-1G',isCO=d.DoorType==='CO';
 const key=(isG1?'1G':'2G')+'|'+d.UprightBrace+'|'+(isCO?'CO':'2S');
 let p=0,q=0;const h=math.h,h0=math.h0,EE=d.EExx,BS=d.BSxx;
 switch(key){
 case '1G|Without|CO':case '1G|Without|2S':p=0;q=0;break;
 case '1G|FlatBar|CO':case '1G|FlatBar|2S':case '2G|FlatBar|CO':case '2G|FlatBar|2S':
  p=0;q=h0(BS-EE-(key==='1G|FlatBar|CO'?262:key==='1G|FlatBar|2S'?298:key==='2G|FlatBar|CO'?378:414),1215);break;
 case '1G|ScrewBar|CO':case '2G|ScrewBar|CO':p=h(EE-324,1240);q=h(BS-EE-(isG1?282:324),1240);break;
 case '1G|ScrewBar|2S':case '2G|ScrewBar|2S':p=h(EE-360,1240);q=h(BS-EE-(isG1?282:360),1240);break;
 default:throw Error(`Upright geometry unsupported ${key}`);
 }
 d.P121=Math.floor(p/10)*10;d.Q121=Math.floor(q/10)*10;
 const temp=d.HBxx-(d.HHxx+d.KUPx),wiring=d.AAxx-d.JJxx-d.AKxx-31;
 const cc=T.CornerPost||[];
 const a=cc.filter(r=>c(r['HB-(HH+KUP)'],v=>temp<=N(v))&&c(r['Door type'],v=>has(v,d.DoorType))&&c(r['Lower space'],v=>N(v)<=wiring)&&c(r['Upper space'],v=>wiring<=N(v)));
 const front=a.find(r=>c(r['Lower EE'],v=>N(v)<=d.EExx)&&c(r['Upper EE'],v=>d.EExx<=N(v)));
 req(front,'CornerPost front');d.CornerPostF=front.CornerPost;
 d.CornerPostR=isG1?'Without':req(a.find(r=>c(r['Lower EE'],v=>N(v)<=d.BSxx-d.EExx)&&c(r['Upper EE'],v=>d.BSxx-d.EExx<=N(v))),'CornerPost rear').CornerPost;
 d.CarPlatform=d.BBxx>1510?'KWR-100HD':'KWR-100D';
 d.VPRubberNum=isG1?(d.CarPlatform==='KWR-100D'?4:6):(/P08|P11/.test(d.CodeNumber)?4:6);
 const lb=T.LinkBeam1||[],bal=d.BalanceWeightNum>0?'With':'Without';
 const link=lb.find(r=>c(r.Entrance,v=>has(v,d.Entrance))&&c(r['Code number'],v=>has(v,d.CodeNumber))&&c(r['Car platform'],v=>has(v,d.CarPlatform))&&c(r['Rubber number'],v=>has(v,d.VPRubberNum))&&c(r['Balance weight'],v=>has(v,bal))&&c(r['Comp. chain'],v=>has(v,'With'))&&c(r['Lower TR [mm]'],v=>N(v)<=d.TRxm)&&c(r['Upper TR [mm]'],v=>d.TRxm<=N(v)));
 req(link,'LinkBeam1');d.LinkBeamCenter=link['Link beam (center)'];d.LinkBeamRear=link['Link beam (rear)'];
 return {sling:row,front,link,key};
 }
 function duty(d){
 buffer(d);
 E.run(611,[d]);E.run(612,[d]);
 if(has(d.CapacityRange,'A')){const selection=selectA(d);E.run(609,[d]);E.run(613,[d]);E.run(605,[d]);return selection;}
 if(has(d.CapacityRange,'B')){E.run(608,[d]);E.run(610,[d]);E.run(613,[d]);E.run(605,[d]);return {large:true};}
 throw Error('unrecognized class');
 }
 root.G11L1_DUTY_SOURCE={balance,buffer,selectA,duty};
})(typeof window!=='undefined'?window:globalThis);
