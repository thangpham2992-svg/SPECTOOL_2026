'use strict';
/* Partial managed IL interpreter for BasicMass1 methods 47--62 from Tech Study G11L1 2.5.0.0.
   Reproduction test only. Throws for missing getters or unsupported IL, never guesses. */
(function(root){
const defs=root.G11L1_MACHINE_IL;
const cache={};
const fnIds={CalculateCabNetArea:19,CalculateHCXX:518,CalculateHCXX_R:519,CalculateEEXX:520,CalculateCAXX:521,CalculateCDXX:522,CalculateCarDimension:523,SelectHatchDoorOpr:605,SelectCarPlatformB:608,SelectCarSafety:609,SelectCarSafetyB:610,SelectCwtSafety:611,SelectCounterweight:612,SelectCarDoorOpr:613,CalculateMaxDecoration:43,CalculateRopesMass:44,CalculateSystemMass:45,CalculateCarMass:46,CalculateMachineMass:47,CalculateMachineMassB:48,CalculateUprightMass:49,CalculateUprightMassB:50,CalculateCrossheadMass:51,CalculateBalustradeMass:52,CalculateBalustradeMassB:53,CalculateSupportFrameInst:54,CalculateSupportFrameInstB:55,CalculateSupportFrameInstBR:56,CalculateCarPlatformMass:57,CalculateCarPlatformMassB:58,CalculateCarSafetyMass:59,CalculateCarSafetyMassB:60,CalculateCarDoorOprMass:61,CalculateOtherMass:62};
const banker=(v,d=0)=>{let t=10**d,x=v*t,nearest=Math.trunc(x),f=Math.abs(x-nearest);if(Math.abs(f-.5)<1e-9){let tie=(nearest%2===0)?nearest:nearest+Math.sign(x);return tie/t;}return Math.round(x)/t;};
function doRun(id,args,depth=0){
 if(depth>10)throw Error('Recursion');let obj=cache[id];if(!obj){const ins=defs[id];if(!ins)throw Error('Unknown '+id);obj=cache[id]={ins,byOff:new Map(ins.map((x,i)=>[x[0],i]))};}
 const {ins,byOff}=obj;let p=0,steps=0;const st=[],locals=[];
 const pop=()=>{if(!st.length)throw Error(`underflow method ${id} @ ${ins[p-1]?.[0]}`);return st.pop();};
 const getLocal=i=>locals[i]??0;
 const br=to=>{let n=(typeof to==='string' ? parseInt(to,16) : Number(to)),idx=byOff.get(n);if(idx===undefined)throw Error(`Bad branch ${id} -> ${to}`);p=idx;};
 const push=v=>{if(typeof v==='number'&&!Number.isFinite(v))throw Error('nonfinite output');st.push(v);};
 while(p<ins.length){if(++steps>150000)throw Error(`limit method ${id}`);const [addr,op,arg]=ins[p++];
  if(op==='ldarg.0'||op==='ldarg.1'||op==='ldarg.2'||op==='ldarg.3'){push(args[Number(op.split('.')[1])]);continue;}
  if(op==='ldarg.s'){push(args[+arg]);continue;}
  if(/^ldloc\.[0123]$/.test(op)){push(getLocal(+op.slice(-1)));continue;}
  if(op==='ldloc.s'){push(getLocal(+arg));continue;}
  if(/^stloc\.[0123]$/.test(op)){locals[+op.slice(-1)]=pop();continue;}
  if(op==='stloc.s'){locals[+arg]=pop();continue;}
  if(op==='ldloca.s'){push({localRef:+arg});continue;}
  if(/^ldc\.i4\.[0-8]$/.test(op)){push(+op.slice(-1));continue;}
  if(op==='ldc.i4.s'||op==='ldc.i4'||op==='ldc.r8'||op==='ldc.r4'){push(Number(arg));continue;}
  if(op==='ldsfld'){const fields={ppth:.001,minError:9999,maxError:0,strError:"****"};if(!(arg in fields))throw Error('unknown static field '+arg);push(fields[arg]);continue;}
  if(op==='ldstr'){push(arg);continue;}
  if(op==='ldnull'){push(null);continue;}
  if(op==='dup'){const a=pop();push(a);push(a);continue;}
  if(op==='pop'){pop();continue;}
  if(['mul','add','sub','div','rem'].includes(op)){const b=pop(),a=pop();push(op==='mul'?a*b:op==='add'?a+b:op==='sub'?a-b:op==='div'?a/b:a%b);continue;}
  if(op==='conv.r8'||op==='conv.i4'){const a=Number(pop());push(op==='conv.r8'?a:Math.trunc(a));continue;}
  if(op==='newobj'){if(arg.endsWith('ref:.ctor')){push([]);continue;}throw Error(`newobj ${arg}`);}
  if(op==='newarr'){let n=pop();if(!Number.isInteger(n)||n<0||n>500)throw Error('bad newarr');push(new Array(n).fill(null));continue;}
  if(op==='stelem.ref'){const val=pop(),j=pop(),ar=pop();if(!Array.isArray(ar))throw Error(`stelem expects array in ${id}`);ar[j]=val;continue;}
  if(op==='call'||op==='callvirt'){
    const toks=arg.split(' '), token=toks[0],nm=toks.slice(1).join(' ');
    if(nm.startsWith('DesignDataEl.get_')){let field=nm.split('get_')[1],a=pop();if(a==null||typeof a!=='object')throw Error('getter owner null '+field+' in method '+id+' @ '+addr);if(a[field]===undefined||a[field]===null)throw Error('missing getter '+field+' in '+id);push(a[field]);continue;}
    if(nm.startsWith('DesignDataEl.set_')){let field=nm.split('set_')[1],v=pop(),a=pop();if(a==null||typeof a!=='object')throw Error('setter owner null '+field);a[field]=v;continue;}
    if(nm.startsWith('CommonStudy.CalculateCabNetArea')){const subargs=Array.from({length:6},()=>pop()).reverse();push(doRun(19,subargs,depth+1));continue;}
    if(nm.startsWith('BasicMass1.')){const name=nm.split('.').pop(),refid=fnIds[name];if(!refid)throw Error('unknown method '+name);const count=[54,55,56].includes(refid)?2:1;const subargs=Array.from({length:count},()=>pop()).reverse();push(doRun(refid,subargs,depth+1));continue;}
    if(nm.endsWith('ref:op_Equality')||nm.endsWith('ref:op_Inequality')){const b=pop(),a=pop();push(nm.endsWith('Equality')?a===b:a!==b);continue;}
    if(nm.endsWith('ref:Round')){const digits=token==='0a000049'?0:pop(),x=pop();push(banker(x,digits));continue;}
    if(nm.endsWith('ref:Concat')){let v;if(token==='0a000087')v=pop().join('');else if(token==='0a000065')v=[pop(),pop(),pop()].reverse().join('');else if(token==='0a000044')v=[pop(),pop()].reverse().join('');else if(token==='0a00004c')v=[pop(),pop(),pop(),pop()].reverse().join('');else throw Error(`unrecognized concat variant ${token}`);push(v);continue;}
    if(nm.endsWith('ref:Add')){const value=pop(),list=pop();if(!Array.isArray(list))throw Error('List.Add owner non-array');list.push(value);continue;}
    if(nm.endsWith('ref:Sum')||nm.endsWith('ref:Min')){const list=pop();if(!Array.isArray(list))throw Error('aggregate non-array');push(nm.endsWith('ref:Min')?Math.min(...list):list.reduce((a,b)=>a+Number(b),0));continue;}
    if(nm.endsWith('ref:get_Length')){push(pop().length);continue;}
    if(nm.endsWith('ref:get_Chars')){const n=pop(),value=pop();push(String(value).charCodeAt(n));continue;}
    if(nm.endsWith('ref:Contains')){const substring=pop(),value=pop();push(String(value).includes(String(substring)));continue;}
    if(nm.endsWith('ref:Floor')){push(Math.floor(pop()));continue;}
    if(nm.endsWith('ref:Ceiling')){push(Math.ceil(pop()));continue;}
    if(nm.endsWith('ref:Sqrt')){push(Math.sqrt(pop()));continue;}
    if(nm.endsWith('ref:Pow')){const b=pop(),a=pop();push(Math.pow(a,b));continue;}
    throw Error(`UNSUPPORTED CALL ${id}@${addr}: ${arg}`);
  }
  if(op==='br'||op==='br.s'){br(arg);continue;}
  if(op==='brtrue'||op==='brtrue.s'){if(pop())br(arg);continue;}
  if(op==='brfalse'||op==='brfalse.s'){if(!pop())br(arg);continue;}
  if(op==='switch'){const key=pop(),cases=JSON.parse(arg);if(Number.isInteger(key)&&key>=0&&key<cases.length)br(Number(cases[key]));continue;}
  const ba=op.replace(/\.s$/,'').replace(/\.un$/,'');if(['bge','bgt','ble','blt','beq','bne'].includes(ba)){const b=pop(),a=pop();const yes=ba==='bge'?a>=b:ba==='bgt'?a>b:ba==='ble'?a<=b:ba==='blt'?a<b:ba==='beq'?a===b:a!==b;if(yes)br(arg);continue;}
  if(op==='cgt.un'){const b=pop(),a=pop();push(a>b);continue;}
  if(op==='ret')return st.length?pop():undefined;
  throw Error(`UNSUPPORTED OPCODE ${id}@${addr}: ${op} ${arg}`);
 }
 throw Error('Method fell off '+id);
}
root.G11L1_MACHINE_SOURCE={run:doRun,round:banker};
})(typeof window!=='undefined'?window:globalThis);
