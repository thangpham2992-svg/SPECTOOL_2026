/* G11L1 .NET IL arithmetic subset, generated from original MethodDefs 67–79.
 * No user input mutation; impossible/error/unrecognized op fail closed.
 */
(function (root) {
  'use strict';
  const definitions = root.G11L1_IL_PURE_DATA || {};
  const cache = {};
  function index(id) {const ins=definitions[String(id)];if(!ins)throw Error('Unknown method '+id);if(!cache[id])cache[id]={ins,offset:new Map(ins.map((x,j)=>[x[0],j]))};return cache[id];}
  function run(id,args,limit=100000) {
   const {ins,offset}=index(id); const st=[],loc=[]; let pos=0,step=0;
   const pop=()=>{if(!st.length)throw Error(`Stack underflow ${id}@${ins[pos]?.[0]}`);return st.pop();};
   const jump=(x)=>{const i=offset.get(typeof x==="string"?parseInt(x,16):Number(x));if(i===undefined)throw Error(`Unknown jump ${id}->${x}`);pos=i;};
   while(pos<ins.length){if(++step>limit)throw Error('Execution step limit '+id); const [addr,op,arg]=ins[pos++];
    const m=op.match(/^(ldarg|ldloc|stloc|ldc\.i4)(?:\.(.+))?$/); if(m){const type=m[1],sub=m[2]||arg;const i=sub==='s'?Number(arg):Number(sub);if(type==='ldarg'){st.push(args[i]);continue;}if(type==='ldloc'){st.push(loc[i]??0);continue;}if(type==='stloc'){loc[i]=pop();continue;}if(type==='ldc.i4'){st.push(Number(i));continue;}}
    if(op==='ldarg.s'){st.push(args[Number(arg)]);continue;}if(op==='ldloc.s'){st.push(loc[Number(arg)]??0);continue;}if(op==='stloc.s'){loc[Number(arg)]=pop();continue;}
    if(op==='ldc.r8'||op==='ldc.r4'){st.push(Number(arg));continue;}
    if(op==='ldstr'){st.push(arg);continue;}
    if(op==='ldsfld'){const v=arg==='minError'?9999:arg==='maxError'?0:arg==='ppth'?0.001:undefined;if(v===undefined)throw Error(`Unknown field ${arg}`);st.push(v);continue;}
    if(op==='conv.r8'||op==='conv.i4'){st.push(Number(pop()));continue;}
    if(op==='add'||op==='sub'||op==='mul'||op==='div'||op==='rem'){const b=pop(),a=pop();st.push(op==='add'?a+b:op==='sub'?a-b:op==='mul'?a*b:op==='div'?a/b:a%b);continue;}
    if(op==='neg'){st.push(-pop());continue;}
    if(op==='pop'){pop();continue;}if(op==='dup'){const a=pop();st.push(a,a);continue;}
    if(op==='call'||op==='callvirt'){
     const nm=arg.split('.').pop().replace(/^ref:/, '');
     if(nm==='op_Equality'||nm==='op_Inequality'){const b=pop(),a=pop();st.push(nm==='op_Equality'?a===b:a!==b);continue;}
     if(nm==='get_Length'){st.push(String(pop()).length);continue;}
     if(nm==='get_Chars'){const j=pop(),s=pop();st.push(String(s).charCodeAt(j));continue;}
     if(nm==='Round'){const precision=pop(),v=pop(),p=10**Number(precision);st.push(Math.round(v*p)/p);continue;}
     if(nm==='Parse'){st.push(Number(pop()));continue;}
     throw Error(`Unsupported call ${id} @ ${addr}: ${arg}`);
    }
    if(op==='br'||op==='br.s'){jump(arg);continue;}
    if(op==='switch'){const x=pop(),a=JSON.parse(arg);if(Number.isInteger(x)&&x>=0&&x<a.length)jump(a[x]);continue;}
    if(op==='brtrue'||op==='brtrue.s'){if(pop())jump(arg);continue;}
    if(op==='brfalse'||op==='brfalse.s'){if(!pop())jump(arg);continue;}
    const bops=['bge','bgt','ble','blt','beq','bne.un'];const key=op.replace(/\.s$/,'').replace(/\.un$/,'');
    if(bops.includes(key)){const b=pop(),a=pop();let yes=key==='bge'?a>=b:key==='bgt'?a>b:key==='ble'?a<=b:key==='blt'?a<b:key==='beq'?a===b:a!==b;if(yes)jump(arg);continue;}
    if(op==='ret')return pop();
    throw Error(`Unsupported opcode ${id}@${addr}: ${op} ${arg}`);
   }
   throw Error(`Method ${id} ended without ret`);
  }
  root.G11L1_IL_PURE={run};
})(typeof window!=='undefined'?window:globalThis);
