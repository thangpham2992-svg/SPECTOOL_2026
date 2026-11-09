'use strict';
/* Workbook-native formula evaluator for SpecForm/DataTable. No eval(), no cached-formula PASS.
   All formula nodes are parsed into a restricted Excel expression tree. */
(()=>{
class XLFailure extends Error {constructor(code, detail=''){super(detail||code);this.name='XLFailure';this.code=code;}}
const xlerr=(code,reason='')=>new XLFailure(code,reason);
const iserr=v=>v instanceof XLFailure;
const alpha='ABCDEFGHIJKLMNOPQRSTUVWXYZ';
function colNum(str){let n=0;for(const c of str.toUpperCase()){n=n*26+(c.charCodeAt(0)-64)}return n;}
function colText(n){let out='';for(;n>0;n=Math.floor((n-1)/26))out=alpha[(n-1)%26]+out;return out;}
function coord(s){const m=/^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s);return m?{col:colNum(m[1]),row:+m[2]}:null;}
function rect(r){let p=String(r).split(':');let a=coord(p[0]),b=coord(p[1]||p[0]);if(!a||!b)throw xlerr('#REF!','Invalid rectangle '+r);return {r1:Math.min(a.row,b.row),r2:Math.max(a.row,b.row),c1:Math.min(a.col,b.col),c2:Math.max(a.col,b.col)};}
class Grid {constructor(sheet,r1,r2,c1,c2,get){this.sheet=sheet;this.r1=r1;this.r2=r2;this.c1=c1;this.c2=c2;this.get=get;this.isGrid=true;}value(r,c){return this.get(this.sheet,colText(c)+r);}*items(){for(let r=this.r1;r<=this.r2;r++)for(let c=this.c1;c<=this.c2;c++)yield this.value(r,c)}values(){return Array.from(this.items());}rows(){let out=[];for(let r=this.r1;r<=this.r2;r++){let row=[];for(let c=this.c1;c<=this.c2;c++)row.push(this.value(r,c));out.push(row)}return out;}}
function scalar(v){let count=0;while((v?.isGrid||Array.isArray(v))&&count++<15)v=v?.isGrid?v.value(v.r1,v.c1):v[0];return v;}
function vals(v){return v?.isGrid?v.values():Array.isArray(v)?v.flat(Infinity):[v];}
function empty(v){return v===null||v===undefined||v==='';}
function xnum(v){v=scalar(v);if(iserr(v))throw v;if(empty(v))return 0;if(typeof v==='boolean')return v?1:0;let n=Number(v);if(!Number.isFinite(n))throw xlerr('#VALUE!','Cannot convert '+String(v)+' to number');return n;}
function rawnum(v){let n=xnum(v);return Number.isFinite(n)?n:0;}
function xstr(v){v=scalar(v);if(iserr(v))throw v;if(empty(v))return '';if(typeof v==='boolean')return v?'TRUE':'FALSE';return String(v);}
function xbool(v){v=scalar(v);if(iserr(v))throw v;if(empty(v))return false;if(typeof v==='boolean')return v;if(typeof v==='number')return v!==0;if(typeof v==='string'){if(/^true$/i.test(v))return true;if(/^false$/i.test(v))return false;return v!=='';}return !!v;}
function cmp(a,b){a=scalar(a);b=scalar(b);if(iserr(a))throw a;if(iserr(b))throw b;if(empty(a))a=(typeof b==='number'?0:'');if(empty(b))b=(typeof a==='number'?0:'');if(typeof a==='boolean')a=+a;if(typeof b==='boolean')b=+b;let an=Number(a),bn=Number(b);if(a!==''&&b!==''&&Number.isFinite(an)&&Number.isFinite(bn))return Math.sign(an-bn);return xstr(a).toLowerCase().localeCompare(xstr(b).toLowerCase(),'en');}
function lex(formula){let src=String(formula||'').trim().replace(/^=/,'');let a=[],i=0;
 const p=(t,v)=>a.push({t,v});
 while(i<src.length){let c=src[i];if(/\s/.test(c)){i++;continue;}
  if(c==='"'){let value='';i++;while(i<src.length){if(src[i]==='"'){if(src[i+1]==='"'){value+='"';i+=2;continue;}i++;break;}value+=src[i++];}p('str',value);continue;}
  if(c==="'"){let start=i++,name='';while(i<src.length){if(src[i]==="'"){if(src[i+1]==="'"){name+="'";i+=2;continue;}i++;break;}name+=src[i++];}if(src[i]==='!'){i++;let j=i;while(i<src.length&&/[\w.$]/.test(src[i]))i++;p('id',name+'!'+src.slice(j,i));}else p('id',name);continue;}
  if(/[0-9]/.test(c)||c==='.'&&/[0-9]/.test(src[i+1]||'')){let match=/^(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?/.exec(src.slice(i));if(!match)throw xlerr('#PARSE!',src.slice(i));p('num',Number(match[0]));i+=match[0].length;continue;}
  if(/[A-Za-z_$]/.test(c)){let start=i;while(i<src.length&&/[A-Za-z_0-9.$]/.test(src[i]))i++;
   if(src[i]==='!'){i++;while(i<src.length&&/[A-Za-z_$0-9.]/.test(src[i]))i++;}
   if(src[i]==='['){let nest=0;do{if(src[i]==='[')nest++;if(src[i]===']')nest--;i++;}while(i<src.length&&nest>0);}
   p('id',src.slice(start,i));continue;
  }
  if(c==='#'){let m=/^#[A-Za-z0-9/?!]+/.exec(src.slice(i));if(m){p('error',m[0]);i+=m[0].length;continue}}
  const op=src.slice(i,i+2);if(['<=','>=','<>'].includes(op)){p('op',op);i+=2;continue;}
  if('+-*/^&%=<>,:();@'.includes(c)){p(c==='('||c===')'||c===','||c===';'?'pun':'op',c);i++;continue;}
  throw xlerr('#PARSE!',`Unexpected '${c}' at ${i} in ${src.slice(0,70)}`);
 }p('eof','');return a;
}
const BP={'=':2,'<>':2,'<':2,'<=':2,'>':2,'>=':2,'&':3,'+':4,'-':4,'*':5,'/':5,'^':7,':':9};
function parse(formula){const t=lex(formula);let i=0;const peek=()=>t[i],take=()=>t[i++];
 function expr(bp=0){let s=take(),node;
  if(s.t==='num'||s.t==='str')node={type:'literal',val:s.v};
  else if(s.t==='error')node={type:'error',code:s.v};
  else if(s.v==='('){node=expr(0);if(take().v!==')')throw xlerr('#PARSE!','Missing )')}
  else if(['+','-','@'].includes(s.v))node={type:'unary',op:s.v,val:expr(8)};
  else if(s.t==='id'){
   if(peek().v==='('){take();const args=[];if(!(peek().t==='pun'&&peek().v===')')){while(true){if(peek().t==='pun'&&peek().v===')'){args.push({type:'literal',val:''});break;}args.push(expr(0));if(peek().t==='pun'&&peek().v===')')break;const sep=take();if(sep.t!=='pun'||![',',';'].includes(sep.v))throw xlerr('#PARSE!','Missing argument separator') }}take();node={type:'call',name:s.v.toUpperCase(),args};}
   else node={type:'name',id:s.v};
  }else throw xlerr('#PARSE!',`Unexpected ${s.v}`);
  while(true){let n=peek();if(n.v==='%'&&8>bp){take();node={type:'percent',val:node};continue;}let b=BP[n.v]||0;if(!b||b<=bp)break;take();let right=expr(n.v==='^'?b-1:b);node={type:'bin',op:n.v,left:node,right};}
  return node;
 }
 const tree=expr(0);if(peek().t!=='eof')throw xlerr('#PARSE!',`Unexpected trailing '${peek().v}'`);return tree;
}
function choice(val,criterion){const original=scalar(criterion);if(typeof original==='boolean')return xbool(val)===original;
 const s=xstr(original),m=/^(<=|>=|<>|=|<|>)(.*)$/.exec(s);let op=m?m[1]:'=',target=m?m[2]:original;
 if(typeof target==='string' && /[?*]/.test(target)){const escaped=target.replace(/[.+^${}()|[\]\\]/g,'\\$&').replace(/\*/g,'.*').replace(/\?/g,'.');const match=new RegExp('^'+escaped+'$','i').test(xstr(val));return op==='<>'?!match:match;}
 const c=cmp(val,target);return op==='='?c===0:op==='<>'?c!==0:op==='>='?c>=0:op==='<='?c<=0:op==='>'?c>0:c<0;
}
class SourceEngine{
 constructor(data, changes={},extra={}){this.d=data;this.changes=changes;this.extra=extra;this.memo=new Map();this.inflight=new Set();this.parsed=new Map();this.unverified=[];this.cacheQueries=new Map();this.callCount=0;this.maxCalls=250000;this.hits=0;}
 ast(formula){if(!this.parsed.has(formula))this.parsed.set(formula,parse(formula));return this.parsed.get(formula)}
 at(sheet,addr){addr=String(addr).replace(/\$/g,'').toUpperCase();let key=sheet+'!'+addr;
  if(sheet==='SpecForm'&&Object.prototype.hasOwnProperty.call(this.changes,addr)){let v=this.changes[addr];if(v===''||v===null)return '';if(/^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(String(v)))return +v;return v;}
  if(this.memo.has(key)){const v=this.memo.get(key);if(iserr(v))throw v;return v;}
  if(this.inflight.has(key))throw xlerr('#CYCLE!',key);
  if(++this.callCount>this.maxCalls)throw xlerr('#LIMIT!','Calculation limit exceeded');
  let cell=this.d.sheets?.[sheet]?.[addr];if(cell===undefined)return '';
  if(typeof cell!=='object'||cell===null)return cell;
  if(!Object.prototype.hasOwnProperty.call(cell,'f'))return cell.v??'';if(cell.f==='')throw xlerr('#UNRESOLVED!','Công thức Excel rỗng tại '+key);
  this.inflight.add(key);
  try{let result=scalar(this.eval(this.ast(cell.f),{sheet,addr}));this.memo.set(key,result);return result;}
  catch(err){const failure=iserr(err)?err:xlerr('#RUNTIME!',String(err?.message||err));this.memo.set(key,failure);throw failure;}
  finally{this.inflight.delete(key)}
 }
 safe(sheet,addr){try{let v=this.at(sheet,addr);return {ok:true,value:scalar(v)};}catch(e){return {ok:false,error:iserr(e)?e.code+': '+e.message:String(e)}}}
 range(sheet,rectangle){let x=rect(rectangle);return new Grid(sheet,x.r1,x.r2,x.c1,x.c2,(s,a)=>this.at(s,a));}
 structured(token,context){let i=token.indexOf('['),tn=token.slice(0,i),spec=token.slice(i);const table=this.d.tables[tn];if(!table)throw xlerr('#NAME?',`Unknown Excel table ${tn}`);
  let bounds=rect(table.ref),start=bounds.r1+1,end=bounds.r2;let cols=table.columns;
  const idx=name=>cols.findIndex(c=>c.toLowerCase()===name.toLowerCase());
  // TblCondition1[[#This Row],[Check1]] or TblFeature_Emergency[[Feature]:[Apply2]]
  const row=/\[#This Row\]/i.test(spec);
  let clean=spec.replace(/\[#This Row\]\s*,?/ig,'');let keys=[...clean.matchAll(/\[([^\[\]#][^\[\]]*)\]/g)].map(m=>m[1]).filter(s=>s&&s!=='#This Row');
  let begin=0,finish=cols.length-1;
  if(keys.length){begin=idx(keys[0]);finish=idx(keys.length>1?keys.at(-1):keys[0]);if(begin<0||finish<0)throw xlerr('#REF!',`Table ${tn} column ${keys.join(':')} missing`);}
  let c1=bounds.c1+Math.min(begin,finish),c2=bounds.c1+Math.max(begin,finish);
  if(row){let r=Number(/\d+$/.exec(context.addr)?.[0]);if(r<start||r>end)throw xlerr('#REF!',`Row ${r} outside ${tn}`);return new Grid(table.sheet,r,r,c1,c2,(s,a)=>this.at(s,a));}
  return new Grid(table.sheet,start,end,c1,c2,(s,a)=>this.at(s,a));
 }
 name(id,ctx){const u=String(id),upper=u.toUpperCase();if(upper==='TRUE')return true;if(upper==='FALSE')return false;if(upper==='NULL')return '';
  let scope=null,base=u;if(u.includes('!')){const part=u.indexOf('!');scope=u.slice(0,part).replace(/^'|'$/g,'');base=u.slice(part+1);}
  const xy=coord(base);if(xy)return this.at(scope||ctx.sheet,colText(xy.col)+xy.row);
  if(base.includes('['))return this.structured(base,ctx);
  const validSheet=scope||ctx.sheet;
  let formula=this.d.names.local?.[validSheet]?.[base]??this.d.names.global?.[base]??this.d.names.local?.SpecFormExport?.[base]??this.d.names.local?.SpecForm?.[base]??this.d.names.local?.DataTable?.[base];
  if(formula===undefined)throw xlerr('#NAME?',`Name ${u} not defined`);
  return this.eval(this.ast(formula),{sheet:validSheet,addr:ctx.addr});
 }
 nodeRef(node,ctx){if(node.type==='name'){
   const u=node.id;if(u.includes('['))return this.structured(u,ctx);
   let scope=ctx.sheet,addr=u;if(u.includes('!')){let pos=u.indexOf('!');scope=u.slice(0,pos).replace(/^'|'$/g,'');addr=u.slice(pos+1)}
   if(coord(addr)){let p=coord(addr);return {sheet:scope,r1:p.row,r2:p.row,c1:p.col,c2:p.col};}
   const name=this.d.names.local?.[scope]?.[addr]||this.d.names.global?.[addr]||this.d.names.local?.SpecForm?.[addr];
   if(name){let m=/^'?([^']+)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(name);if(m)return{sheet:m[1],r1:+m[3],r2:+(m[5]||m[3]),c1:colNum(m[2]),c2:colNum(m[4]||m[2])};}
  }
  if(node.type==='bin'&&node.op===':'){let g=this.eval(node,ctx);if(g?.isGrid)return g;}
  let v=this.eval(node,ctx);if(v?.isGrid)return v;throw xlerr('#REF!',`Not a reference ${node.type}`);
 }
 getRange(left,right,ctx){let a=this.nodeRef(left,ctx),b=this.nodeRef(right,ctx),sh=a.sheet||b.sheet||ctx.sheet;
  return new Grid(sh,Math.min(a.r1,b.r1),Math.max(a.r2,b.r2),Math.min(a.c1,b.c1),Math.max(a.c2,b.c2),(s,r)=>this.at(s,r));
 }
 eval(node,ctx){if(node.type==='literal')return node.val;if(node.type==='error')throw xlerr(node.code);
  if(node.type==='name')return this.name(node.id,ctx);
  if(node.type==='unary'){if(node.op==='@')return scalar(this.eval(node.val,ctx));let val=xnum(this.eval(node.val,ctx));return node.op==='-'?-val:val;}
  if(node.type==='percent')return xnum(this.eval(node.val,ctx))/100;
  if(node.type==='bin'){
   if(node.op===':')return this.getRange(node.left,node.right,ctx);
   const a=this.eval(node.left,ctx),b=this.eval(node.right,ctx);
   if(['=','<>','<','<=','>','>='].includes(node.op)){let x=cmp(a,b);return node.op==='='?x===0:node.op==='<>'?x!==0:node.op==='<'?x<0:node.op==='<='?x<=0:node.op==='>'?x>0:x>=0;}
   if(node.op==='&')return xstr(a)+xstr(b);
   const n=xnum(a),m=xnum(b);if(node.op==='+')return n+m;if(node.op==='-')return n-m;if(node.op==='*')return n*m;if(node.op==='/'){if(m===0)throw xlerr('#DIV/0!');return n/m;}if(node.op==='^')return n**m;
  }
  if(node.type==='call')return this.call(node.name,node.args,ctx);
  throw xlerr('#PARSE!','Unexpected expression');
 }
 call(name,args,ctx){const A=i=>this.eval(args[i]??{type:'literal',val:''},ctx),N=i=>xnum(A(i)),S=i=>xstr(A(i)),B=i=>xbool(A(i));
  name=name.replace(/^_XLFN\./,'');
  if(name==='IF'){return B(0)?A(1):args.length>2?A(2):false;}
  if(name==='IFERROR'||name==='IFNA'){try{return A(0);}catch(e){if(!iserr(e))throw e;if(name==='IFNA'&&e.code!=='#N/A')throw e;return A(1);}}
  if(name==='IFs'||name==='IFS'){for(let i=0;i<args.length-1;i+=2)if(B(i))return A(i+1);throw xlerr('#N/A');}
  if(name==='AND'){for(let i=0;i<args.length;i++){let x=A(i);if(x?.isGrid){for(let r=x.r1;r<=x.r2;r++)for(let c=x.c1;c<=x.c2;c++){if(x.sheet===ctx.sheet&&colText(c)+r===ctx.addr)continue;if(!xbool(x.value(r,c)))return false;}}else for(const v of vals(x))if(!xbool(v))return false;}return true;}
  if(name==='OR'){for(let i=0;i<args.length;i++)for(const v of vals(A(i)))if(xbool(v))return true;return false;}
  if(name==='NOT')return !B(0);
  if(name==='ISNUMBER'){try{const v=scalar(A(0));return typeof v==='number'&&Number.isFinite(v);}catch{return false}}
  if(name==='ISERROR'){try{A(0);return false;}catch(e){return iserr(e)}}
  if(name==='ISBLANK'){try{return empty(A(0));}catch{return false}}
  if(name==='LEFT')return S(0).slice(0,args.length>1?N(1):1);
  if(name==='RIGHT')return S(0).slice(-Math.max(0,args.length>1?N(1):1));
  if(name==='MID')return S(0).substring(Math.max(0,N(1)-1),Math.max(0,N(1)-1)+N(2));
  if(name==='LEN')return S(0).length;
  if(name==='FIND'||name==='SEARCH'){let needle=S(0),hay=S(1),start=args.length>2?N(2)-1:0;let pos=name==='SEARCH'?hay.toLowerCase().indexOf(needle.toLowerCase(),start):hay.indexOf(needle,start);if(pos<0)throw xlerr('#VALUE!','Not found: '+needle);return pos+1;}
  if(name==='SUBSTITUTE'){let s=S(0),from=S(1),to=S(2);if(args.length<=3)return s.split(from).join(to);let occ=N(3),i=0;return s.replace(new RegExp(from.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'g'),matched=>++i===occ?to:matched);}
  if(name==='UPPER')return S(0).toUpperCase();if(name==='LOWER')return S(0).toLowerCase();if(name==='TRIM')return S(0).trim().replace(/\s+/g,' ');
  if(name==='CONCATENATE')return args.map((_,i)=>S(i)).join('');
  if(name==='TEXT'){let value=A(0),fmt=S(1);if(typeof value==='number'&&/^[0#,]+(?:\.[0#]+)?$/.test(fmt)){const decimal=(fmt.split('.')[1]||'').length;let str=value.toFixed(decimal);if(fmt.includes('#'))str=str.replace(/0+$/,'').replace(/\.$/,'');return str;}return xstr(value);}
  if(name==='VALUE')return xnum(A(0));
  if(name==='SUM')return args.reduce((a,_,i)=>a+vals(A(i)).reduce((b,v)=>b+(empty(v)?0:rawnum(v)),0),0);
  if(name==='COUNT')return args.reduce((a,_,i)=>a+vals(A(i)).filter(v=>typeof v==='number').length,0);
  if(name==='COUNTA')return args.reduce((a,_,i)=>a+vals(A(i)).filter(v=>!empty(v)).length,0);
  if(name==='MIN')return Math.min(...args.flatMap((_,i)=>vals(A(i))).filter(v=>typeof v==='number'));
  if(name==='MAX')return Math.max(...args.flatMap((_,i)=>vals(A(i))).filter(v=>typeof v==='number'));
  if(name==='ROUND'){const f=10**N(1),x=N(0)*f;return (Math.sign(x)*Math.round(Math.abs(x)))/f;}
  if(name==='ROUNDUP'){const f=10**N(1),x=N(0)*f;return Math.sign(x)*Math.ceil(Math.abs(x))/f;}
  if(name==='ROUNDDOWN'){const f=10**N(1),x=N(0)*f;return Math.sign(x)*Math.floor(Math.abs(x))/f;}
  if(name==='ABS')return Math.abs(N(0));if(name==='INT')return Math.floor(N(0));
  if(name==='COLUMN'){if(args.length===0)return coord(ctx.addr).col;const ref=this.nodeRef(args[0],ctx);return ref.c1;}
  if(name==='ROW'){if(args.length===0)return coord(ctx.addr).row;const ref=this.nodeRef(args[0],ctx);return ref.r1;}
  if(name==='COLUMNS'){const ref=this.nodeRef(args[0],ctx);return ref.c2-ref.c1+1;}
  if(name==='ROWS'){const ref=this.nodeRef(args[0],ctx);return ref.r2-ref.r1+1;}
  if(name==='INDEX'){const r=A(0),rn=N(1),cn=args.length>2?N(2):1;if(!r?.isGrid)throw xlerr('#REF!','INDEX expects range');if(rn<1||cn<1||rn>r.r2-r.r1+1||cn>r.c2-r.c1+1)throw xlerr('#REF!');return r.value(r.r1+rn-1,r.c1+cn-1);}
  if(name==='MATCH'){let lookup=A(0),r=A(1),mode=args.length>2?N(2):1;let list=vals(r);if(mode===0){let idx=list.findIndex(x=>cmp(x,lookup)===0);if(idx<0)throw xlerr('#N/A','MATCH failed: '+xstr(lookup));return idx+1;}let match=-1;for(let i=0;i<list.length;i++)if((mode===1&&cmp(list[i],lookup)<=0)||(mode===-1&&cmp(list[i],lookup)>=0))match=i;if(match<0)throw xlerr('#N/A');return match+1;}
  if(name==='VLOOKUP'){let lookup=A(0),table=A(1),index=N(2),exact=args.length<=3?false:!B(3);if(!table?.isGrid)throw xlerr('#REF!','VLOOKUP expects range');if(index<1||index>table.c2-table.c1+1)throw xlerr('#REF!','Invalid column');let result=-1;for(let r=table.r1;r<=table.r2;r++){let first=table.value(r,table.c1);if(cmp(first,lookup)===0){result=r;break;}if(!exact&&cmp(first,lookup)<=0)result=r;}if(result<0)throw xlerr('#N/A','VLOOKUP failed: '+xstr(lookup));return table.value(result,table.c1+index-1);}
  if(name==='HLOOKUP'){let lookup=A(0),grid=A(1),r=N(2);if(!grid?.isGrid)throw xlerr('#REF!');let out=grid.rows(),index=out[0].findIndex(v=>cmp(v,lookup)===0);if(index<0||r<1||r>out.length)throw xlerr('#N/A');return out[r-1][index];}
  if(name==='COUNTIF'){let range=vals(A(0)),criterion=A(1);return range.filter(v=>choice(v,criterion)).length;}
  if(name==='COUNTIFS'){let series=[],length=0;for(let i=0;i<args.length;i+=2){let arr=vals(A(i)),criteria=A(i+1);series.push({arr,criteria});length=Math.max(length,arr.length)}let c=0;for(let j=0;j<length;j++)if(series.every(s=>choice(s.arr[j],s.criteria)))c++;return c;}
  if(name==='SUMIF'){let rg=vals(A(0)),criterion=A(1),sum=vals(args.length>2?A(2):A(0)),total=0;for(let i=0;i<rg.length;i++)if(choice(rg[i],criterion)){if(typeof sum[i]==='number')total+=sum[i];}return total;}
  if(name==='SUMIFS'){let sum=vals(A(0)),series=[];for(let i=1;i<args.length;i+=2)series.push({a:vals(A(i)),criteria:A(i+1)});let total=0;for(let j=0;j<sum.length;j++)if(series.every(s=>choice(s.a[j],s.criteria))&&typeof sum[j]==='number')total+=sum[j];return total;}
  if(name==='SUBTOTAL'){let method=N(0),vs=args.slice(1).flatMap((_,i)=>{let val=A(i+1);if(val?.isGrid&&val.sheet===ctx.sheet&&val.r1<=coord(ctx.addr).row&&val.r2>=coord(ctx.addr).row&&val.c1<=coord(ctx.addr).col&&val.c2>=coord(ctx.addr).col){let out=[];for(let r=val.r1;r<=val.r2;r++)for(let c=val.c1;c<=val.c2;c++){if(colText(c)+r===ctx.addr)continue;out.push(val.value(r,c));}return out;}return vals(val);});if([3,103].includes(method))return vs.filter(v=>!empty(v)).length;if([9,109].includes(method))return vs.reduce((s,v)=>s+(typeof v==='number'?v:0),0);throw xlerr('#NAME?','Unsupported SUBTOTAL '+method)}
  if(name==='OFFSET'){let base=this.nodeRef(args[0],ctx),rd=N(1),cd=N(2),h=args.length>3?N(3):base.r2-base.r1+1,w=args.length>4?N(4):base.c2-base.c1+1;return new Grid(base.sheet,base.r1+rd,base.r1+rd+h-1,base.c1+cd,base.c1+cd+w-1,(s,r)=>this.at(s,r));}
  if(name==='INDIRECT'){
   let target=S(0).replace(/\$/g,''),p=target.split('!'),sh=p.length>1?p.shift().replace(/^'|'$/g,''):ctx.sheet,addr=p.at(-1);
   if(/^([A-Z]+\d+)(:[A-Z]+\d+)?$/i.test(addr))return this.range(sh,addr);
   if(addr.includes('['))return this.structured(addr,ctx);
   return this.name(sh+'!'+addr,ctx);
  }
  if(name==='CHAR')return String.fromCharCode(N(0));
  throw xlerr('#UNSUPPORTED!',`Function ${name} is not implemented`);
 }
}
window.SpecExcelSourceEngine={SourceEngine,XLFailure,parse,lex};
})();
