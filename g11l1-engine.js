'use strict';
/* Isolated G11L1P Rev.C workbook evaluator; do not alter SpecTool source engine. */
(()=>{
 const Base=window.SpecExcelSourceEngine.SourceEngine;
 function cnum(s){let n=0;for(const c of s)n=n*26+c.charCodeAt(0)-64;return n;}
 class G11Calc extends Base{
  constructor(data,changes){super(data,changes);this.lastRows={};for(const [sh,cells] of Object.entries(data.sheets||{}))this.lastRows[sh]=Math.max(1,...Object.keys(cells).map(k=>Number(k.match(/\d+$/)?.[0]||0)));}
  // Keep the SpecTool source engine untouched: the G11 calculator is isolated.
  // Tool! values are the *working inputs* of the G11L1P workbook, never
  // modifications to the user's Spectool/SpecForm inputs.
  at(sheet,addr){
   const a=String(addr).replace(/\$/g,'').toUpperCase();
   if(sheet==='Tool' && Object.prototype.hasOwnProperty.call(this.changes,a)){
    const v=this.changes[a];
    if(v===null||v===undefined)return '';
    if(typeof v==='number'||typeof v==='boolean')return v;
    if(/^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(String(v)))return Number(v);
    return v;
   }
   return super.at(sheet,addr);
  }
  call(name,args,ctx){
   if(name.toUpperCase()==='PI')return Math.PI;
   if(name.toUpperCase()==='SUM'){
    let result=0;
    for(const node of args){const v=this.eval(node,ctx);const a=v?.isGrid?v.values():Array.isArray(v)?v.flat(Infinity):[v];
     for(const x of a)if(typeof x==='number'&&Number.isFinite(x))result+=x;
    }
    return result;
   }
   if(name.toUpperCase()==='LOOKUP'&&args.length>=2){
    const val=this.eval(args[0],ctx),v=this.eval(args[1],ctx),r=args[2]?this.eval(args[2],ctx):v;
    if(v?.isGrid&&r?.isGrid){
     const a=v.values(),b=r.values();let i=-1;
     for(let j=0;j<a.length;j++){
      if(a[j]===''||a[j]==null)continue;
      const na=Number(a[j]),nb=Number(val);
      if(Number.isFinite(na)&&Number.isFinite(nb)?na<=nb:String(a[j]).localeCompare(String(val),'en')<=0)i=j;
     }
     if(i>=0)return b[i];
     throw new window.SpecExcelSourceEngine.XLFailure('#N/A','LOOKUP '+val);
    }
   }
   if(name.toUpperCase()==='MATCH' && args.length>1){
    const mode=args.length>2?Number(this.eval(args[2],ctx)):1;
    if(mode===-1){
     const pat=this.eval(args[0],ctx),range=this.eval(args[1],ctx),arr=range?.isGrid?range.values():[];
     if(typeof pat==='number'){
      let found=-1,best=Infinity;
      for(let i=0;i<arr.length;i++){
       if(typeof arr[i]!=='number' || !Number.isFinite(arr[i]))continue;
       if(arr[i]>=pat && arr[i]<best){best=arr[i];found=i;}
      }
      if(found>=0)return found+1;
     }
     const exact=arr.findIndex(x=>typeof x===typeof pat && String(x).toLowerCase()===String(pat).toLowerCase());
     if(exact>=0)return exact+1;
     throw new window.SpecExcelSourceEngine.XLFailure('#N/A','MATCH nearest: '+pat);
    }
    try{
     const pat=this.eval(args[0],ctx),search=this.eval(args[1],ctx);
     const mode=args.length>2?Number(this.eval(args[2],ctx)):1;
     if(mode===0 && typeof pat==='string' && /[?*]/.test(pat) && search?.isGrid){
      const re=new RegExp('^'+pat.replace(/[.+^${}()|[\]\\]/g,'\\$&').replace(/\*/g,'.*').replace(/\?/g,'.')+'$','i');
      let k=0;for(const v of search.items()){k++;if(re.test(String(v)))return k;}
      throw new window.SpecExcelSourceEngine.XLFailure('#N/A','MATCH failed: '+pat);
     }
    }catch(e){if(e.code==='#N/A')throw e;}
   }
   return super.call(name,args,ctx);
  }
  nodeRef(node,ctx){
   if(node.type==='name'){
    let id=String(node.id||'').replace(/\$/g,'');let sh=ctx.sheet,ref=id;
    if(id.includes('!')){let p=id.lastIndexOf('!');sh=id.slice(0,p).replace(/^'|'$/g,'');ref=id.slice(p+1);}
    if(/^[A-Z]{1,3}$/i.test(ref)&&!this.d.names.global?.[ref]&&!this.d.names.local?.[sh]?.[ref]){
     return {sheet:sh,r1:1,r2:Math.min(10000,(this.lastRows[sh]||500)+10),c1:cnum(ref.toUpperCase()),c2:cnum(ref.toUpperCase())};
    }
   }
   return super.nodeRef(node,ctx);
  }
 }
 window.G11L1ExcelEngine=G11Calc;
})();
