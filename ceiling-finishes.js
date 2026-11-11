'use strict';
window.SPEC_CEILING_FINISH={
 options(model,ceiling){
   const c=String(ceiling||'').toUpperCase();
   if(model==='MR'){
     if(['L200S','L210S','N300S'].includes(c))return ['SUS-HL'];
     if(['S300','L310','N120','N130','N140','S200'].includes(c))return [''];
     return ['Y033','Y055','Y073'];
   }
   if(c==='BN1')return ['','No color (BN1)'];
   if(c.startsWith('DL6'))return ['SUS-HL','SUS-M','LA03','LA07'];
   if(['CL2','DL5','DL7'].some(p=>c.startsWith(p)))return ['Y135D','Y136D','Y055B','SUS-HL','SUS-M','LA03'];
   return ['Y135D','Y136D','Y055B'];
 },
 next(model,ceiling,current){const opts=this.options(model,ceiling);return opts.includes(String(current))?String(current):opts[0]??'';}
};
