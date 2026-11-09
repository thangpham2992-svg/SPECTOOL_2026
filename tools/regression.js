'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');const root=__dirname+'/..';
const cx={window:{},console};vm.createContext(cx);
for(const f of ['engine-database.js','spec-conditional.js','source-engine.js','legacy-step1.js','conditional-check.js','source-check.js'])vm.runInContext(fs.readFileSync(root+'/'+f,'utf8'),cx,{filename:f});
const D=cx.window.SPEC_ENGINE_DATABASE,check=cx.window.SPEC_SOURCE_CHECK;
const cases=[
 ['MR','baseline',{},'PASS'],
 ['MR','lower OH',{D41:2000},'FAIL','OH'],
 ['MR','mutually exclusive SDE',{A116:1,A117:1},'FAIL','SDE'],
 ['MR','DSJ1+CO one side',{D50:'SUS-HL with DSJ1',D14:'CO',A116:1},'FAIL','SDE'],
 ['MR','not allowed paint under SUS-HL',{D51:'R001'},'FAIL','LightUp'],
 ['MR','missing Speed',{D15:''},'FAIL','Required'],
 ['MR','travelling height too small',{D39:100},'FAIL','Travel'],
 ['MR','decoration weight exceeds limit',{D61:100},'FAIL','Decoration weight'],
 ['MR','failure count B208',{D50:'SUS-HL with DSJ1',D14:'CO',A116:1},'FAIL','B208'],
 ['MRL','baseline',{},'PASS'],
 ['MRL','OH 2000',{D37:2000},'FAIL','OH/PD'],
 ['MRL','PD 1000',{D36:1000},'FAIL','OH/PD'],
 ['MRL','SDE both and one',{A118:1,A119:1},'FAIL','SDE'],
 ['MRL','DSJ1 door SUS304, overall SUS Standard',{D46:'SUS-HL with DSJ1',D40:'SUS Standard',A118:0,A119:0},'PASS'],
 ['MRL','Large window SUS304, overall SUS Standard',{D46:'SUS-HL with Large window W200xH1300',D40:'SUS Standard',A118:0,A119:0},'PASS'],
 ['MRL','SUS-HL wall painted color',{D44:'RAL 9003'},'FAIL','LightUp'],
 ['MRL','missing Speed',{D17:''},'FAIL','Required'],
 ['MRL','invalid observation',{D14:'NY-999'},'FAIL','Observation'],
 ['MRL','invalid entrance type',{D16:'2D-2G'},'FAIL','D521'],
 ['MRL','invalid side drop',{D23:'Bên hông phòng thang (Side drop) - BWH-720G'},'FAIL','D524'],
 ['MRL','forbidden feature X',{A186:1},'FAIL','Feature unavailable'],
];
let passes=0;
for(const [model,label,changes,status,fragment] of cases){const result=check(model,changes,D[model]);const text=result.errors.map(x=>x.id+' '+x.message).join(' \n ');const ok=result.status===status && (!fragment||text.toLowerCase().includes(fragment.toLowerCase()));
 if(!ok){console.error('FAIL test',model,label,'Expected',status,fragment,'actual',result.status,result.errors.slice(0,10));process.exitCode=1;}else passes++;
 console.log((ok?'OK':'NO'),model,label,result.status,result.conditions.pass+'/'+result.conditions.total,'issues',result.errors.length,'unknown',result.unverified.length);
}
console.log('Result',passes,'/',cases.length);
