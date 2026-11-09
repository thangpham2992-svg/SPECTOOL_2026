import json,sys,pathlib
sys.path.insert(0,'/mnt/data/SpecWebTool_Beta')
from engine import Book
p=pathlib.Path(__file__).resolve().parents[1]
data={model:Book(str(p/'database'/file)).data() for model,file in {'MR':'SpecToolGXLM2P.xlsx','MRL':'SpecToolGQXL2P.xlsx'}.items()}
(p/'embedded-data.js').write_text('window.SPEC_EMBEDDED='+json.dumps(data,ensure_ascii=False,separators=(',',':'))+';',encoding='utf-8')

import base64
a={k:base64.b64encode((p/'database'/f).read_bytes()).decode() for k,f in [('MR','SpecToolGXLM2P.xlsx'),('MRL','SpecToolGQXL2P.xlsx')]}
(p/'workbooks.js').write_text('window.SPEC_WORKBOOKS='+json.dumps(a,separators=(',',':'))+';',encoding='utf-8')
