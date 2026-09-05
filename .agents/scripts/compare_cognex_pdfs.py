import fitz, os, json
paths = [
('native', 'attached_assets/VCCS_VeriWedge(TM)_for_COGNEX_DataMan_TruCheck_2026-09-05_18-2_1788647339776.pdf'),
('adobe', 'attached_assets/VCCS%20VeriWedge(TM)%20for%20COGNEX%20DataMan%20TruCheck_2026-_1788647548572.pdf'),
]
out='.agents/outputs/pdf-compare'
os.makedirs(out,exist_ok=True)
summary={}
for label,path in paths:
 doc=fitz.open(path)
 summary[label]={'pages':len(doc),'metadata':doc.metadata,'page_sizes':[]}
 for i,page in enumerate(doc):
  summary[label]['page_sizes'].append([page.rect.width,page.rect.height])
  pix=page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
  pix.save(f'{out}/{label}-page-{i+1}.png')
  open(f'{out}/{label}-page-{i+1}.txt','w',encoding='utf-8').write(page.get_text())
 print(label, len(doc), summary[label]['page_sizes'], doc.metadata)
open(f'{out}/summary.json','w').write(json.dumps(summary,indent=2))
