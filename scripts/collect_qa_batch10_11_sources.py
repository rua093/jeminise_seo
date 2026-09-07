import csv,json,hashlib,importlib.util
from pathlib import Path
spec=importlib.util.spec_from_file_location('c',Path(__file__).with_name('collect_qa_batch2_sources.py')); m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
ROOT=m.ROOT; RUN_ID=m.RUN_ID; INV=list(csv.DictReader((ROOT/'seo_runs/jeminise.com/20260906_234129/inventory.csv').open(encoding='utf-8-sig')))
def run(batch,start):
 qid=f'20260907_18{batch:02d}00'; qa=ROOT/'seo_runs/jeminise.com/20260906_234129/qa'/qid; qa.mkdir(parents=True,exist_ok=True); (qa/'source_snapshot').mkdir(exist_ok=True); wb=ROOT/'resutls/jeminise.com/20260906_234129/batches'/f'SEO_Product_Optimization_through_batch_{batch:03d}.xlsx'; (qa/'source_snapshot'/wb.name).write_bytes(wb.read_bytes()); rows=INV[start:start+10]; keys={r['product_key'] for r in rows}; handles={r['Handle'] for r in rows}; urls={r['product_url'] for r in rows}; book=m.load_workbook(wb,read_only=True,data_only=True); sub={'products':[r for r in m.rows(book,'SEO_Products') if r['product_key'] in keys],'images':[r for r in m.rows(book,'Image_Audit') if r['Handle'] in handles],'product_evidence':[r for r in m.rows(book,'Product_Evidence') if r['product_url'] in urls],'keyword_map':[r for r in m.rows(book,'Keyword_Map') if r['product_key'] in keys],'buyer_search_research':[r for r in m.rows(book,'Buyer_Search_Research') if r['product_key'] in keys]}; (qa/'submitted_batch_data.json').write_text(json.dumps(sub,ensure_ascii=False,indent=2),encoding='utf8'); live=[]
 for r in rows:
  x={'inventory':r}
  try:x['live']=m.fetch_live(r['product_url']);x['fetch_error']=''
  except Exception as e:x['live']={};x['fetch_error']=str(e)
  live.append(x)
 (qa/'live_source_comparison.json').write_text(json.dumps(live,ensure_ascii=False,indent=2),encoding='utf8'); print(batch,qid,len(sub['products']),len(sub['images']))
for b,s in ((10,90),(11,100)): run(b,s)
