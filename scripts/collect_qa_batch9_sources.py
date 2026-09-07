import csv, json, importlib.util
from pathlib import Path
BASE=Path(__file__).with_name('collect_qa_batch2_sources.py')
spec=importlib.util.spec_from_file_location('collector',BASE); m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
ROOT=m.ROOT; SHOP=m.SHOP; RUN_ID=m.RUN_ID; QA_RUN_ID='20260907_180500'
QA_DIR=ROOT/'seo_runs'/SHOP/RUN_ID/'qa'/QA_RUN_ID; WORKBOOK=ROOT/'resutls'/SHOP/RUN_ID/'batches'/'SEO_Product_Optimization_through_batch_009.xlsx'
def main():
 QA_DIR.mkdir(parents=True,exist_ok=True); snap=QA_DIR/'source_snapshot'; snap.mkdir(exist_ok=True); (snap/WORKBOOK.name).write_bytes(WORKBOOK.read_bytes())
 with m.INVENTORY.open(newline='',encoding='utf-8-sig') as f: batch=list(csv.DictReader(f))[80:90]
 keys={r['product_key'] for r in batch}; handles={r['Handle'] for r in batch}; urls={r['product_url'] for r in batch}
 wb=m.load_workbook(WORKBOOK,read_only=True,data_only=True)
 submitted={'products':[r for r in m.rows(wb,'SEO_Products') if r['product_key'] in keys],'images':[r for r in m.rows(wb,'Image_Audit') if r['Handle'] in handles],'product_evidence':[r for r in m.rows(wb,'Product_Evidence') if r['product_url'] in urls],'keyword_map':[r for r in m.rows(wb,'Keyword_Map') if r['product_key'] in keys],'buyer_search_research':[r for r in m.rows(wb,'Buyer_Search_Research') if r['product_key'] in keys]}
 (QA_DIR/'submitted_batch_data.json').write_text(json.dumps(submitted,ensure_ascii=False,indent=2),encoding='utf-8')
 live=[]
 for row in batch:
  rec={'inventory':row}
  try: rec['live'],rec['fetch_error']=m.fetch_live(row['product_url']),''
  except Exception as e: rec['live'],rec['fetch_error']={},str(e)
  ms=list(m.EVIDENCE_PRODUCTS.glob(f"{row['inventory_position']}_*/evidence_summary.json")); rec['research_snapshot']=json.loads(ms[0].read_text(encoding='utf-8')) if ms else {}
  live.append(rec)
 (QA_DIR/'live_source_comparison.json').write_text(json.dumps(live,ensure_ascii=False,indent=2),encoding='utf-8')
 print('products',len(submitted['products']),'images',len(submitted['images']),'live',len(live))
if __name__=='__main__': main()
