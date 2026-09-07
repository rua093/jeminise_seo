import importlib.util
from pathlib import Path
spec=importlib.util.spec_from_file_location('c',Path(__file__).with_name('inspect_batch8_customizer.py')); m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
m.QA=m.ROOT/'seo_runs/jeminise.com/20260906_234129/qa/20260907_180500'
if __name__=='__main__':
 import concurrent.futures,html,json,re,requests
 rows=json.loads((m.QA/'live_source_comparison.json').read_text(encoding='utf-8')); urls=[x['inventory']['product_url'] for x in rows]
 with concurrent.futures.ThreadPoolExecutor(max_workers=5) as p: pages=list(p.map(lambda u:requests.get(u,timeout=30).text,urls))
 out=[]
 for pos,(u,page) in enumerate(zip(urls,pages),81):
  mm=re.search(r'data-amzcustom-root[^>]*data-config="([^"]+)"',page); cfg=json.loads(html.unescape(mm.group(1))) if mm else {}; txt=json.dumps(cfg,ensure_ascii=False)
  out.append({'inventory_position':pos,'url':u,'customizer_root_present':bool(mm),'config_bytes':len(txt.encode()),'surface_count':len(cfg.get('surfaces',[])),'name_mentions':len(re.findall('name',txt,re.I)),'number_mentions':len(re.findall('number',txt,re.I)),'text_or_input_mentions':len(re.findall('text|input',txt,re.I)),'upload_endpoint_present':'/apps/amazon-customizer/upload' in page,'customize_button_present':'amzcustom-open' in page,'config_excerpt':txt[:2000]})
 (m.QA/'customizer_audit.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8'); print(json.dumps(out,ensure_ascii=False,indent=2))
