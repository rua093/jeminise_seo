import concurrent.futures,html,json,re
from pathlib import Path
import requests
ROOT=Path(__file__).resolve().parents[1]; QA=ROOT/"seo_runs/jeminise.com/20260906_234129/qa/20260907_173014"
def main():
 rows=json.loads((QA/"live_source_comparison.json").read_text(encoding="utf-8")); urls=[x["inventory"]["product_url"] for x in rows]
 with concurrent.futures.ThreadPoolExecutor(max_workers=5) as p: pages=list(p.map(lambda u:requests.get(u,timeout=30).text,urls))
 out=[]
 for pos,(u,page) in enumerate(zip(urls,pages),71):
  m=re.search(r'data-amzcustom-root[^>]*data-config="([^"]+)"',page); cfg=json.loads(html.unescape(m.group(1))) if m else {}; txt=json.dumps(cfg,ensure_ascii=False)
  out.append({"inventory_position":pos,"url":u,"customizer_root_present":bool(m),"config_bytes":len(txt.encode()),"surface_count":len(cfg.get("surfaces",[])),"name_mentions":len(re.findall("name",txt,re.I)),"number_mentions":len(re.findall("number",txt,re.I)),"text_or_input_mentions":len(re.findall("text|input",txt,re.I)),"upload_endpoint_present":"/apps/amazon-customizer/upload" in page,"customize_button_present":"amzcustom-open" in page,"config_excerpt":txt[:2000]})
 (QA/"customizer_audit.json").write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding="utf-8"); print(json.dumps(out,ensure_ascii=False,indent=2))
if __name__ == "__main__": main()
