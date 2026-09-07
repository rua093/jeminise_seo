import hashlib,json
from datetime import datetime
from pathlib import Path
import requests

ROOT=Path(__file__).resolve().parents[1]; QA=ROOT/"seo_runs/jeminise.com/20260906_234129/qa/20260907_173014"; IMAGE_DIR=QA/"images"
def main():
    live=json.loads((QA/"live_source_comparison.json").read_text(encoding="utf-8")); IMAGE_DIR.mkdir(parents=True,exist_ok=True); rows=[]
    for e in live:
        pos=int(e["inventory"]["inventory_position"]); pk=e["inventory"]["product_key"]
        for i,m in enumerate(e["live"]["product_js"]["media"],1):
            url=m["src"]; target=IMAGE_DIR/f"{pos:03d}_{i:02d}{Path(url.split('?',1)[0]).suffix or '.jpg'}"; resp=requests.get(url,timeout=60,headers={"User-Agent":"Mozilla/5.0 independent SEO QA/1.0"}); resp.raise_for_status(); b=resp.content; target.write_bytes(b)
            rows.append({"inventory_position":pos,"product_key":pk,"image_number":i,"media_id":str(m.get("id","")),"source_url":url,"local_path":str(target),"sha256":hashlib.sha256(b).hexdigest(),"bytes":len(b),"downloaded_at":datetime.now().astimezone().isoformat()})
    (QA/"image_download_manifest.json").write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding="utf-8"); print(f"downloaded={len(rows)} bytes={sum(x['bytes'] for x in rows)}")
if __name__ == "__main__": main()
