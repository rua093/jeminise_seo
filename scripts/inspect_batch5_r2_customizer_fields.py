from pathlib import Path
import html
import json
import re
import requests


ROOT = Path(__file__).resolve().parents[1]
QA_DIR = ROOT / "seo_runs/jeminise.com/20260906_234129/qa/20260907_152903"


def walk(node, path="root"):
    if isinstance(node, dict):
        interesting = {k:v for k,v in node.items() if k in {"id","label","name","type","required","minLength","maxLength","placeholder","instructions","accept","componentType","kind"}}
        rendered = json.dumps(node, ensure_ascii=False)
        if interesting and re.search(r"photo|image|upload|name|number|text", rendered, re.I):
            yield {"path":path,"fields":interesting}
        for key,value in node.items():
            yield from walk(value, f"{path}.{key}")
    elif isinstance(node, list):
        for i,value in enumerate(node):
            yield from walk(value, f"{path}[{i}]")


rows = json.loads((QA_DIR / "live_source_comparison.json").read_text(encoding="utf-8"))
output=[]
for entry in rows:
    pos=int(entry["inventory"]["inventory_position"])
    page=requests.get(entry["inventory"]["product_url"],timeout=30).text
    match=re.search(r'data-amzcustom-root[^>]*data-config="([^"]+)"',page)
    config=json.loads(html.unescape(match.group(1))) if match else {}
    found=list(walk(config))
    collections = {key:len(value) for key,value in config.items() if isinstance(value,list)}
    output.append({"inventory_position":pos,"url":entry["inventory"]["product_url"],"collections":collections,"fields":found})
(QA_DIR / "customizer_fields_audit.json").write_text(json.dumps(output,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps(output,ensure_ascii=False,indent=2))
