from pathlib import Path
import html, json, re, requests
ROOT=Path(__file__).resolve().parents[1]
QA_DIR=ROOT/"seo_runs/jeminise.com/20260906_234129/qa/20260907_161639"
def walk(node,path="root"):
    if isinstance(node,dict):
        interesting={k:v for k,v in node.items() if k in {"id","label","name","type","required","minLength","maxLength","placeholder","instructions","accept","componentType","kind"}}
        if interesting and re.search(r"photo|image|upload|name|number|text|date|verse",json.dumps(node,ensure_ascii=False),re.I): yield {"path":path,"fields":interesting}
        for k,v in node.items(): yield from walk(v,f"{path}.{k}")
    elif isinstance(node,list):
        for i,v in enumerate(node): yield from walk(v,f"{path}[{i}]")
rows=json.loads((QA_DIR/"live_source_comparison.json").read_text(encoding="utf-8")); output=[]
for entry in rows:
    page=requests.get(entry["inventory"]["product_url"],timeout=30).text
    match=re.search(r'data-amzcustom-root[^>]*data-config="([^"]+)"',page); config=json.loads(html.unescape(match.group(1))) if match else {}
    output.append({"inventory_position":int(entry["inventory"]["inventory_position"]),"url":entry["inventory"]["product_url"],"collections":{k:len(v) for k,v in config.items() if isinstance(v,list)},"fields":list(walk(config))})
(QA_DIR/"customizer_fields_audit.json").write_text(json.dumps(output,ensure_ascii=False,indent=2),encoding="utf-8")
print(json.dumps(output,ensure_ascii=False,indent=2))
