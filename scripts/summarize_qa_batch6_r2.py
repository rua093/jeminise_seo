from pathlib import Path
import json


q = Path(__file__).resolve().parents[1] / "seo_runs/jeminise.com/20260906_234129/qa/20260907_161639"
d = json.loads((q / "submitted_batch_data.json").read_text(encoding="utf-8"))
for i, p in enumerate(d["products"], 51):
    print("\n", i, p.get("title_proposed"), "\nPRIMARY:", p.get("primary_keyword"), "\nMETA:", p.get("meta_description_seo"), "\nBODY:", p.get("description_proposed_html"))
