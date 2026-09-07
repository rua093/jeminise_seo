import json
import csv
from pathlib import Path

p = Path("seo_runs/jeminise.com/20260906_234129/qa/20260907_161639/submitted_batch_data.json")
data = json.loads(p.read_text(encoding="utf-8"))
with Path("seo_runs/jeminise.com/20260906_234129/inventory.csv").open(encoding="utf-8-sig", newline="") as fh:
    inv = list(csv.DictReader(fh))
pos_by_handle = {row.get("handle") or row.get("Handle"): int(row.get("inventory_position") or row.get("position")) for row in inv}

for row in data["images"]:
    pos = pos_by_handle[row["Handle"]]
    idx = int(row["image_number"])
    print(f"{pos:03d}_{idx:02d}\tOBS={row.get('observed_visual_details','')}\tALT={row.get('alt_proposed','')}")
