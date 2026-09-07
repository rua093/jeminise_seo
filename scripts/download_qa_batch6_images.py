import hashlib
import json
from datetime import datetime
from pathlib import Path

import requests


ROOT = Path(__file__).resolve().parents[1]
QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / "20260907_102634"
IMAGE_DIR = QA_DIR / "images"


def main():
    live = json.loads((QA_DIR / "live_source_comparison.json").read_text(encoding="utf-8"))
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    rows, completed = [], []
    for entry in live:
        pos = int(entry["inventory"]["inventory_position"])
        product_key = entry["inventory"]["product_key"]
        for index, media in enumerate(entry["live"]["product_js"]["media"], 1):
            url = media["src"]
            suffix = Path(url.split("?", 1)[0]).suffix or ".jpg"
            target = IMAGE_DIR / f"{pos:03d}_{index:02d}{suffix}"
            response = requests.get(url, timeout=60, headers={"User-Agent":"Mozilla/5.0 independent SEO QA/1.0"})
            response.raise_for_status()
            target.write_bytes(response.content)
            qa_key = hashlib.sha256(f"{product_key}|{media.get('id')}|GALLERY|{index}".encode()).hexdigest()[:24]
            completed.append(qa_key)
            rows.append({"inventory_position":pos, "product_key":product_key, "image_number":index, "media_id":str(media.get("id", "")), "source_url":url, "local_path":str(target), "sha256":hashlib.sha256(response.content).hexdigest(), "bytes":len(response.content), "downloaded_at":datetime.now().astimezone().isoformat()})
            progress = {"rubric_version":"1.0", "qa_run_id":"20260907_102634", "source_workbook":str(QA_DIR / "source_snapshot" / "SEO_Product_Optimization_through_batch_034.xlsx"), "batch_id":"qa_batch_006", "batch_product_keys":[item["inventory"]["product_key"] for item in live], "current_product_key":product_key, "current_stage":"IMAGE_DOWNLOADED_PENDING_DIRECT_REVIEW", "completed_image_keys":completed, "last_saved_at":datetime.now().astimezone().isoformat(), "awaiting_confirmation":False}
            (QA_DIR / "qa_progress.json").write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")
    (QA_DIR / "image_download_manifest.json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"downloaded={len(rows)} bytes={sum(item['bytes'] for item in rows)}")


if __name__ == "__main__":
    main()
