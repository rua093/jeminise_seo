import csv
import importlib.util
import json
from pathlib import Path


BASE = Path(__file__).with_name("collect_qa_batch2_sources.py")
spec = importlib.util.spec_from_file_location("qa_batch2_collector", BASE)
module = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(module)

module.QA_RUN_ID = "20260907_101050"
module.QA_DIR = module.ROOT / "seo_runs" / module.SHOP / module.RUN_ID / "qa" / module.QA_RUN_ID
module.WORKBOOK = module.QA_DIR / "source_snapshot" / "SEO_Product_Optimization_through_batch_034.xlsx"


def main():
    with module.INVENTORY.open(newline="", encoding="utf-8-sig") as handle:
        batch = list(csv.DictReader(handle))[40:50]
    keys = {r["product_key"] for r in batch}
    handles = {r["Handle"] for r in batch}
    urls = {r["product_url"] for r in batch}
    workbook = module.load_workbook(module.WORKBOOK, read_only=True, data_only=True)
    submitted = {
        "products": [r for r in module.rows(workbook, "SEO_Products") if r["product_key"] in keys],
        "images": [r for r in module.rows(workbook, "Image_Audit") if r["Handle"] in handles],
        "product_evidence": [r for r in module.rows(workbook, "Product_Evidence") if r["product_url"] in urls],
        "keyword_map": [r for r in module.rows(workbook, "Keyword_Map") if r["product_key"] in keys],
        "buyer_search_research": [r for r in module.rows(workbook, "Buyer_Search_Research") if r["product_key"] in keys],
    }
    (module.QA_DIR / "submitted_batch_data.json").write_text(json.dumps(submitted, ensure_ascii=False, indent=2), encoding="utf-8")
    live = []
    for row in batch:
        record = {"inventory": row}
        try:
            record["live"], record["fetch_error"] = module.fetch_live(row["product_url"]), ""
        except Exception as exc:
            record["live"], record["fetch_error"] = {}, str(exc)
        matches = list(module.EVIDENCE_PRODUCTS.glob(f"{row['inventory_position']}_*/evidence_summary.json"))
        if matches:
            record["research_snapshot"] = json.loads(matches[0].read_text(encoding="utf-8"))
        live.append(record)
        checkpoint = {"rubric_version":"1.0","qa_run_id":module.QA_RUN_ID,"source_workbook":str(module.WORKBOOK),"batch_id":"qa_batch_005","batch_product_keys":[x["product_key"] for x in batch],"current_product_key":row["product_key"],"current_stage":"LIVE_SOURCE_COLLECTED","completed_image_keys":[],"awaiting_confirmation":False}
        (module.QA_DIR / "qa_progress.json").write_text(json.dumps(checkpoint, ensure_ascii=False, indent=2), encoding="utf-8")
    (module.QA_DIR / "live_source_comparison.json").write_text(json.dumps(live, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"products={len(submitted['products'])} images={len(submitted['images'])} keywords={len(submitted['keyword_map'])} buyers={len(submitted['buyer_search_research'])} live={len(live)}")
    for item in live:
        print(item["inventory"]["inventory_position"], item["live"].get("status_code"), len(item["live"].get("product_js", {}).get("media", [])), item["fetch_error"])


if __name__ == "__main__":
    main()
