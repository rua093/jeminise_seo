import importlib.util
import json
from pathlib import Path

spec = importlib.util.spec_from_file_location("batch14_collector", Path(__file__).with_name("collect_qa_batch14_sources.py"))
module = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(module)

ROOT = module.ROOT
RUN_ID = "20260907_212312"
QA = ROOT / "seo_runs/jeminise.com/20260906_234129/qa" / RUN_ID
WORKBOOK = ROOT / "resutls/jeminise.com/20260906_234129/batches/SEO_Product_Optimization_through_batch_017.xlsx"


def main():
    QA.mkdir(parents=True, exist_ok=True)
    snapshot = QA / "source_snapshot"
    snapshot.mkdir(exist_ok=True)
    frozen = snapshot / WORKBOOK.name
    frozen.write_bytes(WORKBOOK.read_bytes())
    rows = list(module.csv.DictReader((ROOT / "seo_runs/jeminise.com/20260906_234129/inventory.csv").open(encoding="utf-8-sig")))[160:170]
    keys = {row["product_key"] for row in rows}
    handles = {row["Handle"] for row in rows}
    urls = {row["product_url"] for row in rows}
    book = module.m.load_workbook(WORKBOOK, read_only=True, data_only=True)
    submitted = {
        "products": [row for row in module.m.rows(book, "SEO_Products") if row["product_key"] in keys],
        "images": [row for row in module.m.rows(book, "Image_Audit") if row["Handle"] in handles],
        "product_evidence": [row for row in module.m.rows(book, "Product_Evidence") if row["product_url"] in urls],
        "keyword_map": [row for row in module.m.rows(book, "Keyword_Map") if row["product_key"] in keys],
        "buyer_search_research": [row for row in module.m.rows(book, "Buyer_Search_Research") if row["product_key"] in keys],
    }
    (QA / "submitted_batch_data.json").write_text(json.dumps(submitted, ensure_ascii=False, indent=2), encoding="utf-8")
    live = []
    for row in rows:
        try:
            live.append({"inventory": row, "live": module.m.fetch_live(row["product_url"]), "fetch_error": ""})
        except Exception as exc:
            live.append({"inventory": row, "live": {}, "fetch_error": str(exc)})
    (QA / "live_source_comparison.json").write_text(json.dumps(live, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"products={len(submitted['products'])} images={len(submitted['images'])} live={len(live)}")


if __name__ == "__main__":
    main()
