import csv
import importlib.util
import json
from pathlib import Path


spec = importlib.util.spec_from_file_location("collector17", Path(__file__).with_name("collect_qa_batch17_sources.py"))
collector = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(collector)

ROOT = collector.ROOT
BASE = ROOT / "resutls/jeminise.com/20260906_234129/batches"
CONFIG = {
    18: ("20260907_213100", 170, "SEO_Product_Optimization_through_batch_018.xlsx"),
    19: ("20260907_213200", 180, "SEO_Product_Optimization_through_batch_019.xlsx"),
}


def main():
    inventory = list(csv.DictReader((ROOT / "seo_runs/jeminise.com/20260906_234129/inventory.csv").open(encoding="utf-8-sig")))
    for batch, (run_id, zero_index, source_name) in CONFIG.items():
        qa = ROOT / "seo_runs/jeminise.com/20260906_234129/qa" / run_id
        workbook = BASE / source_name
        qa.mkdir(parents=True, exist_ok=True)
        snapshot = qa / "source_snapshot"
        snapshot.mkdir(exist_ok=True)
        frozen = snapshot / workbook.name
        frozen.write_bytes(workbook.read_bytes())
        rows = inventory[zero_index:zero_index + 10]
        keys, handles, urls = ({r["product_key"] for r in rows}, {r["Handle"] for r in rows}, {r["product_url"] for r in rows})
        book = collector.module.m.load_workbook(workbook, read_only=True, data_only=True)
        submitted = {
            "products": [r for r in collector.module.m.rows(book, "SEO_Products") if r["product_key"] in keys],
            "images": [r for r in collector.module.m.rows(book, "Image_Audit") if r["Handle"] in handles],
            "product_evidence": [r for r in collector.module.m.rows(book, "Product_Evidence") if r["product_url"] in urls],
            "keyword_map": [r for r in collector.module.m.rows(book, "Keyword_Map") if r["product_key"] in keys],
            "buyer_search_research": [r for r in collector.module.m.rows(book, "Buyer_Search_Research") if r["product_key"] in keys],
        }
        (qa / "submitted_batch_data.json").write_text(json.dumps(submitted, ensure_ascii=False, indent=2), encoding="utf-8")
        live = []
        for row in rows:
            try:
                live.append({"inventory": row, "live": collector.module.m.fetch_live(row["product_url"]), "fetch_error": ""})
            except Exception as exc:
                live.append({"inventory": row, "live": {}, "fetch_error": str(exc)})
        (qa / "live_source_comparison.json").write_text(json.dumps(live, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"batch={batch} products={len(submitted['products'])} images={len(submitted['images'])} live={len(live)}")


if __name__ == "__main__":
    main()
