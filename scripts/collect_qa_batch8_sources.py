import csv, importlib.util, json
from pathlib import Path

BASE = Path(__file__).with_name("collect_qa_batch2_sources.py")
spec = importlib.util.spec_from_file_location("qa_batch2_collector8", BASE)
module = importlib.util.module_from_spec(spec); assert spec.loader is not None; spec.loader.exec_module(module)
module.QA_RUN_ID = "20260907_173014"
module.QA_DIR = module.ROOT / "seo_runs" / module.SHOP / module.RUN_ID / "qa" / module.QA_RUN_ID
module.WORKBOOK = module.ROOT / "resutls" / module.SHOP / module.RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_008.xlsx"

def main():
    module.QA_DIR.mkdir(parents=True, exist_ok=True)
    with module.INVENTORY.open(newline="", encoding="utf-8-sig") as fh: batch = list(csv.DictReader(fh))[70:80]
    keys={r["product_key"] for r in batch}; handles={r["Handle"] for r in batch}; urls={r["product_url"] for r in batch}
    wb=module.load_workbook(module.WORKBOOK, read_only=True, data_only=True)
    submitted={"products":[r for r in module.rows(wb,"SEO_Products") if r["product_key"] in keys],"images":[r for r in module.rows(wb,"Image_Audit") if r["Handle"] in handles],"product_evidence":[r for r in module.rows(wb,"Product_Evidence") if r["product_url"] in urls],"keyword_map":[r for r in module.rows(wb,"Keyword_Map") if r["product_key"] in keys],"buyer_search_research":[r for r in module.rows(wb,"Buyer_Search_Research") if r["product_key"] in keys]}
    (module.QA_DIR/"submitted_batch_data.json").write_text(json.dumps(submitted,ensure_ascii=False,indent=2),encoding="utf-8")
    live=[]
    for row in batch:
        rec={"inventory":row}
        try: rec["live"],rec["fetch_error"]=module.fetch_live(row["product_url"]),""
        except Exception as exc: rec["live"],rec["fetch_error"]={},str(exc)
        matches=list(module.EVIDENCE_PRODUCTS.glob(f"{row['inventory_position']}_*/evidence_summary.json")); rec["research_snapshot"]=json.loads(matches[0].read_text(encoding="utf-8")) if matches else {}
        live.append(rec)
    (module.QA_DIR/"live_source_comparison.json").write_text(json.dumps(live,ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"products={len(submitted['products'])} images={len(submitted['images'])} keywords={len(submitted['keyword_map'])} buyers={len(submitted['buyer_search_research'])} live={len(live)}")

if __name__ == "__main__": main()
