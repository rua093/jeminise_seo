from pathlib import Path
import csv
import hashlib
import json


ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_161639"
QA = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
OUT = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "qa" / RUN
SOURCE = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "revisions" / "qa_batch_006_r2" / "SEO_Product_Optimization_qa_batch_006_r2.xlsx"
SNAPSHOT = QA / "source_snapshot" / SOURCE.name


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


dataset = json.loads((QA / "qa_dataset.json").read_text(encoding="utf-8"))
sheet = json.loads((QA / "spreadsheet_validation.json").read_text(encoding="utf-8"))
manifest_path = QA / "qa_manifest.json"
manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
progress_path = QA / "qa_progress.json"
progress = json.loads(progress_path.read_text(encoding="utf-8"))

with (ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "inventory.csv").open(encoding="utf-8-sig", newline="") as fh:
    inventory = list(csv.DictReader(fh))
expected_keys = [r["product_key"] for r in inventory if 51 <= int(r["inventory_position"]) <= 60]

products = dataset["QA_Products"]
images = dataset["QA_Images"]
issues = dataset["QA_Issues"]
criteria = dataset["QA_Criteria"]
issue_ids = {x["issue_id"] for x in issues}
rendered = sorted((QA / "rendered_sheets").glob("*.png"))

tests = {
    "source_hash_matches_snapshot": sha(SOURCE) == sha(SNAPSHOT),
    "source_hash_matches_manifest": sha(SOURCE) == manifest["source_workbook_sha256"],
    "exact_product_keys_51_60": [x["product_key"] for x in products] == expected_keys,
    "product_count_10": len(products) == 10,
    "image_count_78": len(images) == 78,
    "criteria_count_110": len(criteria) == 110,
    "unique_product_keys": len({x["product_key"] for x in products}) == 10,
    "unique_image_keys": len({x["qa_image_key"] for x in images}) == 78,
    "unique_issue_ids": len(issue_ids) == len(issues),
    "product_issue_links_valid": all(all(ref in issue_ids for ref in x["issue_refs"]) for x in products),
    "image_issue_links_valid": all(all(ref in issue_ids for ref in x["issue_refs"]) for x in images),
    "criteria_product_links_valid": all(x["product_key"] in expected_keys for x in criteria),
    "all_assessed_weight_100": all(x["assessed_weight"] == 100 for x in products) and all(x["image_assessed_weight"] == 100 for x in images),
    "all_image_coverage_100": all(x["images_expected"] == x["images_checked"] and x["image_coverage"] == 1 for x in products),
    "spreadsheet_validation_passed": bool(sheet["passed"]),
    "rendered_exact_five_sheets": len(rendered) == 5 and {p.stem for p in rendered} == {"QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues"},
    "logic_100_with_critical": dataset["validation_tests"]["logic_100_with_critical"]["passed"],
    "logic_90_clean": dataset["validation_tests"]["logic_90_full_no_blocker"]["passed"],
    "logic_72_on_80": dataset["validation_tests"]["logic_72_on_80"]["passed"],
}
tests["passed"] = all(tests.values())
assert tests["passed"], tests

(QA / "validation_results.json").write_text(json.dumps(tests, ensure_ascii=False, indent=2), encoding="utf-8")
manifest.update({
    "status": "COMPLETE",
    "source_sha256_at_handoff": sha(SOURCE),
    "source_snapshot_sha256": sha(SNAPSHOT),
    "xlsx_sha256": sha(OUT / "SEO_QA_qa_batch_006_r2.xlsx"),
    "validation_passed": True,
    "rendered_sheets": [str(p) for p in rendered],
})
manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
progress.update({"current_stage": "BATCH_COMPLETE", "awaiting_confirmation": True, "last_saved_at": manifest["completed_at"]})
progress_path.write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"tests": tests, "source_sha256": sha(SOURCE), "xlsx_sha256": manifest["xlsx_sha256"]}, ensure_ascii=False, indent=2))
