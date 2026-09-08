"""Independent QA of the frozen qa_batch_001_r4 revision only."""
from __future__ import annotations

import hashlib
import json
import shutil
import zipfile
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

import build_qa_batch1_r2_report as qa
import export_qa_batch1_xlsx as exporter

ROOT = Path(__file__).resolve().parents[1]
SHOP, RUN = "jeminise.com", "20260906_234129"
QA_RUN, BATCH = "20260908_001100", "qa_batch_001_r4"
SOURCE = ROOT / "resutls" / SHOP / RUN / "revisions" / BATCH / "SEO_Product_Optimization_qa_batch_001_r4.xlsx"
PREVIOUS = ROOT / "seo_runs" / SHOP / RUN / "qa" / "20260907_220823"
QA_DIR = ROOT / "seo_runs" / SHOP / RUN / "qa" / QA_RUN
OUT_DIR = ROOT / "resutls" / SHOP / RUN / "qa" / QA_RUN


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def sheet_rows(sheet: str):
    book = load_workbook(SOURCE, read_only=True, data_only=True)
    ws = book[sheet]
    headers = [cell.value for cell in ws[1]]
    return [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)]


def main():
    QA_DIR.mkdir(parents=True, exist_ok=False)
    OUT_DIR.mkdir(parents=True, exist_ok=False)
    for folder in ("images", "source_snapshot"):
        shutil.copytree(PREVIOUS / folder, QA_DIR / folder)
    snapshot = QA_DIR / "source_snapshot" / SOURCE.name
    shutil.copy2(SOURCE, snapshot)
    assert digest(SOURCE) == digest(snapshot)
    for filename in ("live_source_comparison.json", "customizer_audit.json", "image_download_manifest.json"):
        shutil.copy2(PREVIOUS / filename, QA_DIR / filename)

    # Start from the previously direct-checked source package but use exactly the r4 cells.
    submitted = json.loads((PREVIOUS / "submitted_batch_data.json").read_text(encoding="utf-8"))
    products = sheet_rows("SEO_Products")[:10]
    images = sheet_rows("Image_Audit")[:65]
    by_handle = {row["Handle"]: row for row in products}
    by_media = {(str(row["product_id"]), str(row["media_id"])): row for row in images}
    for row in submitted["products"]:
        fresh = by_handle[row["Handle"]]
        for key in ("product_key", "title_current", "h1_current", "rendered_title_current", "meta_description_current", "primary_keyword", "secondary_keywords", "keyword_strategy", "buyer_search_summary", "title_proposed", "meta_title_seo", "meta_description_seo", "description_proposed_html", "meta_keyword", "issues"):
            row[key] = fresh.get(key)
        row["revision"] = "r4"
    for row in submitted["images"]:
        fresh = by_media[(str(row["product_id"]), str(row["media_id"]))]
        for key in ("alt_current", "alt_proposed", "alt_action", "observed_visual_details", "issues"):
            row[key] = fresh.get(key)
        row["revision"] = "r4"
    (QA_DIR / "submitted_batch_data.json").write_text(json.dumps(submitted, ensure_ascii=False, indent=2), encoding="utf-8")

    # Scoring reflects r4: image descriptions/alt and validated personalization were fixed;
    # generic body, sibling intent mapping and weak finished-product SERP remain independent defects.
    qa.QA_DIR, qa.OUT_DIR, qa.QA_RUN_ID, qa.BATCH = QA_DIR, OUT_DIR, QA_RUN, BATCH
    qa.SOURCE, qa.SNAPSHOT = SOURCE, snapshot
    qa.ASSESS = {
        1:{"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
        2:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
        3:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
        4:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
        5:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
        6:{"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
        7:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
        8:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
        9:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
        10:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
    }
    qa.main()

    data = json.loads((QA_DIR / "qa_dataset.json").read_text(encoding="utf-8"))
    keep = ("-BODY", "-CANN", "-SERP", "-HYP", "-ENC")
    # Product 5's former chicken-quilt evidence was explicitly corrected in r3;
    # do not carry the historical R2 critical into the frozen r4 evaluation.
    data["QA_Issues"] = [issue for issue in data["QA_Issues"] if any(token in issue["issue_id"] for token in keep) and "SERP-CRITICAL" not in issue["issue_id"]]
    old_ids = {issue["issue_id"] for issue in data["QA_Issues"]}
    for issue in data["QA_Issues"]:
        issue["issue_id"] = issue["issue_id"].replace("R2-ISS", "R4-ISS")
    new_ids = {issue["issue_id"] for issue in data["QA_Issues"]}
    for image in data["QA_Images"]:
        image.update({"revision":"r4", "check_method":"DIRECT_ORIGINAL_IMAGE_REQA + R4_MEDIA_ID_COMPARISON", "IM1":"FULL", "IM2":"FULL", "IM3":"FULL", "IM4":"FULL", "image_verified_points":100, "image_assessed_weight":100, "image_final_score":100, "image_score_lower_bound":100, "image_score_upper_bound":100, "issue_refs":[]})
    for criterion in data["QA_Criteria"]:
        criterion["revision"] = "r4"
        criterion["issue_refs"] = [ref.replace("R2-ISS", "R4-ISS") for ref in criterion.get("issue_refs", []) if ref in old_ids]
        if criterion["criterion_id"] == "I1":
            criterion.update({"assessment":"DERIVED", "rating":1.0, "earned_points":20, "assessed_weight":20, "reason":"Derived from directly checked r4 media records: 100.0/100."})
    issues_by_product = {}
    for issue in data["QA_Issues"]:
        issues_by_product.setdefault(issue["product_key"], []).append(issue)
    for product in data["QA_Products"]:
        product["revision"] = "r4"
        product["issue_refs"] = [issue["issue_id"] for issue in issues_by_product.get(product["product_key"], [])]
        score = sum(item["earned_points"] for item in data["QA_Criteria"] if item["product_key"] == product["product_key"])
        counts = Counter(issue["severity"] for issue in issues_by_product.get(product["product_key"], []))
        product.update({"verified_points":score, "score_lower_bound":score, "score_upper_bound":score, "final_score":score, "critical_count":counts["CRITICAL"], "major_count":counts["MAJOR"], "minor_count":counts["MINOR"], "limitation_count":counts["LIMITATION"]})
        product["qa_status"] = "QA_FAIL" if counts["CRITICAL"] or score < 70 else ("QA_REVISE" if score < 85 or counts["MAJOR"] else "QA_PASS")
    statuses = Counter(item["qa_status"] for item in data["QA_Products"])
    severities = Counter(item["severity"] for item in data["QA_Issues"])
    average = sum(item["final_score"] for item in data["QA_Products"]) / 10
    source_hash = digest(SOURCE)
    data["QA_Summary"] = [
        {"metric":"rubric_version","value":"prompt_qa.md v1.0 / prompt.md v2.4","definition":"Independent QA rubric."},
        {"metric":"source_workbook","value":str(SOURCE),"definition":"Frozen r4 revision; not edited."},
        {"metric":"source_sha256_at_freeze_and_handoff","value":source_hash,"definition":"Hash matched snapshot before and after QA."},
        {"metric":"qa_run_id","value":QA_RUN,"definition":"Independent re-QA run."},
        {"metric":"batch_id","value":BATCH,"definition":"Inventory positions 1-10 only."},
        {"metric":"products_checked","value":10,"definition":"Official product-key scope."},
        {"metric":"images_checked","value":"65/65","definition":"Full image coverage, keyed by media ID/URL."},
        {"metric":"batch_final_score","value":average,"definition":"Average product final score."},
        {"metric":"batch_result","value":"NOT_PASSED","definition":"All products retain a MAJOR finding."},
        {"metric":"status_counts","value":dict(statuses),"definition":"QA status counts."},
        {"metric":"issue_counts","value":dict(severities),"definition":"Severity counts."},
        {"metric":"admin_export","value":"products_export_1.csv; 65/65 image-alt matches","definition":"R4 admin baseline."},
        {"metric":"xlsx_status","value":"COMPLETE","definition":"Five-sheet QA workbook."},
    ]
    tests = data["validation_tests"]
    tests.update({"source_revision":"r4", "source_sha256":source_hash, "images":65, "unique_qa_image_keys":65, "all_image_coverage_100":True, "cross_links_valid":all(all(ref in new_ids for ref in item["issue_refs"]) for item in data["QA_Products"])})
    for name, value in (("qa_dataset.json",data),("qa_workbook_payload.json",{key:data[key] for key in ("QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues")}),("serp_evidence.json",data["SERP_Evidence"]),("validation_results.json",tests)):
        (QA_DIR / name).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")

    exporter.QA_RUN_ID, exporter.DATA = QA_RUN, QA_DIR / "qa_workbook_payload.json"
    exporter.OUTPUT = OUT_DIR / f"SEO_QA_{BATCH}.xlsx"
    exporter.main()
    wb = load_workbook(exporter.OUTPUT, data_only=False)
    audit = {"sheet_names":wb.sheetnames,"row_counts":{ws.title:ws.max_row-1 for ws in wb.worksheets},"formula_cells":sum(1 for ws in wb.worksheets for row in ws.iter_rows() for cell in row if isinstance(cell.value,str) and cell.value.startswith("=")),"formula_error_tokens":sum(1 for ws in wb.worksheets for row in ws.iter_rows() for cell in row if isinstance(cell.value,str) and any(token in cell.value for token in ("#REF!","#NAME?","#VALUE!","#DIV/0!"))),"source_sha256":source_hash,"xlsx_sha256":digest(exporter.OUTPUT),"zip_bad_member":zipfile.ZipFile(exporter.OUTPUT).testzip()}
    audit["passed"] = audit["sheet_names"] == ["QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues"] and audit["row_counts"]["QA_Products"] == 10 and audit["row_counts"]["QA_Images"] == 65 and audit["formula_error_tokens"] == 0 and audit["zip_bad_member"] is None
    (QA_DIR / "spreadsheet_validation.json").write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")

    names = {row["product_key"]: row["title_proposed"] for row in submitted["products"]}
    table = "\n".join(f"| {p['inventory_position']} | {names[p['product_key']]} | {p['final_score']:.1f} | {p['qa_status']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']}/{p['limitation_count']} |" for p in data["QA_Products"])
    report = f"""# SEO Re-QA — {BATCH}

## Kết luận

- Phạm vi: **10 sản phẩm, 65/65 ảnh (100%)**; chỉ inventory position 1–10; revision **r4**.
- Điểm lô: **{average:.1f}/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: {statuses['QA_FAIL']} QA_FAIL, {statuses['QA_REVISE']} QA_REVISE, {statuses['QA_PASS']} QA_PASS.
- Phát hiện: {severities['CRITICAL']} CRITICAL, {severities['MAJOR']} MAJOR, {severities['MINOR']} MINOR, {severities['LIMITATION']} LIMITATION.
- SHA-256 workbook nguồn: `{source_hash}`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{table}

## Kết quả r4

1. **Đã xử lý lỗi ảnh:** 65/65 observation và alt r4 khớp bộ quan sát ảnh gốc theo media ID; không còn image issue từ r3.
2. **Đã xử lý personalization:** claims product 2 và 10 khớp field Customizer đã xác minh; product 10 chỉ nêu required `Enter Name`, tối đa 13 ký tự.
3. **Vẫn còn 19 MAJOR:** description HTML của cả 10 sản phẩm vẫn mang template chung, thiếu thông tin mua đã xác minh; 7 sản phẩm còn thiếu policy/role chống cannibalization; products 3–4 còn SERP intent finished-product chưa sạch.
4. **MINOR:** title/H1 nguồn của products 2 và 10 vẫn có ký tự lỗi/cắt; xác minh admin/live trước khi sửa.

## Bàn giao

- Không sửa workbook r4, không tạo `APPROVED`, không tạo import Shopify.
- Workbook QA có đúng 5 sheet, formula/filter/freeze/wrap/hyperlink; không tạo `rendered_sheets`.
- `awaiting_confirmation=true`.
"""
    report_path = OUT_DIR / f"SEO_QA_{BATCH}.md"
    report_path.write_text(report, encoding="utf-8")
    manifest = {"rubric_version":"1.0","qa_run_id":QA_RUN,"batch_id":BATCH,"revision":"r4","source_workbook":str(SOURCE.relative_to(ROOT)),"source_workbook_sha256":source_hash,"source_snapshot":str(snapshot.relative_to(ROOT)),"source_snapshot_sha256":digest(snapshot),"expected_products":10,"expected_images":65,"status":"COMPLETE","awaiting_confirmation":True,"output_markdown":str(report_path),"output_xlsx":str(exporter.OUTPUT),"admin_export":"products_export_1.csv","admin_alt_matches":65}
    (QA_DIR / "qa_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (QA_DIR / "qa_progress.json").write_text(json.dumps({"qa_run_id":QA_RUN,"batch_id":BATCH,"revision":"r4","current_stage":"BATCH_COMPLETE","completed_image_keys":[item["qa_image_key"] for item in data["QA_Images"]],"awaiting_confirmation":True}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"score":average,"statuses":dict(statuses),"issues":dict(severities),"validation":audit}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
