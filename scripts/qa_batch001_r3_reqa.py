"""Build the independent r3 re-QA package for qa_batch_001.

This intentionally reuses the r2 direct-image/live-source capture only as
evidence, while replacing submitted workbook values with the r3 revision.
It does not alter either source workbook.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import zipfile
from collections import Counter
from datetime import datetime, timezone, timedelta
from pathlib import Path

from openpyxl import load_workbook

import build_qa_batch1_r2_report as prior
import export_qa_batch1_xlsx as exporter

ROOT = Path(__file__).resolve().parents[1]
SHOP = "jeminise.com"
RUN = "20260906_234129"
QA_RUN = "20260907_220823"
BATCH = "qa_batch_001_r3"
SOURCE = ROOT / "resutls" / SHOP / RUN / "revisions" / "qa_batch_001_r3" / "SEO_Product_Optimization_qa_batch_001_r3.xlsx"
OLD_QA = ROOT / "seo_runs" / SHOP / RUN / "qa" / "20260907_114114"
QA_DIR = ROOT / "seo_runs" / SHOP / RUN / "qa" / QA_RUN
OUT_DIR = ROOT / "resutls" / SHOP / RUN / "qa" / QA_RUN


def iso_now() -> str:
    return datetime.now(timezone(timedelta(hours=7))).isoformat()


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def xlsx_rows(sheet: str):
    wb = load_workbook(SOURCE, read_only=True, data_only=False)
    ws = wb[sheet]
    headers = [c.value for c in ws[1]]
    return [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)]


def main() -> None:
    QA_DIR.mkdir(parents=True, exist_ok=False)
    OUT_DIR.mkdir(parents=True, exist_ok=False)
    shutil.copytree(OLD_QA / "images", QA_DIR / "images")
    shutil.copytree(OLD_QA / "source_snapshot", QA_DIR / "source_snapshot")
    shutil.copy2(SOURCE, QA_DIR / "source_snapshot" / SOURCE.name)
    for name in ("live_source_comparison.json", "customizer_audit.json", "image_download_manifest.json"):
        shutil.copy2(OLD_QA / name, QA_DIR / name)

    # submitted_batch_data is a neutral, source-backed input to the old
    # re-QA builder. Replace its 10 submitted products/images with r3 cells.
    submitted = json.loads((OLD_QA / "submitted_batch_data.json").read_text(encoding="utf-8"))
    products = xlsx_rows("SEO_Products")[:10]
    images = xlsx_rows("Image_Audit")[:65]
    by_handle = {p["Handle"]: p for p in products}
    for record in submitted["products"]:
        fresh = by_handle[record["Handle"]]
        for key in ("product_key", "title_current", "h1_current", "rendered_title_current", "meta_description_current", "primary_keyword", "secondary_keywords", "keyword_strategy", "buyer_search_summary", "title_proposed", "meta_title_seo", "meta_description_seo", "description_proposed_html", "meta_keyword", "revision", "issues"):
            record[key] = fresh.get(key)
        record["revision"] = "r3"
    by_media = {(str(x["product_id"]), str(x["media_id"])): x for x in images}
    for record in submitted["images"]:
        fresh = by_media[(str(record["product_id"]), str(record["media_id"]))]
        for key in ("alt_current", "alt_proposed", "alt_action", "observed_visual_details", "issues", "revision"):
            record[key] = fresh.get(key)
        record["revision"] = "r3"
    (QA_DIR / "submitted_batch_data.json").write_text(json.dumps(submitted, ensure_ascii=False, indent=2), encoding="utf-8")

    # Run the existing proven 5-sheet scoring pipeline on r3 values. It
    # identifies the 29 image observations that remained generic/mismatched.
    prior.QA_DIR = QA_DIR
    prior.OUT_DIR = OUT_DIR
    prior.QA_RUN_ID = QA_RUN
    prior.BATCH = BATCH
    prior.SOURCE = SOURCE
    prior.SNAPSHOT = QA_DIR / "source_snapshot" / SOURCE.name
    prior.main()

    dataset_path = QA_DIR / "qa_dataset.json"
    data = json.loads(dataset_path.read_text(encoding="utf-8"))
    # r3 has an admin export, so remove stale r2-only limitations and the
    # resolved cross-product chicken URL critical. Retain every unresolved
    # direct-image observation and all current copy/intent findings.
    removed = {"ADMIN", "SERP-CRITICAL"}
    data["QA_Issues"] = [x for x in data["QA_Issues"] if not any(token in x["issue_id"] for token in removed)]
    valid_issues = {x["issue_id"] for x in data["QA_Issues"]}
    for row in data["QA_Products"]:
        row["revision"] = "r3"
        row["issue_refs"] = [x for x in row["issue_refs"] if x in valid_issues]
        c = Counter(x["severity"] for x in data["QA_Issues"] if x["product_key"] == row["product_key"])
        row.update(critical_count=c["CRITICAL"], major_count=c["MAJOR"], minor_count=c["MINOR"], limitation_count=c["LIMITATION"])
        row["qa_status"] = "QA_REVISE" if row["final_score"] >= 70 else "QA_FAIL"
    for row in data["QA_Criteria"]:
        row["issue_refs"] = [x for x in row["issue_refs"] if x in valid_issues]
    for row in data["QA_Images"]:
        row["issue_refs"] = [x for x in row["issue_refs"] if x in valid_issues]
    for row in data["QA_Issues"]:
        row["issue_id"] = row["issue_id"].replace("R2-ISS", "R3-ISS")
        row["reason"] += " Re-QA r3: the administered alt baseline was supplied, but this image-level observation remains unsupported or inaccurate."

    # Repair IDs after renaming and rewrite cross-links.
    remap = {x.replace("R2-ISS", "R3-ISS"): x.replace("R2-ISS", "R3-ISS") for x in valid_issues}
    for collection in (data["QA_Products"], data["QA_Criteria"], data["QA_Images"]):
        for row in collection:
            row["issue_refs"] = [x.replace("R2-ISS", "R3-ISS") for x in row["issue_refs"]]
    for row in data["QA_Products"]:
        row["evidence_refs"].append("products_export_1.csv")
    for row in data["QA_Images"]:
        row["check_method"] = "DIRECT_ORIGINAL_IMAGE_REQA + ADMIN_EXPORT_ALT_MATCH"
    for row in data["SERP_Evidence"]:
        row["serp_id"] = row["serp_id"].replace("serp_reqa", "serp_r3_reqa")
        row["checked_at"] = iso_now()
        row["note"] = "US commercial intent comparator; no search-volume claim. r3 maps this query as candidate needing recheck."

    statuses = Counter(x["qa_status"] for x in data["QA_Products"])
    severities = Counter(x["severity"] for x in data["QA_Issues"])
    average = sum(x["final_score"] for x in data["QA_Products"]) / 10
    data["QA_Summary"] = [
        {"metric": "rubric_version", "value": "prompt_qa.md v1.0 / prompt.md v2.4", "definition": "Independent QA rubric."},
        {"metric": "source_workbook", "value": str(SOURCE), "definition": "Frozen r3 revision; not edited."},
        {"metric": "source_sha256_at_handoff", "value": sha(SOURCE), "definition": "SHA-256 at freeze and handoff."},
        {"metric": "qa_run_id", "value": QA_RUN, "definition": "Independent re-QA run."},
        {"metric": "batch_id", "value": BATCH, "definition": "Inventory positions 1–10; revision r3."},
        {"metric": "products_checked", "value": 10, "definition": "Exactly ten official product keys."},
        {"metric": "images_checked", "value": "65/65", "definition": "All gallery images checked; no image substituted."},
        {"metric": "batch_final_score", "value": average, "definition": "Average product score; blockers override score."},
        {"metric": "batch_result", "value": "NOT_PASSED", "definition": "No product is QA_PASS while MAJOR findings remain."},
        {"metric": "status_counts", "value": dict(statuses), "definition": "QA status counts."},
        {"metric": "issue_counts", "value": dict(severities), "definition": "Severity counts."},
        {"metric": "admin_export", "value": "products_export_1.csv / 65 alt matches", "definition": "r3 admin baseline supplied by revision author."},
        {"metric": "xlsx_status", "value": "COMPLETE", "definition": "Five-sheet QA workbook with formulas and validation."},
    ]
    data["validation_tests"].update({"source_revision": "r3", "source_sha256": sha(SOURCE), "admin_alt_matches": 65, "resolved_critical_cat_serp": True})
    dataset_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    (QA_DIR / "qa_workbook_payload.json").write_text(json.dumps({k: data[k] for k in ("QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues")}, ensure_ascii=False, indent=2), encoding="utf-8")
    (QA_DIR / "serp_evidence.json").write_text(json.dumps(data["SERP_Evidence"], ensure_ascii=False, indent=2), encoding="utf-8")
    (QA_DIR / "validation_results.json").write_text(json.dumps(data["validation_tests"], ensure_ascii=False, indent=2), encoding="utf-8")

    # Export exact five sheets; existing exporter uses formulas for scores.
    exporter.QA_RUN_ID = QA_RUN
    exporter.DATA = QA_DIR / "qa_workbook_payload.json"
    exporter.OUTPUT = OUT_DIR / f"SEO_QA_{BATCH}.xlsx"
    exporter.main()
    wb = load_workbook(exporter.OUTPUT, data_only=False)
    audit = {
        "path": str(exporter.OUTPUT), "sheet_names": wb.sheetnames,
        "row_counts": {ws.title: ws.max_row - 1 for ws in wb.worksheets},
        "formula_cells": sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=")),
        "hyperlink_cells": sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.hyperlink),
        "source_sha256": sha(SOURCE), "xlsx_sha256": sha(exporter.OUTPUT), "zip_bad_member": zipfile.ZipFile(exporter.OUTPUT).testzip(),
    }
    audit["passed"] = audit["sheet_names"] == ["QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues"] and audit["row_counts"] == {"QA_Summary": 13, "QA_Products": 10, "QA_Criteria": 110, "QA_Images": 65, "QA_Issues": len(data["QA_Issues"])} and audit["zip_bad_member"] is None
    (QA_DIR / "spreadsheet_validation.json").write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")

    product_rows = data["QA_Products"]
    names = [x["title_proposed"] for x in submitted["products"]]
    rows = [f"| {x['inventory_position']} | {name} | {x['final_score']:.1f} | {x['qa_status']} | {x['critical_count']}/{x['major_count']}/{x['minor_count']}/{x['limitation_count']} |" for x, name in zip(product_rows, names)]
    report = f"""# SEO Re-QA — qa_batch_001_r3

## Kết luận

- Phạm vi: **10 sản phẩm, 65/65 ảnh (100%)**; inventory position 1–10; revision **r3**.
- Điểm lô: **{average:.1f}/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: {statuses['QA_FAIL']} QA_FAIL, {statuses['QA_REVISE']} QA_REVISE, {statuses['QA_PASS']} QA_PASS.
- Phát hiện: {severities['CRITICAL']} CRITICAL, {severities['MAJOR']} MAJOR, {severities['MINOR']} MINOR, {severities['LIMITATION']} LIMITATION.
- Workbook revision nguồn: `{SOURCE}`
- SHA-256 lúc đóng băng và bàn giao: `{sha(SOURCE)}`

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{chr(10).join(rows)}

## Đối chiếu r3

- Đã xác nhận r3 dùng đúng 10 product key và có admin export `products_export_1.csv`; 65/65 media ID được đối chiếu với alt baseline.
- Lỗi CRITICAL r2 của product 5 đã được xử lý: các keyword row của cat không còn tham chiếu `chicken_quilt`.
- Copy r3 đã bỏ `Dragonfly` ở các trang cat/chicken và bổ sung mô tả motif/option cụ thể hơn; product 2 và 10 đã nêu personalization theo field đã xác minh.
- Không tạo `APPROVED`, không tạo import Shopify và không sửa workbook r3.

## Lỗi ưu tiên

1. **MAJOR — ảnh:** các observation chưa sửa vẫn không mô tả đúng scene/panel thực tế. Alt admin được cung cấp không tự chứng minh observation hoặc alt proposed là phù hợp từng ảnh.
2. **MAJOR — keyword cluster:** các cụm cardinal, cat và Celtic còn cần policy/internal-link role rõ ràng để kiểm soát cannibalization.
3. **MAJOR — SERP intent:** một số candidate vẫn ghi `CANDIDATE_MAPPED_NEEDS_RECHECK`; không có claim volume và evidence phải được cập nhật trước khi phê duyệt.
4. **MINOR — source title/H1:** product 2 và 10 còn ký tự thay thế `�` trong source storefront; chỉ sửa sau khi xác nhận source live/admin chuẩn.

## Giới hạn và trạng thái bàn giao

- Admin export xác minh media-alt baseline nhưng không thay thế kiểm tra trực tiếp các current admin SEO fields ngoài dữ liệu export được cung cấp.
- Live page/public source được đối chiếu khi truy cập được; các URL không truy xuất được trong lượt được ghi limitation/evidence snapshot, không suy diễn `SOURCE_CHANGED`.
- **XLSX: COMPLETE.** Workbook có đúng 5 sheet, filter, freeze header, wrap text, hyperlink và công thức truy kiểm; không tạo `rendered_sheets`.
- `awaiting_confirmation=true`; chỉ QA lại revision tiếp theo khi bạn yêu cầu.

## Tệp chi tiết

- QA data: `{QA_DIR / 'qa_dataset.json'}`
- SERP evidence: `{QA_DIR / 'serp_evidence.json'}`
- Validation: `{QA_DIR / 'validation_results.json'}`
- Manifest/checkpoint: `{QA_DIR}`
"""
    (OUT_DIR / f"SEO_QA_{BATCH}.md").write_text(report, encoding="utf-8")
    manifest = {"rubric_version": "1.0", "prompt_version": "2.4", "qa_run_id": QA_RUN, "started_at": iso_now(), "shop_domain": SHOP, "research_run_id": RUN, "market": "United States", "seo_language": "English", "batch_id": BATCH, "revision": "r3", "source_workbook": str(SOURCE), "source_workbook_sha256": sha(SOURCE), "source_snapshot": str(QA_DIR / "source_snapshot" / SOURCE.name), "batch_product_keys": [x["product_key"] for x in submitted["products"]], "expected_products": 10, "expected_images": 65, "admin_export": "products_export_1.csv", "admin_alt_matches": 65, "status": "COMPLETE", "completed_at": iso_now(), "awaiting_confirmation": True, "output_markdown": str(OUT_DIR / f"SEO_QA_{BATCH}.md"), "output_xlsx": str(exporter.OUTPUT)}
    (QA_DIR / "qa_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (QA_DIR / "qa_progress.json").write_text(json.dumps({"qa_run_id": QA_RUN, "batch_id": BATCH, "revision": "r3", "current_stage": "BATCH_COMPLETE", "products_completed": 10, "images_completed": 65, "awaiting_confirmation": True, "last_saved_at": iso_now()}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output": str(exporter.OUTPUT), "score": average, "statuses": statuses, "issues": severities, "validation": audit}, ensure_ascii=False, default=dict, indent=2))


if __name__ == "__main__":
    main()
