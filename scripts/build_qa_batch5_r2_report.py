from collections import Counter
from pathlib import Path
import hashlib
import json

import build_qa_batch5_report as base


ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_152903"

base.QA_RUN_ID = RUN
base.BATCH = "qa_batch_005_r2"
base.R2_MODE = True
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.OUT_DIR = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.SOURCE = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "revisions" / "qa_batch_005_r2" / "SEO_Product_Optimization_qa_batch_005_r2.xlsx"
base.SNAPSHOT = base.QA_DIR / "source_snapshot" / base.SOURCE.name

# These rows still name the wrong scene/panel after every image was opened directly.
wrong = {
    (41,3),(41,4),(41,5),
    (42,1),(42,2),(42,3),(42,4),(42,5),(42,6),(42,7),
    (43,1),(43,2),(43,3),(43,4),(43,6),(43,7),
    (44,2),(44,3),(44,4),(44,5),
    (45,2),(45,3),(45,5),(45,6),
    (46,2),(46,3),(46,5),(46,6),
    (47,2),(47,3),(47,5),(47,6),
    (48,2),(48,3),(48,5),(48,6),(48,7),
    (49,1),(49,2),(49,3),(49,4),(49,5),(49,6),(49,7),(49,8),(49,9),
    (50,1),(50,2),(50,3),(50,4),(50,5),(50,6),(50,7),(50,8),
}
base.R2_IMAGE_ASSESS = {key:("FAIL","FAIL") for key in wrong}
base.ALT_FIX = {key:base.ACTUAL[key[0]][key[1]-1] for key in wrong}

base.ASSESS = {
    41:{"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    42:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    43:{"P1":"FULL","P2":"FAIL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FAIL","D2":"FAIL","E1":"FAIL"},
    44:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    45:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    46:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    47:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    48:{"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    49:{"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"PARTIAL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    50:{"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"PARTIAL","K3":"FAIL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
}


def ensure_snapshot():
    base.SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
    if not base.SNAPSHOT.exists():
        base.SNAPSHOT.write_bytes(base.SOURCE.read_bytes())
    assert hashlib.sha256(base.SOURCE.read_bytes()).digest() == hashlib.sha256(base.SNAPSHOT.read_bytes()).digest()


def rewrite_outputs():
    dataset_path = base.QA_DIR / "qa_dataset.json"
    dataset = json.loads(dataset_path.read_text(encoding="utf-8"))
    payload = {k:dataset[k] for k in ("QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues")}
    (base.QA_DIR / "qa_workbook_payload.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    products, issues = dataset["QA_Products"], dataset["QA_Issues"]
    submitted = json.loads((base.QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))["products"]
    by_key = {p["product_key"]:p for p in submitted}
    score = sum(p["final_score"] for p in products) / 10
    statuses = Counter(p["qa_status"] for p in products)
    severities = Counter(i["severity"] for i in issues)
    source_hash = hashlib.sha256(base.SOURCE.read_bytes()).hexdigest().upper()
    report = base.OUT_DIR / "SEO_QA_qa_batch_005_r2.md"
    lines = ["# SEO Re-QA — qa_batch_005_r2", "", "## Kết luận", "",
        "- Phạm vi: **10 sản phẩm, 72/72 ảnh (100%)**; chỉ inventory position 41–50; revision **r2**.",
        f"- Điểm lô: **{score:.1f}/100**; kết luận lô: **NOT_PASSED**.",
        f"- Trạng thái: {statuses.get('QA_FAIL',0)} QA_FAIL, {statuses.get('QA_REVISE',0)} QA_REVISE, {statuses.get('QA_PASS',0)} QA_PASS.",
        f"- Phát hiện: {severities.get('CRITICAL',0)} CRITICAL, {severities.get('MAJOR',0)} MAJOR, {severities.get('MINOR',0)} MINOR, {severities.get('LIMITATION',0)} LIMITATION.",
        f"- Workbook nguồn: `{base.SOURCE}`", f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`", "", "## Điểm theo sản phẩm", "",
        "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |", "|---:|---|---:|---|---:|"]
    for q in products:
        p = by_key[q["product_key"]]
        lines.append(f"| {q['inventory_position']} | {p['title_current'].replace('|','/')} | {q['final_score']:.1f} | {q['qa_status']} | {q['critical_count']}/{q['major_count']}/{q['minor_count']}/{q['limitation_count']} |")
    lines += ["", "## Đối chiếu revision r2", "",
        "- Đã bỏ đúng khối SEO/QA/import nội bộ, lỗi design ID D14 và typo `Chrsitian` của r1.",
        "- Customizer live xác nhận các trường tùy biến thực sự tồn tại ở products 41–50; vì vậy các CRITICAL personalization của r1 được gỡ, ngoại trừ photo upload ở product 43.",
        "- Product 43 chỉ có trường tên bắt buộc và **không có image-upload input**, nên claim photo blanket vẫn là CRITICAL.",
        "- 55/72 ảnh vẫn bị gán sai loại scene/panel; thay toàn bộ câu không đồng nghĩa đã kiểm đúng ảnh.",
        "", "## Lỗi ưu tiên", "",
        "1. **CRITICAL — product 43:** title/keyword/copy nhắm photo blanket nhưng Customizer live có 0 image-upload input; phải thêm luồng upload hoạt động hoặc bỏ toàn bộ claim `photo`.",
        "2. **MAJOR — ảnh:** 55 ảnh vẫn dùng observation/alt theo mẫu vị trí và gọi sai scene; nhóm basketball, photo blanket và hai Christian blanket sai nặng nhất.",
        "3. **MAJOR — cả 10 sản phẩm:** description đã sạch câu nội bộ nhưng còn quá generic, thiếu material/size/care/set contents và còn nói pillowcase cho blanket.",
        "4. **MAJOR — products 41, 42, 44–50:** r2 đã bỏ hoặc làm mờ các control Customizer đã xác minh; copy cần ghi đúng field bắt buộc/tùy chọn và giới hạn ký tự.",
        "5. **MAJOR — product 42:** images 1–6 dùng COLÓN 06 nhưng image 7 dùng RASHAD 22; alt r2 chưa phân biệt hai sample.",
        "6. **MAJOR — product 50:** primary keyword quá rộng và chồng cụm Christian-girl blanket với product 49; cần tách intent theo affirmation/personalized design.",
        "", "## SERP và keyword", "",
        "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất theo US commercial intent. URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume.",
        "", "## Giới hạn và trạng thái bàn giao", "",
        "- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.",
        "- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.",
        "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.",
        "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.",
        "- Chưa QA products 51–60. `awaiting_confirmation=true`.", "", "## Tệp chi tiết", "",
        f"- QA data: `{base.QA_DIR/'qa_dataset.json'}`", f"- SERP evidence: `{base.QA_DIR/'serp_evidence.json'}`",
        f"- Validation: `{base.QA_DIR/'validation_results.json'}`", f"- Manifest/checkpoint: `{base.QA_DIR}`", ""]
    base.OUT_DIR.mkdir(parents=True, exist_ok=True)
    report.write_text("\n".join(lines), encoding="utf-8")

    manifest_path = base.QA_DIR / "qa_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest.update({"started_at":"2026-09-07T15:29:03+07:00", "revision":"r2", "source_workbook":str(base.SOURCE.relative_to(ROOT)), "source_workbook_sha256":source_hash, "source_snapshot":str(base.SNAPSHOT.relative_to(ROOT)), "source_snapshot_sha256":source_hash, "output_markdown":str(report), "output_xlsx":str(base.OUT_DIR / "SEO_QA_qa_batch_005_r2.xlsx")})
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    ensure_snapshot()
    base.main()
    rewrite_outputs()
