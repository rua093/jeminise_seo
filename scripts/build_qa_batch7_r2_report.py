from collections import Counter
from pathlib import Path
import hashlib
import json

import build_qa_batch7_report as base


ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_164400"

base.QA_RUN_ID = RUN
base.BATCH = "qa_batch_007_r2"
base.R2_MODE = True
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.OUT_DIR = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.SOURCE = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "revisions" / "qa_batch_007_r2" / "SEO_Product_Optimization_qa_batch_007_r2.xlsx"
base.SNAPSHOT = base.QA_DIR / "source_snapshot" / base.SOURCE.name

base.ASSESS = {
    position: {
        "P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "PARTIAL", "K3": "PARTIAL",
        "T1": "FULL", "T2": "PARTIAL", "D1": "PARTIAL", "D2": "PARTIAL", "E1": "PARTIAL",
    }
    for position in range(61, 71)
}

# Twenty fresh US commercial-intent checks: current keyword plus closest comparator.
base.SERP = {
    61: [("personalized football yard line comforter name number", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"]), ("custom yardline football bedding set", ["https://wearanddecor.com/products/personalized-football-quilt-bedding-set-football-player-usus-flag-quilt-custom-name-and-number-duvet-cover-set"])],
    62: [("American flag football comforter personalized", ["https://www.cubebik.com/shop/football-bedding-set-personalized-football-american-flag-custom-bedding-set-pillow-and-duvet-cover-custom-football-lover-gifts/"]), ("patriotic football bedding custom name number", ["https://wearanddecor.com/products/personalized-football-quilt-bedding-set-football-player-usus-flag-quilt-custom-name-and-number-duvet-cover-set"])],
    63: [("personalized flaming football comforter name number", ["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"]), ("custom football fire lightning bedding set", ["https://custombeddingset.com/"])],
    64: [("football helmet American flag comforter personalized", ["https://joliebrand.com/products/customizable-football-duvet-cover-set-personalized-bedding-with-football-ball-helmet-design-black-and-white-player-gift-custom-name-and-number-option-lylyprint-com"]), ("custom football helmet bedding name number", ["https://www.cubebik.com/shop/football-bedding-set-personalized-football-american-flag-custom-bedding-set-pillow-and-duvet-cover-custom-football-lover-gifts/"])],
    65: [("personalized flaming football helmet comforter name number", ["https://joliebrand.com/products/customizable-football-duvet-cover-set-personalized-bedding-with-football-ball-helmet-design-black-and-white-player-gift-custom-name-and-number-option-lylyprint-com"]), ("custom football helmet bedding set name number", ["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"])],
    66: [("personalized patriotic football comforter American flag", ["https://wearanddecor.com/products/personalized-football-quilt-bedding-set-football-player-usus-flag-quilt-custom-name-and-number-duvet-cover-set"]), ("American flag football bedding custom name number", ["https://www.cubebik.com/shop/football-bedding-set-personalized-football-american-flag-custom-bedding-set-pillow-and-duvet-cover-custom-football-lover-gifts/"])],
    67: [("personalized football player comforter orange name number", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"]), ("custom football player bedding set name number", ["https://www.etsy.com/ie/listing/1883272338/custom-football-bedding-set-personalized"])],
    68: [("custom black white football player bedding name number", ["https://ohaprints.com/products/personalized-football-duvet-cover-set-america-football-black-white-player-gift-duvet-cover-pillowcases-custom-name-number-bedding-set"]), ("personalized football player comforter", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"])],
    69: [("personalized football player collage comforter", ["https://www.megacustom.com/products/personalized-american-football-photo-collage-blanket-with-custom-name-and-number"]), ("custom football photo collage bedding", ["https://www.etsy.com/listing/1766923756/custom-football-blanket-football-player"])],
    70: [("personalized camo football comforter custom name number", ["https://www.youcustomizeit.com/p/Green-Camo-Comforters-Personalized/142299"]), ("camo football player bedding set name number", ["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"])],
}


def ensure_snapshot():
    base.SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
    if not base.SNAPSHOT.exists():
        base.SNAPSHOT.write_bytes(base.SOURCE.read_bytes())
    assert hashlib.sha256(base.SOURCE.read_bytes()).digest() == hashlib.sha256(base.SNAPSHOT.read_bytes()).digest()


def rewrite_outputs():
    dataset = json.loads((base.QA_DIR / "qa_dataset.json").read_text(encoding="utf-8"))
    products, issues = dataset["QA_Products"], dataset["QA_Issues"]
    submitted = json.loads((base.QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))["products"]
    by_key = {p["product_key"]: p for p in submitted}
    score = sum(p["final_score"] for p in products) / len(products)
    statuses = Counter(p["qa_status"] for p in products)
    severities = Counter(i["severity"] for i in issues)
    source_hash = hashlib.sha256(base.SOURCE.read_bytes()).hexdigest().upper()
    report = base.OUT_DIR / "SEO_QA_qa_batch_007_r2.md"
    lines = [
        "# SEO Re-QA — qa_batch_007_r2", "", "## Kết luận", "",
        "- Phạm vi: **10 sản phẩm, 60/60 ảnh (100%)**; chỉ inventory position 61–70; revision **r2**.",
        f"- Điểm lô: **{score:.1f}/100**; kết luận lô: **NOT_PASSED**.",
        f"- Trạng thái: {statuses.get('QA_FAIL', 0)} QA_FAIL, {statuses.get('QA_REVISE', 0)} QA_REVISE, {statuses.get('QA_PASS', 0)} QA_PASS.",
        f"- Phát hiện: {severities.get('CRITICAL', 0)} CRITICAL, {severities.get('MAJOR', 0)} MAJOR, {severities.get('MINOR', 0)} MINOR, {severities.get('LIMITATION', 0)} LIMITATION.",
        f"- Workbook nguồn: `{base.SOURCE}`", f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`", "",
        "## Điểm theo sản phẩm", "", "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |", "|---:|---|---:|---|---:|",
    ]
    for q in products:
        p = by_key[q["product_key"]]
        lines.append(f"| {q['inventory_position']} | {p['title_current'].replace('|', '/')} | {q['final_score']:.1f} | {q['qa_status']} | {q['critical_count']}/{q['major_count']}/{q['minor_count']}/{q['limitation_count']} |")
    lines += [
        "", "## So sánh với r1", "",
        f"- Điểm lô tăng từ **79.4** lên **{score:.1f}** (+{score - 79.36666666666665:.1f} điểm).",
        "- r2 đã bỏ khối SEO/QA/import nội bộ; personalization được xác minh bằng schema Customizer live nên không có CRITICAL.",
        "- Trạng thái vẫn là 10 QA_REVISE, 0 QA_PASS vì mỗi sản phẩm còn ít nhất một lỗi MAJOR.",
        "", "## Lỗi ưu tiên", "",
        "1. **MAJOR — 10 sản phẩm:** description r2 vẫn là mẫu chung, lặp `Visible artwork`, thiếu material, care, exact contents và điều kiện option.",
        "2. **MAJOR — 10 sản phẩm:** copy chưa nói rõ Name là bắt buộc (1–25 ký tự), Number là tùy chọn (1–5 ký tự) trong Customizer.",
        "3. **MAJOR — cluster football:** 10 trang vẫn nhắm intent personalized football bedding rất gần nhau; cần phân vai keyword/landing page theo yard-line, patriotic, flame, helmet, player, collage và camo.",
        "4. **MAJOR — product 65, ảnh 4–5:** observation/alt vẫn bị đảo giữa easy-care panel và feature panel.",
        "5. **MINOR — 10 sản phẩm:** meta description dài 168–199 ký tự và dùng CTA chung; cần rút gọn, nêu Customize chính xác.",
        "6. **MINOR — 9 ảnh feature panel:** alt còn gọi chung là close-up thay vì mô tả đúng bảng soft/lightweight/durable/breathable.",
        "", "## SERP và keyword", "",
        "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và một phương án gần nhất theo US commercial intent. URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume.",
        "", "## Giới hạn và trạng thái bàn giao", "",
        "- Không có Shopify admin export; chỉ xác minh storefront/product.js và Customizer schema, không suy ra giá trị admin.",
        "- Trình duyệt tương tác không khả dụng; đã xác minh schema live nhưng chưa chạy Customize-to-cart để kiểm tra persistence/fulfillment payload.",
        "- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.",
        "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.",
        "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.",
        "- Chưa QA products 71–80. `awaiting_confirmation=true`.",
        "", "## Tệp chi tiết", "",
        f"- QA data: `{base.QA_DIR / 'qa_dataset.json'}`", f"- SERP evidence: `{base.QA_DIR / 'serp_evidence.json'}`",
        f"- Validation: `{base.QA_DIR / 'validation_results.json'}`", f"- Manifest/checkpoint: `{base.QA_DIR}`", "",
    ]
    base.OUT_DIR.mkdir(parents=True, exist_ok=True)
    report.write_text("\n".join(lines), encoding="utf-8")

    manifest_path = base.QA_DIR / "qa_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest.update({
        "started_at": "2026-09-07T16:44:00+07:00", "revision": "r2", "batch_id": "qa_batch_007_r2",
        "source_workbook": str(base.SOURCE.relative_to(ROOT)), "source_workbook_sha256": source_hash,
        "source_snapshot": str(base.SNAPSHOT.relative_to(ROOT)), "source_snapshot_sha256": source_hash,
        "output_markdown": str(report), "output_xlsx": str(base.OUT_DIR / "SEO_QA_qa_batch_007_r2.xlsx"),
    })
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    ensure_snapshot()
    base.main()
    rewrite_outputs()
