from collections import Counter
from pathlib import Path
import hashlib
import json

import build_qa_batch6_report as base


ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_161639"

base.QA_RUN_ID = RUN
base.BATCH = "qa_batch_006_r2"
base.R2_MODE = True
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.OUT_DIR = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.SOURCE = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "revisions" / "qa_batch_006_r2" / "SEO_Product_Optimization_qa_batch_006_r2.xlsx"
base.SNAPSHOT = base.QA_DIR / "source_snapshot" / base.SOURCE.name

# Every row below was reopened from the newly downloaded original image. These
# rows still name the wrong scene/panel rather than merely being concise.
wrong = set(base.WRONG)
base.R2_IMAGE_ASSESS = {key:("FAIL", "FAIL") for key in wrong}

# Product scoring follows prompt_qa.md. I1 is derived from the image rows.
base.ASSESS = {
    51:{"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
    52:{"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
    53:{"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    54:{"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    55:{"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    56:{"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    57:{"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
    58:{"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
    59:{"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
    60:{"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
}

# Fresh URLs read from the 20 re-run US commercial-intent queries.
base.SERP = {
 51:[("personalized God Says I Am blanket Emily",["https://www.etsy.com/listing/4355382705/personalized-christian-blanket-god-says"]),("custom Christian name Bible verse blanket floral",["https://www.etsy.com/listing/1770548286/custom-name-blanket-personalized"])],
 52:[("personalized Christian name blanket God Says I Am",["https://www.etsy.com/listing/4410818176/christian-name-blanket-god-says-i-am"]),("blue floral God Says I Am personalized blanket",["https://www.etsy.com/listing/4388723618/personalized-god-says-i-am-blanket"])],
 53:[("personalized soccer comforter name number",["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"]),("custom exploding soccer ball bedding set",["https://2cooldesigns.com/products/custom-soccer-theme-bedding-personalized-duvet-or-comforter-sets"])],
 54:[("custom fiery soccer ball comforter name",["https://custombeddingset.com/"]),("fire soccer bedding set duvet cover",["https://brightroomy.com/soccer-on-fire-bedding-set/"])],
 55:[("personalized flaming soccer comforter name number",["https://wearanddecor.com/products/personalized-football-quilt-bedding-set-football-ball-black-quilt-custom-name-and-number-duvet-cover-set"]),("personalized soccer bedding name number",["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"])],
 56:[("custom flaming soccer comforter",["https://2cooldesigns.com/products/custom-soccer-theme-bedding-personalized-duvet-or-comforter-sets"]),("soccer duvet cover custom name number",["https://ohaprints.com/products/personalized-soccer-duvet-cover-set-soccer-ball-zigzag-player-gift-black-white-duvet-cover-pillowcases-custom-name-number-bedding-set"])],
 57:[("personalized Proverbs 31 floral butterfly blanket",["https://www.etsy.com/market/proverbs_31_25_blanket"]),("custom name Proverbs 31 25 blanket",["https://paintingprettyhomes.com/product/proverbs-31-blanket/"])],
 58:[("personalized Bible verse affirmation floral blanket",["https://www.etsy.com/listing/1586070079/bible-verse-baby-blanket-personalized"]),("custom name Christian butterfly blanket",["https://www.etsy.com/listing/1880095899/personalized-god-says-you-are-christian"])],
 59:[("personalized floral cross Bible verse blanket name",["https://m.shein.com/us/Custom-1pc-Name-Bible-Verse-Blanket%2C-Personalized-%22You-Are-Chosen%22-Christian-Cross-With-Purple-Floral-Flannel-Shawl-Blanket%2C-Spiritual-Memorial-Blanket-For-Religious-Occasions%2C-Soft-And-Warm-Flannel%2C-Home-Decor%2C-Suitable-For-Sofa%2C-Bed%2C-Couch-p-457394582.html"]),("custom purple floral cross Christian blanket",["https://www.etsy.com/listing/4483937773/woven-cotton-cross-blanket-with-purple"])],
 60:[("personalized cross butterfly Bible affirmation blanket",["https://www.etsy.com/listing/1880095899/personalized-god-says-you-are-christian"]),("custom name pink rose cross blanket",["https://m.shein.com/us/Personalized-Name-Rose-%26-Cross-Flannel-Blanket---Custom-Name-In-Elegant-Script%2C-Heavyweight-280GSM-All-Season-Soft-Throw---Pink-Roses-%26-Black-Cross-Religious-Symbol-Design---Luxury-Home%2C-Office%2C-Travel-Decor---Machine-Washable-Gift-For-Women%2C-Men%2C-Couples---Birthday%2C-Anniversary%2C-Housewarming-Present%2C-Home-Accessory%2C-Contemporary-Style%2C-Premium-Fabric%2C-Home-Decorators-p-378308196.html"])],
}


def ensure_snapshot():
    base.SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
    if not base.SNAPSHOT.exists():
        base.SNAPSHOT.write_bytes(base.SOURCE.read_bytes())
    assert hashlib.sha256(base.SOURCE.read_bytes()).digest() == hashlib.sha256(base.SNAPSHOT.read_bytes()).digest()


def rewrite_outputs():
    dataset_path = base.QA_DIR / "qa_dataset.json"
    dataset = json.loads(dataset_path.read_text(encoding="utf-8"))
    products, issues = dataset["QA_Products"], dataset["QA_Issues"]
    submitted = json.loads((base.QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))["products"]
    by_key = {p["product_key"]:p for p in submitted}
    score = sum(p["final_score"] for p in products) / 10
    statuses = Counter(p["qa_status"] for p in products)
    severities = Counter(i["severity"] for i in issues)
    source_hash = hashlib.sha256(base.SOURCE.read_bytes()).hexdigest().upper()
    report = base.OUT_DIR / "SEO_QA_qa_batch_006_r2.md"
    lines = ["# SEO Re-QA — qa_batch_006_r2", "", "## Kết luận", "",
        "- Phạm vi: **10 sản phẩm, 78/78 ảnh (100%)**; chỉ inventory position 51–60; revision **r2**.",
        f"- Điểm lô: **{score:.1f}/100**; kết luận lô: **NOT_PASSED**.",
        f"- Trạng thái: {statuses.get('QA_FAIL',0)} QA_FAIL, {statuses.get('QA_REVISE',0)} QA_REVISE, {statuses.get('QA_PASS',0)} QA_PASS.",
        f"- Phát hiện: {severities.get('CRITICAL',0)} CRITICAL, {severities.get('MAJOR',0)} MAJOR, {severities.get('MINOR',0)} MINOR, {severities.get('LIMITATION',0)} LIMITATION.",
        f"- Workbook nguồn: `{base.SOURCE}`", f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`", "", "## Điểm theo sản phẩm", "",
        "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |", "|---:|---|---:|---|---:|"]
    for q in products:
        p = by_key[q["product_key"]]
        lines.append(f"| {q['inventory_position']} | {p['title_current'].replace('|','/')} | {q['final_score']:.1f} | {q['qa_status']} | {q['critical_count']}/{q['major_count']}/{q['minor_count']}/{q['limitation_count']} |")
    lines += ["", "## Đối chiếu revision r2", "",
        "- Đã bỏ khối SEO/QA/import nội bộ, lỗi material/design ID sao chép và các CRITICAL personalization của r1.",
        "- Customizer live xác nhận trường tên/text tồn tại ở cả 10 sản phẩm; product 52 còn có lựa chọn màu bắt buộc. Vì vậy personalization không còn là lỗi CRITICAL.",
        "- 31/78 ảnh vẫn gán sai loại scene/panel; các ảnh size, feature, care, fabric và lifestyle tiếp tục bị đảo.",
        "- Description r2 vẫn là mẫu chung, thiếu dữ kiện sản phẩm và ghi pillowcase/sham cho sáu blanket không có lựa chọn này.",
        "", "## Lỗi ưu tiên", "",
        "1. **MAJOR — 31 ảnh:** observation/alt r2 vẫn gọi sai scene hoặc information panel; cần thay theo nội dung ảnh thực tế trong `QA_Images`.",
        "2. **MAJOR — 10 sản phẩm:** description quá generic, thiếu material/size/care/set contents; nhóm blanket còn ghi sai pillowcase/sham.",
        "3. **MAJOR — 10 sản phẩm:** copy r2 không ánh xạ rõ trường Customizer đã xác minh, gồm required/optional, color choice và giới hạn text.",
        "4. **MAJOR — clusters 51–52 và 53–56:** keyword/title chưa phân vai đủ rõ; products 55–56 dùng cùng primary `flaming soccer comforter`.",
        "5. **MINOR — 10 sản phẩm:** meta description lặp cụm `product option options`.",
        "6. **MINOR — product 54:** title/H1 nguồn vẫn bị cắt ở `Machin - Design 5`.",
        "", "## SERP và keyword", "",
        "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và một phương án gần nhất theo US commercial intent. URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume.",
        "", "## Giới hạn và trạng thái bàn giao", "",
        "- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.",
        "- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.",
        "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.",
        "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.",
        "- Chưa QA products 61–70. `awaiting_confirmation=true`.", "", "## Tệp chi tiết", "",
        f"- QA data: `{base.QA_DIR/'qa_dataset.json'}`", f"- SERP evidence: `{base.QA_DIR/'serp_evidence.json'}`",
        f"- Validation: `{base.QA_DIR/'validation_results.json'}`", f"- Manifest/checkpoint: `{base.QA_DIR}`", ""]
    base.OUT_DIR.mkdir(parents=True, exist_ok=True)
    report.write_text("\n".join(lines), encoding="utf-8")

    manifest_path = base.QA_DIR / "qa_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest.update({"started_at":"2026-09-07T16:16:39+07:00", "revision":"r2", "batch_id":"qa_batch_006_r2", "source_workbook":str(base.SOURCE.relative_to(ROOT)), "source_workbook_sha256":source_hash, "source_snapshot":str(base.SNAPSHOT.relative_to(ROOT)), "source_snapshot_sha256":source_hash, "output_markdown":str(report), "output_xlsx":str(base.OUT_DIR / "SEO_QA_qa_batch_006_r2.xlsx")})
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    ensure_snapshot()
    base.main()
    rewrite_outputs()
