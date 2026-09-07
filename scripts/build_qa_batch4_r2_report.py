from collections import Counter
from pathlib import Path
import hashlib
import json

import build_qa_batch4_report as base


ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_144539"

base.QA_RUN_ID = RUN
base.BATCH = "qa_batch_004_r2"
base.R2_MODE = True
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.OUT_DIR = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.SOURCE = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "revisions" / "qa_batch_004_r2" / "SEO_Product_Optimization_qa_batch_004_r2.xlsx"
base.SNAPSHOT = base.QA_DIR / "source_snapshot" / base.SOURCE.name

base.SERP = {
    31: [("Christmas cardinal birdhouse quilt set", ["https://www.etsy.com/listing/4348580665/bird-feeders-and-cardinals-quilt-kit"]), ("cardinal poinsettia quilt set", ["https://www.lakeside.com/products/nordic-cardinal-quilt-set-full-queen-or-king-with-shams", "https://alphaquilt.com/products/tai041124235"])],
    32: [("Christmas cardinal quilt set", ["https://www.lakeside.com/products/nordic-cardinal-quilt-set-full-queen-or-king-with-shams"]), ("snowy cardinal bedding quilt", ["https://www.kohls.com/product/prd-8133602/cf-home-natures-holiday-cardinal-king-quilt-set-with-shams.jsp", "https://www.marcielobedding.com/products/marcelo-3-pcs-winter-cardinals-christmas-quilt-bedspread-set-decor"])],
    33: [("cow sunflower quilt set", ["https://www.target.com/p/-/A-1005538288", "https://www.wayfair.com/bed-bath/pdp/rt-designers-collection-cow-sunflowers-farmhouse-microfiber-quilt-qlko2052.html"]), ("farmhouse cow bedding", ["https://www.walmart.com/ip/19924623339"])],
    34: [("crocodile patchwork quilt set", ["https://vantique.net/products/everglades-alligator-charm-3-piece-quilted-bedding-set-ncu0dv5368"]), ("crocodile bedding", ["https://abracadabrakids.com/products/abracadabra-6-piece-cot-bedding-set-crocodile"])],
    35: [("personalized football comforter", ["https://www.youcustomizeit.com/p/Football-Comforters-Personalized/142169"]), ("custom football bedding name number", ["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"])],
    36: [("grunge football comforter", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"]), ("personalized football quilt bedding", ["https://wearanddecor.com/products/personalized-football-quilt-bedding-set-football-ball-black-quilt-custom-name-and-number-duvet-cover-set"])],
    37: [("cosmic football comforter", []), ("personalized football comforter", ["https://www.youcustomizeit.com/p/Football-Comforters-Personalized/142169"])],
    38: [("personalized patriotic football comforter", ["https://custombeddingset.com/america-football-custom-bedding-set-personalized-us-flag-duvet-cover-bed-sheets-pillow-shams/"]), ("American flag custom football bedding", ["https://wearanddecor.com/products/personalized-football-quilt-bedding-set-football-player-usus-flag-quilt-custom-name-and-number-duvet-cover-set"])],
    39: [("vintage football personalized bedding", ["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"]), ("custom football comforter", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"])],
    40: [("custom patriotic football comforter", ["https://www.cubebik.com/shop/football-bedding-set-personalized-football-american-flag-custom-bedding-set-pillow-and-duvet-cover-custom-football-lover-gifts/"]), ("personalized football comforter", ["https://www.youcustomizeit.com/p/Football-Comforters-Personalized/142169"])],
}

base.ASSESS = {
    31: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    32: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    33: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    34: {"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    35: {"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    36: {"P1":"FULL","P2":"PARTIAL","K1":"FAIL","K2":"PARTIAL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    37: {"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FAIL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    38: {"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"FAIL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    39: {"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"PARTIAL","K3":"PARTIAL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
    40: {"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"FAIL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"PARTIAL"},
}

wrong = {(31,6), (32,2), (32,4), (33,6), (34,6)}
ambiguous = {(31,3),(31,4),(31,5),(32,3),(32,5),(33,3),(33,4),(33,5),(34,3),(34,4),(34,5)}
ambiguous |= {(pos,n) for pos in range(35,41) for n in (3,5)}
base.R2_IMAGE_ASSESS = {key:("FAIL","FAIL") for key in wrong}
base.R2_IMAGE_ASSESS.update({key:("PARTIAL","PARTIAL") for key in ambiguous})
base.R2_IMAGE_SEVERITY = {key:"MAJOR" for key in wrong}
base.R2_IMAGE_SEVERITY.update({key:"MINOR" for key in ambiguous})
base.R2_IMAGE_REASON = {key:"R2 vẫn gán sai loại ảnh so với ảnh gốc đã mở trực tiếp." for key in wrong}
base.R2_IMAGE_REASON.update({key:"Observation/alt r2 vẫn dùng nhãn mẫu chung thay vì mô tả đúng mục đích ảnh." for key in ambiguous})
base.R2_IMAGE_FIX = {(pos,n): base.ACTUAL[pos][n-1] for pos,n in wrong | ambiguous}


def ensure_snapshot():
    base.SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
    if not base.SNAPSHOT.exists():
        base.SNAPSHOT.write_bytes(base.SOURCE.read_bytes())
    assert hashlib.sha256(base.SOURCE.read_bytes()).digest() == hashlib.sha256(base.SNAPSHOT.read_bytes()).digest()


def rewrite_outputs():
    dataset_path = base.QA_DIR / "qa_dataset.json"
    dataset = json.loads(dataset_path.read_text(encoding="utf-8"))
    for product in dataset["QA_Products"]:
        product["revision"] = "r2"
    dataset_path.write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
    payload = {k:dataset[k] for k in ("QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues")}
    (base.QA_DIR / "qa_workbook_payload.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    products, issues = dataset["QA_Products"], dataset["QA_Issues"]
    submitted = json.loads((base.QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))["products"]
    by_key = {p["product_key"]:p for p in submitted}
    score = sum(p["final_score"] for p in products) / 10
    statuses, severities = Counter(p["qa_status"] for p in products), Counter(i["severity"] for i in issues)
    source_hash = hashlib.sha256(base.SOURCE.read_bytes()).hexdigest().upper()
    report = base.OUT_DIR / "SEO_QA_qa_batch_004_r2.md"
    lines = ["# SEO Re-QA — qa_batch_004_r2", "", "## Kết luận", "",
        "- Phạm vi: **10 sản phẩm, 62/62 ảnh (100%)**; chỉ inventory position 31–40; revision **r2**.",
        f"- Điểm lô: **{score:.1f}/100**; kết luận lô: **NOT_PASSED**.",
        f"- Trạng thái: {statuses.get('QA_FAIL',0)} QA_FAIL, {statuses.get('QA_REVISE',0)} QA_REVISE, {statuses.get('QA_PASS',0)} QA_PASS.",
        f"- Phát hiện: {severities.get('CRITICAL',0)} CRITICAL, {severities.get('MAJOR',0)} MAJOR, {severities.get('MINOR',0)} MINOR, {severities.get('LIMITATION',0)} LIMITATION.",
        f"- Workbook nguồn: `{base.SOURCE}`", f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`", "", "## Điểm theo sản phẩm", "",
        "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |", "|---:|---|---:|---|---:|"]
    for q in products:
        p = by_key[q["product_key"]]
        lines.append(f"| {q['inventory_position']} | {p['title_current'].replace('|','/')} | {q['final_score']:.1f} | {q['qa_status']} | {q['critical_count']}/{q['major_count']}/{q['minor_count']}/{q['limitation_count']} |")
    lines += ["", "## Đối chiếu revision r2", "",
        "- Đã bỏ đúng khối SEO/QA nội bộ, thuộc tính `Dragonfly` ở products 33–34 và từ khóa `alligator` không có bằng chứng ở product 34.",
        "- Customizer live xác nhận products 35–40 có `Enter Name` bắt buộc (1–25 ký tự) và `Enter Number` tùy chọn (1–5 ký tự); vì vậy các CRITICAL personalization của r1 được gỡ.",
        "- Tuy nhiên r2 bỏ luôn personalization khỏi keyword/copy của products 35–40, làm mất một thuộc tính mua hàng đã được xác minh.",
        "- 5 ảnh vẫn bị gán sai loại và 23 ảnh còn observation/alt theo mẫu chung.",
        "", "## Lỗi ưu tiên", "",
        "1. **MAJOR — products 35–40:** phải đưa personalization đã xác minh trở lại keyword/copy và nói rõ tên bắt buộc, số tùy chọn; tên/số trong ảnh chỉ là mẫu.",
        "2. **MAJOR — cả 10 sản phẩm:** description đã sạch câu nội bộ nhưng vẫn quá generic, thiếu thông số, thành phần, care và điều kiện mua quan trọng.",
        "3. **MAJOR — ảnh:** products 31, 33, 34 vẫn gọi image 6 là size guide; product 32 vẫn đảo close-up và angled view.",
        "4. **MAJOR — keywords:** product 34 có evidence exact yếu; product 36 dùng cụm không tự nhiên; product 37 SERP lệch intent; products 38 và 40 trùng intent.",
        "5. **MINOR — 23 ảnh:** alt/observation vẫn dùng nhãn vị trí chung thay vì nội dung trực tiếp của ảnh.",
        "", "## SERP và keyword", "",
        "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword r2 và một phương án gần nhất theo US intent. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume.",
        "", "## Giới hạn và trạng thái bàn giao", "",
        "- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.",
        "- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.",
        "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.",
        "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.",
        "- Chưa QA products 41–50. `awaiting_confirmation=true`.", "", "## Tệp chi tiết", "",
        f"- QA data: `{base.QA_DIR/'qa_dataset.json'}`", f"- SERP evidence: `{base.QA_DIR/'serp_evidence.json'}`",
        f"- Validation: `{base.QA_DIR/'validation_results.json'}`", f"- Manifest/checkpoint: `{base.QA_DIR}`", ""]
    report.write_text("\n".join(lines), encoding="utf-8")

    manifest_path = base.QA_DIR / "qa_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest.update({"started_at":"2026-09-07T14:45:39+07:00", "revision":"r2", "source_workbook":str(base.SOURCE.relative_to(ROOT)), "source_workbook_sha256":source_hash, "source_snapshot":str(base.SNAPSHOT.relative_to(ROOT)), "source_snapshot_sha256":source_hash, "output_markdown":str(report), "output_xlsx":str(base.OUT_DIR / "SEO_QA_qa_batch_004_r2.xlsx")})
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    ensure_snapshot()
    base.main()
    rewrite_outputs()
