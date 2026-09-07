from pathlib import Path
import hashlib
import json

import build_qa_batch3_report as base


ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_140123"

base.QA_RUN_ID = RUN
base.QA_BATCH_ID = "qa_batch_003_r2"
base.R2_MODE = True
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.OUT_DIR = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.SOURCE = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "revisions" / "qa_batch_003_r2" / "SEO_Product_Optimization_qa_batch_003_r2.xlsx"
base.SNAPSHOT = base.QA_DIR / "source_snapshot" / base.SOURCE.name

base.SERP = {
    21: [("gingerbread candy cane quilt set", ["https://www.target.com/p/-/A-94992227"]), ("gingerbread Christmas quilt set", ["https://www.potterybarnkids.com/products/gingerbread-quilt/"])],
    22: [("black Christmas tree quilt set", ["https://www.target.com/p/-/A-86506409"]), ("dark Christmas tree quilt bedding", ["https://www.lowes.com/pd/MarCielo-3-Piece-Cotton-Christmas-Tree-King-Size-Lightweight-Quilt-Set/7057959"])],
    23: [("winter cardinal birdhouse quilt set", ["https://www.walmart.com/ip/405776273"]), ("cardinal winter bedding quilt", ["https://www.wayfair.com/bed-bath/pdp/james-home-holiday-cardinal-reversible-bedding-quilt-set-w100620099.html"])],
    24: [("gingerbread Christmas village quilt set", ["https://www.belk.com/p/levtex-home-gingerbread-village-quilt/0480056529258.html"]), ("gingerbread Christmas quilt bedding", ["https://www.potterybarnkids.com/products/gingerbread-quilt/"])],
    25: [("red gingerbread Christmas quilt", ["https://society6.com/products/red-gingerbread-christmas-cookie-pattern_comforter"]), ("snowman cardinal Christmas quilt set", ["https://charmingfavor.com/products/nd230908"])],
    26: [("white Christmas tree quilt set", ["https://www.homedepot.com/p/333170126"]), ("cream Christmas tree quilt bedding set", ["https://www.lowes.com/pd/MarCielo-3-Piece-Cotton-Christmas-Tree-King-Size-Lightweight-Quilt-Set/5016426235"])],
    27: [("gingerbread cardinal Christmas quilt set", ["https://www.walmart.com/ip/20280601271"]), ("snowman cardinal Christmas quilt set", ["https://charmingfavor.com/products/nd230908"])],
    28: [("vintage Christmas tree quilt set", ["https://www.mercari.com/us/item/m52913653160/"]), ("cream Christmas tree quilt bedding", ["https://www.target.com/p/-/A-89827491"])],
    29: [("Christmas cardinal memorial quilt I am always with you", ["https://www.etsy.com/listing/1085801805/i-am-always-with-you-blanket-cardinal"]), ("cardinal remembrance quilt bedding set", ["https://www.walmart.com/c/kp/cardinal-quilt"])],
    30: [("cardinal Christmas quilt set", ["https://www.target.com/p/-/A-81400112"]), ("cardinal wreath Christmas quilt", ["https://jeminise.com/products/winter-cardinal-on-berry-branch-patchwork-christmas-quilt-43ecda62b6-43ecda62b6"])],
}

base.PRODUCT_ASSESS = {
    21: {"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    22: {"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    23: {"P1":"FULL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    24: {"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    25: {"P1":"FAIL","P2":"PARTIAL","K1":"FAIL","K2":"PARTIAL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    26: {"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    27: {"P1":"FAIL","P2":"PARTIAL","K1":"FAIL","K2":"PARTIAL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FULL","D2":"PARTIAL","E1":"PARTIAL"},
    28: {"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    29: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
    30: {"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
}

bad_order = {(21, 3), (21, 4), (29, 2), (29, 4), (30, 2), (30, 4)}
wrong_motif = {(pos, n) for pos in (25, 27) for n in range(1, 6)}
base.R2_IMAGE_ASSESS = {key: ("FAIL", "FAIL") for key in bad_order}
base.R2_IMAGE_ASSESS.update({key: ("FULL", "FAIL") for key in wrong_motif})
base.R2_IMAGE_FIXES = {
    (21, 3): "Close-up of gingerbread and candy cane Christmas quilt print",
    (21, 4): "Matching gingerbread and candy cane Christmas quilt pillow sham",
    (29, 2): "Close-up of cardinal quilt with I am Always With You text",
    (29, 4): "Cardinal memorial Christmas quilt in an angled bedroom view",
    (30, 2): "Close-up of cardinal, pine, holly berries and pinecones on the quilt",
    (30, 4): "Cardinal Christmas quilt with sample name Sophia in an angled bedroom view",
}
for n, suffix in enumerate(("in a front bedroom view", "in an angled bedroom view", "on a matching pillow sham", "in a close-up view", "with the quilt size and components guide"), 1):
    base.R2_IMAGE_FIXES[(25, n)] = f"Snowman and cardinals Christmas quilt {suffix}"
    base.R2_IMAGE_FIXES[(27, n)] = f"Two snowmen and cardinals Christmas quilt {suffix}"
base.R2_IMAGE_REASONS = {key: "R2 vẫn đảo loại ảnh sau khi revision summary báo đã sửa." for key in bad_order}
base.R2_IMAGE_REASONS.update({key: "Alt r2 gọi snowman nhìn thấy trong ảnh là gingerbread, làm sai motif." for key in wrong_motif})


def ensure_snapshot():
    base.SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
    if not base.SNAPSHOT.exists():
        base.SNAPSHOT.write_bytes(base.SOURCE.read_bytes())
    assert hashlib.sha256(base.SOURCE.read_bytes()).hexdigest() == hashlib.sha256(base.SNAPSHOT.read_bytes()).hexdigest()


def rewrite_report():
    report = base.OUT_DIR / "SEO_QA_qa_batch_003_r2.md"
    dataset = json.loads((base.QA_DIR / "qa_dataset.json").read_text(encoding="utf-8"))
    products = dataset["QA_Products"]
    issues = dataset["QA_Issues"]
    submitted = json.loads((base.QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))["products"]
    by_key = {p["product_key"]: p for p in submitted}
    score = sum(p["final_score"] for p in products) / len(products)
    from collections import Counter
    statuses = Counter(p["qa_status"] for p in products)
    sev = Counter(i["severity"] for i in issues)
    source_hash = hashlib.sha256(base.SOURCE.read_bytes()).hexdigest().upper()
    lines = ["# SEO Re-QA — qa_batch_003_r2", "", "## Kết luận", "",
        "- Phạm vi: **10 sản phẩm, 50/50 ảnh (100%)**; chỉ inventory position 21–30; revision **r2**.",
        f"- Điểm lô: **{score:.1f}/100**; kết luận lô: **NOT_PASSED**.",
        f"- Trạng thái: {statuses.get('QA_FAIL',0)} QA_FAIL, {statuses.get('QA_REVISE',0)} QA_REVISE, {statuses.get('QA_PASS',0)} QA_PASS.",
        f"- Phát hiện: {sev.get('CRITICAL',0)} CRITICAL, {sev.get('MAJOR',0)} MAJOR, {sev.get('MINOR',0)} MINOR, {sev.get('LIMITATION',0)} LIMITATION.",
        f"- Workbook nguồn: `{base.SOURCE}`", f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`", "", "## Điểm theo sản phẩm", "",
        "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |", "|---:|---|---:|---|---:|"]
    for q in products:
        p = by_key[q["product_key"]]
        lines.append(f"| {q['inventory_position']} | {p['title_current'].replace('|','/')} | {q['final_score']:.1f} | {q['qa_status']} | {q['critical_count']}/{q['major_count']}/{q['minor_count']}/{q['limitation_count']} |")
    lines += ["", "## Đối chiếu revision r2", "",
        "- Đã sửa đúng product 24 từ snowman thành gingerbread và loại bỏ câu QA/import nội bộ khỏi 10 description.",
        "- Tuy nhiên, r2 tạo lỗi motif mới ở products 25 và 27: ảnh là snowman nhưng keyword/title/meta/alt lại ghi gingerbread.",
        "- Sáu alt/observation được báo sửa ở products 21, 29 và 30 vẫn đảo sai close-up, sham hoặc angled view.",
        "- Live Customizer xác nhận trường `Custom Your Name` tùy chọn ở products 21–28; r2 chưa mô tả điều kiện này.",
        "", "## Lỗi ưu tiên", "",
        "1. **CRITICAL — products 25, 27:** keyword/title/meta/alt nhận diện snowman thành gingerbread; phải sửa toàn bộ chuỗi keyword → copy → alt và chạy lại SERP decision.",
        "2. **MAJOR — 16 ảnh:** 6 ảnh vẫn đảo sai loại ảnh và 10 alt ở products 25/27 sai motif.",
        "3. **MAJOR — cả 10 sản phẩm:** description đã bỏ câu nội bộ nhưng vẫn quá generic, thiếu kích thước, thành phần quilt/sham và điều kiện mua đã xác minh.",
        "4. **MAJOR — products 21–28:** live Customizer có trường tên tùy chọn nhưng r2 bỏ qua tính năng đã xác minh này.",
        "5. **MAJOR — products 21, 23:** SERP exact-match cho finished quilt set vẫn chưa đủ sạch so với pattern/kit hoặc motif rộng hơn.",
        "6. **MINOR — products 21–28:** title/H1 live còn ký tự `�` và từ `Bedsprea` bị cắt.",
        "", "## SERP và keyword", "",
        "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword r2 và một phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume.",
        "", "## Giới hạn và trạng thái bàn giao", "",
        "- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.",
        "- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.",
        "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.",
        "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã được mở lại và render kiểm tra.",
        "- Chưa QA products 31–40 trong lượt này. `awaiting_confirmation=true`.", "", "## Tệp chi tiết", "",
        f"- QA data: `{base.QA_DIR/'qa_dataset.json'}`", f"- SERP evidence: `{base.QA_DIR/'serp_evidence.json'}`",
        f"- Validation: `{base.QA_DIR/'validation_results.json'}`", f"- Manifest/checkpoint: `{base.QA_DIR}`", ""]
    report.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    ensure_snapshot()
    base.main()
    rewrite_report()
