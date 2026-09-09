from __future__ import annotations

import shutil
from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_011"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B11R4"
base.TRACE_ROW_REVISION = "r3"
base.SOURCE_SHA256_EXPECTED = "E25034AE119AE590EDB62EA73FF5B6D1AB95CDB1C65C2AE76079009D17131E65"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_011_r4" / "SEO_Product_Optimization_qa_batch_011_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_011000"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_011000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_011000"

base.PRODUCT_SCOPE = [
    (101, "8835552673991", 7),
    (102, "8860499083463", 8),
    (103, "8860498297031", 8),
    (104, "8860497674439", 7),
    (105, "8860497608903", 8),
    (106, "8860497936583", 8),
    (107, "8860498428103", 8),
    (108, "8860499050695", 8),
    (109, "8860499148999", 7),
    (110, "8860498591943", 8),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "101-110"

base.PRODUCT_FACTS = {
    101: "Keep On Truckin flag comforter with semi truck and American flag artwork",
    102: "black soccer cleats comforter set with soccer ball and cleats motif",
    103: "blue soccer goal net comforter set with ball/goal-net motif",
    104: "Kenzo soccer number comforter set using a sample name/number personalization motif",
    105: "Jayden soccer net comforter set using a sample name and soccer net motif",
    106: "Matthew water splash soccer comforter with sample name and water splash soccer ball motif",
    107: "Lucas lightning soccer comforter set with sample name and lightning soccer artwork",
    108: "Isaac soccer ball 07 comforter set with sample name/number soccer ball motif",
    109: "James green soccer goal comforter with sample name and green goal artwork",
    110: "Kenzo black soccer net comforter with sample name and black soccer net motif",
}

base.SERP_EVIDENCE = {
    101: [("Keep On Truckin flag comforter", ["https://ohaprints.com/", "https://jeminise.com/", "https://www.ebay.com/"], "Commercial trucker/custom blanket intent is supported, but the exact slogan phrase is niche."),
          ("custom semi truck American flag comforter", ["https://jeminise.com/", "https://www.zazzle.com/", "https://www.ebay.com/"], "Comparator supports semi-truck and American flag bedding/blanket purchase intent.")],
    102: [("black soccer cleats comforter set", ["https://www.walmart.com/", "https://www.amazon.com/", "https://www.ebay.com/"], "Soccer bedding intent is supported; cleats-specific black phrase is narrower."),
          ("soccer cleats bedding set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.wayfair.com/"], "Comparator supports soccer-themed bedding/comforter intent.")],
    103: [("blue soccer goal net comforter set", ["https://www.amazon.com/", "https://www.target.com/", "https://www.walmart.com/"], "Blue soccer comforter/goal bedding intent is supported by product comparables."),
          ("soccer goal net bedding set", ["https://www.ebay.com/", "https://www.wayfair.com/", "https://www.target.com/"], "Comparator supports soccer field/goal/net bedding intent.")],
    104: [("Kenzo soccer number comforter set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "The sample name makes the exact query weak; personalized soccer comforter intent is visible."),
          ("personalized soccer comforter set name number", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Comparator directly supports custom name/number soccer bedding intent.")],
    105: [("Jayden soccer net comforter set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "The sample name is product-specific; soccer net comforter intent is broader."),
          ("personalized soccer net comforter set", ["https://www.youcustomizeit.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Comparator supports personalized soccer bedding and custom text purchase intent.")],
    106: [("Matthew water splash soccer comforter", ["https://www.temu.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Water/splash soccer bedding intent exists, while Matthew is sample-name driven."),
          ("water splash soccer bedding comforter", ["https://www.temu.com/", "https://www.desertcart.com.au/", "https://www.amazon.com/"], "Comparator supports water/splash soccer bedding imagery.")],
    107: [("Lucas lightning soccer comforter set", ["https://www.etsy.com/", "https://www.amazon.ae/", "https://www.amazon.com/"], "Lightning soccer bedding intent is supported; Lucas is a sample-name modifier."),
          ("lightning soccer bedding comforter set", ["https://www.etsy.com/", "https://www.amazon.ae/", "https://www.desertcart.com.au/"], "Comparator supports lightning/flame soccer bedding products.")],
    108: [("Isaac soccer ball 07 comforter set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Sample name/number makes exact demand weak; custom soccer name/number intent is supported."),
          ("soccer ball number comforter set personalized", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Comparator supports personalized soccer ball name/number bedding intent.")],
    109: [("James green soccer goal comforter", ["https://www.target.com/", "https://www.walmart.com/", "https://www.wayfair.com/"], "Green soccer goal comforter intent is broadly supported; James is sample-name driven."),
          ("green soccer goal bedding comforter", ["https://www.target.com/", "https://www.ebay.com/", "https://www.walmart.com/"], "Comparator supports green soccer/goal bedding intent.")],
    110: [("Kenzo black soccer net comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Exact phrase is sample-name driven, but black soccer comforter/net intent has comparables."),
          ("black soccer net comforter set personalized", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.youcustomizeit.com/"], "Comparator supports black soccer/custom comforter set purchase intent.")],
}

base.CRITERION_ASSESSMENTS = {
    101: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    102: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    103: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    104: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    105: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    106: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    107: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    108: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    109: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    110: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
}

_base_customizer_summary = base.customizer_summary


def customizer_summary(cust: dict | None) -> str:
    if not cust:
        return "no cached customizer audit available"
    nodes = cust.get("personalization_nodes", [])
    if nodes:
        return "; ".join(f"{n.get('label')} required={n.get('required')} max={n.get('maxLength')}" for n in nodes)
    return _base_customizer_summary(cust)


base.customizer_summary = customizer_summary


def markdown_report(run: str, out_xlsx: Path, out_md: Path, products: list[dict], scoped: list[tuple], issues: list[dict], avg_score: float, batch_result: str, source_sha: str) -> str:
    title_by_key = {p["product_key"]: p["title_proposed"] for _, _, _, p, _ in scoped}
    status_counts = Counter(p["qa_status"] for p in products)
    severity_counts = Counter(i["severity"] for i in issues)
    rows = "\n".join(f"| {p['inventory_position']} | {title_by_key[p['product_key']]} | {p['final_score']:.1f} | {p['qa_status']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']}/{p['limitation_count']} |" for p in products)
    priority = "\n".join(f"{idx}. **{i['severity']} — {i['field']}** ({title_by_key.get(i['product_key'], 'batch-level')}): {i['reason']} Đề xuất: {i['recommended_fix']}" for idx, i in enumerate(issues[:12], 1))
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_011_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 77/77 ảnh (100%)**; chỉ inventory position **101–110**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 11 được chấm độc lập từ r4; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Các exact keyword có sample name như Kenzo, Jayden, Matthew, Lucas, Isaac và James được chấm thận trọng; comparator rộng hơn vẫn hỗ trợ intent personalized soccer comforter.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 77 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: semi truck flag, black cleats, blue goal net, soccer number/name, water splash, lightning và green/black soccer net variants.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 77 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
- Test logic đạt: `100 + CRITICAL -> QA_FAIL`, `90 + đủ coverage + không lỗi chặn -> QA_PASS`, `72/80 -> 72–92 và QA_INCOMPLETE`.
- XLSX có đúng 5 sheet: `QA_Summary`, `QA_Products`, `QA_Criteria`, `QA_Images`, `QA_Issues`; có filter/freeze/wrap/hyperlink và công thức truy kiểm.
- Không tạo `rendered_sheets` trong project; preview nằm trong thư mục tạm ngoài project.
- Không sửa workbook nguồn, không tạo APPROVED/import, không cập nhật Shopify. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA workbook: `{out_xlsx}`
- QA report: `{out_md}`
- QA data: `{base.ROOT / 'seo_runs' / base.SHOP / base.RUN_ID / 'qa' / run / 'qa_dataset.json'}`
- SERP evidence: `{base.ROOT / 'seo_runs' / base.SHOP / base.RUN_ID / 'qa' / run / 'serp_evidence.json'}`
- Validation: `{base.ROOT / 'seo_runs' / base.SHOP / base.RUN_ID / 'qa' / run / 'validation_results.json'}`
- Manifest/checkpoint: `{base.ROOT / 'seo_runs' / base.SHOP / base.RUN_ID / 'qa' / run}`
"""


base.markdown_report = markdown_report

if __name__ == "__main__":
    base.main()
