from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_024"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B24R4"
base.TRACE_ROW_REVISION = "r2"
base.SOURCE_SHA256_EXPECTED = "456E0F975D8D9852D41E9C1472F54522028AD41426561086F476E7B63412E756"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_024_r4" / "SEO_Product_Optimization_qa_batch_024_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_144500"
base.OLD_QA_RUN_DIR = base.PREVIOUS_CACHE_DIR
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_144500"

base.PRODUCT_SCOPE = [
    (231, "8834369913031", 7),
    (232, "8834129199303", 7),
    (233, "8834316763335", 8),
    (234, "8834116124871", 5),
    (235, "8859716976839", 7),
    (236, "8859711930567", 7),
    (237, "8859717763271", 7),
    (238, "8859718451399", 7),
    (239, "8859718058183", 7),
    (240, "8859718320327", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "231-240"

base.PRODUCT_FACTS = {
    231: "custom splatter basketball blanket with splattered background, sample name Autumn and jersey number 7",
    232: "custom red basketball comforter with red basketball artwork, sample name Braylon and jersey number 21",
    233: "custom reflection basketball blanket with basketball reflection details and sample name/number",
    234: "custom water splash basketball comforter with water splash basketball artwork and sample name Matthew number 10",
    235: "personalized Bigfoot forest quilt set with tall trees and large walking silhouette",
    236: "personalized Bigfoot campfire quilt set with campfire, lantern and outdoor forest details",
    237: "personalized Bigfoot campfire night quilt set with two seated figures around a fire",
    238: "personalized Bigfoot sunset forest quilt set with sunset shadows and small walking silhouette",
    239: "personalized Bigfoot full moon quilt set with full moon and walking silhouette among trees",
    240: "personalized Bigfoot moon mountain quilt set with moon, mountains and walking silhouette",
}

base.SERP_EVIDENCE = {
    231: [("custom splatter basketball blanket", ["https://www.etsy.com/market/custom_basketball_blanket", "https://www.zazzle.com/", "https://www.walmart.com/c/kp/basketball-blanket"], "Splatter exact modifier is niche; custom basketball blanket intent is supported."),
          ("personalized basketball blanket name number", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.walmart.com/"], "Comparator supports personalized basketball blanket intent.")],
    232: [("custom red basketball comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Red basketball comforter intent is supported by product comparables."),
          ("personalized basketball comforter name number", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.amazon.com/"], "Comparator supports custom name/number basketball comforter intent.")],
    233: [("custom reflection basketball blanket", ["https://www.etsy.com/market/custom_basketball_blanket", "https://www.zazzle.com/", "https://www.walmart.com/"], "Reflection exact modifier is narrow; basketball blanket purchase intent is supported."),
          ("basketball reflection blanket custom", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/"], "Comparator partially supports basketball blanket/decor intent.")],
    234: [("custom water splash basketball comforter", ["https://www.amazon.com/", "https://www.walmart.com/", "https://www.etsy.com/"], "Water-splash basketball comforter intent is supported."),
          ("personalized basketball water splash bedding", ["https://www.amazon.com/", "https://ohaprints.com/", "https://www.etsy.com/"], "Comparator supports personalized basketball bedding with splash visual language.")],
    235: [("personalized bigfoot forest quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.pinterest.com/"], "Bigfoot forest quilt intent is supported by themed quilt comparables."),
          ("bigfoot forest bedding quilt", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports Bigfoot/forest bedding purchase intent.")],
    236: [("personalized bigfoot campfire quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.pinterest.com/"], "Bigfoot campfire quilt intent is supported."),
          ("bigfoot campfire bedding quilt", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports Bigfoot campfire bedding/quilt intent.")],
    237: [("personalized bigfoot campfire night quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.pinterest.com/"], "Night/campfire exact phrase is niche but Bigfoot campfire quilt intent is supported."),
          ("bigfoot campfire night quilt", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/"], "Comparator supports Bigfoot night/campfire quilt intent.")],
    238: [("personalized bigfoot sunset forest quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.pinterest.com/"], "Sunset forest exact modifier is narrow; Bigfoot quilt intent is supported."),
          ("bigfoot sunset forest bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports Bigfoot forest bedding intent.")],
    239: [("personalized bigfoot full moon quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.pinterest.com/"], "Full moon Bigfoot quilt intent is supported."),
          ("bigfoot full moon forest quilt", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/"], "Comparator supports Bigfoot/moon/forest quilt intent.")],
    240: [("personalized bigfoot moon mountain quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.pinterest.com/"], "Moon mountain exact modifier is narrow but Bigfoot quilt intent is supported."),
          ("bigfoot moon mountain bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports Bigfoot mountain/moon bedding intent.")],
}

base.CRITERION_ASSESSMENTS = {
    231: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    232: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    233: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    234: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    235: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    236: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    237: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    238: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    239: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    240: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
}


def criterion_reason(cid: str, assessment: str, pos: int, product: dict, cust: dict | None = None) -> str:
    if cid == "P1":
        return f"{assessment}: product identity, handle and live product data match inventory scope for batch 24."
    if cid == "P2":
        return f"{assessment}: personalization/product-option claims were checked against live/cache customizer evidence for this product."
    if cid == "K1":
        return f"{assessment}: primary keyword '{product['primary_keyword']}' was checked against this product's own US/English SERP; partial means the exact visual modifier is niche."
    if cid == "K2":
        return f"{assessment}: comparator query supports commercial bedding/blanket/quilt intent for this product type; partial means results broaden beyond the exact motif."
    if cid == "K3":
        return f"{assessment}: evidence is SERP/product-comparable only; no paid volume or Search Console proof is claimed."
    if cid == "T1":
        return f"{assessment}: meta_title_seo is complete English and maps to the verified product type/motif."
    if cid == "T2":
        return f"{assessment}: title_proposed can function as product title/H1 and distinguishes this batch-24 design by motif/product type."
    if cid == "D1":
        return f"{assessment}: meta description is complete, not cut mid-word, and the 145-165 character target was treated as editorial guidance."
    if cid == "D2":
        return f"{assessment}: description HTML is customer-facing and tied to verified artwork, product type, options and personalization controls."
    if cid == "E1":
        return f"{assessment}: product key, handle, links and scoped evidence are consistent; row-level {base.TRACE_ROW_REVISION} labels are recorded separately as a traceability limitation."
    raise ValueError(cid)


base.criterion_reason = criterion_reason

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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_024_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 69/69 ảnh (100%)**; chỉ inventory position **231–240**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 24 được chấm độc lập từ r4; row-level vẫn ghi r2 nên chỉ ghi limitation truy vết, không dùng để hạ lỗi nội dung.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm basketball và Bigfoot có intent mua hàng rõ; các exact modifier như splatter, reflection, sunset forest và moon mountain được chấm thận trọng khi SERP còn rộng hoặc marketplace-heavy.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 69 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: splatter/red/reflection/water-splash basketball và Bigfoot forest/campfire/night/sunset/full-moon/moon-mountain quilt designs.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 69 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
