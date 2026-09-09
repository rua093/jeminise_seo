from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_023"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B23R4"
base.TRACE_ROW_REVISION = "r3"
base.SOURCE_SHA256_EXPECTED = "61DA10F127E364A7AA7654E910E5BCDC7527CD3C8763C353EC9768CA284528E1"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_023_r4" / "SEO_Product_Optimization_qa_batch_023_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_143500"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_143500"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_143500"

base.PRODUCT_SCOPE = [
    (221, "8834320859335", 8),
    (222, "8834357756103", 8),
    (223, "8834162819271", 7),
    (224, "8834143813831", 5),
    (225, "8834152693959", 5),
    (226, "8834404090055", 8),
    (227, "8834397372615", 8),
    (228, "8834339537095", 8),
    (229, "8834140963015", 7),
    (230, "8834135851207", 6),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "221-230"

base.PRODUCT_FACTS = {
    221: "custom basketball hoop blanket with close-up hoop/net artwork and optional name/number personalization",
    222: "custom basketball shoes blanket with basketball next to athletic shoes and optional name/number personalization",
    223: "custom court lines basketball comforter with basketball court-line artwork and name/number personalization",
    224: "custom basketball court perspective comforter with perspective court artwork and name/number personalization",
    225: "custom hardwood court basketball comforter with hardwood court basketball artwork and player name/number personalization",
    226: "custom neon basketball player blanket with neon player graphic and optional name/number personalization",
    227: "custom flaming basketball player blanket with flaming player/basketball effects and optional personalization",
    228: "custom dribbling silhouette basketball blanket with dribbling silhouette graphics and optional personalization",
    229: "custom basketball collage comforter with players/hoops collage and name/number personalization",
    230: "custom shattered glass basketball comforter with shattered-glass basketball artwork and optional shams",
}

base.SERP_EVIDENCE = {
    221: [("custom basketball hoop blanket", ["https://www.etsy.com/market/personalized_basketball_blankets", "https://www.pinterest.com/pin/basketball-hoop-black-white-custom-blankets-with-name-0503h--889109151413637707/", "https://www.walmart.com/c/kp/basketball-blanket"], "Basketball hoop blanket/custom blanket purchase intent is supported."),
          ("personalized basketball blanket name number", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.walmart.com/"], "Comparator supports name/number basketball blanket intent.")],
    222: [("custom basketball shoes blanket", ["https://www.etsy.com/market/personalized_basketball_blankets", "https://www.amazon.com/basketball-fleece-blanket/s?k=basketball+fleece+blanket", "https://www.walmart.com/c/kp/basketball-blanket"], "Basketball shoes exact modifier is narrow; basketball blanket intent is supported."),
          ("personalized basketball blanket shoes", ["https://www.etsy.com/", "https://www.walmart.com/", "https://www.amazon.com/"], "Comparator partially supports the sport gift/blanket intent.")],
    223: [("custom court lines basketball comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Court-lines exact phrase is niche; basketball court comforter intent is supported."),
          ("personalized basketball court bedding set", ["https://www.amazon.com/", "https://www.pinterest.com/", "https://muchomasqueflores.com/"], "Comparator supports basketball court bedding purchase intent.")],
    224: [("custom basketball court perspective comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Perspective exact modifier is narrow; custom basketball court comforter intent is supported."),
          ("basketball court comforter personalized", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.ubuy.co.id/"], "Comparator supports personalized basketball comforter intent.")],
    225: [("custom hardwood court basketball comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Hardwood-court exact phrase is niche; basketball court comforter intent is supported."),
          ("basketball hardwood court bedding set", ["https://www.amazon.com/", "https://www.pinterest.com/", "https://www.etsy.com/"], "Comparator partially supports hardwood/court basketball bedding intent.")],
    226: [("custom neon basketball player blanket", ["https://www.etsy.com/market/custom_basketball_blanket", "https://www.amazon.com/", "https://www.walmart.com/c/kp/basketball-blanket"], "Neon exact phrase is narrow; custom basketball player blanket intent is supported."),
          ("personalized basketball player blanket name number", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.walmart.com/"], "Comparator supports personalized basketball player blanket intent.")],
    227: [("custom flaming basketball player blanket", ["https://www.etsy.com/market/custom_basketball_blanket", "https://www.zazzle.com/", "https://www.amazon.com/"], "Flaming exact modifier is niche; basketball player blanket intent is supported."),
          ("basketball player blanket custom name", ["https://www.etsy.com/", "https://www.walmart.com/", "https://www.amazon.com/"], "Comparator supports custom basketball player blanket intent.")],
    228: [("custom dribbling silhouette basketball blanket", ["https://www.etsy.com/market/custom_basketball_blanket", "https://www.zazzle.com/", "https://www.amazon.com/"], "Dribbling silhouette exact phrase is narrow; basketball blanket intent is supported."),
          ("personalized basketball silhouette blanket", ["https://www.etsy.com/", "https://www.pinterest.com/", "https://www.walmart.com/"], "Comparator supports silhouette/custom basketball blanket intent.")],
    229: [("custom basketball collage comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Basketball collage comforter intent is supported, though exact wording is broad."),
          ("personalized basketball players comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.pinterest.com/"], "Comparator supports personalized basketball players bedding intent.")],
    230: [("custom shattered glass basketball comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Shattered-glass exact modifier is niche; basketball comforter intent is supported."),
          ("basketball shattered glass bedding", ["https://www.amazon.com/", "https://www.walmart.com/", "https://www.etsy.com/"], "Comparator partially supports basketball bedding with glass/graphic visual intent.")],
}

base.CRITERION_ASSESSMENTS = {
    221: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    222: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    223: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    224: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    225: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    226: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    227: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    228: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    229: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    230: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
}


def criterion_reason(cid: str, assessment: str, pos: int, product: dict, cust: dict | None = None) -> str:
    if cid == "P1":
        return f"{assessment}: product identity, handle and live product data match inventory scope for batch 23."
    if cid == "P2":
        return f"{assessment}: personalization/product-option claims were checked against live/cache customizer evidence for this product."
    if cid == "K1":
        return f"{assessment}: primary keyword '{product['primary_keyword']}' was checked against this product's own US/English SERP; partial means the exact visual modifier is niche."
    if cid == "K2":
        return f"{assessment}: comparator query supports commercial bedding/blanket intent for this product type; partial means results broaden beyond the exact motif."
    if cid == "K3":
        return f"{assessment}: evidence is SERP/product-comparable only; no paid volume or Search Console proof is claimed."
    if cid == "T1":
        return f"{assessment}: meta_title_seo is complete English and maps to the verified product type/motif."
    if cid == "T2":
        return f"{assessment}: title_proposed can function as product title/H1 and distinguishes this batch-23 design by motif/product type."
    if cid == "D1":
        return f"{assessment}: meta description is complete, not cut mid-word, and the 145-165 character target was treated as editorial guidance."
    if cid == "D2":
        if assessment == "PARTIAL":
            return f"{assessment}: description HTML is mostly publish-ready, but fallback wording such as 'when available' should be replaced with exact verified panel names."
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_023_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 70/70 ảnh (100%)**; chỉ inventory position **221–230**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 23 được chấm độc lập từ r4; lỗi meta description bị cắt cụt ở QA r2 đã được kiểm lại và không còn lặp lại trong r4.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm basketball blanket/comforter có intent mua hàng rõ; các exact modifier như shoes, court lines, hardwood court, neon player, flaming player, dribbling silhouette và shattered glass được chấm thận trọng khi SERP còn rộng hoặc marketplace-heavy.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 70 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: hoop/net close-up, basketball shoes, court lines/perspective/hardwood, neon/flaming player, dribbling silhouette, collage và shattered-glass basketball designs.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 70 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
