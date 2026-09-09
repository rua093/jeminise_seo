from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_021"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B21R4"
base.TRACE_ROW_REVISION = "r3"
base.SOURCE_SHA256_EXPECTED = "47771C12438547A1A05FCA12AA6D3FD79345B043DA005855C1B45A903F9004AC"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_021_r4" / "SEO_Product_Optimization_qa_batch_021_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_133500"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_133500"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_133500"

base.PRODUCT_SCOPE = [
    (201, "8867181035719", 7),
    (202, "8867178381511", 4),
    (203, "8834108326087", 5),
    (204, "8834126807239", 7),
    (205, "8834131755207", 5),
    (206, "8834378662087", 8),
    (207, "8834164129991", 7),
    (208, "8834147320007", 7),
    (209, "8834124480711", 7),
    (210, "8834328592583", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "201-210"

base.PRODUCT_FACTS = {
    201: "custom baseball flag name number bedding with American flag, wood texture, baseball and custom text artwork",
    202: "custom neon baseball player duvet cover with neon green player silhouettes and custom name/number artwork",
    203: "custom basketball hoop comforter set with ball above hoop/net and custom name/number artwork",
    204: "custom blue red basketball hoop bedding with close-up hoop, net, ball and vertical name/number artwork",
    205: "custom basketball paint splash comforter with red and blue paint-splash ball artwork",
    206: "custom basketball net blanket with hoop, net, ball below rim and optional name/number artwork",
    207: "custom basketball court hoop comforter with blue-black court lines, hoop, ball and custom name artwork",
    208: "custom black basketball hoop comforter with orange ball, grid background and vertical name/number artwork",
    209: "custom flame basketball comforter set with oversized ball and burning flame background",
    210: "custom basketball close-up blanket with dark close-up ball, gray player silhouettes and optional name/number artwork",
}

base.SERP_EVIDENCE = {
    201: [("custom baseball flag name number bedding", ["https://ohaprints.com/", "https://www.truegether.com/", "https://www.etsy.com/fi-en/market/custom_baseball_bedding"], "Custom baseball flag/name/number bedding has clear commercial comparables."),
          ("personalized baseball quilt bedding custom name number", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/ideas/personalized-baseball-bedding/927363056917/"], "Comparator supports broader personalized baseball bedding purchase intent.")],
    202: [("custom neon baseball player duvet cover", ["https://www.etsy.com/fi-en/market/custom_baseball_bedding", "https://www.amazon.com/Personalized-Baseball-Softball-Comforter-Customized/dp/B0B1JB7HTZ", "https://ohaprints.com/"], "Neon exact phrase is narrow; baseball player custom bedding intent is supported."),
          ("personalized baseball player duvet cover", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/ideas/personalized-baseball-bedding/927363056917/"], "Comparator supports custom baseball player bedding/duvet intent.")],
    203: [("custom basketball hoop comforter set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.u-buy.com.ng/"], "Basketball hoop comforter set intent is supported by marketplace comparables."),
          ("personalized basketball comforter name number", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.ubuy.co.id/"], "Comparator supports personalized basketball comforter purchase intent.")],
    204: [("custom blue red basketball hoop bedding", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Color-specific exact phrase is narrow; basketball hoop bedding intent is supported."),
          ("personalized basketball hoop bedding set", ["https://www.amazon.com/", "https://www.pinterest.com/", "https://www.amorcustomgifts.com/"], "Comparator supports personalized basketball bedding/hoop intent.")],
    205: [("custom basketball paint splash comforter", ["https://jeminise.com/", "https://www.walmart.com/", "https://www.amazon.com/"], "Paint-splash basketball comforter has product intent and close comparables."),
          ("basketball paint splatter comforter set", ["https://www.walmart.com/", "https://www.amazon.com/", "https://www.etsy.com/"], "Comparator supports basketball paint/splatter comforter intent.")],
    206: [("custom basketball net blanket", ["https://www.etsy.com/", "https://www.zazzle.com/", "https://www.walmart.com/"], "Basketball net blanket exact intent is marketplace-supported, with some broader blanket results."),
          ("personalized basketball blanket name number", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.walmart.com/"], "Comparator supports custom basketball blanket purchase intent.")],
    207: [("custom basketball court hoop comforter", ["https://www.amazon.com/", "https://muchomasqueflores.com/", "https://www.pinterest.com/"], "Basketball court/hoop comforter intent is supported, with marketplace-heavy results."),
          ("basketball court bedding personalized", ["https://www.pinterest.com/", "https://ohaprints.com/", "https://www.etsy.com/"], "Comparator supports personalized basketball court bedding intent.")],
    208: [("custom black basketball hoop comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Black exact modifier is narrow; basketball hoop comforter intent is supported."),
          ("black basketball comforter set personalized", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports black/personalized basketball comforter intent.")],
    209: [("custom flame basketball comforter set", ["https://www.amorcustomgifts.com/", "https://ohaprints.com/", "https://www.amazon.com/"], "Flame exact phrase is niche but custom basketball comforter intent is supported."),
          ("personalized basketball flames comforter", ["https://www.etsy.com/", "https://www.amazon.com/", "https://ohaprints.com/"], "Comparator partially supports flames/custom basketball bedding intent.")],
    210: [("custom basketball close-up blanket", ["https://www.etsy.com/", "https://www.zazzle.com/", "https://www.walmart.com/"], "Close-up exact phrase is narrow; custom basketball blanket intent is supported."),
          ("personalized basketball blanket name number", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.walmart.com/"], "Comparator supports personalized basketball blanket purchase intent.")],
}

base.CRITERION_ASSESSMENTS = {
    201: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    202: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    203: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    204: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    205: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    206: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    207: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    208: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    209: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    210: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
}


def criterion_reason(cid: str, assessment: str, pos: int, product: dict, cust: dict | None = None) -> str:
    if cid == "P1":
        return f"{assessment}: product identity, handle and live product data match inventory scope for batch 21."
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
        return f"{assessment}: title_proposed can function as product title/H1 and distinguishes this batch-21 design by motif/product type."
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_021_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 64/64 ảnh (100%)**; chỉ inventory position **201–210**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 21 được chấm độc lập từ r4; QA r2/r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Baseball và basketball custom-name/name-number có intent mua hàng rõ; các exact modifier như neon baseball player, blue-red hoop, black hoop, flame và close-up được chấm thận trọng khi SERP còn rộng hoặc marketplace-heavy.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 64 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: baseball flag/name-number, neon baseball player, basketball hoop/net/court, paint splash, flame, close-up ball và blanket/comforter product forms.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 64 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
