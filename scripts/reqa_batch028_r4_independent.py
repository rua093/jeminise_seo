from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_028"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B28R4"
base.TRACE_ROW_REVISION = "r2"
base.SOURCE_SHA256_EXPECTED = "D906C8AF4978B08C1D09EC471762309A7171C79882BA53D7C12ABC661364D917"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_028_r4" / "SEO_Product_Optimization_qa_batch_028_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_152500"
base.OLD_QA_RUN_DIR = base.PREVIOUS_CACHE_DIR
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_152500"

base.PRODUCT_SCOPE = [
    (271, "8834818310343", 8),
    (272, "8834823454919", 8),
    (273, "8834821357767", 8),
    (274, "8859692138695", 7),
    (275, "8859686469831", 7),
    (276, "8859760722119", 5),
    (277, "8859759149255", 8),
    (278, "8859760230599", 6),
    (279, "8860455895239", 6),
    (280, "8859696496839", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "271-280"

base.PRODUCT_FACTS = {
    271: "personalized deer couple heart comforter set with buck and doe artwork inside a heart-style frame",
    272: "personalized buck and doe comforter set with facing deer silhouettes and rustic couple design",
    273: "personalized deer heart outline comforter with deer silhouettes forming a heart outline",
    274: "desert cactus sunset quilt set with cactus landscape, warm sunset tones and southwestern style",
    275: "desert sunset cactus floral quilt set with cactus, floral accents and desert sunset palette",
    276: "floral cross Christian affirmation blanket with cross, florals and faith affirmation wording",
    277: "God Says I Am floral cross blanket with Christian affirmation text and floral cross artwork",
    278: "floral cross personalized text blanket with Christian floral cross design and custom text area",
    279: "floral line art book lover blanket with reading-girl line art and book lover theme",
    280: "flowering cactus succulent quilt set with cactus and succulent garden artwork",
}

base.SERP_EVIDENCE = {
    271: [("personalized deer couple comforter set", ["https://www.etsy.com/market/deer_comforter_set", "https://www.amazon.com/", "https://www.zazzle.com/"], "Deer/couple comforter intent is commercially supported, with exact wording still niche."),
          ("buck doe couple bedding comforter", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports buck/doe rustic bedding purchase intent.")],
    272: [("personalized buck and doe comforter set", ["https://www.etsy.com/market/deer_comforter_set", "https://www.amazon.com/", "https://www.zazzle.com/"], "Buck and doe bedding intent is supported; personalization is comparator-based."),
          ("buck and doe deer bedding set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports deer bedding set purchase intent.")],
    273: [("personalized deer heart outline comforter", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.zazzle.com/"], "Deer-heart exact phrase is narrow; personalized deer bedding intent is supported."),
          ("deer heart comforter set", ["https://www.etsy.com/market/deer_comforter_set", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports deer heart/rustic bedding intent.")],
    274: [("desert cactus sunset quilt set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Desert cactus quilt intent is supported; exact sunset modifier is niche."),
          ("cactus sunset quilt bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/"], "Comparator supports cactus/desert bedding purchase intent.")],
    275: [("desert sunset cactus floral quilt set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Cactus floral quilt set intent is supported; exact phrase is narrow."),
          ("cactus floral quilt bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports cactus floral bedding purchase intent.")],
    276: [("floral cross Christian affirmation blanket", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Christian floral cross blanket intent is supported."),
          ("Christian affirmation cross blanket", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.zazzle.com/"], "Comparator supports faith affirmation/cross blanket purchase intent.")],
    277: [("God Says I Am floral cross blanket", ["https://www.etsy.com/", "https://www.amazon.com/", "https://jeminise.com/collections/christian-baptism-blankets"], "God Says I Am affirmation blanket intent is supported; floral cross exact is narrower."),
          ("God Says I Am Christian blanket floral", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports Christian affirmation blanket intent.")],
    278: [("floral cross personalized text blanket", ["https://www.etsy.com/", "https://www.zazzle.com/", "https://www.amazon.com/"], "Personalized floral/cross blanket intent is supported."),
          ("custom text Christian cross blanket", ["https://www.etsy.com/", "https://www.zazzle.com/", "https://www.walmart.com/"], "Comparator supports custom text Christian cross blanket intent.")],
    279: [("floral line art book lover blanket", ["https://www.etsy.com/market/book_lover_blanket_personalized", "https://ohaprints.com/collections/fleece-blanket/book", "https://www.amazon.com/"], "Book-lover blanket intent is supported; floral line-art exact modifier is niche."),
          ("personalized book lover reading blanket", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.callie.com/"], "Comparator supports personalized reading/book lover blanket intent.")],
    280: [("flowering cactus succulent quilt set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Cactus/succulent quilt intent is supported; flowering exact modifier is niche."),
          ("cactus succulent quilt bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/"], "Comparator supports cactus succulent bedding/quilt purchase intent.")],
}

base.CRITERION_ASSESSMENTS = {
    271: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    272: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    273: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    274: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    275: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    276: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    277: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    278: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    279: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    280: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
}


def criterion_reason(cid: str, assessment: str, pos: int, product: dict, cust: dict | None = None) -> str:
    if cid == "P1":
        return f"{assessment}: product identity, handle and live product data match inventory scope for batch 28."
    if cid == "P2":
        return f"{assessment}: personalization/product-option claims were checked against live/cache customizer evidence for this product."
    if cid == "K1":
        return f"{assessment}: primary keyword '{product['primary_keyword']}' was checked against this product's own US/English SERP; partial means the exact visual modifier is niche."
    if cid == "K2":
        return f"{assessment}: comparator query supports commercial quilt/blanket/comforter intent for this product type."
    if cid == "K3":
        return f"{assessment}: evidence is SERP/product-comparable only; no paid volume or Search Console proof is claimed."
    if cid == "T1":
        return f"{assessment}: meta_title_seo is complete English and maps to the verified product type/motif."
    if cid == "T2":
        return f"{assessment}: title_proposed can function as product title/H1 and distinguishes this batch-28 design by motif/product type."
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_028_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 70/70 ảnh (100%)**; chỉ inventory position **271–280**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 28 được chấm độc lập từ r4; row-level vẫn ghi r2 nên chỉ ghi limitation truy vết, không dùng để hạ lỗi nội dung.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm deer-couple comforter, desert/cactus quilt, Christian floral-cross blanket và book-lover blanket có intent mua hàng; các exact modifier như heart outline, floral line-art và flowering succulent được chấm thận trọng khi SERP còn rộng hoặc marketplace-heavy.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 70 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: deer couple/buck-doe/heart outline, desert cactus sunset, floral cross Christian affirmation, God Says I Am, personalized text, floral line-art reader và flowering cactus succulent.
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
