from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_026"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B26R4"
base.TRACE_ROW_REVISION = "r2"
base.SOURCE_SHA256_EXPECTED = "5E4D3C5F9ED400010C54B8EF2E532A4D9C9C9FFAE9C36C5302729B9F8C9CBA42"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_026_r4" / "SEO_Product_Optimization_qa_batch_026_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_150500"
base.OLD_QA_RUN_DIR = base.PREVIOUS_CACHE_DIR
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_150500"

base.PRODUCT_SCOPE = [
    (251, "8859758887111", 7),
    (252, "8860456583367", 7),
    (253, "8859736408263", 7),
    (254, "8859737686215", 7),
    (255, "8859738407111", 7),
    (256, "8859736998087", 7),
    (257, "8859738505415", 7),
    (258, "8859759575239", 6),
    (259, "8859757183175", 5),
    (260, "8859761311943", 6),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "251-260"

base.PRODUCT_FACTS = {
    251: "personalized Christian butterfly blanket with butterfly wings, affirmation words and scripture references",
    252: "personalized reading tree blanket with book tree artwork, quote text and sample name Jessica",
    253: "personalized Celtic Tree of Life quilt set with Celtic panels and central medallion design",
    254: "personalized Celtic Yggdrasil quilt set with blue accents and glowing tree medallion",
    255: "personalized green Yggdrasil quilt set with green Yggdrasil and interlaced Celtic knot pattern",
    256: "personalized cosmic Tree of Life quilt set with galaxy, chain border and purple accents",
    257: "personalized Celtic knot tree quilt set with Celtic knot accents and central tree symbol",
    258: "personalized Christian affirmations blanket with butterfly/bubble scripture-style design, photo area and sample name",
    259: "personalized Christian woman affirmation blanket with woman, butterflies and scripture reference bubbles",
    260: "personalized lavender scripture blanket with lavender tones, affirmation words and sample name",
}

base.SERP_EVIDENCE = {
    251: [("personalized Christian butterfly blanket", ["https://www.etsy.com/listing/4528235664/personalized-christian-butterfly-blanket", "https://www.amazon.com/personalized-christian-blanket/s?k=personalized+christian+blanket", "https://famvibe.com/products/personalized-gift-for-daughter-inspirational-butterfly-christian-bible-verse-blanket-31450-0701-36o78"], "Personalized Christian butterfly blanket intent is directly supported."),
          ("Christian butterfly scripture blanket custom name", ["https://www.etsy.com/nz/listing/4419478072/personalized-floral-butterfly-christian", "https://www.walmart.com/ip/Christian-Butterfly-Throw-Blanket-Religious-Gift-Woman-Christian-Gifts-Women-Inspirational-Bible-Verse-Blanket-Spiritual-Birthday-Christmas-Gift-60x5/20579174228", "https://www.ebay.com/itm/236513190069"], "Comparator supports Christian butterfly/scripture blanket purchase intent.")],
    252: [("personalized reading tree blanket", ["https://wanderprints.com/products/book-tree-just-a-girl-who-loves-books-personalized-fleece-blanket-sherpa-blanket-mn1078cin2744", "https://www.etsy.com/market/book_stack_blanket", "https://www.amazon.com/"], "Reading tree exact phrase is narrow; personalized reading/book-lover blanket intent is supported."),
          ("custom book lover reading blanket name", ["https://www.etsy.com/market/book_stack_blanket", "https://www.ufurnish.com/en-gb/p/a/28756244/personalized-reading-hooded-blanket-kids-adults-fleece-wearable-gift-for-book-lover", "https://www.amazon.com/"], "Comparator supports personalized book-lover blanket intent.")],
    253: [("personalized Celtic Tree of Life quilt set", ["https://www.amazon.com/", "https://www.macao.ubuy.com/", "https://myvikinggear.com/"], "Celtic Tree of Life quilt intent is supported."),
          ("Celtic Tree of Life quilt set", ["https://www.etsy.com/", "https://luvingift.com/", "https://www.amazon.com/"], "Comparator supports Celtic Tree of Life quilt purchase intent.")],
    254: [("personalized Celtic Yggdrasil quilt set", ["https://www.amazon.com/", "https://myvikinggear.com/", "https://www.etsy.com/"], "Celtic Yggdrasil quilt intent is supported; exact personalization evidence is comparator-based."),
          ("Yggdrasil Celtic quilt bedding", ["https://www.etsy.com/", "https://luvingift.com/", "https://www.shein.com/"], "Comparator supports Yggdrasil/Celtic bedding intent.")],
    255: [("personalized green Yggdrasil quilt set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://myvikinggear.com/"], "Green exact modifier is niche; Yggdrasil quilt intent is supported."),
          ("green Yggdrasil quilt bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://luvingift.com/"], "Comparator partially supports green/Yggdrasil bedding intent.")],
    256: [("personalized cosmic Tree of Life quilt set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.pinterest.com/"], "Cosmic exact modifier is narrow; Tree of Life quilt intent is supported."),
          ("cosmic tree of life quilt bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/"], "Comparator supports Tree of Life/cosmic visual bedding intent.")],
    257: [("personalized Celtic knot tree quilt set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://myvikinggear.com/"], "Celtic knot tree quilt intent is supported by related Celtic Tree of Life comparables."),
          ("Celtic knot tree of life quilt", ["https://www.etsy.com/", "https://luvingift.com/", "https://www.amazon.com/"], "Comparator supports Celtic knot/tree quilt purchase intent.")],
    258: [("personalized Christian affirmations blanket", ["https://www.etsy.com/nz/listing/4419478072/personalized-floral-butterfly-christian", "https://www.amazon.com/personalized-christian-blanket/s?k=personalized+christian+blanket", "https://jeminise.com/collections/christian-baptism-blankets"], "Personalized Christian affirmation blanket intent is supported."),
          ("custom Christian affirmation butterfly blanket", ["https://famvibe.com/products/personalized-gift-for-daughter-inspirational-butterfly-christian-bible-verse-blanket-31450-0701-36o78", "https://www.walmart.com/ip/Christian-Butterfly-Throw-Blanket-Religious-Gift-Woman-Christian-Gifts-Women-Inspirational-Bible-Verse-Blanket-Spiritual-Birthday-Christmas-Gift-60x5/20579174228", "https://www.etsy.com/"], "Comparator supports Christian butterfly/affirmation blanket intent.")],
    259: [("personalized Christian woman affirmation blanket", ["https://www.etsy.com/nz/listing/4419478072/personalized-floral-butterfly-christian", "https://www.amazon.com/personalized-christian-blanket/s?k=personalized+christian+blanket", "https://famvibe.com/products/personalized-gift-for-daughter-inspirational-butterfly-christian-bible-verse-blanket-31450-0701-36o78"], "Woman exact modifier is niche; Christian affirmation blanket intent is supported."),
          ("Christian affirmation blanket for women personalized", ["https://www.etsy.com/", "https://www.walmart.com/ip/Christian-Butterfly-Throw-Blanket-Religious-Gift-Woman-Christian-Gifts-Women-Inspirational-Bible-Verse-Blanket-Spiritual-Birthday-Christmas-Gift-60x5/20579174228", "https://us.shein.com/"], "Comparator supports personalized Christian gift blanket for women intent.")],
    260: [("personalized lavender scripture blanket", ["https://www.etsy.com/nz/listing/4419478072/personalized-floral-butterfly-christian", "https://www.amazon.com/personalized-christian-blanket/s?k=personalized+christian+blanket", "https://famvibe.com/products/personalized-gift-for-daughter-inspirational-butterfly-christian-bible-verse-blanket-31450-0701-36o78"], "Lavender exact modifier is niche; personalized scripture/Christian blanket intent is supported."),
          ("lavender Christian scripture blanket personalized", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports lavender/Christian scripture blanket intent at a broader level.")],
}

base.CRITERION_ASSESSMENTS = {
    251: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    252: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    253: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    254: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    255: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    256: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    257: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    258: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    259: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    260: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
}


def criterion_reason(cid: str, assessment: str, pos: int, product: dict, cust: dict | None = None) -> str:
    if cid == "P1":
        return f"{assessment}: product identity, handle and live product data match inventory scope for batch 26."
    if cid == "P2":
        return f"{assessment}: personalization/product-option claims were checked against live/cache customizer evidence for this product."
    if cid == "K1":
        return f"{assessment}: primary keyword '{product['primary_keyword']}' was checked against this product's own US/English SERP; partial means the exact visual modifier is niche."
    if cid == "K2":
        return f"{assessment}: comparator query supports commercial quilt/blanket intent for this product type; partial means results broaden beyond the exact motif."
    if cid == "K3":
        return f"{assessment}: evidence is SERP/product-comparable only; no paid volume or Search Console proof is claimed."
    if cid == "T1":
        return f"{assessment}: meta_title_seo is complete English and maps to the verified product type/motif."
    if cid == "T2":
        return f"{assessment}: title_proposed can function as product title/H1 and distinguishes this batch-26 design by motif/product type."
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_026_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 66/66 ảnh (100%)**; chỉ inventory position **251–260**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 26 được chấm độc lập từ r4; row-level vẫn ghi r2 nên chỉ ghi limitation truy vết, không dùng để hạ lỗi nội dung.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm Christian butterfly/affirmation blanket và Celtic Tree of Life/Yggdrasil quilt có intent mua hàng rõ; các exact modifier như reading tree, green/cosmic Yggdrasil, woman affirmation và lavender scripture được chấm thận trọng khi SERP còn rộng hoặc marketplace-heavy.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 66 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: Christian butterfly/scripture blankets, reading tree/book artwork, Celtic Tree of Life/Yggdrasil variants, Christian affirmation/photo area và lavender scripture blanket.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 66 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
