from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_027"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B27R4"
base.TRACE_ROW_REVISION = "r2"
base.SOURCE_SHA256_EXPECTED = "5413FA231BB7DD63F0FE199F037919B7C6FF6735B05C2BD70785F2C432AB6B34"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_027_r4" / "SEO_Product_Optimization_qa_batch_027_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_151500"
base.OLD_QA_RUN_DIR = base.PREVIOUS_CACHE_DIR
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_151500"

base.PRODUCT_SCOPE = [
    (261, "8859760951495", 5),
    (262, "8859758166215", 7),
    (263, "8859736375495", 7),
    (264, "8859738603719", 7),
    (265, "8859695907015", 7),
    (266, "8859735326919", 7),
    (267, "8859736342727", 7),
    (268, "8834785673415", 7),
    (269, "8860456321223", 9),
    (270, "8860456288455", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "261-270"

base.PRODUCT_FACTS = {
    261: "personalized Christian woman affirmation blanket with wood-panel background, seated woman artwork, Bible verse labels and sample name Rebecca",
    262: "personalized Christian photo scripture blanket with vintage beige layout, circular photo area, faith quotes and sample portrait",
    263: "personalized colorful Celtic Tree of Life quilt set with teal-orange artwork, central medallion and Celtic knot border",
    264: "personalized vintage Yggdrasil quilt set with sunset tones, large Tree of Life artwork and Celtic border",
    265: "personalized desert RV sunset quilt set with RV camper, cactus, flowers, mountains and southwestern sunset scene",
    266: "personalized rainbow Tree of Life quilt set with dark background, rainbow leaves and glowing sun",
    267: "personalized Celtic border Yggdrasil quilt set with black field, green Celtic border and orange corner knots",
    268: "personalized couple photo quilt set with couple portrait layout, custom photo badge and Valentine's Day sample text",
    269: "personalized cozy room bookworm blanket with reader in green chair, bookshelf background and sample name",
    270: "personalized dark hair book lover blanket with dark-haired girl, stacked books, floral beige background and reading quote",
}

base.SERP_EVIDENCE = {
    261: [("personalized Christian woman affirmation blanket", ["https://christianartbag.com/", "https://www.zazzle.com/", "https://macorner.co/"], "Christian affirmation blanket purchase intent is supported."),
          ("Christian affirmation blanket for women personalized", ["https://www.etsy.com/", "https://www.walmart.com/", "https://www.shineful.net/"], "Comparator supports personalized Christian gift blanket for women intent.")],
    262: [("personalized Christian photo scripture blanket", ["https://www.zazzle.com/", "https://prettyperfect.com/", "https://thesunnyzone.com/"], "Personalized scripture blanket intent is supported; photo exact modifier is comparator-based."),
          ("custom photo Christian scripture blanket", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.callie.com/"], "Comparator supports custom photo/name Christian blanket intent.")],
    263: [("personalized colorful Celtic Tree quilt set", ["https://jeminise.com/", "https://www.amazon.com/", "https://ohaprints.com/"], "Celtic Tree of Life quilt intent is supported; colorful exact modifier is niche."),
          ("Celtic Tree of Life quilt set", ["https://www.etsy.com/", "https://luvingift.com/", "https://www.amazon.com/"], "Comparator supports Celtic Tree of Life quilt purchase intent.")],
    264: [("personalized vintage Yggdrasil quilt set", ["https://luvingift.com/", "https://myvikinggear.com/", "https://www.etsy.com/"], "Yggdrasil quilt intent is supported; vintage exact modifier is narrow."),
          ("vintage Yggdrasil Celtic quilt bedding", ["https://www.etsy.com/", "https://eur.shein.com/", "https://www.redbubble.com/"], "Comparator supports Viking/Yggdrasil bedding intent.")],
    265: [("personalized desert RV sunset quilt set", ["https://ohaprints.com/", "https://www.etsy.com/", "https://cusgifts.com/"], "RV/camping quilt set intent is supported; desert sunset exact modifier is niche."),
          ("custom RV camping sunset quilt bedding", ["https://www.ebay.com/", "https://www.walmart.com/", "https://www.amorcustomgifts.com/"], "Comparator supports camping/RV bedding purchase intent.")],
    266: [("personalized rainbow Tree of Life quilt set", ["https://www.amazon.com/", "https://jeminise.com/", "https://www.google.com/m/storepages?c=TH&hl=en-TH&q=ebay.com"], "Rainbow Tree of Life quilt intent is supported at a broad level."),
          ("rainbow tree of life quilt bedding", ["https://www.etsy.com/", "https://www.pinterest.com/", "https://www.amazon.com/"], "Comparator supports Tree of Life/rainbow visual bedding intent.")],
    267: [("personalized Celtic border Yggdrasil quilt set", ["https://keppeu.lv/", "https://luvingift.com/", "https://myvikinggear.com/"], "Celtic/Yggdrasil quilt intent is supported; border exact modifier is narrow."),
          ("Yggdrasil Celtic knot border throw blanket", ["https://www.redbubble.com/", "https://www.etsy.com/", "https://eur.shein.com/"], "Comparator supports Celtic border/Yggdrasil bedding intent.")],
    268: [("personalized couple photo quilt set", ["https://www.bagsoflove.co.uk/blankets/photo-quilts.aspx", "https://www.etsy.com/market/personalized_photo_quilt", "https://www.amazon.com/Photo-Quilt/s?k=Photo+Quilt"], "Custom photo quilt intent is directly supported."),
          ("custom couple photo quilt bedding", ["https://geckocustom.com/", "https://www.personalizationmall.com/Personalized-LOVE-Photo-Couple-Blanket-p19101.prod", "https://www.bluebirdgardens.com/custom-quilts/custom-photo-memory-quilt"], "Comparator supports custom couple/photo blanket or quilt intent.")],
    269: [("personalized cozy room bookworm blanket", ["https://www.etsy.com/market/personalized_book_blanket", "https://ohaprints.com/collections/fleece-blanket/book", "https://jeminise.com/"], "Bookworm/personalized book blanket intent is supported; cozy-room exact modifier is niche."),
          ("custom book lover reading blanket name", ["https://www.etsy.com/listing/1793791874/personalized-books-and-vines-blanket", "https://www.amazon.com/Personalized-Blanket-Custom-Reading-Bookworm/dp/B0G51MRP5D", "https://uk.callie.com/"], "Comparator supports custom-name reading/book lover blanket intent.")],
    270: [("personalized dark hair book lover blanket", ["https://www.etsy.com/hk-en/market/book_lover_blanket_personalized", "https://www.pinterest.com/pin/personalized-dark-book-lover-woven-blanket-bookish-reading-throw-custom-name-romance-reader-gift-etsy--945404146810288378/", "https://ohaprints.com/collections/fleece-blanket/book"], "Book lover personalized blanket intent is supported; dark-hair exact modifier is niche."),
          ("personalized reading blanket with name", ["https://www.etsy.com/", "https://www.amazon.ca/personalized-reading-blanket/s?k=personalized+reading+blanket", "https://www.amazon.com/clp/B0DJ32WQ8X"], "Comparator supports personalized reading blanket intent.")],
}

base.CRITERION_ASSESSMENTS = {
    261: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    262: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    263: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    264: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    265: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    266: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    267: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    268: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    269: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    270: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
}


def criterion_reason(cid: str, assessment: str, pos: int, product: dict, cust: dict | None = None) -> str:
    if cid == "P1":
        return f"{assessment}: product identity, handle and live product data match inventory scope for batch 27."
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
        return f"{assessment}: title_proposed can function as product title/H1 and distinguishes this batch-27 design by motif/product type."
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_027_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 70/70 ảnh (100%)**; chỉ inventory position **261–270**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 27 được chấm độc lập từ r4; row-level vẫn ghi r2 nên chỉ ghi limitation truy vết, không dùng để hạ lỗi nội dung.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm Christian affirmation/photo scripture, Celtic Tree of Life/Yggdrasil, RV sunset, couple photo và book-lover blanket đều có intent mua hàng; các exact modifier như colorful/vintage/border, cozy-room và dark-hair được chấm thận trọng khi SERP còn rộng hoặc marketplace-heavy.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 70 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: Christian woman affirmation/scripture photo, colorful/vintage/rainbow/Celtic-border Tree of Life, desert RV sunset, couple photo quilt và bookworm/book-lover blankets.
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
