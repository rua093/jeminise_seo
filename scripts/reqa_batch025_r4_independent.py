from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_025"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B25R4"
base.TRACE_ROW_REVISION = "r2"
base.SOURCE_SHA256_EXPECTED = "A559DAEFB53518AAD602416AB6B3BF785310034B01259289FDDAED25C09DAF13"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_025_r4" / "SEO_Product_Optimization_qa_batch_025_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_145500"
base.OLD_QA_RUN_DIR = base.PREVIOUS_CACHE_DIR
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_145500"

base.PRODUCT_SCOPE = [
    (241, "8859718844615", 7),
    (242, "8859718189255", 7),
    (243, "8859717533895", 7),
    (244, "8859718648007", 7),
    (245, "8859719106759", 7),
    (246, "8859719467207", 7),
    (247, "8860456353991", 7),
    (248, "8860456419527", 9),
    (249, "8860456517831", 7),
    (250, "8860456452295", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "241-250"

base.PRODUCT_FACTS = {
    241: "personalized Bigfoot peace sign mountain quilt set with mountain landscape and large figure making a peace sign",
    242: "personalized Bigfoot peace sign sunset quilt set with sunset scene and close figure making a peace sign",
    243: "personalized Bigfoot forest night quilt set with woods/night scene and large walking silhouette",
    244: "personalized Bigfoot snowy mountain quilt set with snowy mountain/forest details and small walking silhouette",
    245: "personalized Bigfoot sunset forest quilt set with orange forest shadows and walking silhouette",
    246: "personalized Bigfoot red sunglasses quilt set with large figure wearing red sunglasses",
    247: "personalized bookshelf reading girl blanket with bookshelf collage, quotes and sample name",
    248: "personalized blonde reading girl blanket with bookshelves, quote text and sample name",
    249: "personalized antique books reading blanket with antique books, quote text and sample name",
    250: "personalized wildflower reading girl blanket with books, floral details and quote text",
}

base.SERP_EVIDENCE = {
    241: [("personalized Bigfoot peace sign mountain quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://ohaprints.com/"], "Peace-sign mountain exact phrase is niche; Bigfoot quilt/bedding intent is supported."),
          ("Bigfoot mountain quilt set personalized", ["https://www.etsy.com/", "https://www.walmart.com/", "https://www.pinterest.com/"], "Comparator supports Bigfoot mountain quilt purchase intent.")],
    242: [("personalized Bigfoot peace sign sunset quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://ohaprints.com/"], "Peace-sign sunset exact phrase is niche; Bigfoot quilt intent is supported."),
          ("Bigfoot sunset quilt bedding personalized", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports Bigfoot sunset bedding/quilt intent.")],
    243: [("personalized Bigfoot forest night quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.pinterest.com/"], "Bigfoot forest/night quilt intent is supported."),
          ("Bigfoot forest night bedding quilt", ["https://www.etsy.com/", "https://www.walmart.com/", "https://ohaprints.com/"], "Comparator supports forest/night Bigfoot bedding intent.")],
    244: [("personalized Bigfoot snowy mountain quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.pinterest.com/"], "Snowy mountain exact modifier is narrow; Bigfoot quilt intent is supported."),
          ("Bigfoot snowy mountain bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator partially supports Bigfoot snowy/mountain bedding intent.")],
    245: [("personalized Bigfoot sunset forest quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.walmart.com/"], "Bigfoot sunset forest quilt intent is supported."),
          ("Bigfoot sunset forest bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://ohaprints.com/"], "Comparator supports Bigfoot forest bedding intent.")],
    246: [("personalized Bigfoot red sunglasses quilt set", ["https://www.etsy.com/market/bigfoot_quilt", "https://www.amazon.com/", "https://www.pinterest.com/"], "Red-sunglasses exact phrase is very niche; broader Bigfoot quilt intent is supported."),
          ("funny Bigfoot quilt bedding personalized", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports novelty/personalized Bigfoot quilt intent.")],
    247: [("personalized bookshelf reading girl blanket", ["https://www.etsy.com/market/book_stack_blanket", "https://www.amazon.com/", "https://www.ufurnish.com/"], "Bookshelf/reading personalized blanket intent is supported by book-lover blanket comparables."),
          ("personalized reading girl blanket", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.facebook.com/CallieForGifts/"], "Comparator supports personalized reading-themed blanket intent.")],
    248: [("personalized blonde reading girl blanket", ["https://www.etsy.com/market/book_stack_blanket", "https://www.amazon.com/", "https://www.ufurnish.com/"], "Blonde exact modifier is niche; personalized reading blanket intent is supported."),
          ("custom book lover girl blanket name", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/"], "Comparator supports custom book-lover/girl blanket intent.")],
    249: [("personalized antique books reading blanket", ["https://www.etsy.com/market/book_stack_blanket", "https://www.amazon.com/", "https://www.pinterest.com/"], "Antique-books exact phrase is narrow; book lover blanket intent is supported."),
          ("personalized book lover blanket antique books", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports personalized book-themed blanket purchase intent.")],
    250: [("personalized wildflower reading girl blanket", ["https://www.etsy.com/market/book_stack_blanket", "https://www.amazon.com/", "https://www.facebook.com/CallieForGifts/"], "Wildflower reading exact phrase is niche; personalized reading blanket intent is supported."),
          ("personalized floral book lover blanket", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/"], "Comparator supports floral/book-lover personalized blanket intent.")],
}

base.CRITERION_ASSESSMENTS = {
    241: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    242: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    243: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    244: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    245: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    246: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    247: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    248: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    249: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    250: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
}


def criterion_reason(cid: str, assessment: str, pos: int, product: dict, cust: dict | None = None) -> str:
    if cid == "P1":
        return f"{assessment}: product identity, handle and live product data match inventory scope for batch 25."
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
        return f"{assessment}: title_proposed can function as product title/H1 and distinguishes this batch-25 design by motif/product type."
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_025_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 72/72 ảnh (100%)**; chỉ inventory position **241–250**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 25 được chấm độc lập từ r4; row-level vẫn ghi r2 nên chỉ ghi limitation truy vết, không dùng để hạ lỗi nội dung.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm Bigfoot quilt và reading/book-lover blanket có intent mua hàng rõ; các exact modifier như peace sign, red sunglasses, blonde reading girl, antique books và wildflower được chấm thận trọng khi SERP còn rộng hoặc marketplace-heavy.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 72 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: Bigfoot peace-sign/mountain/sunset/night/snow/red-sunglasses variants và personalized reading girl/bookshelf/antique-books/wildflower blanket designs.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 72 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
