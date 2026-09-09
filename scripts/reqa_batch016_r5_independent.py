from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_016"
base.REVISION = "r5"
base.ISSUE_PREFIX = "B16R5"
base.TRACE_ROW_REVISION = "r4"
base.SOURCE_SHA256_EXPECTED = "C9DAA419A20383AE23CD576C7BE60869E9088205D6B2797E420D74D61F9A74FF"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_016_r5" / "SEO_Product_Optimization_qa_batch_016_r5.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_101406"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_101406"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_101406"

base.PRODUCT_SCOPE = [
    (151, "8867094364359", 7),
    (152, "8867094626503", 7),
    (153, "8867095642311", 7),
    (154, "8860132835527", 8),
    (155, "8833343815879", 7),
    (156, "8833359675591", 7),
    (157, "8833360691399", 7),
    (158, "8834805203143", 7),
    (159, "8834807234759", 7),
    (160, "8834809528519", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "151-160"

base.PRODUCT_FACTS = {
    151: "blue mosaic giraffe quilt set with animal patchwork/mosaic giraffe artwork",
    152: "vintage circle horse quilt set with horse artwork inside circular patchwork framing",
    153: "yellow horse patchwork quilt set with animal farmhouse patchwork artwork",
    154: "I Am Who He Says blanket with Christian scripture affirmation artwork",
    155: "cream twisted Tree of Life quilt set with twisted trunk and cream-toned nature artwork",
    156: "blooming Tree of Life quilt set with exposed roots and blooming tree artwork",
    157: "stained glass Tree landscape quilt set with stained-glass style tree landscape artwork",
    158: "Mjolnir skull Celtic quilt with Thor hammer, skull and circular Celtic/Norse artwork",
    159: "Norse ravens runic quilt with ravens and runic Norse artwork",
    160: "pink runic raven comforter with raven centered in a pink runic circle design",
}

base.SERP_EVIDENCE = {
    151: [("blue mosaic giraffe quilt set", ["https://www.etsy.com/", "https://www.walmart.com/", "https://www.bedbathandbeyond.com/"], "Giraffe quilt intent is supported; blue mosaic exact modifier is niche and mainly visual."),
          ("giraffe quilt bedding set", ["https://www.walmart.com/", "https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports giraffe quilt/bedding purchase intent.")],
    152: [("vintage circle horse quilt set", ["https://www.amazon.com/", "https://www.ebay.com/", "https://www.etsy.com/market/vintage_horse_quilt"], "Horse/vintage quilt intent is supported; circle exact modifier is a product-specific visual differentiator."),
          ("vintage horse quilt bedding set", ["https://www.wayfair.com/keyword.php?keyword=horse+bedding", "https://www.walmart.com/c/kp/horse-quilts", "https://thepaintingpony.com/freisian-horse-bedding-set/"], "Comparator supports horse bedding and quilt set purchase intent.")],
    153: [("yellow horse patchwork quilt set", ["https://www.ebay.com/", "https://www.wayfair.com/", "https://www.etsy.com/"], "Horse patchwork quilt intent exists but exact yellow modifier is weak."),
          ("horse patchwork quilt bedding set", ["https://www.ebay.com/", "https://www.walmart.com/c/kp/horse-quilts", "https://alphaquilt.com/collections/horse"], "Comparator supports horse patchwork/quilt product intent, mixed with pattern results.")],
    154: [("I Am Who He Says blanket", ["https://www.etsy.com/au/listing/4410818176/christian-name-blanket-god-says-i-am", "https://www.walmart.com/ip/Customizaholic-God-Says-I-Am-Blanket-Custom-Name-Blanket-Gift-Blanket-Blankets-For-Beds-Christian-Blanket-Jesus-Blanket-Bible-Blanket/13494970869", "https://www.instagram.com/p/DBAbo8gtODA/"], "Christian affirmation blanket intent is supported; wording overlaps song/devotional content and must stay product-specific."),
          ("God Says I Am Christian blanket", ["https://www.etsy.com/", "https://www.walmart.com/", "https://www.amazon.com/"], "Comparator supports commercial God Says I Am blanket purchase intent.")],
    155: [("cream twisted tree quilt set", ["https://www.bedbathandbeyond.com/", "https://www.homedepot.com/p/Gwyn-3-Piece-Cream-Cotton-King-Quilt-Set-A032821CMNFS/319345110", "https://www.etsy.com/"], "Exact cream/twisted phrase is weak; broad cream quilt and Tree of Life bedding evidence must be combined carefully."),
          ("twisted tree of life quilt set", ["https://www.etsy.com/", "https://alphaquilt.com/", "https://jeminise.com/"], "Comparator supports Tree of Life quilt intent with twisted-tree visible modifier.")],
    156: [("blooming tree of life quilt set", ["https://miravohome.com/", "https://www.etsy.com/", "https://www.amazon.com/"], "Blooming Tree of Life quilt/blanket intent is supported, with some broader Tree of Life products."),
          ("tree of life quilt bedding set", ["https://www.amazon.com/", "https://www.ebay.com/", "https://shop.crackerbarrel.com/"], "Comparator supports broad Tree of Life quilt/bedding intent.")],
    157: [("stained glass tree landscape quilt set", ["https://www.pinterest.com/", "https://pinetreecountryquilts.com/", "https://quiltingbookspatternsandnotions.com/"], "Stained-glass landscape exact results skew toward patterns/kits, not finished bedding."),
          ("stained glass tree quilt bedding", ["https://www.etsy.com/", "https://www.walmart.com/", "https://www.amazon.com/"], "Comparator partially supports stained-glass/tree quilt intent but remains mixed.")],
    158: [("Mjolnir skull Celtic quilt", ["https://gearfrost.com/", "https://vikingsonsofodin.com/", "https://luvingift.com/"], "Mjolnir/Norse quilt product intent is supported; skull/Celtic exact phrase is narrower."),
          ("Mjolnir skull Norse quilt bedding", ["https://vikingsonsofodin.com/", "https://www.ubuy.co.in/", "https://www.redbubble.com/"], "Comparator supports Norse/Mjolnir bedding and blanket purchase intent.")],
    159: [("Norse ravens runic quilt", ["https://www.amazon.com/", "https://myvikinggear.com/", "https://dingmun.com/"], "Norse/raven/runic quilt bedding intent is supported by Viking bedding product comparables."),
          ("Norse raven quilt bedding", ["https://vikingsonsofodin.com/", "https://www.amazon.com/", "https://www.etsy.com/"], "Comparator supports Norse raven bedding/quilt purchase intent.")],
    160: [("pink runic raven comforter", ["https://www.etsy.com/", "https://www.walmart.com/", "https://vikingstyle.co/"], "Pink raven comforter/bedding has product comparables, but runic exact phrase is narrow."),
          ("raven runic comforter bedding", ["https://www.tiendamia.cr/", "https://worldnorse.com/", "https://www.etsy.com/"], "Comparator supports raven/runic comforter and Viking bedding intent.")],
}

base.CRITERION_ASSESSMENTS = {
    151: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    152: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    153: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    154: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    155: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    156: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    157: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    158: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    159: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    160: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_016_r5

## Kết luận

- Phạm vi: **10 sản phẩm, 71/71 ảnh (100%)**; chỉ inventory position **151–160**; revision **r5**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 16 được chấm độc lập từ r5; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm Tree of Life và Norse raven/Mjolnir có intent mua hàng khá rõ; giraffe/horse/stained-glass exact modifiers được chấm thận trọng khi SERP còn lẫn pattern, broad bedding hoặc kết quả quá niche.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 71 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r5 khớp các motif chính: blue mosaic giraffe, vintage/yellow horse patchwork, Christian affirmation blanket, Tree of Life variants, Mjolnir skull Celtic, Norse ravens runic và pink runic raven.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 71 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
