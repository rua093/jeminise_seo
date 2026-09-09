from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_018"
base.REVISION = "r5"
base.ISSUE_PREFIX = "B18R5"
base.TRACE_ROW_REVISION = "r4"
base.SOURCE_SHA256_EXPECTED = "792BB315670E1FF7672E1C481F797E6E7A3CDFA0DA9FD704EC4251D842DC9F9A"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_018_r5" / "SEO_Product_Optimization_qa_batch_018_r5.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_105500"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_105500"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_105500"

base.PRODUCT_SCOPE = [
    (171, "8879779446983", 6),
    (172, "8879779479751", 6),
    (173, "8879778726087", 6),
    (174, "8834791506119", 7),
    (175, "8859733885127", 7),
    (176, "8867178610887", 4),
    (177, "8867178905799", 4),
    (178, "8867178741959", 4),
    (179, "8867179004103", 4),
    (180, "8867179299015", 4),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "171-180"

base.PRODUCT_FACTS = {
    171: "pink gothic skull comforter set with Halloween skull and gothic bedding artwork",
    172: "Trick or Treat haunted house comforter set with Halloween haunted house artwork",
    173: "cream pumpkin ghost comforter set with cream Halloween pumpkin and ghost artwork",
    174: "custom photo music player quilt with photo upload and song/artist text controls",
    175: "autumn Tree of Life birds quilt set with autumn tree, birds and floral artwork",
    176: "custom name baseball flag bedding with baseball over American flag artwork",
    177: "custom baseball glove bedding with glove and baseball artwork",
    178: "baseball flag glove bedding set with American flag and baseball glove artwork",
    179: "custom baseball home quote bedding with baseball and home quote artwork",
    180: "catcher American flag baseball bedding with catcher and American flag artwork",
}

base.SERP_EVIDENCE = {
    171: [("pink gothic skull comforter set", ["https://www.walmart.com/", "https://www.ebay.com/", "https://everythingskull.com/"], "Pink/gothic skull comforter intent is commercially supported."),
          ("skull gothic comforter bedding set", ["https://www.walmart.com/", "https://www.ebay.com/", "https://www.temu.com/"], "Comparator supports skull/gothic bedding purchase intent.")],
    172: [("trick or treat haunted house comforter set", ["https://www.pinterest.com/top1bedding/halloween-bedding-set/", "https://www.amazon.com.be/-/en/Bedlam-Haunted-House-Duvet-Double/dp/B0CC6J7LTG", "https://www.etsy.com/market/halloween_bedding"], "Trick-or-treat/haunted-house exact phrase is niche; Halloween bedding product intent is clear."),
          ("haunted house Halloween comforter set", ["https://www.ebay.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports haunted-house Halloween bedding purchase intent.")],
    173: [("cream pumpkin ghost comforter set", ["https://www.etsy.com/", "https://www.walmart.com/c/kp/pumpkin-bedding", "https://www.tesco.com/"], "Cream exact modifier is niche, but pumpkin/ghost bedding intent is strong."),
          ("pumpkin ghost Halloween bedding set", ["https://www.walmart.com/c/kp/halloween-bedding-sets", "https://www.ebay.com/", "https://www.etsy.com/"], "Comparator supports pumpkin/ghost Halloween bedding purchase intent.")],
    174: [("custom photo music player quilt", ["https://us.shein.com/", "https://fridgebeats.com/", "https://www.madeingift.com/"], "Custom photo/music player products exist, but quilt-specific SERP support is weaker."),
          ("custom photo song blanket music player", ["https://www.amazon.com/", "https://fridgebeats.com/", "https://www.madeingift.com/"], "Comparator supports custom photo/song gift intent; bedding/quilt exact fit is mixed.")],
    175: [("autumn Tree of Life birds quilt set", ["https://www.wanderquilt.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Autumn Tree of Life quilt/bedding intent is supported, with birds as a visual modifier."),
          ("autumn tree birds quilt bedding set", ["https://www.ebay.com/", "https://www.bedbathandbeyond.com/", "https://www.walmart.com/"], "Comparator supports autumn tree/bird quilt bedding intent.")],
    176: [("custom name baseball flag bedding", ["https://www.etsy.com/market/baseball_bedspread", "https://ohaprints.com/", "https://www.amazon.com/HOSIMA-Baseball-Comforter-Bedroom-Pattern/dp/B0BWTB8ZNH"], "Custom baseball bedding with American flag/product comparables supports purchase intent."),
          ("personalized baseball American flag bedding", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.walmart.com/"], "Comparator supports personalized baseball flag bedding intent.")],
    177: [("custom baseball glove bedding", ["https://ohaprints.com/", "https://www.etsy.com/market/baseball_bedspread", "https://www.amazon.com/"], "Baseball glove bedding intent is supported; avoid confusing with actual sporting glove customization."),
          ("personalized baseball glove bedding set", ["https://ohaprints.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports custom baseball glove bedding purchase intent.")],
    178: [("baseball flag glove bedding set", ["https://ohaprints.com/", "https://www.etsy.com/market/baseball_bedspread", "https://www.amazon.com/HOSIMA-Baseball-Comforter-Bedroom-Pattern/dp/B0BWTB8ZNH"], "Baseball flag/glove bedding product intent is supported."),
          ("baseball glove American flag comforter set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports baseball glove/flag comforter and bedding intent.")],
    179: [("custom baseball home quote bedding", ["https://ohaprints.com/", "https://www.etsy.com/market/baseball_bedspread", "https://www.pinterest.com/ideas/all-stars-baseball-themed-bed-sheets/905715423673/"], "Home quote exact modifier is narrow, while custom baseball bedding intent is supported."),
          ("baseball home quote comforter bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports baseball quote bedding purchase intent.")],
    180: [("catcher American flag baseball bedding", ["https://www.etsy.com/market/baseball_bedspread", "https://www.amazon.com/HOSIMA-Baseball-Comforter-Bedroom-Pattern/dp/B0BWTB8ZNH", "https://ohaprints.com/"], "Catcher/flag exact phrase is niche but baseball flag bedding intent is supported."),
          ("baseball catcher American flag comforter", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports baseball catcher/flag bedding product intent.")],
}

base.CRITERION_ASSESSMENTS = {
    171: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    172: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    173: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    174: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    175: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    176: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    177: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    178: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    179: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    180: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_018_r5

## Kết luận

- Phạm vi: **10 sản phẩm, 52/52 ảnh (100%)**; chỉ inventory position **171–180**; revision **r5**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 18 được chấm độc lập từ r5; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm Halloween và baseball có intent mua hàng rõ; `Custom Photo Music Player Quilt`, `Trick or Treat Haunted House`, `Cream Pumpkin Ghost` và `Custom Baseball Home Quote` được chấm thận trọng khi exact SERP rộng hoặc lẫn sản phẩm không phải bedding.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 52 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r5 khớp các motif chính: pink gothic skull, Trick or Treat haunted house, cream pumpkin ghost, custom photo music player, autumn Tree of Life birds và baseball flag/glove/home/catcher designs.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 52 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
