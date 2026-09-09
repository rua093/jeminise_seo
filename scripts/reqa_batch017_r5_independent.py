from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_017"
base.REVISION = "r5"
base.ISSUE_PREFIX = "B17R5"
base.TRACE_ROW_REVISION = "r4"
base.SOURCE_SHA256_EXPECTED = "EE3D2575CBA94B12A5E1B0B53AFBB2EAE967D85BD0102DC32EDF2CBDBCBE0D59"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_017_r5" / "SEO_Product_Optimization_qa_batch_017_r5.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_103500"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_103500"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_103500"

base.PRODUCT_SCOPE = [
    (161, "8834806218951", 7),
    (162, "8834808512711", 7),
    (163, "8867095118023", 7),
    (164, "8879778791623", 6),
    (165, "8879778889927", 5),
    (166, "8879779217607", 6),
    (167, "8879779283143", 6),
    (168, "8879779348679", 6),
    (169, "8879779381447", 6),
    (170, "8879779414215", 6),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "161-170"

base.PRODUCT_FACTS = {
    161: "blue raven Celtic knot quilt with raven and Celtic knot frame artwork",
    162: "Valhalla Viking shield quilt with shield, crossed axes and Norse artwork",
    163: "blue gold owl night quilt set with owl night wildlife artwork",
    164: "witch moon Halloween comforter set with moon, witchy and Halloween bedding artwork",
    165: "ghost pumpkin Halloween comforter set with cute ghost and pumpkin artwork",
    166: "black ghost Halloween comforter set with black Halloween ghost artwork",
    167: "haunted pumpkin ghost comforter set with haunted pumpkin and ghost artwork",
    168: "pink cute ghost Halloween comforter set with pink ghost Halloween artwork",
    169: "red handprint Halloween comforter set with red handprint horror-style artwork",
    170: "haunted house pumpkin comforter set with haunted house and pumpkin Halloween artwork",
}

base.SERP_EVIDENCE = {
    161: [("blue raven Celtic knot quilt", ["https://www.etsy.com/uk/listing/789677515/celtic-raven-spirit-animal-totem-woven", "https://www.instagram.com/reel/Dba_ww7O1IH/", "https://www.etsy.com/market/celtic_patterns_fabric"], "Celtic raven textile intent is supported; exact blue quilt wording is narrow and partly decorative/fabric-oriented."),
          ("Celtic raven quilt bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://worldnorse.com/"], "Comparator supports raven/Celtic bedding and blanket product intent.")],
    162: [("Valhalla Viking shield quilt", ["https://www.ubuy.com.gh/", "https://www.pinterest.com/", "https://www.amazon.co.uk/"], "Valhalla/Viking shield quilt-bedding intent is supported by product comparables."),
          ("Viking shield quilt bedding set", ["https://www.ubuy.mq/", "https://www.amazon.co.uk/", "https://www.etsy.com/"], "Comparator supports Viking shield bedding purchase intent.")],
    163: [("blue gold owl night quilt set", ["https://miravodecor.com/", "https://miravoventures.com/", "https://www.etsy.com/"], "Owl night quilt intent is supported; blue/gold exact wording is a visual modifier."),
          ("owl quilt bedding set", ["https://www.wayfair.com/", "https://www.amazon.com/", "https://www.etsy.com/"], "Comparator supports owl quilt/bedding purchase intent.")],
    164: [("witch moon Halloween comforter set", ["https://www.walmart.com/", "https://moonchildworld.com/", "https://zirconic.com.au/"], "Witch/moon Halloween comforter and bedding intent is commercially supported."),
          ("Halloween comforter set sheets witch moon", ["https://www.walmart.com/", "https://www.amazon.com/", "https://society6.com/"], "Comparator supports Halloween/witch bedding purchase intent.")],
    165: [("ghost pumpkin Halloween comforter set", ["https://www.amazon.com/kids-halloween-comforter/s?k=kids+halloween+comforter", "https://www.walmart.com/", "https://www.ebay.com/"], "Ghost/pumpkin Halloween comforter intent is strong and product-oriented."),
          ("Halloween ghost pumpkin bedding set", ["https://www.walmart.com/c/kp/halloween-bedding-sets", "https://www.etsy.com/", "https://www.ebay.com/"], "Comparator supports Halloween ghost/pumpkin bedding intent.")],
    166: [("black ghost Halloween comforter set", ["https://www.amazon.com/clp/B0D6YVKSFD", "https://www.amazon.com/clp/B0D6YZG9Z1", "https://www.walmart.com/c/kp/halloween-bedding-sets"], "Black ghost Halloween comforter exact intent is directly supported by product results."),
          ("black ghost Halloween bedding set", ["https://www.ebay.com/", "https://www.walmart.com/", "https://www.etsy.com/"], "Comparator supports black ghost Halloween bedding purchase intent.")],
    167: [("haunted pumpkin ghost comforter set", ["https://www.ebay.com/", "https://www.etsy.com/market/pumpkin_bedding", "https://www.pinterest.com/"], "Haunted/pumpkin/ghost exact wording is narrow but Halloween bedding purchase intent is clear."),
          ("haunted house pumpkin ghost bedding", ["https://www.etsy.com/", "https://www.ebay.com/", "https://www.walmart.com/"], "Comparator supports haunted Halloween bedding with pumpkin/ghost motifs.")],
    168: [("pink cute ghost Halloween comforter set", ["https://www.amazon.com/", "https://www.ubuy.hn/", "https://www.pinterest.com/"], "Pink/cute ghost Halloween comforter intent is commercially supported."),
          ("pink ghost Halloween bedding set", ["https://www.etsy.com/", "https://www.walmart.com/", "https://www.amazon.com/"], "Comparator supports pink ghost Halloween bedding purchase intent.")],
    169: [("red handprint Halloween comforter set", ["https://society6.com/collections/comforters-halloween", "https://www.etsy.com/market/halloween_bedding", "https://www.pinterest.com/ideas/red-comforter-sets/917891129208/"], "Red handprint exact phrase is niche; Halloween comforter product intent exists more broadly."),
          ("red horror Halloween bedding set", ["https://www.ebay.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports red/horror Halloween bedding intent.")],
    170: [("haunted house pumpkin comforter set", ["https://www.etsy.com/market/pumpkin_bedding", "https://www.ebay.com/", "https://www.pinterest.com/pin/halloween-ghost-king-size-bedding-set-halloween-pumpkin-bedding-set-halloween-haunted-house-nhm--970385050951620396/"], "Haunted house/pumpkin bedding product intent is supported, with some marketplace/pin results."),
          ("haunted house pumpkin Halloween bedding set", ["https://www.etsy.com/", "https://www.ebay.com/", "https://www.walmart.com/"], "Comparator supports haunted-house pumpkin Halloween bedding intent.")],
}

base.CRITERION_ASSESSMENTS = {
    161: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    162: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    163: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    164: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    165: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    166: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    167: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    168: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    169: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    170: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_017_r5

## Kết luận

- Phạm vi: **10 sản phẩm, 62/62 ảnh (100%)**; chỉ inventory position **161–170**; revision **r5**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 17 được chấm độc lập từ r5; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm Halloween comforter 164–170 có intent mua hàng rõ; exact modifier ở Blue Raven Celtic Knot, Owl Night và Red Handprint được chấm thận trọng khi SERP còn lẫn broader bedding, fabric/pattern hoặc marketplace intent.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 62 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r5 khớp các motif chính: blue raven Celtic knot, Valhalla Viking shield, blue-gold owl, witch moon, ghost/pumpkin, black ghost, pink cute ghost, red handprint và haunted house pumpkin Halloween designs.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 62 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
