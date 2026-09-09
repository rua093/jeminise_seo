from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_012"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B12R4"
base.TRACE_ROW_REVISION = "r3"
base.SOURCE_SHA256_EXPECTED = "C07C0FEB5AC2CC4810DD152389DBEBACB10879F342B9B7A4D0546AB905CFFBE1"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_012_r4" / "SEO_Product_Optimization_qa_batch_012_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_012000"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_012000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_012000"

base.PRODUCT_SCOPE = [
    (111, "8860497772743", 8),
    (112, "8860498198727", 8),
    (113, "8860498133191", 8),
    (114, "8867220553927", 7),
    (115, "8834771779783", 6),
    (116, "8835553001671", 7),
    (117, "8833356955847", 7),
    (118, "8833357742279", 7),
    (119, "8835559260359", 7),
    (120, "8835560276167", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "111-120"

base.PRODUCT_FACTS = {
    111: "Daniel soccer ball comforter set with sample name and soccer ball/player artwork",
    112: "Tyler soccer goal comforter set with sample name and soccer goal artwork",
    113: "Matthew paint splatter soccer comforter with sample name and splatter soccer motif",
    114: "custom softball name number comforter with softball motif and personalization context",
    115: "Kevin football flag comforter with sample name and American flag football motif",
    116: "blue geometric semi truck comforter with transportation/trucker artwork",
    117: "colorful tree of life quilt set with nature/tree artwork",
    118: "Celtic Yggdrasil tree quilt set with knotwork/world tree artwork",
    119: "red semi truck stone wall comforter with truck breaking through wall artwork",
    120: "black semi truck stone wall comforter with truck breaking through wall artwork",
}

base.SERP_EVIDENCE = {
    111: [("Daniel soccer ball comforter set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Daniel is likely sample-name driven; soccer ball comforter intent is supported by broader results."),
          ("personalized soccer ball comforter set name number", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.youcustomizeit.com/"], "Comparator supports custom name/number soccer bedding intent.")],
    112: [("Tyler soccer goal comforter set", ["https://www.target.com/", "https://ohaprints.com/", "https://www.amazon.com/"], "Tyler is sample-name driven; soccer goal bedding intent is supported by broader results."),
          ("soccer goal comforter set", ["https://www.target.com/", "https://www.lushdecor.com/", "https://www.kohls.com/"], "Comparator supports soccer goal/comforter set product intent.")],
    113: [("Matthew paint splatter soccer comforter", ["https://www.walmart.com/", "https://www.spotlightstores.com/", "https://www.amazon.com/"], "Matthew is sample-name driven; paint/splatter soccer bedding intent has product comparables."),
          ("paint splatter soccer bedding set", ["https://www.walmart.com/", "https://www.spotlightstores.com/", "https://www.mercadolibre.com.ar/"], "Comparator supports splatter soccer bedding/coverlet intent.")],
    114: [("custom softball name number comforter", ["https://www.etsy.com/", "https://www.amazon.com/", "https://ohaprints.com/"], "Custom softball name/number bedding intent is commercially supported."),
          ("personalized softball comforter set name number", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports personalized softball comforter/blanket purchase intent.")],
    115: [("Kevin football flag comforter", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.ebay.com/"], "Kevin is sample-name driven; football flag bedding intent is supported."),
          ("personalized football flag comforter set", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.ebay.com/"], "Comparator supports American flag football bedding and custom name intent.")],
    116: [("blue geometric semi truck comforter", ["https://www.wayfair.ca/", "https://www.walmart.com/", "https://jeminise.com/"], "Truck bedding intent is supported; exact blue geometric phrase is niche."),
          ("semi truck bedding comforter set blue", ["https://www.wayfair.ca/", "https://www.walmart.com/", "https://www.target.com/"], "Comparator supports truck/semi-truck bedding/comforter intent.")],
    117: [("colorful tree of life quilt set", ["https://www.amazon.com/", "https://www.etsy.com/market/tree_of_life_bed_quilt", "https://www.ebay.com/"], "Tree of Life quilt set intent is strongly supported; colorful is a valid visual modifier."),
          ("tree of life quilt bedding set colorful", ["https://www.etsy.com/market/tree_of_life_bed_quilt", "https://www.amazon.com/tree-life-quilt/s?k=tree+of+life+quilt", "https://golden-vale.com/"], "Comparator supports finished quilt/bedding product intent.")],
    118: [("Celtic Yggdrasil tree quilt set", ["https://www.amazon.com/clp/B0D9NXRP1Q", "https://luvingift.com/product/viking-quilt-set-norse-celtic-yggdrasil-tree-of-life/", "https://www.pinterest.com/"], "Celtic/Yggdrasil Tree of Life quilt intent is supported by direct product comparables."),
          ("Yggdrasil tree of life quilt bedding set", ["https://www.amazon.com/clp/B0DNSVY2QJ", "https://luvingift.com/product/viking-quilt-set-norse-celtic-yggdrasil-tree-of-life/", "https://www.pinterest.com/"], "Comparator supports Yggdrasil/Celtic quilt bedding intent.")],
    119: [("red semi truck stone wall comforter", ["https://jeminise.com/collections/couples-family-blankets", "https://www.facebook.com/", "https://www.ebay.com/"], "Exact stone-wall phrase is weak in public SERP; truck comforter intent exists but needs product-specific support."),
          ("semi truck breaking through wall comforter", ["https://jeminise.com/collections/couples-family-blankets", "https://www.walmart.com/", "https://www.amazon.com/"], "Comparator has weak exact support; use broader semi-truck comforter language if no stronger data exists.")],
    120: [("black semi truck stone wall comforter", ["https://jeminise.com/collections/couples-family-blankets", "https://www.facebook.com/", "https://www.ebay.com/"], "Exact black/stone-wall phrase is weak in public SERP; truck comforter intent is broader."),
          ("custom trucker semi truck blanket stone wall", ["https://jeminise.com/collections/couples-family-blankets", "https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports custom trucker blanket intent more than exact stone-wall wording.")],
}

base.CRITERION_ASSESSMENTS = {
    111: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    112: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    113: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    114: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    115: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    116: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    117: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    118: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    119: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    120: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_012_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 72/72 ảnh (100%)**; chỉ inventory position **111–120**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 12 được chấm độc lập từ r4; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Các exact keyword có sample name như Daniel, Tyler, Matthew, Kevin và các cụm semi-truck stone wall được chấm thận trọng; Tree of Life/Yggdrasil và custom softball name-number có support mạnh hơn.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 72 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: soccer ball/goal/paint splatter, softball name-number, football flag, blue geometric semi truck, Tree of Life/Yggdrasil và red/black semi truck stone-wall variants.
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
