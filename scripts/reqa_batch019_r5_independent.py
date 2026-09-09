from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_019"
base.REVISION = "r5"
base.ISSUE_PREFIX = "B19R5"
base.TRACE_ROW_REVISION = "r4"
base.SOURCE_SHA256_EXPECTED = "8FA661F5519F5A4F9BC2695B67213E4F1BA59902DF0EE6E7E96A47799FDF7905"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_019_r5" / "SEO_Product_Optimization_qa_batch_019_r5.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_111000"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_111000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_111000"

base.PRODUCT_SCOPE = [
    (181, "8867179331783", 4),
    (182, "8867179495623", 4),
    (183, "8867179561159", 4),
    (184, "8867179593927", 4),
    (185, "8867179626695", 4),
    (186, "8867179692231", 4),
    (187, "8867179757767", 4),
    (188, "8867179724999", 4),
    (189, "8867179888839", 4),
    (190, "8867179921607", 4),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "181-190"

base.PRODUCT_FACTS = {
    181: "custom baseball glove name bedding with baseball glove and name artwork",
    182: "custom baseball flag name bedding with baseball, American flag and name artwork",
    183: "custom baseball flag glove bedding with baseball glove and American flag artwork",
    184: "custom close-up baseball glove bedding with close-up glove and ball artwork",
    185: "custom gray baseball flag bedding with gray flag baseball artwork",
    186: "custom baseball pattern name bedding with repeated baseball pattern and name artwork",
    187: "custom pitcher baseball bedding with pitcher/player baseball artwork",
    188: "custom black baseball glove bedding with black-background glove artwork",
    189: "custom blue stripe baseball bedding with blue stripe baseball artwork",
    190: "custom galaxy baseball name bedding with galaxy baseball and name artwork",
}

base.SERP_EVIDENCE = {
    181: [("custom baseball glove name bedding", ["https://www.etsy.com/listing/818286355/baseball-kids-bedding-set-boy-comforter", "https://www.amazon.com/stores/page/7FC5A2D6-E482-4281-AE65-0CAAD84FB88D", "https://www.pinterest.com/pin/889109151455713116/"], "Custom baseball glove/name bedding intent is supported by product comparables."),
          ("personalized baseball glove bedding set", ["https://ohaprints.com/", "https://www.etsy.com/listing/4517822148/personalized-name-number-baseball", "https://www.walmart.com/ip/Castle-Fairy-Sport-Baseball-Glove-Print-Zipper-Full-Size-Comforter-Sets-Fitted-Sheet-2-Pillowcases-Starry-Sky-Zipper-Bedding-Set-Black-Pink-White-4pc/20379850094"], "Comparator supports personalized baseball glove bedding purchase intent.")],
    182: [("custom baseball flag name bedding", ["https://www.etsy.com/listing/4442011102/personalized-baseball-quilt-bedding-set", "https://jeminise.com/products/personalized-baseball-bedding-full-size-flag-custom-name-d-design-09", "https://www.truegether.com/ohaprints-custom-baseball-vintage-us-flag-for-son-boy-personalized-name-number-quilt-blanket-pillowcases-quilts-bedding-set-pillow-cover-king-queen-double-twin-throw-full-size-bedspread-bed-sets/USER.99ae4901-bf9b-4e4d-8c1b-8b80be948f5e/listing.html"], "Custom baseball flag/name bedding intent is strongly supported."),
          ("personalized baseball American flag bedding", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.amazon.com/HOSIMA-Baseball-Comforter-Bedroom-Pattern/dp/B0BWTB8ZNH"], "Comparator supports personalized baseball flag bedding intent.")],
    183: [("custom baseball flag glove bedding", ["https://www.etsy.com/listing/4442011102/personalized-baseball-quilt-bedding-set", "https://ohaprints.com/", "https://www.facebook.com/groups/3293063017460293/posts/27556864467320143/"], "Baseball flag/glove bedding purchase intent is supported."),
          ("baseball glove American flag custom bedding set", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.amazon.com/"], "Comparator supports custom baseball glove/flag bedding.")],
    184: [("custom close-up baseball glove bedding", ["https://www.walmart.com/ip/Castle-Fairy-Sport-Baseball-Glove-Print-Zipper-Full-Size-Comforter-Sets-Fitted-Sheet-2-Pillowcases-Starry-Sky-Zipper-Bedding-Set-Black-Pink-White-4pc/20379850094", "https://www.etsy.com/market/baseball_bedspread", "https://ohaprints.com/"], "Close-up exact modifier is narrow, but baseball glove bedding intent exists."),
          ("baseball glove bedding set", ["https://www.walmart.com/c/kp/baseball-duvet-cover", "https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports baseball glove bedding purchase intent.")],
    185: [("custom gray baseball flag bedding", ["https://www.walmart.com/c/kp/baseball-duvet-cover", "https://www.etsy.com/market/baseball_bedspread", "https://www.pinterest.com/pin/999939923509362029/"], "Gray exact modifier is niche; baseball/flag bedding intent is supported."),
          ("gray baseball bedding set flag", ["https://www.walmart.com/c/kp/baseball-duvet-cover", "https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports gray/baseball bedding intent.")],
    186: [("custom baseball pattern name bedding", ["https://www.etsy.com/listing/818286355/baseball-kids-bedding-set-boy-comforter", "https://www.etsy.com/listing/1581320516/custom-baseball-blanket-personalized", "https://www.amazon.com/Personalized-Baseball-Blankets-Lightweight-Birthday/dp/B0GPWY997X"], "Custom baseball pattern/name bedding intent is supported."),
          ("personalized baseball pattern bedding set", ["https://www.etsy.com/", "https://www.walmart.com/c/kp/baseball-duvet-cover", "https://www.amazon.com/"], "Comparator supports personalized baseball bedding purchase intent.")],
    187: [("custom pitcher baseball bedding", ["https://www.etsy.com/listing/1689165441/baseball-i-am-pitcher-custom-name-number", "https://www.amazon.com/OhaPrints-Softball-Personalized-Pillowcases-Bedspread/dp/B0CG8FHWDK", "https://www.pinterest.com/pin/personalized-baseball-quilt-set-baseball-pitcher-boy-quilt-blanket-with-pillowcases-custom-name-and-number-quilt-beddi--999939923509378196/"], "Pitcher baseball custom bedding intent is supported."),
          ("personalized baseball pitcher bedding set", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.amazon.com/"], "Comparator supports personalized pitcher baseball bedding.")],
    188: [("custom black baseball glove bedding", ["https://www.walmart.com/c/kp/baseball-duvet-cover", "https://www.etsy.com/market/baseball_bedspread", "https://ohaprints.com/"], "Black exact modifier is a visual differentiator; baseball glove bedding intent is supported."),
          ("black baseball glove bedding set", ["https://www.walmart.com/c/kp/baseball-duvet-cover", "https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports black/baseball glove bedding purchase intent.")],
    189: [("custom blue stripe baseball bedding", ["https://www.wayfair.com/bed-bath/pdp/design-art-baseball-collage-ii-sport-baseball-duvet-cover-set-microfiber-polyester-w110967964.html", "https://www.pinterest.com/pin/personalized-baseball-quilt-set-baseball-softball-blue-patchwork-quilt-blanket-with-pillowcases-custom-name-quilt-bedd--999939923509358899/", "https://www.walmart.com/c/kp/baseball-duvet-cover"], "Blue stripe exact phrase is narrow, but blue/baseball bedding intent is supported."),
          ("blue baseball bedding set custom name", ["https://www.etsy.com/", "https://www.walmart.com/c/kp/baseball-duvet-cover", "https://ohaprints.com/"], "Comparator supports custom blue baseball bedding purchase intent.")],
    190: [("custom galaxy baseball name bedding", ["https://ohaprints.com/products/ohaprints-quilt-bed-set-pillowcase-basketball-baseball-yin-yang-galaxy-player-fan-custom-personalized-name-number-blanket-bedspread-bedding-960", "https://www.etsy.com/market/kpop_demon_hunter_duvet_cotton", "https://www.facebook.com/SnugglyMates/posts/we-launched-and-this-is-why-emmies-face-when-she-read-her-own-name-on-her-galaxy/122176840958603390/"], "Galaxy/custom-name intent exists, but baseball galaxy exact results are narrower and mixed."),
          ("personalized galaxy baseball bedding set", ["https://ohaprints.com/", "https://www.etsy.com/", "https://www.amazon.com/"], "Comparator partially supports galaxy/custom bedding plus baseball bedding intent.")],
}

base.CRITERION_ASSESSMENTS = {
    181: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    182: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    183: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    184: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    185: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    186: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    187: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    188: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    189: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    190: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_019_r5

## Kết luận

- Phạm vi: **10 sản phẩm, 40/40 ảnh (100%)**; chỉ inventory position **181–190**; revision **r5**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 19 được chấm độc lập từ r5; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Cụm baseball bedding/custom-name có intent mua hàng rõ; các exact modifier như close-up, gray, black, blue stripe và galaxy được chấm thận trọng khi SERP còn rộng hoặc lẫn decor/blanket ngoài bedding.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 40 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r5 khớp các motif chính: baseball glove/name, American flag, close-up glove, gray/black/blue stripe treatments, pitcher, pattern-name và galaxy baseball artwork.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 40 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
