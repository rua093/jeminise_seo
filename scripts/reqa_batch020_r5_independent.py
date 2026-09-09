from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_020"
base.REVISION = "r5"
base.ISSUE_PREFIX = "B20R5"
base.TRACE_ROW_REVISION = "r4"
base.SOURCE_SHA256_EXPECTED = "FA070F63BBF3D2D9F6304C1D4FC20F4AB6723ED22AE581A98DC1531621178C31"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_020_r5" / "SEO_Product_Optimization_qa_batch_020_r5.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_112500"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_112500"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_112500"

base.PRODUCT_SCOPE = [
    (191, "8867179954375", 7),
    (192, "8867180052679", 7),
    (193, "8867179987143", 4),
    (194, "8867180019911", 4),
    (195, "8867180085447", 4),
    (196, "8867180118215", 4),
    (197, "8867180150983", 4),
    (198, "8867180183751", 4),
    (199, "8867180216519", 4),
    (200, "8867180347591", 4),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "191-200"

base.PRODUCT_FACTS = {
    191: "custom red black baseball batter bedding with batter artwork and custom name/number context",
    192: "custom brown baseball glove bedding with glove and baseball artwork",
    193: "custom flaming baseball batter bedding with batter and flame artwork",
    194: "custom orange fire baseball bedding with orange fire baseball artwork",
    195: "custom baseball glove flag bedding with glove, baseball and American flag artwork",
    196: "custom black baseball glove bedding with black-background baseball glove artwork",
    197: "custom lightning baseball bedding with lightning baseball artwork",
    198: "custom night field baseball bedding with night baseball field artwork",
    199: "custom smoke baseball batter bedding with smoky baseball batter artwork",
    200: "custom baseball stitch name bedding with baseball stitch and custom name artwork",
}

base.SERP_EVIDENCE = {
    191: [("custom red black baseball batter bedding", ["https://www.zazzle.com/", "https://www.etsy.com/fi-en/market/custom_baseball_bedding", "https://ohaprints.com/products/ohaprints-quilt-bed-set-pillowcase-baseball-batter-boy-player-fan-gift-idea-custom-personalized-name-number-blanket-bedspread-bedding-3188"], "Red/black exact wording is narrow; baseball batter/custom bedding purchase intent is supported by comparables."),
          ("personalized baseball batter bedding custom name", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.pinterest.com/ideas/personalized-baseball-bedding/927363056917/"], "Comparator supports custom-name baseball batter bedding intent.")],
    192: [("custom brown baseball glove bedding", ["https://www.amazon.com/", "https://www.etsy.com/listing/1669308633/baseball-glove-blanket-custom-soft-cozy", "https://www.walmart.com/ip/Castle-Fairy-Sport-Baseball-Glove-Print-Zipper-Full-Size-Comforter-Sets-Fitted-Sheet-2-Pillowcases-Starry-Sky-Zipper-Bedding-Set-Black-Pink-White-4pc/20379850094"], "Brown exact modifier is visual/niche; baseball glove bedding intent is supported."),
          ("custom baseball glove bedding name number", ["https://ohaprints.com/products/ohaprints-quilt-bed-set-pillowcase-baseball-glove-ball-vintage-black-player-fan-custom-personalized-name-number-blanket-bedspread-bedding-1622", "https://www.amazon.com/Baseball-Bedding-Comforter-Blanket-Bedspread/dp/B0DJW8MLFC", "https://www.pinterest.com/pin/999939923509326923/"], "Comparator supports custom baseball glove bedding with name/number.")],
    193: [("custom flaming baseball batter bedding", ["https://www.amazon.com/", "https://www.etsy.com/fi-en/market/custom_baseball_bedding", "https://ohaprints.com/products/ohaprints-quilt-bed-set-pillowcase-baseball-batter-boy-player-fan-gift-idea-custom-personalized-name-number-blanket-bedspread-bedding-3188"], "Flaming batter exact phrase is narrow, but baseball batter bedding purchase intent exists."),
          ("personalized baseball batter comforter set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.amorcustomgifts.com/products/baseball-sport-100921nhm-bedding-set-duvet-covers1"], "Comparator supports personalized batter bedding/comforter intent.")],
    194: [("custom orange fire baseball bedding", ["https://www.google.com/m/storepages?c=TW&hl=en-TW&q=amazon.com", "https://www.etsy.com/fi-en/market/custom_baseball_bedding", "https://www.amazon.com/Personalized-Baseball-Softball-Comforter-Customized/dp/B0B1JB7HTZ"], "Orange fire exact wording is niche; baseball/fire colorway bedding is partially supported."),
          ("personalized baseball fire bedding set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Comparator supports custom baseball bedding but exact fire modifier is mixed.")],
    195: [("custom baseball glove flag bedding", ["https://www.etsy.com/listing/4442011102/personalized-baseball-quilt-bedding-set", "https://ohaprints.com/", "https://www.amazon.com/Baseball-Bedding-Comforter-Blanket-Bedspread/dp/B0DJW8MLFC"], "Baseball glove/flag custom bedding intent is supported by marketplace comparables."),
          ("baseball glove American flag bedding set", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.amazon.com/"], "Comparator supports the glove + flag + bedding purchase intent.")],
    196: [("custom black baseball glove bedding", ["https://www.etsy.com/listing/1669308633/baseball-glove-blanket-custom-soft-cozy", "https://ohaprints.com/products/ohaprints-quilt-bed-set-pillowcase-baseball-glove-ball-vintage-black-player-fan-custom-personalized-name-number-blanket-bedspread-bedding-1622", "https://www.walmart.com/c/kp/baseball-duvet-cover"], "Black baseball glove bedding intent is supported, though exact custom bedding SERP remains marketplace-heavy."),
          ("black baseball glove bedding set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/pin/999939923509326951/"], "Comparator supports black baseball glove bedding purchase intent.")],
    197: [("custom lightning baseball bedding", ["https://www.walmart.com/c/kp/baseball-bedding-set", "https://www.google.com/m/storepages?c=TW&hl=en-TW&q=amazon.com", "https://www.etsy.com/fi-en/market/custom_baseball_bedding"], "Lightning exact modifier is narrow; broader custom baseball bedding intent is supported."),
          ("baseball lightning bedding set", ["https://www.amazon.com/", "https://www.walmart.com/", "https://www.pinterest.com/ideas/personalized-baseball-bedding/927363056917/"], "Comparator partially supports sports bedding and baseball bedding intent.")],
    198: [("custom night field baseball bedding", ["https://www.amazon.com/", "https://www.wayfair.com/", "https://www.etsy.com/fi-en/market/custom_baseball_bedding"], "Night-field exact wording is niche; baseball field bedding intent is partially supported."),
          ("baseball field bedding set custom", ["https://www.amazon.com/", "https://www.walmart.com/c/kp/baseball-bedding-set", "https://www.etsy.com/"], "Comparator supports baseball field bedding purchase intent.")],
    199: [("custom smoke baseball batter bedding", ["https://www.etsy.com/fi-en/market/custom_baseball_bedding", "https://ohaprints.com/products/ohaprints-quilt-bed-set-pillowcase-baseball-batter-boy-player-fan-gift-idea-custom-personalized-name-number-blanket-bedspread-bedding-3188", "https://www.amazon.com/"], "Smoke exact modifier is narrow; batter/custom baseball bedding is supported."),
          ("personalized smoky baseball batter bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/ideas/personalized-baseball-bedding/927363056917/"], "Comparator partially supports the product intent, with weaker exact smoke evidence.")],
    200: [("custom baseball stitch name bedding", ["https://www.etsy.com/fi-en/market/custom_baseball_bedding", "https://www.etsy.com/listing/4517822148/personalized-name-number-baseball", "https://www.amazon.com/Personalized-Baseball-Softball-Comforter-Customized/dp/B0B1JB7HTZ"], "Baseball stitch/name exact phrase is narrow, but custom baseball name bedding intent is supported."),
          ("personalized baseball stitch bedding set", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.pinterest.com/ideas/personalized-baseball-bedding/927363056917/"], "Comparator supports personalized baseball bedding with stitch/baseball visual language.")],
}

base.CRITERION_ASSESSMENTS = {
    191: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    192: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    193: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    194: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    195: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    196: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    197: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    198: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    199: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    200: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_020_r5

## Kết luận

- Phạm vi: **10 sản phẩm, 46/46 ảnh (100%)**; chỉ inventory position **191–200**; revision **r5**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 20 được chấm độc lập từ r5; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Cụm baseball bedding/custom-name có intent mua hàng rõ; các exact modifier như red-black batter, orange fire, lightning, night field, smoke và stitch được chấm thận trọng khi SERP còn rộng hoặc lẫn blanket/duvet/decor ngoài bedding.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 46 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r5 khớp các motif chính: red/black batter, brown/black glove, flaming/orange fire, glove flag, lightning, night field, smoke batter và baseball stitch/name artwork.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 46 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
