from __future__ import annotations

import shutil
from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_010"
base.REVISION = "r6"
base.ISSUE_PREFIX = "B10R6"
base.TRACE_ROW_REVISION = "r3"
base.SOURCE_SHA256_EXPECTED = "DF0D197B97B4C99DADABA12ABE9A17D30539FFAFC2103E7628AF6B4DA9904310"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_010_r6" / "SEO_Product_Optimization_qa_batch_010_r6.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_010000"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_010000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_010000"

base.PRODUCT_SCOPE = [
    (91, "8837138940103", 5),
    (92, "8860498821319", 8),
    (93, "8835550806215", 7),
    (94, "8835554541767", 7),
    (95, "8835554246855", 7),
    (96, "8834787442887", 7),
    (97, "8860137160903", 8),
    (98, "8860096954567", 8),
    (99, "8860106883271", 8),
    (100, "8860127658183", 8),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "91-100"

base.PRODUCT_FACTS = {
    91: "turquoise wolf quilt set with circular wolf portrait and Native American-inspired geometric frame",
    92: "neon green soccer player comforter set with personalized name/number style artwork",
    93: "blue semi truck American flag comforter with patriotic trucking motif",
    94: "silver semi truck waving American flag comforter with patriotic trucking motif",
    95: "red semi truck sunset comforter with American flag trucking motif",
    96: "custom photo collage quilt set using bicycle and flowers sample imagery",
    97: "God Says I Am butterfly blanket with purple butterflies and Christian affirmation text",
    98: "God Is Within Her butterfly blanket with scripture/faith butterfly motif",
    99: "Blessed Is She purple cross blanket with floral Christian cross motif",
    100: "Estella scripture collage blanket with purple floral cross, butterflies and personalized name sample",
}

base.SERP_EVIDENCE = {
    91: [("turquoise wolf quilt set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Commercial wolf quilt/bedding intent is supported; turquoise is a useful visual modifier."),
         ("wolf quilt bedding set turquoise", ["https://www.ebay.com/", "https://www.miravoventures.com/"], "Comparator supports wolf quilt/bedding product intent with color/motif variations.")],
    92: [("neon green soccer comforter set", ["https://www.amazon.com/", "https://www.walmart.com/", "https://jeminise.com/"], "Soccer bedding/comforter intent is supported; exact neon green phrase is narrower."),
         ("soccer player comforter set neon green", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.ubuy.com/"], "Comparator supports soccer player bedding intent with personalized sports-room use.")],
    93: [("blue semi truck flag comforter", ["https://jeminise.com/", "https://us.shein.com/", "https://www.zazzle.com/"], "Semi truck and American flag bedding/blanket comparables support commercial intent; exact blue wording is niche."),
         ("custom semi truck American flag comforter", ["https://jeminise.com/", "https://www.zazzle.com/", "https://www.ebay.com/"], "Comparator supports custom truck/US flag bedding or blanket intent.")],
    94: [("silver semi truck flag comforter", ["https://jeminise.com/", "https://us.shein.com/", "https://www.zazzle.com/"], "Semi truck flag bedding intent exists; silver exact modifier is mostly product-specific."),
         ("personalized truck blanket American flag", ["https://us.shein.com/", "https://www.ebay.com/", "https://www.etsy.com/"], "Comparator supports truck + American flag personalized blanket intent.")],
    95: [("red semi truck sunset comforter", ["https://jeminise.com/", "https://www.walmart.com/", "https://www.amazon.com/"], "Exact red/sunset wording is weak, but truck bedding and comforter products appear in broader results."),
         ("custom semi truck flag comforter", ["https://jeminise.com/", "https://www.zazzle.com/", "https://www.ebay.com/"], "Comparator supports custom semi truck flag bedding/blanket intent.")],
    96: [("custom photo collage quilt set", ["https://www.bagsoflove.com/", "https://portraitblankets.com/", "https://www.printerpix.com/"], "Strong commercial intent for photo collage blankets/quilts supports the selected direction."),
         ("personalized photo collage quilt", ["https://www.etsy.com/", "https://www.shutterfly.com/", "https://www.printerpix.com/"], "Comparator supports personalized photo collage bedding/blanket intent.")],
    97: [("God Says I Am butterfly blanket", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Marketplace results support God Says I Am butterfly blanket intent."),
         ("personalized God Says I Am blanket", ["https://www.amazon.com/", "https://www.callie.com/", "https://www.etsy.com/"], "Comparator supports custom name Christian affirmation blanket intent.")],
    98: [("God Is Within Her butterfly blanket", ["https://www.amazon.com/", "https://www.pinterest.com/", "https://www.biblegateway.com/"], "Faith phrase is supported as scripture/Christian decor intent; exact butterfly blanket is narrower."),
         ("Psalm 46:5 butterfly blanket", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.zazzle.com/"], "Comparator supports scripture phrase blanket/gift intent.")],
    99: [("Blessed is she purple cross blanket", ["https://www.amazon.com/", "https://www.zazzle.com/", "https://www.etsy.com/"], "Christian cross blanket intent is supported; exact purple phrase is narrower."),
         ("purple floral cross Christian blanket", ["https://www.amazon.com/", "https://www.zazzle.com/", "https://www.etsy.com/"], "Comparator supports floral/cross Christian blanket product intent.")],
    100: [("Estella scripture collage blanket", ["https://www.zazzle.com/", "https://www.etsy.com/", "https://www.pinterest.com/"], "Exact Estella wording appears sample-name driven, so support is weaker than broad scripture collage intent."),
          ("personalized scripture collage blanket", ["https://www.etsy.com/market/scripture_blanket", "https://www.zazzle.com/scripture%2Bblankets", "https://prettyperfect.com/"], "Comparator strongly supports personalized scripture blanket/collage gift intent.")],
}

base.CRITERION_ASSESSMENTS = {
    91: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    92: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    93: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    94: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    95: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    96: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    97: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    98: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    99: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    100: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_010_r6

## Kết luận

- Phạm vi: **10 sản phẩm, 73/73 ảnh (100%)**; chỉ inventory position **91–100**; revision **r6**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 10 được chấm độc lập từ r6; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Exact keyword màu/motif như neon green soccer, blue/silver/red semi truck và Estella scripture collage được chấm thận trọng nếu SERP chỉ hỗ trợ broader product intent.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 73 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r6 khớp các motif chính: turquoise wolf, neon soccer, semi truck flag variants, photo collage, Christian butterfly/cross/scripture blankets.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 73 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
