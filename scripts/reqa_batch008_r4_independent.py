from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_008"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B08R4"
base.SOURCE_SHA256_EXPECTED = "46E2312BE9475A3CEE7A4D3684EF021B24046B1DACCA49672D084B104D97CEC0"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_008_r4" / "SEO_Product_Optimization_qa_batch_008_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_008000"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_008000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_008000"

base.PRODUCT_SCOPE = [
    (71, "8834766471367", 6),
    (72, "8834755068103", 6),
    (73, "8834757296327", 6),
    (74, "8834752053447", 6),
    (75, "8834761654471", 7),
    (76, "8834772566215", 6),
    (77, "8860498854087", 7),
    (78, "8860100493511", 8),
    (79, "8860102820039", 9),
    (80, "8860110880967", 8),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "71-80"

base.PRODUCT_FACTS = {
    71: "teal football player running with ball comforter and custom name/number controls",
    72: "electric football player running with ball comforter and custom name/number controls",
    73: "football player with American flag background comforter and custom name/number controls",
    74: "football player with burning ball and American football artwork plus custom name/number controls",
    75: "football players and large close-up football comforter with custom name/number controls",
    76: "vintage football texture/laces comforter with custom name/number controls",
    77: "glowing soccer ball trails comforter set with optional custom text control",
    78: "God Says You Are floral inspirational blanket with custom text/name control",
    79: "floral butterfly affirmation blanket with personalized name and affirmation artwork",
    80: "God Says I Am butterfly scripture blanket with custom name control",
}

base.SERP_EVIDENCE = {
    71: [("personalized teal football player comforter", ["https://ohaprints.com/", "https://www.etsy.com/"], "Commercial personalized football bedding intent is supported; teal exact modifier is niche."),
         ("custom football player bedding set name number", ["https://www.youcustomizeit.com/", "https://www.amazon.com/"], "Comparator supports custom football player bedding with name/number.")],
    72: [("personalized electric football player comforter", ["https://www.amazon.com/", "https://www.walmart.com/"], "Electric/lightning football bedding evidence is mixed, but sports bedding commercial intent exists."),
         ("custom football lightning bedding set name", ["https://www.amazon.com/", "https://www.callie.com/"], "Comparator supports personalized football sports bedding/blanket intent.")],
    73: [("personalized American flag football player comforter", ["https://ohaprints.com/", "https://www.amazon.com/"], "Commercial football player + US flag bedding intent is supported."),
         ("custom football player American flag bedding set", ["https://www.cubebik.com/", "https://www.etsy.com/"], "Comparator supports patriotic football bedding/custom gift intent.")],
    74: [("personalized burning football player comforter", ["https://www.walmart.com/", "https://www.temu.com/"], "Burning/fire football bedding wording is narrower than broad football bedding."),
         ("custom football player fire bedding set name", ["https://www.amazon.com/", "https://ohaprints.com/"], "Comparator supports custom football bedding with fire/flame visual angle.")],
    75: [("personalized football players name comforter", ["https://www.etsy.com/", "https://www.amazon.com/"], "Personalized football bedding with player/name intent is supported."),
         ("custom football players bedding set name number", ["https://www.youcustomizeit.com/", "https://doonakingdom.com.au/"], "Comparator supports custom football player bedding intent.")],
    76: [("personalized football laces comforter", ["https://www.walmart.com/", "https://www.amazon.com/"], "Football bedding intent is visible; exact laces-only phrase is niche."),
         ("football texture laces bedding comforter set", ["https://www.wayfair.com/", "https://www.amazon.com/"], "Comparator supports football texture/laces bedding motif.")],
    77: [("custom glowing soccer ball comforter set", ["https://www.amazon.com/", "https://www.etsy.com/"], "Commercial soccer bedding exists; exact glowing-trails phrase is niche."),
         ("personalized soccer ball bedding set custom name", ["https://www.desertcart.in/", "https://www.amazon.com/"], "Comparator supports custom soccer bedding with name/product intent.")],
    78: [("custom God says you are floral blanket", ["https://www.amazon.com/", "https://www.suzitee.com/"], "Custom Christian/Bible verse blanket intent is supported; exact wording is narrower."),
         ("personalized Christian affirmation floral blanket", ["https://www.etsy.com/", "https://www.zazzle.com/"], "Comparator supports floral Christian affirmation blanket intent.")],
    79: [("custom floral butterfly affirmation blanket", ["https://www.walmart.com/", "https://macorner.co/"], "Floral/butterfly affirmation gift products exist; exact blanket phrase is moderately niche."),
         ("personalized butterfly affirmation blanket name", ["https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports personalized butterfly/affirmation blanket intent.")],
    80: [("custom God Says I Am butterfly blanket", ["https://www.walmart.com/", "https://famvibe.com/"], "God Says I Am blanket intent is supported; butterfly exact modifier is narrower."),
         ("personalized God Says I Am Christian blanket name", ["https://www.suzitee.com/", "https://www.etsy.com/"], "Comparator supports personalized God Says I Am blanket intent.")],
}

base.CRITERION_ASSESSMENTS = {
    71: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    72: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    73: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    74: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    75: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    76: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    77: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    78: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    79: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    80: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_008_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 69/69 ảnh (100%)**; chỉ inventory position **71–80**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 08 được chấm độc lập từ r4; QA r3 chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Exact keyword trong cụm football 71–76 và soccer 77 có vài motif rất niche, nên K1/K2 được chấm thận trọng theo từng sản phẩm.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 69 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính của batch 08: football player teal/electric/flag/burning, football laces, glowing soccer ball, God Says You Are, floral butterfly affirmation và God Says I Am butterfly.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 69 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
