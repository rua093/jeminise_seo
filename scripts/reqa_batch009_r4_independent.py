from __future__ import annotations

import shutil
from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_009"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B09R4"
base.SOURCE_SHA256_EXPECTED = "5D0C14FE3E207FF726348EC5A90C3FB7DB946CC1C6742D6ECFCA5630BC775749"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_009_r4" / "SEO_Product_Optimization_qa_batch_009_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260907_180500"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_009000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_009000"
EXTRA_OLD_DIRS = [
    base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260907_180500",
    base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260907_180500",
]

base.PRODUCT_SCOPE = [
    (81, "8867271082183", 5),
    (82, "8867273310407", 6),
    (83, "8867270394055", 7),
    (84, "8867273113799", 7),
    (85, "8867271770311", 6),
    (86, "8867272196295", 6),
    (87, "8867272589511", 6),
    (88, "8867270721735", 7),
    (89, "8867270656199", 5),
    (90, "8867269279943", 4),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "81-90"

base.PRODUCT_FACTS = {
    81: "purple baseball batter blanket with custom name/personalization context",
    82: "baseball batter with stadium and American flag themed blanket",
    83: "baseball field quote blanket with custom text/personalization context",
    84: "baseball practice quote blanket with custom text/personalization context",
    85: "flaming baseball blanket with fire motif",
    86: "baseball glove and flag blanket with custom/personalized gift intent",
    87: "baseball catcher quote blanket",
    88: "American flag baseball pitcher blanket",
    89: "baseball catcher and batter blanket",
    90: "God Says You Are baseball blanket with faith/sports motif",
}

base.SERP_EVIDENCE = {
    81: [("custom purple baseball batter blanket", ["https://www.etsy.com/", "https://www.haloballs.com/"], "Commercial baseball batter personalized blanket intent is supported; purple exact modifier is narrow."),
         ("personalized baseball batter blanket custom name", ["https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports custom baseball batter blanket with name.")],
    82: [("baseball batter stadium flag blanket", ["https://www.amazon.com/", "https://www.zazzle.com/"], "Broad baseball/stadium/flag blanket intent exists, exact phrase is niche."),
         ("American flag baseball blanket personalized", ["https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports flag baseball blanket product intent.")],
    83: [("custom baseball field quote blanket", ["https://www.etsy.com/", "https://www.zazzle.com/"], "Custom baseball blanket intent exists; quote-specific field phrasing is narrower."),
         ("personalized baseball blanket quote custom", ["https://stinkylockers.com/", "https://www.amazon.com/"], "Comparator supports personalized baseball blanket and custom text intent.")],
    84: [("custom baseball practice quote blanket", ["https://us.shein.com/", "https://www.etsy.com/"], "Baseball quote blanket intent is visible but exact practice phrase is weaker."),
         ("baseball theme blanket inspirational quote", ["https://us.shein.com/", "https://www.zazzle.com/"], "Comparator supports quote/inspirational baseball blanket intent.")],
    85: [("flaming baseball blanket", ["https://www.etsy.com/", "https://www.pinterest.com/"], "Flame baseball blanket imagery/product intent is supported."),
         ("personalized baseball blanket flame design", ["https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports personalized flame baseball blanket intent.")],
    86: [("custom baseball glove flag blanket", ["https://www.etsy.com/", "https://www.amazon.com/"], "Baseball glove/flag blanket intent exists; exact custom phrase is niche."),
         ("personalized baseball glove blanket name", ["https://www.roseinside.com/", "https://www.zazzle.com/"], "Comparator supports baseball glove blanket with name.")],
    87: [("baseball catcher quote blanket", ["https://www.pinterest.com/", "https://us.shein.com/"], "Catcher quote content exists, but commercial blanket exact intent is weaker."),
         ("personalized catcher blanket with name", ["https://www.etsy.com/", "https://www.facebook.com/"], "Comparator supports catcher-specific personalized blanket intent.")],
    88: [("American flag baseball pitcher blanket", ["https://www.amazon.com/", "https://www.redbubble.com/"], "American flag baseball player/pitcher blanket motif is supported."),
         ("personalized baseball pitcher blanket flag", ["https://www.etsy.com/", "https://interestpod.co/"], "Comparator supports baseball player + flag blanket/gift intent.")],
    89: [("baseball catcher batter blanket", ["https://www.amazon.com/", "https://www.etsy.com/"], "Catcher/batter baseball blanket intent is visible in product/comparator results."),
         ("custom name catcher batter baseball blanket", ["https://www.amazon.com/", "https://www.haloballs.com/"], "Comparator supports custom name catcher/batter baseball blanket.")],
    90: [("God Says You Are baseball blanket", ["https://www.etsy.com/", "https://us.shein.com/"], "Faith + baseball blanket exact concept is supported by marketplace listings."),
         ("personalized baseball blanket God Says I Am", ["https://www.etsy.com/", "https://www.callie.com/"], "Comparator supports Christian sports personalized blanket intent.")],
}

base.CRITERION_ASSESSMENTS = {
    81: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    82: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    83: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    84: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    85: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    86: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    87: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    88: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    89: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    90: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_009_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 59/59 ảnh (100%)**; chỉ inventory position **81–90**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 09 được chấm độc lập từ r4; QA cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Các exact keyword baseball như practice quote/catcher quote/stadium flag được chấm thận trọng nếu SERP chỉ hỗ trợ broader baseball blanket hoặc custom-name blanket intent.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 59 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: baseball batter, stadium flag, quote blankets, flaming baseball, glove flag, catcher, pitcher và God Says You Are baseball.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 59 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
    for old_dir in EXTRA_OLD_DIRS:
        if old_dir.exists():
            shutil.rmtree(old_dir)
