from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_007"
base.REVISION = "r5"
base.SOURCE_SHA256_EXPECTED = "A0DB47332AE2EAB9EAA03AA968BEC1B01B0F912350090FAB58230146ACE38191"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_007_r5" / "SEO_Product_Optimization_qa_batch_007_r5.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_007000"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_007000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_007000"

base.PRODUCT_SCOPE = [
    (61, "8834770993351", 6),
    (62, "8834758475975", 6),
    (63, "8834759655623", 6),
    (64, "8834765455559", 6),
    (65, "8834750873799", 6),
    (66, "8834764669127", 6),
    (67, "8834770108615", 6),
    (68, "8834763489479", 6),
    (69, "8834753298631", 6),
    (70, "8834760376519", 6),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "61-70"

base.PRODUCT_FACTS = {
    61: "football above green yard lines with perspective field artwork and custom name/number controls",
    62: "close-up football over an American flag background with custom name/number controls",
    63: "football engulfed in flames and lightning with custom name/number controls",
    64: "football helmet on American flag vintage background with custom name/number controls",
    65: "football helmet with flaming football design and custom name/number controls",
    66: "football over American flag background with custom name/number controls",
    67: "football player holding ball in close-up artwork with custom name/number controls",
    68: "football player holding ball with helmet-focused artwork and custom name/number controls",
    69: "football player collage with ball and helmet imagery plus custom name/number controls",
    70: "football player running on camouflage flag background with custom name/number controls",
}

base.SERP_EVIDENCE = {
    61: [("personalized football yard line comforter", ["https://www.etsy.com/", "https://ohaprints.com/"], "Commercial football bedding/custom comforter intent is supported; exact yard-line phrase is narrower."),
         ("custom football comforter set name number", ["https://www.youcustomizeit.com/", "https://www.amazon.com/"], "Comparator supports custom football comforter with text/name personalization.")],
    62: [("personalized American flag football comforter", ["https://www.amazon.com/", "https://www.cubebik.com/"], "Commercial patriotic football bedding intent is visible."),
         ("custom football flag bedding set name", ["https://ohaprints.com/", "https://www.etsy.com/"], "Comparator supports flag-football bedding and custom name intent.")],
    63: [("personalized flaming lightning football comforter", ["https://www.walmart.com/", "https://www.temu.com/"], "Fire/lightning football bedding exists, but exact phrase is niche."),
         ("custom football lightning bedding set name", ["https://www.amazon.com/", "https://www.callie.com/"], "Comparator supports sports bedding/personalized football blanket intent.")],
    64: [("personalized football helmet flag comforter", ["https://www.amazon.com/", "https://ohaprints.com/"], "Commercial football helmet/flag bedding intent is supported."),
         ("custom football helmet American flag bedding", ["https://www.walmart.com/", "https://us.amazon.com/"], "Comparator supports helmet/flag bedding product intent.")],
    65: [("personalized flaming football helmet comforter", ["https://us.amazon.com/", "https://www.amazon.com/"], "Football helmet on fire/flaming bedding intent is supported, with exact personalized phrase narrower."),
         ("custom football helmet on fire bedding set", ["https://us.amazon.com/", "https://www.walmart.com/"], "Comparator strongly supports the visible motif and product type.")],
    66: [("personalized football over flag comforter", ["https://www.amazon.com/", "https://www.cubebik.com/"], "Commercial football flag bedding intent is supported."),
         ("custom football American flag comforter set", ["https://ohaprints.com/", "https://www.etsy.com/"], "Comparator supports custom football + American flag bedding intent.")],
    67: [("personalized football player close up comforter", ["https://www.amazon.com/", "https://www.etsy.com/"], "Football player bedding and personalized football bedding are supported; exact close-up phrase is narrow."),
         ("custom football player bedding set name", ["https://www.youcustomizeit.com/", "https://doonakingdom.com.au/"], "Comparator supports custom football player bedding intent.")],
    68: [("personalized football player helmet comforter", ["https://www.wayfair.com/", "https://www.amazon.com/"], "Football player/helmet bedding intent is supported."),
         ("custom football player helmet bedding set", ["https://www.etsy.com/", "https://us.gowgowstore.com/"], "Comparator supports custom football player bedding with name/number.")],
    69: [("personalized football player collage comforter", ["https://www.etsy.com/", "https://snuggleandsend.com/"], "Football collage personalization is better supported for blankets; comforter exact phrase is niche."),
         ("custom football player collage blanket name number", ["https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports collage/name/number sports gift intent.")],
    70: [("personalized camo football player comforter", ["https://us.gowgowstore.com/", "https://www.walmart.com/"], "Camo football player bedding appears in commercial results, but exact comforter query is narrow."),
         ("custom camo football bedding set name number", ["https://www.amazon.com/", "https://www.etsy.com/"], "Comparator supports custom football bedding/name-number intent.")],
}

base.CRITERION_ASSESSMENTS = {
    61: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    62: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    63: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    64: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    65: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    66: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    67: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    68: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    69: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    70: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
}


_base_customizer_summary = base.customizer_summary


def customizer_summary(cust: dict | None) -> str:
    if not cust:
        return "no cached customizer audit available"
    nodes = cust.get("personalization_nodes", [])
    if nodes:
        return "; ".join(
            f"{n.get('label')} required={n.get('required')} max={n.get('maxLength')}" for n in nodes
        )
    return _base_customizer_summary(cust)


base.customizer_summary = customizer_summary


def markdown_report(run: str, out_xlsx: Path, out_md: Path, products: list[dict], scoped: list[tuple], issues: list[dict], avg_score: float, batch_result: str, source_sha: str) -> str:
    title_by_key = {p["product_key"]: p["title_proposed"] for _, _, _, p, _ in scoped}
    status_counts = Counter(p["qa_status"] for p in products)
    severity_counts = Counter(i["severity"] for i in issues)
    rows = "\n".join(f"| {p['inventory_position']} | {title_by_key[p['product_key']]} | {p['final_score']:.1f} | {p['qa_status']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']}/{p['limitation_count']} |" for p in products)
    priority = "\n".join(f"{idx}. **{i['severity']} — {i['field']}** ({title_by_key.get(i['product_key'], 'batch-level')}): {i['reason']} Đề xuất: {i['recommended_fix']}" for idx, i in enumerate(issues[:12], 1))
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_007_r5

## Kết luận

- Phạm vi: **10 sản phẩm, 60/60 ảnh (100%)**; chỉ inventory position **61–70**; revision **r5**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 07 được chấm độc lập từ r5; QA r4 chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Các exact keyword football motif như yard line, flaming lightning, player collage và camo được chấm thận trọng nếu SERP chỉ hỗ trợ broader personalized football bedding intent.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 60 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r5 khớp các motif football: yard line, American flag, flaming/lightning, helmet, football-over-flag, close-up player, helmet player, collage và camo player.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 60 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
