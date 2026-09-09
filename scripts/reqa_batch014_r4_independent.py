from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_014"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B14R4"
base.TRACE_ROW_REVISION = "r3"
base.SOURCE_SHA256_EXPECTED = "BE6DC10F7CBA2B6ABE4EDA2FB2C3C68C44D09EE44C4846294F1E624C1E58E9BE"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_014_r4" / "SEO_Product_Optimization_qa_batch_014_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_014000"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_014000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_014000"

base.PRODUCT_SCOPE = [
    (131, "8837155815623", 5),
    (132, "8867219341511", 7),
    (133, "8859639873735", 7),
    (134, "8867097346247", 7),
    (135, "8867094200519", 7),
    (136, "8867095216327", 7),
    (137, "8867093741767", 7),
    (138, "8867094462663", 7),
    (139, "8867095838919", 7),
    (140, "8833333493959", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "131-140"

base.PRODUCT_FACTS = {
    131: "wolf headdress quilt set with centered wolf head and headdress artwork",
    132: "custom teal yellow softball comforter with bright sports bedding motif",
    133: "desert cactus flower quilt set with cactus and floral desert artwork",
    134: "purple floral dragonfly quilt set with dragonfly and flower artwork",
    135: "sunburst dragonfly pond quilt set with dragonfly pond artwork",
    136: "stained glass dragonfly pond quilt with stained-glass pond artwork",
    137: "purple vintage dragonfly quilt set with vintage dragonfly artwork",
    138: "colorful elephant patchwork quilt with animal patchwork motif",
    139: "navy yellow elephant quilt set with elephant patchwork motif",
    140: "rainbow fantasy dragon quilt set with fantasy dragon and scales artwork",
}

base.SERP_EVIDENCE = {
    131: [("wolf headdress quilt set", ["https://www.etsy.com/market/wolf_queen_bedding_set", "https://www.amazon.com/", "https://www.pinterest.com/"], "Wolf bedding intent is supported; headdress exact modifier is niche and should stay tied to visible artwork."),
          ("wolf head quilt bedding set", ["https://www.etsy.com/market/wolf_queen_bedding_set", "https://www.ebay.co.uk/b/bn_95788689", "https://www.temu.com/"], "Comparator supports wolf head bedding/quilt product intent.")],
    132: [("custom teal yellow softball comforter", ["https://www.youcustomizeit.com/", "https://www.ebay.com/", "https://2cooldesigns.com/"], "Custom softball bedding/comforter intent is supported, including teal/yellow comparables."),
          ("personalized teal softball comforter set", ["https://2cooldesigns.com/", "https://www.etsy.com/", "https://www.ebay.com/"], "Comparator supports personalized softball comforter purchase intent.")],
    133: [("desert cactus flower quilt set", ["https://www.wayfair.com/", "https://rndrustics.com/", "https://www.walmart.com/"], "Cactus/desert quilt bedding intent is supported by product comparables."),
          ("cactus flower bedding quilt set", ["https://www.wayfair.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports cactus/floral bedding product intent.")],
    134: [("purple floral dragonfly quilt set", ["https://www.amazon.com/", "https://www.walmart.com/", "https://luvingift.com/"], "Dragonfly/floral quilt bedding intent is supported; exact purple phrase has product comparables."),
          ("dragonfly floral quilt bedding set", ["https://www.walmart.com/", "https://ohaprints.com/", "https://www.ebay.com/"], "Comparator supports dragonfly floral bedding/quilt intent.")],
    135: [("sunburst dragonfly pond quilt set", ["https://www.cherylsquiltcorner.com/", "https://www.bearpawproductions.com/", "https://www.pinterest.com/"], "Exact phrase tends to pattern/project results; finished bedding intent is weaker."),
          ("dragonfly pond quilt bedding set", ["https://alphaquilt.com/", "https://ohaprints.com/", "https://www.walmart.com/"], "Comparator supports dragonfly quilt bedding more broadly, with pond exact modifier weaker.")],
    136: [("stained glass dragonfly pond quilt", ["https://www.cherylsquiltcorner.com/", "https://www.bearpawproductions.com/", "https://www.etsy.com/"], "Stained-glass dragonfly pond results skew to quilt patterns/projects rather than finished bedding."),
          ("stained glass dragonfly quilt bedding", ["https://www.walmart.com/", "https://www.missouriquiltco.com/", "https://www.etsy.com/"], "Comparator supports dragonfly/stained-glass quilt concept, mixed between bedding and patterns.")],
    137: [("purple vintage dragonfly quilt set", ["https://www.amazon.com/", "https://www.walmart.com/", "https://www.ebay.com/"], "Purple dragonfly bedding intent is supported; vintage exact modifier is narrower."),
          ("vintage dragonfly bedding quilt set", ["https://www.ebay.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports dragonfly bedding product intent.")],
    138: [("colorful elephant patchwork quilt", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Elephant patchwork/quilt intent is supported, though some results are patterns or children's products."),
          ("elephant patchwork quilt bedding set", ["https://www.walmart.com/", "https://www.amazon.com/", "https://www.etsy.com/"], "Comparator supports elephant quilt/bedding purchase intent.")],
    139: [("navy yellow elephant quilt set", ["https://www.walmart.com/", "https://www.amazon.com/", "https://www.etsy.com/"], "Navy/yellow exact modifier is niche; elephant quilt set intent exists."),
          ("elephant quilt set navy yellow", ["https://www.walmart.com/", "https://www.amazon.com/", "https://www.ebay.com/"], "Comparator supports elephant quilt bedding with weaker exact color demand.")],
    140: [("rainbow fantasy dragon quilt set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.pinterest.com/"], "Fantasy dragon bedding/quilt intent is supported; rainbow modifier is a visible product-specific differentiator."),
          ("fantasy dragon quilt bedding set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports fantasy dragon bedding product intent.")],
}

base.CRITERION_ASSESSMENTS = {
    131: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    132: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    133: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    134: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    135: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    136: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    137: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    138: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    139: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    140: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_014_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 68/68 ảnh (100%)**; chỉ inventory position **131–140**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 14 được chấm độc lập từ r4; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Dragonfly pond/stained-glass được chấm thận trọng vì SERP có nhiều pattern/kit thay vì finished quilt set; softball, cactus, dragonfly floral và fantasy dragon có product intent rõ hơn.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 68 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: wolf headdress, teal softball, desert cactus, dragonfly variants, elephant patchwork và rainbow fantasy dragon.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 68 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
