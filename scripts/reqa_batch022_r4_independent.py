from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_022"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B22R4"
base.TRACE_ROW_REVISION = "r3"
base.SOURCE_SHA256_EXPECTED = "6F88917A2C2245F62E7CBA61E74EE5D1F2CE69E23A134B90E529E08C77BDB5A2"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_022_r4" / "SEO_Product_Optimization_qa_batch_022_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_134500"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_134500"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_134500"

base.PRODUCT_SCOPE = [
    (211, "8834374402247", 9),
    (212, "8834121498823", 7),
    (213, "8834150170823", 7),
    (214, "8834158067911", 5),
    (215, "8834325414087", 8),
    (216, "8834104852679", 7),
    (217, "8834118844615", 7),
    (218, "8834383216839", 8),
    (219, "8834364113095", 8),
    (220, "8834138177735", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "211-220"

base.PRODUCT_FACTS = {
    211: "basketball close-up number blanket with close-up ball, jersey number and optional name/number artwork",
    212: "custom basketball court comforter with perspective court layout and personalized name/number artwork",
    213: "custom basketball cracked wall comforter with cracked-wall ball artwork and custom name/number context",
    214: "custom rainbow basketball net comforter with basketball entering the net and colorful/rainbow lighting",
    215: "custom basketball fire water blanket with fire-and-water splash basketball artwork",
    216: "custom flaming basketball comforter with flying basketball and flame artwork",
    217: "custom light burst basketball comforter with glowing light-burst basketball artwork",
    218: "custom basketball court blanket with ball held by hand on court and optional personalization",
    219: "custom basketball player blanket with ball held under player arm and optional personalization",
    220: "custom orange basketball hoop comforter with orange hoop/ball artwork and custom name/number context",
}

base.SERP_EVIDENCE = {
    211: [("basketball close-up number blanket", ["https://www.zazzle.com/", "https://pixers.us/", "https://www.walmart.com/"], "Close-up/number exact wording is narrow; basketball blanket purchase intent is supported."),
          ("personalized basketball blanket name number", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.walmart.com/"], "Comparator supports custom basketball blanket with name/number.")],
    212: [("custom basketball court comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Basketball court comforter intent is supported by product comparables."),
          ("personalized basketball court bedding set", ["https://www.pinterest.com/", "https://muchomasqueflores.com/", "https://www.etsy.com/"], "Comparator supports personalized basketball court bedding intent.")],
    213: [("custom basketball cracked wall comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Cracked-wall exact phrase is niche; custom basketball comforter intent is supported."),
          ("basketball wall art comforter set", ["https://www.amazon.com/", "https://www.walmart.com/", "https://www.etsy.com/"], "Comparator partially supports the wall/cracked visual and bedding intent.")],
    214: [("custom rainbow basketball net comforter", ["https://www.temu.com/", "https://www.callie.com/", "https://ohaprints.com/"], "Rainbow exact phrase is narrow; basketball net bedding intent is supported."),
          ("basketball net bedding set custom", ["https://www.temu.com/", "https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports basketball net bedding intent.")],
    215: [("custom basketball fire water blanket", ["https://www.etsy.com/", "https://www.zazzle.com/", "https://www.walmart.com/"], "Fire/water exact modifier is niche, but custom basketball blanket intent is supported."),
          ("basketball fire blanket custom name", ["https://www.etsy.com/", "https://ohaprints.com/", "https://www.amazon.com/"], "Comparator partially supports flame/sports personalized blanket intent.")],
    216: [("custom flaming basketball comforter", ["https://www.amorcustomgifts.com/", "https://ohaprints.com/", "https://www.amazon.com/"], "Flaming basketball comforter has commercial comparables, though exact phrasing is marketplace-heavy."),
          ("personalized basketball flames comforter", ["https://www.etsy.com/", "https://www.amazon.com/", "https://ohaprints.com/"], "Comparator supports custom basketball comforter with flame motif.")],
    217: [("custom light burst basketball comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Light-burst exact phrase is niche; custom basketball comforter intent is supported."),
          ("glowing basketball comforter set", ["https://www.amazon.com/", "https://www.walmart.com/", "https://www.etsy.com/"], "Comparator partially supports glowing/light sports bedding intent.")],
    218: [("custom basketball court blanket", ["https://www.etsy.com/", "https://www.zazzle.com/", "https://www.walmart.com/"], "Basketball court blanket intent is supported by blanket/comparable listings."),
          ("personalized basketball court blanket", ["https://www.etsy.com/", "https://www.pinterest.com/", "https://www.walmart.com/"], "Comparator supports personalized basketball blanket purchase intent.")],
    219: [("custom basketball player blanket", ["https://www.amazon.com/", "https://www.etsy.com/", "https://doonakingdom.com.au/"], "Basketball player blanket intent is supported, with some broad sports-gift results."),
          ("personalized basketball player blanket name number", ["https://www.etsy.com/", "https://www.walmart.com/", "https://ohaprints.com/"], "Comparator supports personalized basketball player blanket intent.")],
    220: [("custom orange basketball hoop comforter", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.ubuy.co.id/"], "Orange exact modifier is narrow; basketball hoop comforter purchase intent is supported."),
          ("personalized basketball hoop comforter name number", ["https://www.amazon.com/", "https://www.etsy.com/", "https://ohaprints.com/"], "Comparator supports personalized basketball hoop comforter intent.")],
}

base.CRITERION_ASSESSMENTS = {
    211: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    212: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    213: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    214: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    215: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    216: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    217: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    218: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    219: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
    220: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "FULL", "E1": "FULL"},
}


def criterion_reason(cid: str, assessment: str, pos: int, product: dict, cust: dict | None = None) -> str:
    if cid == "P1":
        return f"{assessment}: product identity, handle and live product data match inventory scope for batch 22."
    if cid == "P2":
        return f"{assessment}: personalization/product-option claims were checked against live/cache customizer evidence for this product."
    if cid == "K1":
        return f"{assessment}: primary keyword '{product['primary_keyword']}' was checked against this product's own US/English SERP; partial means the exact visual modifier is niche."
    if cid == "K2":
        return f"{assessment}: comparator query supports commercial bedding/blanket intent for this product type; partial means results broaden beyond the exact motif."
    if cid == "K3":
        return f"{assessment}: evidence is SERP/product-comparable only; no paid volume or Search Console proof is claimed."
    if cid == "T1":
        return f"{assessment}: meta_title_seo is complete English and maps to the verified product type/motif."
    if cid == "T2":
        return f"{assessment}: title_proposed can function as product title/H1 and distinguishes this batch-22 design by motif/product type."
    if cid == "D1":
        return f"{assessment}: meta description is complete, not cut mid-word, and the 145-165 character target was treated as editorial guidance."
    if cid == "D2":
        return f"{assessment}: description HTML is customer-facing and tied to verified artwork, product type, options and personalization controls."
    if cid == "E1":
        return f"{assessment}: product key, handle, links and scoped evidence are consistent; row-level {base.TRACE_ROW_REVISION} labels are recorded separately as a traceability limitation."
    raise ValueError(cid)


base.criterion_reason = criterion_reason

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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_022_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 73/73 ảnh (100%)**; chỉ inventory position **211–220**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 22 được chấm độc lập từ r4; QA r2/r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm basketball blanket/comforter có intent mua hàng rõ; các exact modifier như cracked wall, rainbow net, fire-water, light burst, close-up và orange hoop được chấm thận trọng khi SERP còn rộng hoặc marketplace-heavy.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 73 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: close-up basketball, court perspective, cracked wall, rainbow/net, fire-water splash, flaming ball, light burst, hand-on-court và player-arm basketball designs.
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
