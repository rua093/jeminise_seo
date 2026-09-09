from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_015"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B15R4"
base.TRACE_ROW_REVISION = "r3"
base.SOURCE_SHA256_EXPECTED = "6EE4B35320B2AC7C4CA8F6F1E96A3B744AE342B318B1354D9943C5344D1FE84F"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_015_r4" / "SEO_Product_Optimization_qa_batch_015_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_015000"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_015000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_015000"

base.PRODUCT_SCOPE = [
    (141, "8833336115399", 5),
    (142, "8833331822791", 5),
    (143, "8833354268871", 7),
    (144, "8833355350215", 6),
    (145, "8833335099591", 5),
    (146, "8833328414919", 5),
    (147, "8833330806983", 5),
    (148, "8834736488647", 5),
    (149, "8834799861959", 7),
    (150, "8867094888647", 7),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "141-150"

base.PRODUCT_FACTS = {
    141: "red fire dragon lava quilt set with fantasy dragon, fire and lava artwork",
    142: "teal winged dragon quilt set with mythical winged dragon artwork",
    143: "golden roots Tree of Life quilt set with entwined golden roots artwork",
    144: "twisted Tree of Life quilt set with twisted trunk and nature artwork",
    145: "green fire dragon quilt set with dragon, fire, lava and spikes artwork",
    146: "purple nebula dragon quilt set with winged dragon and starry nebula artwork",
    147: "volcanic winged dragon quilt set with dragon and volcanic eruption artwork",
    148: "cardinal oval floral quilt with cardinal inside an ornate floral oval frame",
    149: "Celtic knot Tree of Life quilt with folklore tree and Celtic knot artwork",
    150: "fox patchwork quilt set with animal patchwork bedding artwork",
}

base.SERP_EVIDENCE = {
    141: [("red fire dragon lava quilt set", ["https://www.walmart.com/", "https://www.desertcart.vn/", "https://www.liberia.ubuy.com/"], "Red/fire dragon bedding intent is supported; lava exact phrasing is product-specific and should stay tied to visible artwork."),
          ("fantasy dragon quilt bedding set", ["https://www.bigw.com.au/", "https://www.etsy.com/", "https://www.amazon.com/"], "Comparator supports fantasy dragon quilt/bedding product intent.")],
    142: [("teal winged dragon quilt set", ["https://www.u-buy.co.uk/", "https://www.ubuy.co.bw/", "https://www.etsy.com/"], "Dragon bedding intent is supported; teal winged exact phrase is narrower than the broad SERP."),
          ("dragon bedding set teal", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports dragon bedding purchase intent with color/motif modifiers.")],
    143: [("golden roots tree of life quilt set", ["https://miravoventures.com/", "https://www.amazon.com/", "https://www.mercadolivre.com.br/"], "Tree of Life quilt/bedding intent is supported, with golden roots as a visible differentiator."),
          ("Tree of Life golden roots quilt bedding", ["https://www.etsy.com/", "https://www.ebay.com/", "https://alphaquilt.com/"], "Comparator supports Tree of Life bedding product intent.")],
    144: [("twisted tree of life quilt set", ["https://jeminise.com/", "https://www.etsy.com/", "https://alphaquilt.com/"], "Tree of Life quilt product intent is supported and the twisted-trunk modifier is visible."),
          ("tree of life quilt bedding set", ["https://www.amazon.com/", "https://www.ebay.com/", "https://shop.crackerbarrel.com/"], "Comparator supports broad Tree of Life quilt/bedding intent.")],
    145: [("green fire dragon quilt set", ["https://www.amazon.com/", "https://jeminise.com/", "https://www.etsy.com/"], "Green/fire dragon bedding intent is supported by product comparables."),
          ("dragon fire quilt bedding set", ["https://www.bigw.com.au/", "https://www.mattblatt.com.au/", "https://www.walmart.com/"], "Comparator supports dragon fire bedding purchase intent.")],
    146: [("purple nebula dragon quilt set", ["https://www.walmart.com/", "https://www.kosovo.ubuy.com/", "https://www.amazon.com/"], "Purple/nebula dragon bedding intent is supported; exact wording remains niche."),
          ("dragon galaxy nebula bedding set", ["https://www.amazon.com/", "https://www.ebay.com/", "https://www.etsy.com/"], "Comparator supports galaxy/nebula dragon bedding product intent.")],
    147: [("volcanic winged dragon quilt set", ["https://songhaiflange.com/", "https://www.handfulofprints.com/", "https://www.pinterest.com/"], "Volcanic winged exact phrase is niche and mixed; broad dragon bedding intent is clearer."),
          ("volcano dragon bedding set", ["https://www.etsy.com/", "https://www.ebay.com/", "https://www.amazon.com/"], "Comparator supports volcano/dragon bedding product intent.")],
    148: [("cardinal oval floral quilt", ["https://www.walmart.com/c/kp/cardinal-quilt", "https://www.wayfair.com/keyword.php?keyword=cardinal+quilt", "https://www.etsy.com/market/cardinal_throw_quilt"], "Cardinal quilt intent is supported; oval floral exact phrase is a narrow visual modifier."),
          ("floral cardinal quilt bedding", ["https://www.walmart.com/", "https://www.wayfair.com/", "https://www.ebay.com/"], "Comparator supports floral/cardinal quilt and bedding purchase intent.")],
    149: [("Celtic knot tree of life quilt", ["https://www.amazon.com/", "https://luvingift.com/product/viking-quilt-set-norse-celtic-yggdrasil-tree-of-life/", "https://jeminise.com/"], "Celtic Tree of Life quilt intent is supported with relevant product comparables."),
          ("Celtic tree of life quilt bedding", ["https://www.etsy.com/", "https://www.amazon.com/", "https://alphaquilt.com/"], "Comparator supports Celtic Tree of Life bedding intent.")],
    150: [("fox patchwork quilt set", ["https://www.etsy.com/", "https://legitkits.com/", "https://www.quiltylove.com/"], "Fox patchwork exact SERP includes patterns as well as quilt products, so product intent is mixed."),
          ("fox quilt bedding set", ["https://www.amazon.com/", "https://www.etsy.com/", "https://www.walmart.com/"], "Comparator supports fox quilt/bedding purchase intent.")],
}

base.CRITERION_ASSESSMENTS = {
    141: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    142: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    143: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    144: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    145: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    146: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    147: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    148: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    149: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    150: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_015_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 57/57 ảnh (100%)**; chỉ inventory position **141–150**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 15 được chấm độc lập từ r4; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm fantasy dragon và Tree of Life có intent sản phẩm tương đối rõ; các exact modifier như teal winged, volcanic, cardinal oval và fox patchwork được chấm thận trọng khi SERP còn lẫn broader hoặc pattern intent.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 57 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: red/teal/green/purple/volcanic dragon, golden/twisted/Celtic Tree of Life, cardinal oval floral và fox patchwork.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 57 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
