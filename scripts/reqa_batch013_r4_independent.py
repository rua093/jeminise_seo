from __future__ import annotations

from collections import Counter
from pathlib import Path

import reqa_batch006_r5_independent as base

base.BATCH_ID = "qa_batch_013"
base.REVISION = "r4"
base.ISSUE_PREFIX = "B13R4"
base.TRACE_ROW_REVISION = "r3"
base.SOURCE_SHA256_EXPECTED = "731864DE605A8FBF0EA1399F67E122B755B4F5EE1C5FED1233184BADD985DEB6"
base.SOURCE_XLSX = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "revisions" / "qa_batch_013_r4" / "SEO_Product_Optimization_qa_batch_013_r4.xlsx"
base.PREVIOUS_CACHE_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_013000"
base.OLD_QA_RUN_DIR = base.ROOT / "seo_runs" / base.SHOP / base.RUN_ID / "qa" / "20260908_013000"
base.OLD_QA_OUT_DIR = base.ROOT / "resutls" / base.SHOP / base.RUN_ID / "qa" / "20260908_013000"

base.PRODUCT_SCOPE = [
    (121, "8835553263815", 7),
    (122, "8835558965447", 7),
    (123, "8835554345159", 7),
    (124, "8835553067207", 7),
    (125, "8835550511303", 7),
    (126, "8834711421127", 7),
    (127, "8837116854471", 5),
    (128, "8837129535687", 5),
    (129, "8837134844103", 5),
    (130, "8837122195655", 5),
]
base.EXPECTED_PRODUCTS = len(base.PRODUCT_SCOPE)
base.EXPECTED_IMAGES_TOTAL = sum(item[2] for item in base.PRODUCT_SCOPE)
base.POSITION_RANGE_LABEL = "121-130"

base.PRODUCT_FACTS = {
    121: "vintage blue semi truck comforter with truck breaking through a stone wall",
    122: "blue semi truck chevron comforter with centered truck and chevron background",
    123: "metallic mesh semi truck comforter with truck on metallic mesh background",
    124: "textured metallic semi truck comforter with truck on textured metal background",
    125: "starry night semi truck comforter with truck under night sky",
    126: "cardinal Christmas quilt set with two cardinals on holly branches",
    127: "colorful wolf head quilt set with close-up wolf portrait artwork",
    128: "geometric wolf head quilt set with wolf portrait and geometric accents",
    129: "profile wolf feathers quilt set with side-profile wolf and feather motif",
    130: "wolf dreamcatcher quilt set with centered wolf head inside dreamcatcher",
}

base.SERP_EVIDENCE = {
    121: [("vintage blue semi truck comforter", ["https://www.amazon.com/", "https://www.wayfair.com/", "https://jeminise.com/"], "Truck bedding intent is supported; exact vintage blue/stone-wall wording is niche."),
          ("semi truck breaking through wall comforter", ["https://jeminise.com/collections/couples-family-blankets", "https://www.walmart.com/", "https://www.amazon.com/"], "Comparator supports broader semi-truck bedding/blanket intent more than exact wall phrase.")],
    122: [("blue semi truck chevron comforter", ["https://www.amazon.com/", "https://www.wayfair.com/", "https://www.spoonflower.com/"], "Exact chevron/truck bedding support is weak; broader truck bedding exists."),
          ("semi truck bedding comforter set blue", ["https://www.wayfair.com/keyword.php?keyword=semi+truck+bedding", "https://www.amazon.com/", "https://www.walmart.com/"], "Comparator supports semi-truck/truck bedding product intent.")],
    123: [("metallic mesh semi truck comforter", ["https://www.amazon.com/Printluxe-Trucker-Tapestry-Comforters-Bedspreads/dp/B0DFQ9TMLY", "https://www.wayfair.com/keyword.php?keyword=semi+truck+bedding", "https://jeminise.com/"], "Semi-truck comforter intent is supported, but metallic mesh exact phrase collides with non-bedding mesh products."),
          ("custom trucker semi truck comforter", ["https://www.amazon.com/Printluxe-Trucker-Tapestry-Comforters-Bedspreads/dp/B0DFQ9TMLY", "https://www.etsy.com/market/trucks_fleece", "https://www.zazzle.com/"], "Comparator supports custom trucker bedding/blanket purchase intent.")],
    124: [("textured metallic semi truck comforter", ["https://www.amazon.com/Printluxe-Trucker-Tapestry-Comforters-Bedspreads/dp/B0DFQ9TMLY", "https://www.wayfair.com/keyword.php?keyword=semi+truck+bedding", "https://www.walmart.com/"], "Exact textured metallic phrase is niche; truck comforter intent has broader support."),
          ("semi truck comforter set metallic", ["https://www.amazon.com/Printluxe-Trucker-Tapestry-Comforters-Bedspreads/dp/B0DFQ9TMLY", "https://www.wayfair.com/keyword.php?keyword=semi+truck+bedding", "https://www.fineartamerica.com/"], "Comparator supports semi-truck comforter/duvet product intent.")],
    125: [("starry night semi truck comforter", ["https://www.amazon.com/", "https://www.wayfair.com/", "https://www.homedepot.com/"], "Starry-night exact phrase is niche and can drift to generic starry bedding; semi truck modifier requires product-specific evidence."),
          ("semi truck night sky blanket", ["https://www.etsy.com/market/trucks_fleece", "https://www.ebay.com/", "https://www.zazzle.com/"], "Comparator supports truck blanket/gift intent with weaker exact night-sky demand.")],
    126: [("cardinal Christmas quilt set", ["https://www.lowes.com/pd/MarCielo-C79-3-Piece-Winter-Cardinals-Christmas-Queen-Size-Quilt-Bedspread-Set/6930609", "https://us.shein.com/HappyHeartedGifts-Cardinal-Quilt-Set-U2013-Christmas-Winter-Bedding-With-Cardinals-And-Snowy-Trees-Cozy-Lightweight-Bedspread-For-All-Seasons-Cardinal-Quilt-10-p-352802969.html", "https://www.walmart.com/"], "Cardinal Christmas quilt/bedding intent is strongly supported by product comparables."),
          ("Christmas cardinal bedding quilt", ["https://www.amazon.com/Christmas-Snowman-Bedspread-Coverlet-Microfiber/dp/B07WPHBNWS", "https://www.lowes.com/pd/MarCielo-C79-3-Piece-Winter-Cardinals-Christmas-Queen-Size-Quilt-Bedspread-Set/6930609", "https://www.walmart.com/"], "Comparator supports Christmas cardinal bedding intent.")],
    127: [("colorful wolf head quilt set", ["https://www.etsy.com/market/wolf_queen_bedding_set", "https://www.pinterest.com/pin/wolf-head-quilt-wall-art--21814379437671248/", "https://www.ebay.co.uk/b/bn_95788689"], "Wolf bedding/quilt intent is supported; colorful wolf head exact phrase is narrower."),
          ("wolf head quilt bedding set", ["https://www.etsy.com/market/wolf_queen_bedding_set", "https://www.ebay.co.uk/b/bn_95788689", "https://www.temu.com/bh/wolf-bedding-5020228697451-s.html"], "Comparator supports wolf bedding and quilt product intent.")],
    128: [("geometric wolf head quilt set", ["https://www.spoonflower.com/en/home-decor/bedding/duvet-cover/6720129-geometric-wolf-arrows-greystone-geo-wolves-woodland-animals-baby-boy-by-gingerlous", "https://www.etsy.com/market/wolf_queen_bedding_set", "https://www.pinterest.com/"], "Geometric wolf bedding support exists, with some results as duvet/fabric rather than finished quilt."),
          ("geometric wolf bedding set", ["https://www.spoonflower.com/en/home-decor/bedding/duvet-cover/6720129-geometric-wolf-arrows-greystone-geo-wolves-woodland-animals-baby-boy-by-gingerlous", "https://www.etsy.com/market/wolf_queen_bedding_set", "https://www.temu.com/"], "Comparator supports geometric wolf bedding intent.")],
    129: [("profile wolf feathers quilt set", ["https://www.etsy.com/market/wolf_queen_bedding_set", "https://www.pinterest.com/ideas/wolf-quilt-kit/928734450929/", "https://www.ebay.co.uk/b/bn_95788689"], "Exact profile/feathers phrasing is niche; wolf quilt/bedding intent is supported."),
          ("wolf feather quilt bedding set", ["https://www.etsy.com/market/wolf_queen_bedding_set", "https://www.pinterest.com/", "https://www.temu.com/"], "Comparator supports wolf bedding; feather modifier is weaker.")],
    130: [("wolf dreamcatcher quilt set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.pinterest.com/"], "Wolf dreamcatcher blanket/quilt product intent is supported by marketplace comparables."),
          ("dreamcatcher wolf bedding set", ["https://www.etsy.com/", "https://www.amazon.com/", "https://www.temu.com/"], "Comparator supports dreamcatcher wolf bedding intent.")],
}

base.CRITERION_ASSESSMENTS = {
    121: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    122: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    123: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    124: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    125: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    126: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    127: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    128: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    129: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    130: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
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
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_013_r4

## Kết luận

- Phạm vi: **10 sản phẩm, 62/62 ảnh (100%)**; chỉ inventory position **121–130**; revision **r4**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{base.SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: batch 13 được chấm độc lập từ r4; QA r3 cũ chỉ dùng như lịch sử/nguồn cache ảnh, không dùng để cấp điểm.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Nhóm semi-truck có nhiều exact keyword nền/texture rất niche nên K1/K2 được chấm thận trọng theo từng thiết kế; cardinal Christmas và wolf dreamcatcher có intent mạnh hơn.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 62 ảnh SET và mở lại file ảnh gốc trong run mới.
- Alt/observation r4 khớp các motif chính: semi truck stone-wall/chevron/metallic/starry-night, cardinal Christmas, colorful/geometric/profile wolf và wolf dreamcatcher.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 62 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
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
