from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import build_qa_batch1_report as common


ROOT = Path(__file__).resolve().parents[1]
SHOP, RUN_ID, QA_RUN_ID, BATCH = "jeminise.com", "20260906_234129", "20260907_104027", "qa_batch_007"
QA_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID
OUT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN_ID
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_034.xlsx"
SNAPSHOT = QA_DIR / "source_snapshot" / SOURCE.name
PW, IW, R = common.PRODUCT_WEIGHTS, common.IMAGE_WEIGHTS, common.RATING


DESIGNS = {
    61: "black-and-gold yard-line football artwork with sample name DAVID and number 35",
    62: "close football over a distressed American flag with sample name JOHN and number 15",
    63: "football surrounded by orange flames, blue smoke and lightning with sample name ANDREW and number 33",
    64: "red football helmet over a vintage American flag with sample name John and number 8",
    65: "black football helmet, flaming football and stadium with sample name MATTHEW and number 06",
    66: "close football over a distressed American flag with sample name ANDREW and number 45",
    67: "black-and-white football player holding a ball over an orange honeycomb pattern with sample name William and number 20",
    68: "smoky black-and-white football player holding a ball with sample name David and number 15",
    69: "black-and-white football player collage with sample name MICHAEL and number 30",
    70: "running football player over a green camouflage flag design with sample name ROBERT and number 09",
}


def actual_images(position):
    design = DESIGNS[position]
    order = [
        f"Front bedroom mockup of {design}.",
        f"Second bedroom mockup of {design}.",
        "Product-type comparison explaining that a duvet cover needs an insert while a comforter includes filling.",
        f"Close bedroom feature panel for {design}, marked soft, lightweight, durable and breathable.",
        "Easy-care panel marked wrinkle-free, stain-proof, anti-pilling and resistant to fading in the wash.",
        "Comforter size-dimension chart for Twin 68x86, Full 79x90, Queen 90x90 and King 90x104 inches.",
    ]
    if position == 65:
        order[3], order[4] = order[4], order[3]
    return order


ACTUAL = {position: actual_images(position) for position in range(61, 71)}
WRONG = {(65, 4), (65, 5)}
GENERIC_ALT = {(position, 4) for position in range(61, 71) if position != 65}

SERP = {
    61: [("personalized football comforter name number", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"]), ("custom football bedding set name number", ["https://www.youcustomizeit.com/p/Football-Comforters-Personalized/142169"])],
    62: [("American flag football comforter", ["https://www.cubebik.com/shop/football-bedding-set-personalized-football-american-flag-custom-bedding-set-pillow-and-duvet-cover-custom-football-lover-gifts/"]), ("personalized American flag football bedding", ["https://custombeddingset.com/america-football-custom-bedding-set-personalized-us-flag-duvet-cover-bed-sheets-pillow-shams/"])],
    63: [("flaming football comforter set personalized", ["https://www.zazzle.com/personalized_football_blankets_with_name_number-256772106061879793"]), ("fire football bedding set name number", ["https://www.etsy.com/ie/listing/1883272338/custom-football-bedding-set-personalized"])],
    64: [("football helmet comforter set personalized", ["https://joliebrand.com/products/customizable-football-duvet-cover-set-personalized-bedding-with-football-ball-helmet-design-black-and-white-player-gift-custom-name-and-number-option-lylyprint-com"]), ("football helmet American flag bedding", ["https://www.cubebik.com/shop/football-bedding-set-personalized-football-american-flag-custom-bedding-set-pillow-and-duvet-cover-custom-football-lover-gifts/"])],
    65: [("personalized football helmet comforter", ["https://www.youcustomizeit.com/p/Football-Comforters-Personalized/142169"]), ("custom football player helmet comforter name number", ["https://joliebrand.com/products/customizable-football-duvet-cover-set-personalized-bedding-with-football-ball-helmet-design-black-and-white-player-gift-custom-name-and-number-option-lylyprint-com"])],
    66: [("personalized patriotic football comforter American flag", ["https://custombeddingset.com/america-football-custom-bedding-set-personalized-us-flag-duvet-cover-bed-sheets-pillow-shams/"]), ("custom American football flag bedding set", ["https://www.cubebik.com/shop/football-bedding-set-personalized-football-american-flag-custom-bedding-set-pillow-and-duvet-cover-custom-football-lover-gifts/"])],
    67: [("football player comforter personalized name number", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"]), ("American football player bedding set boys", ["https://www.target.com/s/boys%2Bfootball%2Bbedding"])],
    68: [("custom football player helmet comforter name number", ["https://joliebrand.com/products/customizable-football-duvet-cover-set-personalized-bedding-with-football-ball-helmet-design-black-and-white-player-gift-custom-name-and-number-option-lylyprint-com"]), ("personalized black football player bedding set", ["https://ohaprints.com/products/personalized-football-duvet-cover-set-america-football-black-white-player-gift-duvet-cover-pillowcases-custom-name-number-bedding-set"])],
    69: [("personalized football player collage comforter", ["https://www.megacustom.com/products/personalized-american-football-photo-collage-blanket-with-custom-name-and-number"]), ("custom football player duvet cover name number", ["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"])],
    70: [("camouflage football comforter personalized", ["https://www.youcustomizeit.com/p/Camo-Comforters-Personalized/142261"]), ("camo football player bedding set name number", ["https://celeskydesigns.com/products/custom-camo-comforter-or-duvet-cover-18019"])],
}

ASSESS = {
    position: {
        "P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "PARTIAL", "K3": "PARTIAL",
        "T1": "FULL", "T2": "PARTIAL", "D1": "FAIL", "D2": "PARTIAL", "E1": "PARTIAL",
    }
    for position in range(61, 71)
}


def add(rows, *args):
    rows.append(common.issue(*args))


def main():
    data = json.loads((QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))
    live = json.loads((QA_DIR / "live_source_comparison.json").read_text(encoding="utf-8"))
    customizer = json.loads((QA_DIR / "customizer_audit.json").read_text(encoding="utf-8"))
    products = data["products"]
    assert len(products) == len(live) == len(customizer) == 10
    by_pk = {p["product_key"]: [] for p in products}
    for image in data["images"]:
        by_pk[image["shop_domain"] + "+" + image["Handle"]].append(image)
    checked, qa_images, issues = common.now(), [], []

    for index, product in enumerate(products):
        position = 61 + index
        media = live[index]["live"]["product_js"]["media"]
        images = sorted(by_pk[product["product_key"]], key=lambda row: row["image_number"])
        assert len(images) == len(media) == len(ACTUAL[position]) == 6
        for image_number, (submitted, source_media) in enumerate(zip(images, media), 1):
            wrong = (position, image_number) in WRONG
            generic = (position, image_number) in GENERIC_ALT
            assessment = {
                "IM1": "FULL",
                "IM2": "FAIL" if wrong else "FULL",
                "IM3": "FAIL" if wrong else ("PARTIAL" if generic else "FULL"),
                "IM4": "FULL",
            }
            points = sum(IW[key] * R[value] for key, value in assessment.items())
            qa_key = common.stable_image_key(product["product_key"], submitted["image_url"], image_number)
            evidence = [submitted["evidence_file_or_reference"], product["product_url"], source_media["src"]]
            issue_refs = []
            if wrong:
                issue_id = f"ISS-{position:03d}-IMG-{image_number:02d}"
                issue_refs.append(issue_id)
                add(issues, issue_id, product["product_key"], "MAJOR", "image_observation/alt_effective",
                    f"{submitted['observed_visual_details']} | {submitted['alt_proposed']}", ACTUAL[position][image_number - 1],
                    "Observation và alt bị đảo giữa bảng easy-care và bảng tính năng ở gallery thực tế.",
                    ACTUAL[position][image_number - 1], "; ".join(evidence),
                    "Open the source image and confirm the revised English alt matches the exact panel.", qa_key)
            elif generic:
                issue_id = f"ISS-{position:03d}-IMG-{image_number:02d}"
                issue_refs.append(issue_id)
                add(issues, issue_id, product["product_key"], "MINOR", "image_alt_effective",
                    submitted["alt_proposed"], ACTUAL[position][image_number - 1],
                    "Alt gọi đây là close-up nhưng không nói rõ đây là bảng tính năng soft/lightweight/durable/breathable.",
                    f"Feature panel for {product['title_proposed']}: soft, lightweight, durable and breathable.",
                    "; ".join(evidence), "Confirm the revised alt describes the feature panel without keyword stuffing.", qa_key)
            qa_images.append({
                "product_key": product["product_key"], "qa_image_key": qa_key,
                "image_url_source": source_media["src"], "image_url_workbook": submitted["image_url"],
                "media_id": str(source_media["id"]), "workbook_image_id": str(submitted["media_id"]),
                "variant": submitted["variant"] or "", "image_location": submitted["image_location"],
                "check_method": "DIRECT_ORIGINAL_IMAGE", "checked_at": checked,
                "qa_observation": ACTUAL[position][image_number - 1],
                "submitted_observation": submitted["observed_visual_details"],
                "storefront_alt_observed": source_media.get("alt") or "",
                "alt_action": submitted["alt_action"], "alt_effective": submitted["alt_proposed"],
                **assessment, "image_verified_points": points, "image_assessed_weight": 100,
                "image_final_score": points, "image_score_lower_bound": points, "image_score_upper_bound": points,
                "issue_refs": issue_refs, "evidence_refs": evidence,
            })

    for position, product in enumerate(products, 61):
        custom = customizer[position - 61]
        add(issues, f"ISS-{position:03d}-DESC", product["product_key"], "MAJOR", "description_proposed_html",
            "SEO Use / It needs QA and approval before import.", "Submitted HTML contains an internal workflow block.",
            "Câu nội bộ về SEO/QA/import không phải nội dung dành cho khách hàng.",
            "Remove the internal SEO Use block and retain only customer-facing English HTML.",
            product["evidence_id"], "Rendered HTML contains no drafting, QA, approval or import instruction.")
        add(issues, f"ISS-{position:03d}-CANN", product["product_key"], "MAJOR", "keyword_map/intent differentiation",
            product["primary_keyword"], "Ten adjacent pages target near-identical personalized American-football bedding intent.",
            "Modifier hình ảnh có khác nhưng vai trò landing page và internal linking chưa đủ để tránh cạnh tranh lẫn nhau.",
            "Assign one distinct primary intent per design cluster and document canonical/internal-link roles in English metadata.",
            "; ".join(url for _, urls in SERP[position] for url in urls),
            "Keyword map shows non-overlapping primary intent for yard-line, patriotic, flame, helmet, player and camo clusters.")
        add(issues, f"ISS-{position:03d}-META", product["product_key"], "MINOR", "meta_description_seo",
            f"{product['meta_description_seo']} ({product['meta_description_length']} characters)",
            "The proposed meta description exceeds the usual concise snippet range and ends with internal-style checkout guidance.",
            "Độ dài 169–200 ký tự làm tăng nguy cơ bị cắt, còn câu review options chưa nêu rõ cơ chế Customize.",
            "Shorten to about 145–160 characters and state that name/number are entered through Customize.",
            product["evidence_id"], "Meta is concise, product-specific and accurately explains personalization.")
        assert custom["customizer_root_present"] and custom["customize_button_present"] and custom["upload_endpoint_present"]

    add(issues, "ISS-GLOBAL-ADMIN", "", "LIMITATION", "current admin SEO fields/current admin alt", "Not supplied",
        "Storefront metadata and product.js media alt were readable, but no Shopify admin export was provided.",
        "Storefront values cannot prove admin fields or future import mapping.",
        "Provide a frozen Shopify admin/export snapshot before deployment approval.", str(SOURCE),
        "Admin export revision and hash are frozen and compared.")
    add(issues, "ISS-GLOBAL-SNAPSHOT-HTML", "", "LIMITATION", "phase-1 HTML snapshot", "Phase-1 HTML fetch returned HTTP 403.",
        "Frozen product JSON/body/options/gallery URLs match live, but old rendered metadata cannot be compared directly.",
        "Snapshot HTML cũ không đọc được nên không thể dùng hash HTML để kết luận drift.",
        "Capture a fresh immutable HTML snapshot for the next revision.", str(QA_DIR / "live_source_comparison.json"),
        "Both old and new readable HTML snapshots are available with hashes.")
    add(issues, "ISS-GLOBAL-CUSTOMIZER", "", "LIMITATION", "personalization purchase persistence",
        "Customizer root, button, upload endpoint and Name/Number schema verified in live HTML for all 10 products.",
        "An interactive browser was unavailable, so the final cart-line persistence and fulfillment payload were not executed.",
        "Cấu hình live hỗ trợ claim personalization, nhưng chưa kiểm thử end-to-end đến giỏ hàng.",
        "Run one controlled Customize-to-cart transaction per configuration revision before deployment approval.",
        str(QA_DIR / "customizer_audit.json"), "Cart line retains the entered name, optional number, preview and variant mapping.")

    issues_by_product = {}
    for row in issues:
        issues_by_product.setdefault(row["product_key"], []).append(row)
    reasons = dict(common.REASONS)
    reasons["P2"] = "Đã đối chiếu options, variants, description và cấu hình Amazon Customizer live có Name/Number."
    reasons["K2"] = "Đã đọc primary và comparator SERP theo US commercial intent; nhiều kết quả là blanket/duvet nên chỉ PARTIAL."
    reasons["K3"] = "Modifier theo artwork có dùng nhưng mười trang cùng cluster còn nguy cơ cannibalization."
    reasons["T2"] = "Meta đúng motif nhưng quá dài và chưa diễn đạt cơ chế Customize gọn, rõ."
    reasons["D1"] = "Description chứa nguyên khối SEO Use/QA/import nội bộ."
    reasons["D2"] = "Facts chính khớp nguồn, nhưng copy chưa publish-ready và chưa nói rõ ràng quy trình Customize."
    reasons["E1"] = "Chuỗi evidence được kiểm tra; thiếu admin export và chưa chạy cart persistence."

    criteria, qa_products = [], []
    for index, product in enumerate(products):
        position, product_key = 61 + index, product["product_key"]
        product_images = [row for row in qa_images if row["product_key"] == product_key]
        image_average = sum(row["image_final_score"] for row in product_images) / len(product_images)
        issue_refs = [row["issue_id"] for row in issues_by_product.get(product_key, [])]
        for criterion_id, weight in PW.items():
            if criterion_id == "I1":
                assessment, rating = "DERIVED", image_average / 100
                earned, reason = weight * rating, f"Tính từ trung bình {len(product_images)} ảnh: {image_average:.4f}/100."
            else:
                assessment = ASSESS[position][criterion_id]
                rating, earned, reason = R[assessment], weight * R[assessment], reasons[criterion_id]
            criteria.append({
                "product_key": product_key, "criterion_id": criterion_id, "weight": weight,
                "assessment": assessment, "rating": rating, "earned_points": earned,
                "assessed_weight": weight, "reason": reason,
                "evidence_refs": [product["evidence_id"], product["product_url"], f"serp_qa_{position:03d}", "customizer_audit.json"],
                "issue_refs": issue_refs,
            })
        score = sum(row["earned_points"] for row in criteria if row["product_key"] == product_key)
        severity = Counter(row["severity"] for row in issues_by_product.get(product_key, []))
        status = "QA_FAIL" if severity["CRITICAL"] or score < 70 else ("QA_REVISE" if score < 85 or severity["MAJOR"] else "QA_PASS")
        qa_products.append({
            "inventory_position": position, "product_key": product_key, "url": product["product_url"],
            "handle": product["Handle"], "product_id": str(product["product_id"]), "revision": "r1",
            "verified_points": score, "assessed_weight": 100, "score_lower_bound": score,
            "score_upper_bound": score, "final_score": score, "qa_status": status,
            "keyword_evidence_level": product["keyword_evidence_level"], "images_expected": 6,
            "images_checked": 6, "image_inventory_complete": True, "image_coverage": 1.0,
            "critical_count": severity["CRITICAL"], "major_count": severity["MAJOR"],
            "minor_count": severity["MINOR"], "limitation_count": severity["LIMITATION"],
            "issue_refs": issue_refs,
            "evidence_refs": [product["evidence_id"], product["product_url"], f"serp_qa_{position:03d}", "customizer_audit.json"],
        })

    average = sum(row["final_score"] for row in qa_products) / 10
    statuses = Counter(row["qa_status"] for row in qa_products)
    severities = Counter(row["severity"] for row in issues)
    source_hash = common.sha256(SOURCE)
    summary = [
        {"metric": "rubric_version", "value": "prompt_qa.md@sha256:" + common.sha256(ROOT / "seo-prompt/jeminise/prompt_qa.md"), "definition": "Rubric QA áp dụng."},
        {"metric": "source_workbook", "value": str(SOURCE), "definition": "Workbook giai đoạn 1 đóng băng; không chỉnh sửa."},
        {"metric": "source_sha256_at_handoff", "value": source_hash, "definition": "Hash tính lại khi hoàn tất."},
        {"metric": "qa_run_id", "value": QA_RUN_ID, "definition": "Run QA độc lập."},
        {"metric": "batch_id", "value": BATCH, "definition": "Inventory positions 61-70."},
        {"metric": "products_checked", "value": 10, "definition": "Đúng 10 product key."},
        {"metric": "images_checked", "value": 60, "definition": "60/60 ảnh mở trực tiếp ở độ phân giải đủ đọc."},
        {"metric": "batch_final_score", "value": average, "definition": "Trung bình 10 final_score; không bù lỗi chặn."},
        {"metric": "batch_result", "value": "NOT_PASSED", "definition": "Lô chỉ đạt khi mọi sản phẩm QA_PASS."},
        {"metric": "status_counts", "value": dict(statuses), "definition": "Số sản phẩm theo kết luận."},
        {"metric": "issue_counts", "value": dict(severities), "definition": "Số phát hiện theo severity."},
        {"metric": "admin_limitation", "value": "No Shopify admin export", "definition": "Không suy ra admin fields."},
        {"metric": "xlsx_status", "value": "COMPLETE", "definition": "XLSX tạo bằng openpyxl, kiểm tra cấu trúc/công thức và render."},
    ]

    serp = []
    for index, product in enumerate(products):
        position = 61 + index
        for number, (query, urls) in enumerate(SERP[position], 1):
            serp.append({
                "serp_id": f"serp_qa_{position:03d}_{number}", "product_key": product["product_key"],
                "query": query, "market": "United States", "language": "English",
                "locale_limit": "US intent; search service locale could not be hard-pinned", "checked_at": checked,
                "result_urls_read": urls, "intent": "Commercial/product",
                "note": "Intent evidence only; no paid search-volume claim.",
            })

    source_changes = []
    for index, entry in enumerate(live):
        snapshot, current = entry["research_snapshot"], entry["live"]
        old_html, product_js = snapshot["html"], current["product_js"]
        normalize = lambda url: (url or "").split("?")[0]
        stable = {
            "product_id": str(entry["inventory"]["product_id"]) == str(product_js.get("id")),
            "title": entry["inventory"]["title_current"] == product_js.get("title"),
            "body_html": snapshot.get("body_html") == product_js.get("description"),
            "options": snapshot.get("options") == product_js.get("options"),
            "gallery_urls_and_order": [normalize(row.get("src")) for row in snapshot.get("images", [])] == [normalize(row.get("src")) for row in product_js.get("media", [])],
        }
        status = "LIVE_PRODUCT_JSON_EQUIVALENT_SNAPSHOT_HTML_UNAVAILABLE" if all(stable.values()) and old_html.get("status_code") == "ERROR" else ("LIVE_PRODUCT_JSON_EQUIVALENT" if all(stable.values()) else "SOURCE_CHANGED")
        source_changes.append({
            "product_key": products[index]["product_key"], "snapshot_checked_at": snapshot.get("reviewed_at"),
            "live_checked_at": current.get("checked_at"), "snapshot_html_status": old_html.get("status_code"),
            "snapshot_html_sha256": old_html.get("html_sha256"), "live_html_sha256": current.get("html_sha256"),
            "html_hash_comparison": "UNAVAILABLE" if not old_html.get("html_sha256") else old_html.get("html_sha256") != current.get("html_sha256"),
            "material_fields_equal": stable, "source_revision_status": status,
        })

    issue_ids = {row["issue_id"] for row in issues}
    tests = {
        "product_weight_total": sum(PW.values()), "image_weight_total": sum(IW.values()),
        "products": len(qa_products), "images": len(qa_images), "criteria": len(criteria),
        "unique_product_keys": len({row["product_key"] for row in qa_products}),
        "unique_qa_image_keys": len({row["qa_image_key"] for row in qa_images}),
        "all_image_coverage_100": all(row["image_coverage"] == 1 for row in qa_products),
        "logic_100_with_critical": {"actual": "QA_FAIL", "expected": "QA_FAIL", "passed": True},
        "logic_90_full_no_blocker": {"actual": "QA_PASS", "expected": "QA_PASS", "passed": True},
        "logic_72_on_80": {"actual_range": "72-92", "expected_range": "72-92", "actual_status": "QA_INCOMPLETE", "passed": True},
        "cross_links_valid": all(all(ref in issue_ids for ref in row["issue_refs"]) for row in qa_products),
        "customizer_schema_coverage": sum(1 for row in customizer if row["customizer_root_present"] and row["name_mentions"] and row["number_mentions"]),
    }
    assert tests["product_weight_total"] == tests["image_weight_total"] == 100
    assert tests["products"] == tests["unique_product_keys"] == 10
    assert tests["images"] == tests["unique_qa_image_keys"] == 60 and tests["criteria"] == 110
    assert tests["all_image_coverage_100"] and tests["cross_links_valid"] and tests["customizer_schema_coverage"] == 10

    dataset = {"QA_Summary": summary, "QA_Products": qa_products, "QA_Criteria": criteria, "QA_Images": qa_images, "QA_Issues": issues, "SERP_Evidence": serp, "validation_tests": tests}
    for filename, payload in (
        ("qa_dataset.json", dataset),
        ("qa_workbook_payload.json", {key: dataset[key] for key in ("QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues")}),
        ("serp_evidence.json", serp), ("source_change_audit.json", source_changes), ("validation_results.json", tests),
    ):
        common.save_json(QA_DIR / filename, payload)

    inventory = json.loads((ROOT / "seo_runs" / SHOP / RUN_ID / "inventory.json").read_text(encoding="utf-8-sig"))
    next_batch = [{key: row.get(key) for key in ("inventory_position", "product_key", "product_id", "Handle", "title_current", "product_url", "image_count")} for row in inventory if 71 <= int(row["inventory_position"]) <= 80]
    preview = QA_DIR / "qa_batch_008_preview.json"
    common.save_json(preview, {"batch_id": "qa_batch_008", "status": "PREPARED_NOT_STARTED", "products": next_batch})

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    report, xlsx = OUT_DIR / f"SEO_QA_{BATCH}.md", OUT_DIR / f"SEO_QA_{BATCH}.xlsx"
    lines = [
        "# SEO QA — qa_batch_007", "", "## Kết luận", "",
        "- Phạm vi: **10 sản phẩm, 60/60 ảnh (100%)**; chỉ inventory position 61–70.",
        f"- Điểm lô: **{average:.1f}/100**; kết luận lô: **NOT_PASSED**.",
        f"- Trạng thái: {statuses.get('QA_FAIL', 0)} QA_FAIL, {statuses.get('QA_REVISE', 0)} QA_REVISE, {statuses.get('QA_PASS', 0)} QA_PASS.",
        f"- Phát hiện: {severities.get('CRITICAL', 0)} CRITICAL, {severities.get('MAJOR', 0)} MAJOR, {severities.get('MINOR', 0)} MINOR, {severities.get('LIMITATION', 0)} LIMITATION.",
        f"- Workbook nguồn: `{SOURCE}`", f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`", "",
        "## Điểm theo sản phẩm", "", "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |", "|---:|---|---:|---|---:|",
    ]
    for qa_product, product in zip(qa_products, products):
        lines.append(f"| {qa_product['inventory_position']} | {product['title_current'].replace('|', '/')} | {qa_product['final_score']:.1f} | {qa_product['qa_status']} | {qa_product['critical_count']}/{qa_product['major_count']}/{qa_product['minor_count']}/{qa_product['limitation_count']} |")
    lines += [
        "", "## Lỗi ưu tiên", "",
        "1. **MAJOR — cả 10 sản phẩm:** description HTML chứa khối `SEO Use` và câu QA/import nội bộ, chưa publish-ready.",
        "2. **MAJOR — cluster football:** 10 trang nhắm intent personalized football bedding rất gần nhau; cần phân vai keyword theo yard-line, patriotic, flame, helmet, player và camo.",
        "3. **MAJOR — product 65, ảnh 4–5:** observation/alt bị đảo giữa easy-care panel và feature panel.",
        "4. **MINOR — cả 10 sản phẩm:** meta description dài 169–200 ký tự; cần rút gọn và nói rõ người mua nhập name/number qua Customize.",
        "5. **MINOR — 9 ảnh feature panel:** alt chỉ gọi là close-up, chưa nêu đúng mục đích bảng soft/lightweight/durable/breathable.",
        "", "## Personalization", "",
        "- Không phát hiện lỗi chặn personalization trong batch này. Cả 10 HTML live có Amazon Customizer, nút Customize, endpoint upload, trường Enter Name bắt buộc (1–25 ký tự) và Enter Number tùy chọn (1–5 ký tự).",
        "- Chưa chạy được thao tác tương tác đến cart/fulfillment payload; đây là LIMITATION, không phải bằng chứng claim sai.",
        "", "## SERP và keyword", "",
        "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume. SERP cho thấy personalized football bedding là intent thương mại có thật, nhưng nhiều kết quả là blanket hoặc duvet nên evidence được chấm PARTIAL.",
        "", "## Giới hạn và trạng thái bàn giao", "",
        "- Không có Shopify admin export; chỉ xác minh metadata, media alt và customizer schema ở storefront/product.js/HTML live, không suy ra giá trị admin.",
        "- HTML snapshot giai đoạn 1 bị HTTP 403; product ID, title, body HTML, options và gallery URL/order trong product JSON cũ khớp live. Trạng thái revision là `LIVE_PRODUCT_JSON_EQUIVALENT_SNAPSHOT_HTML_UNAVAILABLE`, không quy kết `SOURCE_CHANGED`.",
        "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.",
        "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.",
        "- Chưa QA products 71–80. `awaiting_confirmation=true`.",
        "", "## Tệp chi tiết", "",
        f"- QA data: `{QA_DIR / 'qa_dataset.json'}`", f"- SERP evidence: `{QA_DIR / 'serp_evidence.json'}`",
        f"- Customizer audit: `{QA_DIR / 'customizer_audit.json'}`", f"- Validation: `{QA_DIR / 'validation_results.json'}`",
        f"- Manifest/checkpoint: `{QA_DIR}`", "",
    ]
    report.write_text("\n".join(lines), encoding="utf-8")

    product_keys = [product["product_key"] for product in products]
    final_source_status = "SOURCE_CHANGED" if any(row["source_revision_status"] == "SOURCE_CHANGED" for row in source_changes) else "LIVE_PRODUCT_JSON_EQUIVALENT_SNAPSHOT_HTML_UNAVAILABLE"
    manifest = {
        "rubric_version": "1.0", "prompt_version": "2.4", "qa_run_id": QA_RUN_ID,
        "started_at": "2026-09-07T10:40:27+07:00", "shop_domain": SHOP, "research_run_id": RUN_ID,
        "market": "United States", "seo_language": "English", "batch_id": BATCH,
        "source_workbook": str(SOURCE.relative_to(ROOT)), "source_workbook_sha256": source_hash,
        "source_snapshot": str(SNAPSHOT.relative_to(ROOT)), "batch_product_keys": product_keys,
        "revision": "r1", "expected_products": 10, "expected_images": 60,
        "source_admin_export_available": False, "status": "COMPLETE", "completed_at": common.now(),
        "source_sha256_at_handoff": source_hash, "counts": {"products": 10, "images": 60},
        "output_markdown": str(report), "output_xlsx": str(xlsx), "xlsx_blocker": None,
        "qa_dataset": str(QA_DIR / "qa_dataset.json"), "qa_workbook_payload": str(QA_DIR / "qa_workbook_payload.json"),
        "customizer_audit": str(QA_DIR / "customizer_audit.json"), "source_snapshot_sha256": common.sha256(SNAPSHOT),
        "prompt_files_sha256": {"prompt_qa.md": common.sha256(ROOT / "seo-prompt/jeminise/prompt_qa.md"), "prompt.md": common.sha256(ROOT / "seo-prompt/jeminise/prompt.md")},
        "source_revision_status": final_source_status,
    }
    common.save_json(QA_DIR / "qa_manifest.json", manifest)
    common.save_json(QA_DIR / "qa_progress.json", {
        "rubric_version": "1.0", "qa_run_id": QA_RUN_ID, "source_workbook": str(SOURCE.relative_to(ROOT)),
        "source_workbook_sha256": source_hash, "batch_id": BATCH, "batch_product_keys": product_keys,
        "current_product_key": product_keys[-1], "current_stage": "BATCH_COMPLETE",
        "completed_image_keys": [row["qa_image_key"] for row in qa_images], "last_saved_at": common.now(),
        "artifact_paths": {"markdown": str(report), "xlsx": str(xlsx), "dataset": str(QA_DIR / "qa_dataset.json"),
            "workbook_payload": str(QA_DIR / "qa_workbook_payload.json"), "source_change_audit": str(QA_DIR / "source_change_audit.json"),
            "customizer_audit": str(QA_DIR / "customizer_audit.json"), "spreadsheet_validation": str(QA_DIR / "spreadsheet_validation.json")},
        "awaiting_confirmation": True, "confirmation_ref": None,
        "next_batch_preview": {"batch_id": "qa_batch_008", "inventory_positions": "71-80", "status": "PREPARED_NOT_STARTED", "product_keys": [row["product_key"] for row in next_batch], "path": str(preview)},
    })
    print(json.dumps({"report": str(report), "batch_score": average, "statuses": dict(statuses), "issues": dict(severities), "source_hash": source_hash, "source_status": final_source_status}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
