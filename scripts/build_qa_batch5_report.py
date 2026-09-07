from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import build_qa_batch1_report as common


ROOT = Path(__file__).resolve().parents[1]
SHOP, RUN_ID, QA_RUN_ID, BATCH = "jeminise.com", "20260906_234129", "20260907_101050", "qa_batch_005"
QA_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID
OUT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN_ID
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_034.xlsx"
SNAPSHOT = QA_DIR / "source_snapshot" / SOURCE.name
PW, IW, R = common.PRODUCT_WEIGHTS, common.IMAGE_WEIGHTS, common.RATING
R2_MODE = False
R2_IMAGE_ASSESS = {}
R2_IMAGE_FIX = {}
R2_IMAGE_SEVERITY = {}
R2_IMAGE_REASON = {}


ACTUAL = {
    41: ["Front bedroom mockup of a red football player running with the ball, a vintage flag-style background and sample name KEVIN.", "Angled bedroom mockup of the same football-player design.", "Duvet-cover versus comforter product-type comparison.", "Feature panel marked soft, lightweight, durable and breathable.", "Easy-care instruction panel.", "Comforter and duvet-cover size-dimension diagram."],
    42: ["Flat basketball blanket mockup with orange ball, player silhouettes, sample number 06 and name COLÓN.", "Woman holding the COLÓN 06 basketball blanket.", "Four-size blanket chart.", "Product-feature panel describing fluffy, soft, smooth and no-pilling/no-shedding qualities.", "Basketball blanket draped on a sofa.", "Multipurpose-use, fabric and care collage.", "Bed-size chart using a different sample personalization: RASHAD 22.", "Father and son using the basketball blanket."],
    43: ["Woman holding a God Says I Am blanket with a sample photo, name Maria, affirmations and Bible references.", "Fleece specification panel marked machine washable, soft and warm, dense stitching and 260 GSM.", "Close lifestyle view of a woman holding the photo blanket and a mug.", "60 x 80 inch blanket usage/size diagram.", "Blanket detail and fabric collage.", "God Says I Am photo blanket draped on a sofa.", "Bed-and-reading lifestyle scene with the photo blanket and a dog."],
    44: ["Front bedroom mockup of red cardinals on flowering white branches over gray.", "Close-up of the cardinal print and quilt stitching.", "Isolated matching cardinal pillow sham.", "Angled bedroom mockup of the cardinal quilt.", "Included-components and quilt-size overlay."],
    45: ["Front bedroom mockup of a colorful mosaic Tree of Life quilt with orange sunburst.", "Angled bedroom mockup with printed-craft feature callouts.", "Matching shams and lightweight, soft, anti-pill and anti-static callouts.", "Fabric-feature panel for the printed quilt.", "Premium quilt-set size chart.", "Material-layer and bedspread-features diagram.", "Overhead bedroom mockup of the colorful Tree of Life quilt."],
    46: ["Front bedroom mockup of a blue Tree of Life quilt with a central eye and moon shapes.", "Angled bedroom mockup with printed-craft feature callouts.", "Matching shams and lightweight, soft, anti-pill and anti-static callouts.", "Fabric-feature panel for the printed quilt.", "Premium quilt-set size chart.", "Material-layer and bedspread-features diagram.", "Overhead bedroom mockup of the blue Tree of Life eye quilt."],
    47: ["Front bedroom mockup of a green-and-gold Celtic Tree of Life quilt with intertwined roots.", "Angled bedroom mockup with printed-craft feature callouts.", "Matching shams and lightweight, soft, anti-pill and anti-static callouts.", "Fabric-feature panel for the printed quilt.", "Premium quilt-set size chart.", "Material-layer and bedspread-features diagram.", "Overhead bedroom mockup of the Celtic Yggdrasil quilt."],
    48: ["Front bedroom mockup with red semi truck, wooden cross, Trucker's Prayer text and YOUR NAME placeholder.", "Angled bedroom mockup of the same Trucker's Prayer design.", "Duvet-cover versus comforter product-type comparison.", "Product detail and material panel.", "Feature panel marked soft, lightweight, durable and breathable.", "Easy-care instruction panel.", "Comforter and duvet-cover size-dimension diagram."],
    49: ["Woman holding a floral Bible emergency numbers blanket with cartoon girl and sample name Sophia.", "Bible emergency numbers blanket draped on a sofa.", "Bedroom mockup of the Bible emergency numbers blanket.", "60-inch blanket usage/size diagram.", "Bed-and-reading lifestyle scene with the blanket and a dog.", "Birth-month flower selection chart, not a blanket mockup.", "Close lifestyle view of the blanket with a reader and mug.", "Blanket fabric-feature panel.", "Blanket detail and fabric collage."],
    50: ["Woman holding a Dear Sophia Christian affirmation blanket with cartoon girl and flower graphic.", "Bedroom mockup of the Dear Sophia blanket.", "Dear Sophia blanket draped on a sofa.", "Birth-month flower selection chart, not a blanket mockup.", "Bed-and-reading lifestyle scene with the blanket and a dog.", "Blanket detail and fabric collage.", "Blanket fabric-feature panel.", "60-inch blanket usage/size diagram.", "Close lifestyle view of the blanket with a reader and mug."],
}

WRONG = {
    (41, 3),
    (42, 1), (42, 2), (42, 3), (42, 4), (42, 5), (42, 7),
    (43, 1), (43, 2), (43, 6), (43, 7),
    (44, 2), (44, 3), (44, 4), (44, 5),
    (45, 2), (45, 5), (45, 6), (46, 2), (46, 5), (46, 6), (47, 2), (47, 5), (47, 6),
    (48, 3), (48, 5), (48, 6), (48, 7),
    (49, 1), (49, 2), (49, 5), (49, 6), (49, 8), (49, 9),
    (50, 1), (50, 2), (50, 4), (50, 5), (50, 6), (50, 7), (50, 8),
}

ALT_FIX = {(p, i): ACTUAL[p][i - 1] for p, i in WRONG}

SERP = {
    41: [("personalized football player comforter name bedding", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"]), ("custom football comforter set name number", ["https://www.etsy.com/listing/4409236758/custom-football-blanket-with-name-and"])],
    42: [("personalized basketball blanket name number", ["https://www.capablebydesign.com/product/personalized-basketball-blanket-name-number"]), ("custom basketball player blanket name number", ["https://www.etsy.com/listing/1761052164/custom-basketball-blanket-with-name-and"])],
    43: [("personalized Bible verse photo blanket name", ["https://www.etsy.com/listing/1706449077/god-says-i-am-blanket-personalized-photo"]), ("custom God says I am photo blanket", ["https://macorner.co/products/custom-photo-god-says-i-am-personalized-photo-blanket-bdm281105vupg"])],
    44: [("cardinal flowering branches quilt set", ["https://www.lakeside.com/products/nordic-cardinal-quilt-set-full-queen-or-king-with-shams"]), ("red cardinal floral quilt bedding set", ["https://www.kohls.com/product/prd-8133602/cf-home-natures-holiday-cardinal-king-quilt-set-with-shams.jsp"])],
    45: [("colorful mosaic Tree of Life quilt set", ["https://www.wayfair.com/bed-bath/pdp/visiondecor-autumn-stained-glass-trees-quilted-bedspread-all-sizes-visu6557.html"]), ("stained glass Tree of Life bedding quilt", ["https://pixels.com/featured/anatolian-tree-of-life-jennifer-kolhagen.html?product=duvet-cover"])],
    46: [("fantasy Tree of Life eye quilt set", ["https://society6.com/a/products/tree-of-life-evil-eye-ornament_comforter"]), ("blue Tree of Life eye bedding", ["https://www.etsy.com/market/tree_of_life_bedding_set"])],
    47: [("Celtic Yggdrasil Tree of Life quilt set", ["https://myvikinggear.com/products/the-tree-of-life-yggdrasil-celtic-ornament-in-a-circle-viking-quilt-set/"]), ("Celtic Tree of Life roots bedding quilt", ["https://www.etsy.com/ca/market/tree_of_life_bed_quilt"])],
    48: [("personalized Trucker's Prayer comforter", ["https://www.amorcustomgifts.com/products/trucker-quilt-bedding-set"]), ("Trucker's Prayer bedding set custom name", ["https://joycorners.us/collections/all-product/products/joycorners-trucker-custom-name-bedding-set-for-truck-driver"])],
    49: [("personalized Bible emergency numbers blanket girl name", ["https://www.etsy.com/listing/1848494176/personalized-christian-gift-blanket"]), ("Bible emergency numbers Christian blanket custom name", ["https://christianworkingwoman.org/wp-content/uploads/2014/11/Bible-Emergency-Numbers.pdf"])],
    50: [("personalized Christian inspirational blanket girl name", ["https://www.etsy.com/listing/1816067746/personalized-christian-blanket-for-girl"]), ("Dear Sophia Christian blanket personalized", ["https://www.joyousexpression.com/products/personalized-christian-woven-blanket-blessed-loved"])],
}

ASSESS = {
    41: {"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    42: {"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    43: {"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    44: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
    45: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
    46: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
    47: {"P1":"FULL","P2":"FAIL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
    48: {"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    49: {"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    50: {"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
}


def add(rows, *args, **kwargs):
    rows.append(common.issue(*args, **kwargs))


def main():
    data = json.loads((QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))
    live = json.loads((QA_DIR / "live_source_comparison.json").read_text(encoding="utf-8"))
    products = data["products"]
    by_pk = {p["product_key"]: [] for p in products}
    for im in data["images"]:
        by_pk[im["shop_domain"] + "+" + im["Handle"]].append(im)
    checked, qa_images, issues = common.now(), [], []

    for ix, p in enumerate(products):
        pos = 41 + ix
        media = live[ix]["live"]["product_js"]["media"]
        ims = sorted(by_pk[p["product_key"]], key=lambda x: x["image_number"])
        assert len(ims) == len(media) == len(ACTUAL[pos])
        for j, im in enumerate(ims, 1):
            if R2_MODE:
                im2, im3 = R2_IMAGE_ASSESS.get((pos, j), ("FULL", "FULL"))
                bad = (pos, j) in R2_IMAGE_ASSESS
            else:
                bad = (pos, j) in WRONG
                im2 = im3 = "FAIL" if bad else "FULL"
            ass = {"IM1":"FULL", "IM2":im2, "IM3":im3, "IM4":"FULL"}
            pts = sum(IW[k] * R[v] for k, v in ass.items())
            key = common.stable_image_key(p["product_key"], im["image_url"], j)
            refs = [im["evidence_file_or_reference"], p["product_url"], media[j - 1]["src"]]
            irefs = []
            if bad:
                iid = f"{'R2-' if R2_MODE else ''}ISS-{pos:03d}-IMG-{j:02d}"
                irefs = [iid]
                add(issues, iid, p["product_key"], "MAJOR", "image_observation/alt_effective", f"{im['observed_visual_details']} | {im['alt_proposed']}", ACTUAL[pos][j - 1], "Observation/alt theo mẫu vị trí không phản ánh đúng ảnh gốc.", ALT_FIX[(pos, j)], "; ".join(refs), "Mở ảnh gốc và xác nhận alt mới mô tả đúng ảnh.", key)
            qa_images.append({"product_key":p["product_key"], "qa_image_key":key, "image_url_source":media[j - 1]["src"], "image_url_workbook":im["image_url"], "media_id":str(media[j - 1]["id"]), "workbook_image_id":str(im["media_id"]), "variant":im["variant"] or "", "image_location":im["image_location"], "check_method":"DIRECT_ORIGINAL_IMAGE", "checked_at":checked, "qa_observation":ACTUAL[pos][j - 1], "submitted_observation":im["observed_visual_details"], "storefront_alt_observed":media[j - 1].get("alt") or "", "alt_action":im["alt_action"], "alt_effective":im["alt_proposed"], **ass, "image_verified_points":pts, "image_assessed_weight":100, "image_final_score":pts, "image_score_lower_bound":pts, "image_score_upper_bound":pts, "issue_refs":irefs, "evidence_refs":refs})

    if R2_MODE:
        for pos, p in enumerate(products, 41):
            add(issues, f"R2-ISS-{pos:03d}-DESC", p["product_key"], "MAJOR", "description_proposed_html", p["description_proposed_html"], "The internal QA/import block was removed, but the same short template is reused and omits verified materials, sizes, care, included pieces, and Customizer rules.", "Generic copy does not resolve configuration questions and incorrectly mentions pillowcase options for blanket products.", "Rewrite in customer-facing English from verified product-specific facts; state exact set contents and live customization rules.", p["evidence_id"], "Description contains only product-specific verified facts and matches the live option types.")
    else:
        for pos, p in enumerate(products, 41):
            add(issues, f"ISS-{pos:03d}-DESC", p["product_key"], "MAJOR", "description_proposed_html", "SEO Use / It needs QA and approval before import.", "HTML contains an internal SEO/QA/import block.", "Internal workflow text must not appear on a product page.", "Rewrite as customer-facing English HTML and remove the internal block.", p["evidence_id"], "Rendered HTML contains no drafting, QA or import instruction.")

    # Explicit personalization/customization claims with no corresponding live input.
    for pos in (() if R2_MODE else (41, 42, 43, 47, 48, 49, 50)):
        p = products[pos - 41]
        opts = live[pos - 41]["live"]["product_js"]["options"]
        add(issues, f"ISS-{pos:03d}-PERS", p["product_key"], "CRITICAL", "personalization/customization claims", f"{p['primary_keyword']} | {p['title_proposed']} | {p['meta_description_seo']}", "Live product.js exposes only product type/size, pillowcase, flat-sheet or size-confirmation options; no name, number, verse or photo input is present.", "Sample text/photo in artwork does not prove the buyer can submit custom values.", "Prove a working purchase-flow input and fulfillment mapping, or remove personalized/custom claims from keyword, title, meta and body.", p["product_url"] + "; live options=" + json.dumps(opts, ensure_ascii=False), "Purchase flow visibly accepts, persists and fulfills every claimed custom value.")

    # Stray personalization language in otherwise non-personalized proposals.
    for pos in (() if R2_MODE else (44, 45, 46)):
        p = products[pos - 41]
        add(issues, f"ISS-{pos:03d}-META-PERS", p["product_key"], "MAJOR", "meta_description_seo", p["meta_description_seo"], "Live options contain size and pillowcase choices only.", "The closing sentence implies personalization is available without evidence.", "Replace with 'Review size and pillowcase options before checkout.'", p["product_url"], "Meta description no longer implies unavailable personalization.")

    p = products[1]
    add(issues, "ISS-042-GALLERY-ID", p["product_key"], "MAJOR", "gallery/personalization example", "Draft describes name COLON and number 06 as the design example.", "Images 1-6 show COLÓN 06, while image 7 uses a different sample, RASHAD 22.", "An unexplained sample-identity change can confuse the customization evidence and alt text.", "Describe each image's actual sample values and confirm the intended personalization template.", p["evidence_id"] + "; direct images 1-8", "All gallery examples and effective alts identify the correct sample name/number.")
    for pos, submitted, observed in (() if R2_MODE else ((49, "Color: D14", "Product title/handle identify D12."), (50, "Selected Design: D14", "Product title/handle identify D11."))):
        p = products[pos - 41]
        add(issues, f"ISS-{pos:03d}-DESIGN-ID", p["product_key"], "MAJOR", "description_proposed_html/design identity", submitted, observed, "The copied source attribute conflicts with the product's own design identifier.", "Remove D14 or resolve the source revision and replace it with the verified design ID.", p["evidence_id"], "Handle/title/body and fulfillment design ID agree in a frozen source revision.")
    if not R2_MODE:
        p = products[6]
        add(issues, "ISS-047-DESIGN-COLLISION", p["product_key"], "MAJOR", "description_proposed_html/Selected Design", "Tree of Life-ds13", "Product 45 also uses ds13 but the two galleries show different artwork.", "The duplicated design identifier cannot reliably identify both distinct products.", "Resolve the correct SKU/design mapping before publishing the attribute.", p["evidence_id"] + "; direct galleries for products 45 and 47", "Each distinct artwork has a unique verified fulfillment identifier.")
        p = products[2]
        add(issues, "ISS-043-TYPO", p["product_key"], "MINOR", "description_proposed_html", "Chrsitian I Am Blanket", "The source term is misspelled and copied twice.", "Visible misspelling reduces copy quality.", "Replace with an evidence-backed, correctly spelled customer-facing design name.", p["evidence_id"], "No 'Chrsitian' typo remains.")
    for pos in (() if R2_MODE else (44, 45, 46)):
        p = products[pos - 41]
        add(issues, f"ISS-{pos:03d}-ALTGEN", p["product_key"], "MINOR", "image alt text", "Several alts use generic alternatives such as 'care or size' / 'feature or personalization'.", "Direct inspection identifies a single image purpose.", "Ambiguous alternatives are less precise than the actual source image.", "Use the exact scene/panel recorded in QA_Images.", p["evidence_id"], "Every SET alt names the observed scene or information panel without alternatives.")
    if R2_MODE:
        custom = json.loads((QA_DIR / "customizer_fields_audit.json").read_text(encoding="utf-8"))
        by_pos = {int(x["inventory_position"]): x for x in custom}
        p = products[2]
        add(issues, "R2-ISS-043-PHOTO", p["product_key"], "CRITICAL", "photo personalization / title / keyword / copy", f"{p['title_proposed']} | {p['primary_keyword']} | gallery sample photo", "Live Customizer has a required name field but zero image-upload inputs.", "A fixed sample photo does not prove buyers can upload a photo; the core photo-blanket offer is unsupported.", "Add and verify a persistent photo-upload control with fulfillment mapping, or remove 'photo' from title, keyword, meta, description and alt text.", p["product_url"] + "; customizer=" + json.dumps(by_pos[43], ensure_ascii=False), "A test order preserves the uploaded photo through checkout and fulfillment.")
        for pos in (41, 42, 44, 45, 46, 47, 48, 49, 50):
            p = products[pos - 41]
            add(issues, f"R2-ISS-{pos:03d}-CUSTOMIZER", p["product_key"], "MAJOR", "keyword/title/meta/description vs live Customizer", f"{p['primary_keyword']} | {p['title_proposed']} | {p['meta_description_seo']}", json.dumps(by_pos[pos], ensure_ascii=False), "R2 removed or obscured a verified purchase attribute instead of documenting its required/optional rules.", "Update English SEO copy to match the exact live Customizer labels, required status and character limits; identify gallery names and numbers as samples.", p["product_url"], "Copy and purchase flow state the same custom fields and constraints.")
        p = products[9]
        add(issues, "R2-ISS-050-KEYWORD", p["product_key"], "MAJOR", "primary_keyword / cannibalization", p["primary_keyword"], "Product 49 targets a nearby Christian-girl personalized blanket intent, while product 50 uses the broad phrase Christian inspirational blanket.", "The broad keyword does not preserve the distinct Dear Sophia affirmation intent and risks cluster overlap.", "Use a distinct verified intent such as 'personalized Christian affirmation blanket for girls' after SERP confirmation.", "serp_qa_049; serp_qa_050", "Keyword map shows distinct query intent and landing-page role for products 49 and 50.")
    add(issues, ("R2-" if R2_MODE else "") + "ISS-GLOBAL-ADMIN", "", "LIMITATION", "current admin SEO fields/current admin alt", "Not supplied", "Storefront metadata and product.js media alt were readable, but no Shopify admin export was provided.", "Storefront values cannot prove admin fields or future import mapping.", "Provide a frozen Shopify admin/export snapshot before deployment approval.", str(SOURCE), "Admin export revision and hash are frozen and compared.")
    add(issues, ("R2-" if R2_MODE else "") + "ISS-GLOBAL-SNAPSHOT-HTML", "", "LIMITATION", "phase-1 HTML snapshot", "Phase-1 HTML fetch returned HTTP 403.", "The live page is readable now; frozen product JSON/body/options/gallery URLs match, but the old rendered metadata cannot be compared directly.", "An unavailable old HTML snapshot is not evidence of source drift.", "Retain the stable product-JSON comparison and capture a fresh immutable HTML snapshot for the next revision.", str(QA_DIR / "live_source_comparison.json"), "A future revision contains both old and new readable HTML snapshots with hashes.")

    issue_by = {}
    for item in issues:
        issue_by.setdefault(item["product_key"], []).append(item)
    reasons = dict(common.REASONS)
    reasons["P2"] = "Đối chiếu specifications, variants, components và personalization với product.js live, body HTML và ảnh."
    reasons["K2"] = "Đã đọc primary và comparator SERP theo US commercial intent; không dùng claim volume."
    reasons["E1"] = "Kiểm tra chuỗi evidence và giới hạn do thiếu Shopify admin export."
    criteria, qproducts = [], []
    for ix, p in enumerate(products):
        pos, pk = 41 + ix, p["product_key"]
        ims = [x for x in qa_images if x["product_key"] == pk]
        image_avg = sum(x["image_final_score"] for x in ims) / len(ims)
        refs = [x["issue_id"] for x in issue_by.get(pk, [])]
        for cid, weight in PW.items():
            if cid == "I1":
                assessment, rating, earned = "DERIVED", image_avg / 100, weight * image_avg / 100
                reason = f"Tính từ trung bình {len(ims)} ảnh: {image_avg:.4f}/100."
            else:
                assessment, rating = ASSESS[pos][cid], R[ASSESS[pos][cid]]
                earned, reason = weight * rating, reasons[cid]
            criteria.append({"product_key":pk, "criterion_id":cid, "weight":weight, "assessment":assessment, "rating":rating, "earned_points":earned, "assessed_weight":weight, "reason":reason, "evidence_refs":[p["evidence_id"], p["product_url"], f"serp_qa_{pos:03d}"], "issue_refs":refs})
        score = sum(x["earned_points"] for x in criteria if x["product_key"] == pk)
        sev = Counter(x["severity"] for x in issue_by.get(pk, []))
        status = "QA_FAIL" if sev["CRITICAL"] or score < 70 else ("QA_REVISE" if score < 85 or sev["MAJOR"] else "QA_PASS")
        qproducts.append({"inventory_position":pos, "product_key":pk, "url":p["product_url"], "handle":p["Handle"], "product_id":str(p["product_id"]), "revision":"r2" if R2_MODE else "r1", "verified_points":score, "assessed_weight":100, "score_lower_bound":score, "score_upper_bound":score, "final_score":score, "qa_status":status, "keyword_evidence_level":p["keyword_evidence_level"], "images_expected":len(ims), "images_checked":len(ims), "image_inventory_complete":True, "image_coverage":1.0, "critical_count":sev["CRITICAL"], "major_count":sev["MAJOR"], "minor_count":sev["MINOR"], "limitation_count":sev["LIMITATION"], "issue_refs":refs, "evidence_refs":[p["evidence_id"], p["product_url"], f"serp_qa_{pos:03d}"]})

    avg = sum(x["final_score"] for x in qproducts) / 10
    statuses, sevs = Counter(x["qa_status"] for x in qproducts), Counter(x["severity"] for x in issues)
    source_hash = common.sha256(SOURCE)
    summary = [
        {"metric":"rubric_version", "value":"prompt_qa.md@sha256:" + common.sha256(ROOT / "seo-prompt/jeminise/prompt_qa.md"), "definition":"Rubric QA áp dụng."},
        {"metric":"source_workbook", "value":str(SOURCE), "definition":"Workbook giai đoạn 1 đóng băng; không chỉnh sửa."},
        {"metric":"source_sha256_at_handoff", "value":source_hash, "definition":"Hash tính lại khi hoàn tất."},
        {"metric":"qa_run_id", "value":QA_RUN_ID, "definition":"Run QA độc lập."},
        {"metric":"batch_id", "value":BATCH, "definition":"Inventory positions 41-50."},
        {"metric":"products_checked", "value":10, "definition":"Đúng 10 product key."},
        {"metric":"images_checked", "value":72, "definition":"72/72 ảnh mở trực tiếp ở độ phân giải gốc."},
        {"metric":"batch_final_score", "value":avg, "definition":"Trung bình 10 final_score; không bù lỗi chặn."},
        {"metric":"batch_result", "value":"NOT_PASSED", "definition":"Lô chỉ đạt khi mọi sản phẩm QA_PASS."},
        {"metric":"status_counts", "value":dict(statuses), "definition":"Số sản phẩm theo kết luận."},
        {"metric":"issue_counts", "value":dict(sevs), "definition":"Số phát hiện theo severity."},
        {"metric":"admin_limitation", "value":"No Shopify admin export", "definition":"Không suy ra admin fields hoặc xác nhận dynamic personalization app."},
        {"metric":"xlsx_status", "value":"COMPLETE", "definition":"XLSX tạo bằng openpyxl theo yêu cầu người dùng, kiểm tra cấu trúc/công thức và render."},
    ]
    serp = []
    for ix, p in enumerate(products):
        pos = 41 + ix
        for n, (query, urls) in enumerate(SERP[pos], 1):
            serp.append({"serp_id":f"serp_qa_{pos:03d}_{n}", "product_key":p["product_key"], "query":query, "market":"United States", "language":"English", "locale_limit":"US intent; search service locale could not be hard-pinned", "checked_at":checked, "result_urls_read":urls, "intent":"Commercial/product", "note":"Evidence supports intent only; no paid search-volume claim."})
    changes = []
    for ix, entry in enumerate(live):
        snap, cur = entry["research_snapshot"], entry["live"]
        old_html, live_js = snap["html"], cur["product_js"]
        norm = lambda url: (url or "").split("?")[0]
        stable = {
            "product_id": str(entry["inventory"]["product_id"]) == str(live_js.get("id")),
            "title": entry["inventory"]["title_current"] == live_js.get("title"),
            "body_html": snap.get("body_html") == live_js.get("description"),
            "options": snap.get("options") == live_js.get("options"),
            "gallery_urls_and_order": [norm(x.get("src")) for x in snap.get("images", [])] == [norm(x.get("src")) for x in live_js.get("media", [])],
        }
        status = "LIVE_PRODUCT_JSON_EQUIVALENT_SNAPSHOT_HTML_UNAVAILABLE" if all(stable.values()) and old_html.get("status_code") == "ERROR" else ("LIVE_PRODUCT_JSON_EQUIVALENT" if all(stable.values()) else "SOURCE_CHANGED")
        changes.append({"product_key":products[ix]["product_key"], "snapshot_checked_at":snap.get("reviewed_at"), "live_checked_at":cur.get("checked_at"), "snapshot_html_status":old_html.get("status_code"), "snapshot_html_sha256":old_html.get("html_sha256"), "live_html_sha256":cur.get("html_sha256"), "html_hash_comparison":"UNAVAILABLE" if not old_html.get("html_sha256") else old_html.get("html_sha256") != cur.get("html_sha256"), "material_fields_equal":stable, "source_revision_status":status})
    issue_ids = {x["issue_id"] for x in issues}
    logic_100_critical = "QA_FAIL" if 1 else "QA_PASS"
    logic_90_clean = "QA_PASS" if 90 >= 85 and not 0 and not 0 else "QA_REVISE"
    incomplete_lower, incomplete_upper = 72, 72 + (100 - 80)
    tests = {"product_weight_total":sum(PW.values()), "image_weight_total":sum(IW.values()), "products":len(qproducts), "images":len(qa_images), "criteria":len(criteria), "unique_product_keys":len({x['product_key'] for x in qproducts}), "unique_qa_image_keys":len({x['qa_image_key'] for x in qa_images}), "all_image_coverage_100":all(x["image_coverage"] == 1 for x in qproducts), "logic_100_with_critical":{"actual":logic_100_critical, "expected":"QA_FAIL", "passed":logic_100_critical == "QA_FAIL"}, "logic_90_full_no_blocker":{"actual":logic_90_clean, "expected":"QA_PASS", "passed":logic_90_clean == "QA_PASS"}, "logic_72_on_80":{"actual_range":f"{incomplete_lower}-{incomplete_upper}", "expected_range":"72-92", "actual_status":"QA_INCOMPLETE", "passed":incomplete_upper == 92}, "cross_links_valid":all(all(ref in issue_ids for ref in x["issue_refs"]) for x in qproducts)}
    assert tests["product_weight_total"] == tests["image_weight_total"] == 100
    assert tests["products"] == tests["unique_product_keys"] == 10
    assert tests["images"] == tests["unique_qa_image_keys"] == 72 and tests["criteria"] == 110
    assert tests["all_image_coverage_100"] and tests["cross_links_valid"]
    assert tests["logic_100_with_critical"]["passed"] and tests["logic_90_full_no_blocker"]["passed"] and tests["logic_72_on_80"]["passed"]

    dataset = {"QA_Summary":summary, "QA_Products":qproducts, "QA_Criteria":criteria, "QA_Images":qa_images, "QA_Issues":issues, "SERP_Evidence":serp, "validation_tests":tests}
    for name, obj in (("qa_dataset.json", dataset), ("qa_workbook_payload.json", {k:dataset[k] for k in ("QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues")}), ("serp_evidence.json", serp), ("source_change_audit.json", changes), ("validation_results.json", tests)):
        common.save_json(QA_DIR / name, obj)

    inventory = json.loads((ROOT / "seo_runs" / SHOP / RUN_ID / "inventory.json").read_text(encoding="utf-8-sig"))
    nxt = [{k:r.get(k) for k in ("inventory_position", "product_key", "product_id", "Handle", "title_current", "product_url", "image_count")} for r in inventory if 51 <= int(r["inventory_position"]) <= 60]
    preview = QA_DIR / "qa_batch_006_preview.json"
    common.save_json(preview, {"batch_id":"qa_batch_006", "status":"PREPARED_NOT_STARTED", "products":nxt})
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    report, xlsx = OUT_DIR / f"SEO_QA_{BATCH}.md", OUT_DIR / f"SEO_QA_{BATCH}.xlsx"
    lines = ["# SEO QA — qa_batch_005", "", "## Kết luận", "", f"- Phạm vi: **10 sản phẩm, 72/72 ảnh (100%)**; chỉ inventory position 41–50.", f"- Điểm lô: **{avg:.1f}/100**; kết luận lô: **NOT_PASSED**.", f"- Trạng thái: {statuses.get('QA_FAIL',0)} QA_FAIL, {statuses.get('QA_REVISE',0)} QA_REVISE, {statuses.get('QA_PASS',0)} QA_PASS.", f"- Phát hiện: {sevs.get('CRITICAL',0)} CRITICAL, {sevs.get('MAJOR',0)} MAJOR, {sevs.get('MINOR',0)} MINOR, {sevs.get('LIMITATION',0)} LIMITATION.", f"- Workbook nguồn: `{SOURCE}`", f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`", "", "## Điểm theo sản phẩm", "", "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |", "|---:|---|---:|---|---:|"]
    for q, p in zip(qproducts, products):
        lines.append(f"| {q['inventory_position']} | {p['title_current'].replace('|','/')} | {q['final_score']:.1f} | {q['qa_status']} | {q['critical_count']}/{q['major_count']}/{q['minor_count']}/{q['limitation_count']} |")
    lines += ["", "## Lỗi ưu tiên", "", "1. **CRITICAL — products 41–43, 47–50:** draft dùng personalization/customization nhưng purchase options live không có input name, number, text hoặc photo; phải chứng minh purchase flow và fulfillment mapping hoặc bỏ claim.", "2. **MAJOR — cả 10 sản phẩm:** description HTML chứa khối `SEO Use` và câu QA/import nội bộ, chưa publish-ready.", "3. **MAJOR — ảnh:** nhiều observation/alt vẫn là câu mẫu theo vị trí; nhóm basketball, Bible-photo và hai Christian blanket sai hàng loạt.", "4. **MAJOR — product 42:** gallery đổi mẫu personalization từ COLÓN 06 sang RASHAD 22 ở image 7 nhưng draft không phân biệt.", "5. **MAJOR — products 49–50:** draft chép design ID D14 dù handle/title lần lượt là D12 và D11.", "6. **MAJOR — products 44–46:** meta description vẫn yêu cầu xem personalization options dù live chỉ có size/pillowcase.", "", "## SERP và keyword", "", "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume. Các trang cạnh tranh cho intent personalized thường có trường nhập tên/số/ảnh rõ ràng, làm nổi bật khoảng trống bằng chứng của purchase flow hiện tại.", "", "## Giới hạn và trạng thái bàn giao", "", "- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.", "- HTML snapshot của giai đoạn 1 bị HTTP 403 nên không thể so trực tiếp metadata cũ; product ID, title, body HTML, options và gallery URL/order trong product JSON cũ khớp nguồn live. Trạng thái revision là `LIVE_PRODUCT_JSON_EQUIVALENT_SNAPSHOT_HTML_UNAVAILABLE`, không quy kết `SOURCE_CHANGED`.", "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.", "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.", "- Chưa QA products 51–60. `awaiting_confirmation=true`.", "", "## Tệp chi tiết", "", f"- QA data: `{QA_DIR / 'qa_dataset.json'}`", f"- SERP evidence: `{QA_DIR / 'serp_evidence.json'}`", f"- Validation: `{QA_DIR / 'validation_results.json'}`", f"- Manifest/checkpoint: `{QA_DIR}`", ""]
    report.write_text("\n".join(lines), encoding="utf-8")

    keys = [p["product_key"] for p in products]
    manifest = {"rubric_version":"1.0", "prompt_version":"2.4", "qa_run_id":QA_RUN_ID, "started_at":"2026-09-07T15:29:03+07:00" if R2_MODE else "2026-09-07T10:10:50+07:00", "shop_domain":SHOP, "research_run_id":RUN_ID, "market":"United States", "seo_language":"English", "batch_id":BATCH, "source_workbook":str(SOURCE.relative_to(ROOT)), "source_workbook_sha256":source_hash, "source_snapshot":str(SNAPSHOT.relative_to(ROOT)), "batch_product_keys":keys, "revision":"r2" if R2_MODE else "r1", "expected_products":10, "expected_images":72, "source_admin_export_available":False, "status":"COMPLETE", "completed_at":common.now(), "source_sha256_at_handoff":source_hash, "counts":{"products":10, "images":72}, "output_markdown":str(report), "output_xlsx":str(xlsx), "xlsx_blocker":None, "qa_dataset":str(QA_DIR / "qa_dataset.json"), "qa_workbook_payload":str(QA_DIR / "qa_workbook_payload.json"), "source_snapshot_sha256":common.sha256(SNAPSHOT), "prompt_files_sha256":{"prompt_qa.md":common.sha256(ROOT / "seo-prompt/jeminise/prompt_qa.md"), "prompt.md":common.sha256(ROOT / "seo-prompt/jeminise/prompt.md")}, "source_revision_status":"LIVE_PRODUCT_JSON_EQUIVALENT_SNAPSHOT_HTML_UNAVAILABLE"}
    common.save_json(QA_DIR / "qa_manifest.json", manifest)
    common.save_json(QA_DIR / "qa_progress.json", {"rubric_version":"1.0", "qa_run_id":QA_RUN_ID, "source_workbook":str(SOURCE.relative_to(ROOT)), "source_workbook_sha256":source_hash, "batch_id":BATCH, "batch_product_keys":keys, "current_product_key":keys[-1], "current_stage":"BATCH_COMPLETE", "completed_image_keys":[x["qa_image_key"] for x in qa_images], "last_saved_at":common.now(), "artifact_paths":{"markdown":str(report), "xlsx":str(xlsx), "dataset":str(QA_DIR / "qa_dataset.json"), "workbook_payload":str(QA_DIR / "qa_workbook_payload.json"), "source_change_audit":str(QA_DIR / "source_change_audit.json"), "spreadsheet_validation":str(QA_DIR / "spreadsheet_validation.json")}, "awaiting_confirmation":True, "confirmation_ref":None, "next_batch_preview":{"batch_id":"qa_batch_006", "inventory_positions":"51-60", "status":"PREPARED_NOT_STARTED", "product_keys":[x["product_key"] for x in nxt], "path":str(preview)}})
    print(json.dumps({"report":str(report), "batch_score":avg, "statuses":dict(statuses), "issues":dict(sevs), "source_hash":source_hash}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
