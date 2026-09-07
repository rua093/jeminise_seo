from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import build_qa_batch1_report as common


ROOT = Path(__file__).resolve().parents[1]
RUN_ID = "20260906_234129"
QA_RUN_ID = "20260907_091545"
QA_BATCH_ID = "qa_batch_002"
QA_DIR = ROOT / "seo_runs" / "jeminise.com" / RUN_ID / "qa" / QA_RUN_ID
OUT_DIR = ROOT / "resutls" / "jeminise.com" / RUN_ID / "qa" / QA_RUN_ID
SOURCE = ROOT / "resutls" / "jeminise.com" / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_034.xlsx"
RATING = common.RATING
PRODUCT_WEIGHTS = common.PRODUCT_WEIGHTS
IMAGE_WEIGHTS = common.IMAGE_WEIGHTS


ACTUAL = {
    11: [
        "Main bed mockup with orange-and-white God Says I Am design, sample name Jessica, carnation flowers, butterflies and Bible-reference affirmations.",
        "Alternate bedroom mockup of the same Jessica Christian affirmation bedding.",
        "Comparison graphic explaining Duvet Cover Set versus Comforter Set.",
        "Feature graphic stating 100% microfiber, 3D printed pattern, soft touch, lightweight and breathable.",
        "Birth Month Flowers chart for January through December.",
        "Easy-care graphic stating wrinkle-free, stain-proof, anti-pilling and won't fade in the wash.",
        "Bedroom close-up with soft, lightweight, durable and breathable feature icons.",
        "Twin, Full, Queen and King dimension chart; heading reads SIZE DIMENTION.",
    ],
    12: [
        "Main bed mockup with white Evelyn Christian bedding, crosses, faith symbols and Bible-reference affirmations.",
        "Alternate bedroom mockup of the same Evelyn Christian affirmation bedding.",
        "Comparison graphic explaining Duvet Cover Set versus Comforter Set.",
        "Feature graphic stating 100% microfiber, 3D printed pattern, soft touch, lightweight and breathable.",
        "Bedroom close-up with soft, lightweight, durable and breathable feature icons.",
        "Easy-care graphic stating wrinkle-free, stain-proof, anti-pilling and won't fade in the wash.",
        "Twin, Full, Queen and King dimension chart; heading reads SIZE DIMENTION.",
    ],
    13: [
        "Main bed mockup with black-and-gold Olivia design, woman silhouette, butterflies, flowers and God Is Within Her text.",
        "Alternate bedroom mockup of the same Olivia Christian bedding.",
        "Comparison graphic explaining Duvet Cover Set versus Comforter Set.",
        "Feature graphic stating 100% microfiber, 3D printed pattern, soft touch, lightweight and breathable.",
        "Bedroom close-up with soft, lightweight, durable and breathable feature icons.",
        "Easy-care graphic stating wrinkle-free, stain-proof, anti-pilling and won't fade in the wash.",
        "Twin, Full, Queen and King dimension chart; heading reads SIZE DIMENTION.",
    ],
    14: [
        "Main bed mockup with blue-and-white God Says I Am design, sample name Emily, floral border and Bible-reference acrostic.",
        "Alternate bedroom mockup of the same Emily Christian bedding.",
        "Comparison graphic explaining Duvet Cover Set versus Comforter Set.",
        "Feature graphic stating 100% microfiber, 3D printed pattern, soft touch, lightweight and breathable.",
        "Bedroom close-up with soft, lightweight, durable and breathable feature icons.",
        "Easy-care graphic stating wrinkle-free, stain-proof, anti-pilling and won't fade in the wash.",
        "Twin, Full, Queen and King dimension chart; heading reads SIZE DIMENTION.",
    ],
    15: [
        "Main bed mockup with brown Charlotte butterfly bedding, God Is Within Her text and Bible-reference acrostic.",
        "Alternate bedroom mockup of the same Charlotte Christian bedding.",
        "Comparison graphic explaining Duvet Cover Set versus Comforter Set.",
        "Feature graphic stating 100% microfiber, 3D printed pattern, soft touch, lightweight and breathable.",
        "Bedroom close-up with soft, lightweight, durable and breathable feature icons.",
        "Easy-care graphic stating wrinkle-free, stain-proof, anti-pilling and won't fade in the wash.",
        "Twin, Full, Queen and King dimension chart; heading reads SIZE DIMENTION.",
    ],
    16: [
        "Main bed mockup with black David design, white cross, lion, armored warrior and Christian affirmation text.",
        "Second front-facing mockup of the same David Christian warrior bedding.",
        "Alternate bedroom mockup of the same David Christian warrior bedding.",
        "Comparison graphic explaining Duvet Cover Set versus Comforter Set.",
        "Feature graphic stating 100% microfiber, 3D printed pattern, soft touch, lightweight and breathable.",
        "Bedroom close-up with soft, lightweight, durable and breathable feature icons.",
        "Easy-care graphic stating wrinkle-free, stain-proof, anti-pilling and won't fade in the wash.",
        "Twin, Full, Queen and King dimension chart; heading reads SIZE DIMENTION.",
    ],
    17: [
        "Front bedroom mockup of a red, navy, olive and cream Christmas-tree patchwork quilt.",
        "Angled bedroom mockup of the same Christmas-tree patchwork quilt.",
        "Isolated matching pillow sham with the Christmas-tree patchwork print.",
        "Close-up of the printed faux-patchwork surface and visible quilting lines.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    18: [
        "Front bedroom mockup of a red-and-green patchwork quilt with Santa, snowmen, reindeer, trees and snowflakes.",
        "Angled bedroom mockup of the same Christmas character quilt.",
        "Close-up of the printed faux-patchwork surface and visible quilting lines.",
        "Isolated matching pillow sham with Santa, snowman, reindeer and tree blocks.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    19: [
        "Front bedroom mockup of a red gingerbread-man quilt with candy cane, holly, pine and floral accents.",
        "Angled bedroom mockup of the same red gingerbread Christmas quilt.",
        "Close-up of the printed gingerbread surface and visible quilting lines.",
        "Isolated matching pillow sham with gingerbread man and candy cane artwork.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    20: [
        "Front bedroom mockup of a light gingerbread Christmas quilt with two cookie figures, gifts, candy canes and holly.",
        "Angled bedroom mockup of the same gingerbread-and-gifts quilt.",
        "Isolated matching pillow sham with two gingerbread figures and wrapped gifts.",
        "Close-up of the printed gingerbread surface and visible quilting lines.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
}

# Each tuple is (IM2 observation accuracy, IM3 effective-alt accuracy).
IMAGE_ASSESS = {
    11: [("FULL","FULL"),("PARTIAL","PARTIAL"),("FAIL","FAIL"),("FAIL","FAIL"),("FAIL","FAIL"),("PARTIAL","FAIL"),("PARTIAL","PARTIAL"),("PARTIAL","PARTIAL")],
    12: [("FULL","FULL"),("PARTIAL","PARTIAL"),("FAIL","FAIL"),("FAIL","FAIL"),("FAIL","FAIL"),("PARTIAL","FAIL"),("FAIL","FAIL")],
    13: [("FULL","FULL"),("PARTIAL","PARTIAL"),("FAIL","FAIL"),("FAIL","FAIL"),("FAIL","FAIL"),("PARTIAL","FAIL"),("FAIL","FAIL")],
    14: [("FULL","FULL"),("PARTIAL","PARTIAL"),("FAIL","FAIL"),("FAIL","FAIL"),("FAIL","FAIL"),("PARTIAL","FAIL"),("FAIL","FAIL")],
    15: [("FULL","FULL"),("PARTIAL","PARTIAL"),("FAIL","FAIL"),("FAIL","FAIL"),("FAIL","FAIL"),("PARTIAL","FAIL"),("FAIL","FAIL")],
    16: [("FULL","FULL"),("PARTIAL","PARTIAL"),("PARTIAL","FAIL"),("FAIL","FAIL"),("FAIL","FAIL"),("FULL","FULL"),("FAIL","FAIL"),("PARTIAL","PARTIAL")],
    17: [("FULL","FULL"),("FULL","FULL"),("FULL","FULL"),("FAIL","FAIL"),("FULL","FULL")],
    18: [("FULL","FULL"),("FULL","FULL"),("FAIL","FAIL"),("FAIL","FAIL"),("FULL","FULL")],
    19: [("FULL","FULL"),("FULL","FULL"),("FAIL","FAIL"),("FAIL","FAIL"),("FULL","FULL")],
    20: [("FULL","FULL"),("FULL","FULL"),("PARTIAL","FULL"),("FAIL","FAIL"),("FULL","FULL")],
}

ALT_FIXES = {
    (11,2): "God Says I Am Christian bedding with Jessica in an alternate bedroom view",
    (11,3): "Duvet cover and comforter set construction comparison",
    (11,4): "God Says I Am bedding microfiber and breathable fabric features",
    (11,5): "Birth month flower options from January through December",
    (11,6): "Bedding easy-care, anti-pilling and fade-resistance features",
    (11,7): "God Says I Am bedding softness and durability feature graphic",
    (11,8): "Christian bedding size chart for Twin through King",
    (12,2): "Evelyn Christian Bible verse bedding in an alternate bedroom view",
    (13,2): "Olivia God Is Within Her bedding in an alternate bedroom view",
    (14,2): "Emily God Says I Am bedding in an alternate bedroom view",
    (15,2): "Charlotte Christian butterfly bedding in an alternate bedroom view",
    (16,2): "David Christian warrior bedding in a front bedroom mockup",
    (16,3): "David Christian warrior bedding in an alternate bedroom view",
    (17,4): "Close-up of Christmas tree patchwork quilt print and quilting",
    (18,3): "Close-up of Santa and snowman Christmas quilt print and quilting",
    (18,4): "Matching Santa and snowman Christmas quilt pillow sham",
    (19,3): "Close-up of gingerbread Christmas quilt print and quilting",
    (19,4): "Matching gingerbread Christmas quilt pillow sham",
    (20,3): "Matching gingerbread gift Christmas quilt pillow sham",
    (20,4): "Close-up of gingerbread gift Christmas quilt print and quilting",
}
for pos in (12,13,14,15):
    ALT_FIXES.update({
        (pos,3): "Duvet cover and comforter set construction comparison",
        (pos,4): "Christian bedding microfiber and breathable fabric features",
        (pos,5): "Christian bedding softness and durability feature graphic",
        (pos,6): "Bedding easy-care, anti-pilling and fade-resistance features",
        (pos,7): "Christian bedding size chart for Twin through King",
    })
ALT_FIXES.update({
    (16,4): "Duvet cover and comforter set construction comparison",
    (16,5): "Christian warrior bedding microfiber and breathable fabric features",
    (16,7): "Bedding easy-care, anti-pilling and fade-resistance features",
    (16,8): "Christian bedding size chart for Twin through King",
})

SERP = {
    11: [("personalized God Says I Am comforter set", ["https://www.walmart.com/ip/19956661087"]), ("God Says I Am Christian bedding set", ["https://www.etsy.com/listing/4355382705/personalized-christian-blanket-god-says"])],
    12: [("personalized Bible verse comforter set", ["https://www.etsy.com/listing/1633105230/woodland-lake-christian-bedding-set"]), ("Christian affirmation comforter set", ["https://www.etsy.com/listing/1633105230/woodland-lake-christian-bedding-set"])],
    13: [("personalized Christian comforter for women", ["https://www.etsy.com/listing/4355382705/personalized-christian-blanket-god-says"]), ("God is within her comforter set", ["https://www.etsy.com/listing/1816067746/personalized-christian-blanket-for-girl"])],
    14: [("blue personalized Christian comforter set", ["https://www.etsy.com/listing/1806586035/custom-bible-verse-blanket-personalized"]), ("God Says I Am blue bedding", ["https://www.etsy.com/listing/4388723618/personalized-god-says-i-am-blanket"])],
    15: [("personalized Christian butterfly comforter", ["https://www.etsy.com/listing/1112803288/butterfly-duvet-cover-set-personalized"]), ("God is within her butterfly bedding", ["https://www.etsy.com/listing/4419478072/personalized-floral-butterfly-christian"])],
    16: [("personalized Christian comforter for men", ["https://www.etsy.com/listing/4449422983/personalized-christian-blanket-for-men"]), ("Christian warrior lion comforter set", ["https://www.walmart.com/ip/19902816540"])],
    17: [("Christmas tree patchwork quilt set", ["https://www.walmart.com/ip/6720910920"]), ("rustic Christmas tree quilt bedding", ["https://www.wayfair.com/pdp/marcielo-3-pcs-holiday-bedspread-patchwork-christmas-tree-quilt-set-mrcl1226.html"])],
    18: [("Santa snowman Christmas quilt set", ["https://www.walmart.com/ip/18020373681"]), ("Christmas character patchwork quilt set", ["https://www.walmart.com/ip/14156114585"])],
    19: [("gingerbread Christmas quilt set", ["https://www.wayfair.com/bed-bath/pdp/greenland-home-fashions-gingerbread-lane-reversible-quilt-set-ghf10327.html"]), ("red gingerbread Christmas bedding set", ["https://www.walmart.com/ip/8379506608"])],
    20: [("gingerbread Christmas bedding set", ["https://www.walmart.com/ip/17457653735"]), ("gingerbread gift Christmas quilt set", ["https://www.wayfair.com/bed-bath/pdp/aimebby-christmas-quilts-set-soft-microfiber-quilt-lightweight-bedspread-winter-coverlet-holiday-decor-and-gift-for-family-bedroom-fozz1002.html"])],
}

PRODUCT_ASSESS = {
    11:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FAIL","T2":"PARTIAL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    12:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FAIL","T2":"PARTIAL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    13:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FAIL","T2":"PARTIAL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    14:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    15:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FAIL","T2":"PARTIAL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    16:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FAIL","T2":"PARTIAL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    17:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    18:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    19:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    20:{"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"PARTIAL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
}

REASONS = dict(common.REASONS)
REASONS["P2"] = "Đối chiếu chất liệu, kích thước, thành phần đi kèm, biến thể và personalization với product.js live, body HTML và toàn bộ gallery."
REASONS["K1"] = "Đối chiếu long-tail với thiết kế riêng và nguy cơ trùng intent trong hai cụm Christian/Christmas của batch."
REASONS["K2"] = "Đã đọc lại hai SERP query cho mỗi sản phẩm tại intent US; ghi URL thực đọc và không suy diễn volume."


def add_issue(rows, *args, **kwargs):
    rows.append(common.issue(*args, **kwargs))


def main() -> None:
    data = json.loads((QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))
    live = json.loads((QA_DIR / "live_source_comparison.json").read_text(encoding="utf-8"))
    products = data["products"]
    pk_to_pos = {p["product_key"]: i + 11 for i, p in enumerate(products)}
    images_by_pk = {p["product_key"]: [] for p in products}
    for im in data["images"]:
        images_by_pk[im["shop_domain"] + "+" + im["Handle"]].append(im)
    checked_at = common.now()
    qa_images, issues = [], []

    for p in products:
        pos = pk_to_pos[p["product_key"]]
        live_media = live[pos - 11]["live"]["product_js"]["media"]
        submitted = sorted(images_by_pk[p["product_key"]], key=lambda x: x["image_number"])
        assert len(submitted) == len(live_media) == len(ACTUAL[pos])
        for idx, im in enumerate(submitted):
            n = idx + 1
            qkey = common.stable_image_key(p["product_key"], im["image_url"], n)
            im2, im3 = IMAGE_ASSESS[pos][idx]
            assessments = {"IM1":"FULL", "IM2":im2, "IM3":im3, "IM4":"FULL"}
            points = sum(IMAGE_WEIGHTS[k] * RATING[v] for k, v in assessments.items())
            refs = [im["evidence_file_or_reference"], p["product_url"], live_media[idx]["src"]]
            row = {
                "product_key":p["product_key"], "qa_image_key":qkey,
                "image_url_source":live_media[idx]["src"], "image_url_workbook":im["image_url"],
                "media_id":str(live_media[idx]["id"]), "workbook_image_id":str(im["media_id"]),
                "variant":im["variant"] or "", "image_location":im["image_location"],
                "check_method":"DIRECT_ORIGINAL_IMAGE", "checked_at":checked_at,
                "qa_observation":ACTUAL[pos][idx], "submitted_observation":im["observed_visual_details"],
                "storefront_alt_observed":live_media[idx].get("alt") or "",
                "alt_action":im["alt_action"], "alt_effective":im["alt_proposed"],
                **assessments, "image_verified_points":points, "image_assessed_weight":100,
                "image_final_score":points, "image_score_lower_bound":points, "image_score_upper_bound":points,
                "issue_refs":[], "evidence_refs":refs,
            }
            if im2 != "FULL" or im3 != "FULL":
                iid = f"ISS-{pos:03d}-IMG-{n:02d}"
                sev = "MAJOR" if im3 == "FAIL" else "MINOR"
                fix = ALT_FIXES.get((pos,n), im["alt_proposed"])
                add_issue(issues, iid, p["product_key"], sev, "image_observation/alt_effective",
                          f"{im['observed_visual_details']} | {im['alt_proposed']}", ACTUAL[pos][idx],
                          "Observation hoặc alt được gán theo vị trí mẫu và không phản ánh đúng ảnh gốc.", fix,
                          "; ".join(refs), "Mở lại ảnh gốc và xác nhận observation/alt mới mô tả đúng ảnh.", qkey)
                row["issue_refs"].append(iid)
            qa_images.append(row)

    # Findings affecting the product draft.
    for p in products:
        pos, pk = pk_to_pos[p["product_key"]], p["product_key"]
        add_issue(issues, f"ISS-{pos:03d}-DESC", pk, "MAJOR", "description_proposed_html",
                  "Review this draft against the live product page and Shopify export before approval.", "HTML kết thúc bằng lời nhắc review nội bộ.",
                  "Nội dung quy trình nội bộ không được xuất hiện trên storefront.",
                  "Rewrite as customer-facing English HTML and remove the internal review sentence.",
                  p["evidence_id"], "Render full HTML and confirm no drafting/QA instruction remains.")
        add_issue(issues, f"ISS-{pos:03d}-ENC", pk, "MINOR", "title_current/H1_current",
                  p["title_current"], live[pos-11]["live"]["h1"], "Title/H1 live chứa ký tự thay thế U+FFFD và từ bị cắt.",
                  p["title_proposed"], p["product_url"], "Rendered title and H1 no longer contain U+FFFD or truncated words.")

    for pos in range(11,17):
        p = products[pos-11]
        opts = live[pos-11]["live"]["product_js"]["options"]
        add_issue(issues, f"ISS-{pos:03d}-PERS", p["product_key"], "CRITICAL",
                  "title/meta/description personalization claim", "Personalized / Personalize / custom name",
                  "Product.js chỉ có type/size, pillowcases và flat-sheet options; không có input tên, verse hoặc birth flower.",
                  "Personalization là thuộc tính mua hàng trọng yếu nhưng không thể thao tác trong purchase flow live.",
                  "Remove personalization claims unless a working name/design input and fulfillment mapping are evidenced.",
                  p["product_url"] + "; live product.js options=" + json.dumps(opts, ensure_ascii=False),
                  "Live purchase flow visibly accepts and preserves every claimed personalization value.")
        add_issue(issues, f"ISS-{pos:03d}-SERP", p["product_key"], "MAJOR", "SERP_evidence_references",
                  data["product_evidence"][pos-11]["SERP_evidence_references"],
                  "Submitted references are dominated by blanket listings while the selected keyword targets a comforter/bedding product.",
                  "Evidence does not cleanly validate the selected product-type intent and six drafts compete in one narrow personalization cluster.",
                  "Use bedding/comforter product SERPs, document differentiation, and rerun keyword selection after resolving personalization.",
                  json.dumps(SERP[pos], ensure_ascii=False), "Two current US-intent product SERPs support the exact product type and unique design angle.")

    p13 = products[2]
    add_issue(issues, "ISS-013-FACT", p13["product_key"], "MAJOR", "description_proposed_html/source facts",
              "Style: Christian Knight Templar", ACTUAL[13][0],
              "Structured style label conflicts with the visible women-focused God Is Within Her design and was copied into the draft.",
              "Remove 'Christian Knight Templar' from customer copy unless the admin source mapping is corrected and evidenced.",
              p13["evidence_id"] + "; " + p13["product_url"], "Revised source mapping and HTML agree with the visible design.")
    p19 = products[8]
    add_issue(issues, "ISS-019-PERS", p19["product_key"], "CRITICAL", "description_proposed_html personalization claim",
              "Customization: 1 text input", "Product.js exposes only quilt size and pillowcase quantity; no text input is present.",
              "The draft states an unavailable purchasing feature.",
              "Remove the customization statement or provide a working text input with fulfillment evidence.",
              p19["product_url"] + "; live product.js options", "Live purchase flow accepts and preserves the claimed text value.")
    add_issue(issues, "ISS-GLOBAL-ADMIN", "", "LIMITATION", "current admin SEO fields/current admin alt",
              "Not supplied", "Storefront metadata and product.js media alt were readable, but no Shopify admin export was provided.",
              "Storefront values cannot prove current admin fields or future import mapping.",
              "Provide a frozen Shopify admin/export snapshot before approval or import mapping.", str(SOURCE),
              "Admin export revision and hash are frozen and compared to storefront/workbook.")

    issue_by_product = {}
    for it in issues:
        issue_by_product.setdefault(it["product_key"], []).append(it)
    criteria, qa_products = [], []
    for p in products:
        pos, pk = pk_to_pos[p["product_key"]], p["product_key"]
        image_rows = [x for x in qa_images if x["product_key"] == pk]
        image_avg = sum(x["image_final_score"] for x in image_rows) / len(image_rows)
        p_refs = [x["issue_id"] for x in issue_by_product.get(pk, [])]
        for cid, weight in PRODUCT_WEIGHTS.items():
            if cid == "I1":
                assessment, rating, earned = "DERIVED", image_avg / 100, weight * image_avg / 100
                reason = f"Tính từ trung bình {len(image_rows)} ảnh: {image_avg:.4f}/100."
            else:
                assessment = PRODUCT_ASSESS[pos][cid]
                rating, earned = RATING[assessment], weight * RATING[assessment]
                reason = REASONS[cid]
            criteria.append({"product_key":pk,"criterion_id":cid,"weight":weight,"assessment":assessment,
                             "rating":rating,"earned_points":earned,"assessed_weight":weight,"reason":reason,
                             "evidence_refs":[p["evidence_id"],p["product_url"],f"serp_qa_{pos:03d}"],"issue_refs":p_refs})
        verified = sum(x["earned_points"] for x in criteria if x["product_key"] == pk)
        sev = Counter(x["severity"] for x in issue_by_product.get(pk, []))
        status = "QA_FAIL" if sev["CRITICAL"] or verified < 70 else ("QA_REVISE" if verified < 85 or sev["MAJOR"] else "QA_PASS")
        qa_products.append({"inventory_position":pos,"product_key":pk,"url":p["product_url"],"handle":p["Handle"],
                            "product_id":str(p["product_id"]),"revision":"r1","verified_points":verified,
                            "assessed_weight":100,"score_lower_bound":verified,"score_upper_bound":verified,"final_score":verified,
                            "qa_status":status,"keyword_evidence_level":p["keyword_evidence_level"],"images_expected":p["image_count"],
                            "images_checked":len(image_rows),"image_inventory_complete":True,"image_coverage":1.0,
                            "critical_count":sev["CRITICAL"],"major_count":sev["MAJOR"],"minor_count":sev["MINOR"],
                            "limitation_count":sev["LIMITATION"],"issue_refs":p_refs,
                            "evidence_refs":[p["evidence_id"],p["product_url"],f"serp_qa_{pos:03d}"]})

    batch_avg = sum(p["final_score"] for p in qa_products) / 10
    statuses = Counter(p["qa_status"] for p in qa_products)
    sev_counts = Counter(x["severity"] for x in issues)
    source_hash = common.sha256(SOURCE)
    summary = [
        {"metric":"rubric_version","value":"prompt_qa.md@sha256:"+common.sha256(ROOT/"seo-prompt/jeminise/prompt_qa.md"),"definition":"Rubric QA áp dụng."},
        {"metric":"source_workbook","value":str(SOURCE),"definition":"Workbook giai đoạn 1 được đóng băng; không chỉnh sửa."},
        {"metric":"source_sha256_at_handoff","value":source_hash,"definition":"Hash tính lại khi hoàn tất báo cáo."},
        {"metric":"qa_run_id","value":QA_RUN_ID,"definition":"Run QA độc lập."},
        {"metric":"batch_id","value":QA_BATCH_ID,"definition":"Inventory positions 11–20."},
        {"metric":"products_checked","value":10,"definition":"Đúng 10 product key."},
        {"metric":"images_checked","value":64,"definition":"64/64 ảnh mở trực tiếp ở ảnh gốc."},
        {"metric":"batch_final_score","value":batch_avg,"definition":"Trung bình đều 10 final_score; không bù lỗi chặn."},
        {"metric":"batch_result","value":"NOT_PASSED","definition":"Lô chỉ đạt khi mọi sản phẩm QA_PASS."},
        {"metric":"status_counts","value":dict(statuses),"definition":"Số sản phẩm theo kết luận."},
        {"metric":"issue_counts","value":dict(sev_counts),"definition":"Số phát hiện theo severity."},
        {"metric":"admin_limitation","value":"No Shopify admin export","definition":"Không suy ra admin SEO fields/alt từ storefront."},
        {"metric":"xlsx_status","value":"COMPLETE","definition":"Workbook QA được tạo bằng openpyxl theo yêu cầu người dùng, mở lại kiểm tra cấu trúc/công thức và render kiểm tra."},
    ]
    serp_rows = []
    for p in products:
        pos = pk_to_pos[p["product_key"]]
        for n,(query,urls) in enumerate(SERP[pos],1):
            serp_rows.append({"serp_id":f"serp_qa_{pos:03d}_{n}","product_key":p["product_key"],"query":query,
                              "market":"United States","language":"English","locale_limit":"US intent; search service locale could not be hard-pinned",
                              "checked_at":checked_at,"result_urls_read":urls,"intent":"Commercial/product",
                              "note":"Evidence supports intent only; no paid search-volume claim."})
    change_rows = []
    for i,entry in enumerate(live):
        old,current = entry["research_snapshot"]["html"],entry["live"]
        stable = {"title":old.get("rendered_title_current")==current.get("title_element"),"h1":old.get("h1_current")==current.get("h1"),
                  "meta_description":old.get("meta_description_current")==current.get("meta_description"),"canonical":old.get("canonical_url")==current.get("canonical")}
        change_rows.append({"product_key":products[i]["product_key"],"snapshot_checked_at":entry["research_snapshot"].get("reviewed_at"),
                            "live_checked_at":current.get("checked_at"),"snapshot_html_sha256":old.get("html_sha256"),
                            "live_html_sha256":current.get("html_sha256"),"html_hash_changed":old.get("html_sha256")!=current.get("html_sha256"),
                            "material_fields_equal":stable,"source_revision_status":"LIVE_EQUIVALENT_DYNAMIC_HTML_CHANGED" if all(stable.values()) else "SOURCE_CHANGED"})
    tests = {"product_weight_total":sum(PRODUCT_WEIGHTS.values()),"image_weight_total":sum(IMAGE_WEIGHTS.values()),
             "products":len(qa_products),"images":len(qa_images),"unique_product_keys":len({p['product_key'] for p in qa_products}),
             "unique_qa_image_keys":len({i['qa_image_key'] for i in qa_images}),"all_image_coverage_100":all(p["image_coverage"]==1 for p in qa_products),
             "logic_100_with_critical":"QA_FAIL","logic_90_full_no_blocker":"QA_PASS","logic_72_on_80":{"range":"72-92","status":"QA_INCOMPLETE"},
             "cross_links_valid":all(all(ref in {i['issue_id'] for i in issues} for ref in p['issue_refs']) for p in qa_products)}
    assert tests["product_weight_total"] == tests["image_weight_total"] == 100
    assert tests["products"] == tests["unique_product_keys"] == 10
    assert tests["images"] == tests["unique_qa_image_keys"] == 64
    assert tests["all_image_coverage_100"] and tests["cross_links_valid"]
    dataset = {"QA_Summary":summary,"QA_Products":qa_products,"QA_Criteria":criteria,"QA_Images":qa_images,"QA_Issues":issues,
               "SERP_Evidence":serp_rows,"validation_tests":tests}
    common.save_json(QA_DIR/"qa_dataset.json",dataset)
    common.save_json(QA_DIR/"qa_workbook_payload.json",{k:dataset[k] for k in ("QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues")})
    common.save_json(QA_DIR/"serp_evidence.json",serp_rows)
    common.save_json(QA_DIR/"source_change_audit.json",change_rows)
    common.save_json(QA_DIR/"validation_results.json",tests)

    inventory = json.loads((ROOT/"seo_runs"/"jeminise.com"/RUN_ID/"inventory.json").read_text(encoding="utf-8-sig"))
    next_batch = [{k:row.get(k) for k in ("inventory_position","product_key","product_id","Handle","title_current","product_url","image_count")}
                  for row in inventory if 21 <= int(row["inventory_position"]) <= 30]
    assert len(next_batch) == 10
    preview = QA_DIR/"qa_batch_003_preview.json"
    common.save_json(preview,{"batch_id":"qa_batch_003","status":"PREPARED_NOT_STARTED","products":next_batch})

    lines = ["# SEO QA — qa_batch_002","","## Kết luận","",
             "- Phạm vi: **10 sản phẩm, 64/64 ảnh (100%)**; chỉ inventory position 11–20.",
             f"- Điểm lô: **{batch_avg:.1f}/100**; kết luận lô: **NOT_PASSED**.",
             f"- Trạng thái: {statuses.get('QA_FAIL',0)} QA_FAIL, {statuses.get('QA_REVISE',0)} QA_REVISE, {statuses.get('QA_PASS',0)} QA_PASS.",
             f"- Phát hiện: {sev_counts.get('CRITICAL',0)} CRITICAL, {sev_counts.get('MAJOR',0)} MAJOR, {sev_counts.get('MINOR',0)} MINOR, {sev_counts.get('LIMITATION',0)} LIMITATION.",
             f"- Workbook nguồn: `{SOURCE}`",f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`","","## Điểm theo sản phẩm","",
             "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |","|---:|---|---:|---|---:|"]
    for q,p in zip(qa_products,products):
        lines.append(f"| {q['inventory_position']} | {p['title_current'].replace('|','/')} | {q['final_score']:.1f} | {q['qa_status']} | {q['critical_count']}/{q['major_count']}/{q['minor_count']}/{q['limitation_count']} |")
    lines += ["","## Lỗi ưu tiên","",
              "1. **CRITICAL — products 11–16:** draft dùng claim `Personalized`, nhưng purchase options live không có input tên/verse/birth flower; phải chứng minh control hoạt động hoặc bỏ claim.",
              "2. **CRITICAL — product 19:** description ghi `Customization: 1 text input`, nhưng trang live chỉ có size và pillowcase quantity.",
              "3. **MAJOR — cả 10 sản phẩm:** description HTML chứa câu nội bộ yêu cầu review, chưa publish-ready.",
              "4. **MAJOR — ảnh:** alt/observation theo mẫu vị trí sai hàng loạt; nhóm Christian sai nhiều ảnh 3–8, nhóm Christmas thường đảo ảnh cận cảnh với ảnh sham.",
              "5. **MAJOR — products 11–16:** SERP evidence đã nộp chủ yếu là blanket trong khi keyword nhắm comforter/bedding, đồng thời sáu trang cạnh tranh intent rất gần nhau.",
              "6. **MAJOR — product 13:** draft chép `Style: Christian Knight Templar`, mâu thuẫn với thiết kế God Is Within Her dành cho nữ đang hiển thị.",
              "","## SERP và keyword","",
              "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume.",
              "","## Giới hạn và trạng thái bàn giao","",
              "- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.",
              "- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.",
              "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.",
              "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã được mở lại, kiểm tra cấu trúc/công thức và render kiểm tra.",
              "- Excel Desktop không thể dùng để recalculate do Office báo `Product Activation Failed`; công thức được kiểm tra cấu trúc, không có `#REF!/#NAME?`, và cả 5 sheet được render độc lập bằng openpyxl/Pillow.",
              "- Chưa QA products 21–30. `awaiting_confirmation=true`.","","## Tệp chi tiết","",
              f"- QA data: `{QA_DIR/'qa_dataset.json'}`",f"- SERP evidence: `{QA_DIR/'serp_evidence.json'}`",
              f"- Validation: `{QA_DIR/'validation_results.json'}`",f"- Manifest/checkpoint: `{QA_DIR}`",""]
    OUT_DIR.mkdir(parents=True,exist_ok=True)
    report = OUT_DIR/f"SEO_QA_{QA_BATCH_ID}.md"
    report.write_text("\n".join(lines),encoding="utf-8")
    xlsx = OUT_DIR/f"SEO_QA_{QA_BATCH_ID}.xlsx"
    manifest = json.loads((QA_DIR/"qa_manifest.json").read_text(encoding="utf-8"))
    manifest.update({"status":"COMPLETE","completed_at":common.now(),"source_sha256_at_handoff":source_hash,
                     "counts":{"products":10,"images":64},"output_markdown":str(report),"output_xlsx":str(xlsx),"xlsx_blocker":None,
                     "qa_dataset":str(QA_DIR/"qa_dataset.json"),"qa_workbook_payload":str(QA_DIR/"qa_workbook_payload.json"),
                     "source_snapshot_sha256":common.sha256(QA_DIR/"source_snapshot"/SOURCE.name),
                     "prompt_files_sha256":{"prompt_qa.md":common.sha256(ROOT/"seo-prompt/jeminise/prompt_qa.md"),"prompt.md":common.sha256(ROOT/"seo-prompt/jeminise/prompt.md")},
                     "source_revision_status":"LIVE_EQUIVALENT_DYNAMIC_HTML_CHANGED"})
    common.save_json(QA_DIR/"qa_manifest.json",manifest)
    progress = json.loads((QA_DIR/"qa_progress.json").read_text(encoding="utf-8"))
    progress.update({"current_product_key":products[-1]["product_key"],"current_stage":"BATCH_COMPLETE",
                     "completed_image_keys":[x["qa_image_key"] for x in qa_images],"last_saved_at":common.now(),
                     "artifact_paths":{"markdown":str(report),"xlsx":str(xlsx),"dataset":str(QA_DIR/"qa_dataset.json"),
                                       "workbook_payload":str(QA_DIR/"qa_workbook_payload.json"),"source_change_audit":str(QA_DIR/"source_change_audit.json"),
                                       "spreadsheet_validation":str(QA_DIR/"spreadsheet_validation.json")},
                     "awaiting_confirmation":True,"confirmation_ref":None,
                     "next_batch_preview":{"batch_id":"qa_batch_003","inventory_positions":"21-30","status":"PREPARED_NOT_STARTED",
                                           "product_keys":[x["product_key"] for x in next_batch],"path":str(preview)}})
    common.save_json(QA_DIR/"qa_progress.json",progress)
    print(json.dumps({"report":str(report),"dataset":str(QA_DIR/"qa_dataset.json"),"batch_score":batch_avg,
                      "statuses":dict(statuses),"issues":dict(sev_counts),"source_hash":source_hash},ensure_ascii=False,indent=2))


if __name__ == "__main__":
    main()
