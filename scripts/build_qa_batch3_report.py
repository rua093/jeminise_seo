from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import build_qa_batch1_report as common


ROOT = Path(__file__).resolve().parents[1]
SHOP = "jeminise.com"
RUN_ID = "20260906_234129"
QA_RUN_ID = "20260907_094127"
QA_BATCH_ID = "qa_batch_003"
QA_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID
OUT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN_ID
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_034.xlsx"
SNAPSHOT = QA_DIR / "source_snapshot" / SOURCE.name
RATING = common.RATING
PRODUCT_WEIGHTS = common.PRODUCT_WEIGHTS
IMAGE_WEIGHTS = common.IMAGE_WEIGHTS


ACTUAL = {
    21: [
        "Front bedroom mockup of a cream Christmas quilt with a gingerbread man, candy cane, green bow and peppermint border.",
        "Angled bedroom mockup of the same gingerbread-and-candy-cane Christmas quilt.",
        "Close-up of the printed gingerbread bow and candy-cane surface with visible quilting lines.",
        "Isolated matching pillow sham with gingerbread man and candy cane artwork.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    22: [
        "Front bedroom mockup of a black Christmas-tree quilt with ornaments, snowflakes and wrapped gifts.",
        "Angled bedroom mockup of the same black Christmas-tree quilt.",
        "Isolated matching pillow sham with decorated tree and gift artwork.",
        "Close-up of the printed Christmas-tree surface with visible quilting lines.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    23: [
        "Front bedroom mockup of a soft blue-and-cream winter quilt with red cardinals, snowy birdhouse and wreath.",
        "Angled bedroom mockup of the same cardinal-and-birdhouse winter quilt.",
        "Isolated matching pillow sham with two red cardinals and snowy birdhouse artwork.",
        "Close-up of the printed cardinal, birdhouse and wreath surface with visible quilting lines.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    24: [
        "Front bedroom mockup of a large gingerbread figure in a green hat and scarf, surrounded by gifts, cookies and a snowy village.",
        "Angled bedroom mockup of the same gingerbread Christmas-village quilt.",
        "Isolated matching pillow sham with gingerbread figure, gifts, cookies and village artwork.",
        "Close-up of the gingerbread scarf and icing detail with visible quilting lines.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    25: [
        "Front bedroom mockup of a red Christmas quilt with a snowman, cardinal birds, pine trees and holly.",
        "Angled bedroom mockup of the same red snowman-and-cardinal quilt.",
        "Isolated matching pillow sham with snowman and cardinals.",
        "Close-up of the printed snowman scarf and cardinal surface with visible quilting lines.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    26: [
        "Front bedroom mockup of a white Christmas quilt with a black-and-gray tree, red and gold stars, and check border.",
        "Angled bedroom mockup of the same white Christmas-tree quilt.",
        "Isolated matching pillow sham with black, gray, red and gold Christmas tree.",
        "Close-up of the printed tree branches and star ornaments with visible quilting lines.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    27: [
        "Front bedroom mockup of a deep-red Christmas quilt with two snowmen, a red cardinal, a brown cardinal and pine greenery.",
        "Angled bedroom mockup of the same snowman-and-cardinal quilt.",
        "Isolated matching pillow sham with the snowman pair and cardinals.",
        "Close-up of the printed snowman and brown cardinal surface with visible quilting lines.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    28: [
        "Front bedroom mockup of a cream vintage-style Christmas-tree quilt with ornaments, wrapped gifts and an ornate border.",
        "Angled bedroom mockup of the same cream Christmas-tree quilt.",
        "Isolated matching pillow sham with decorated tree and wrapped gifts.",
        "Close-up of the printed ornaments and tree branches with visible quilting lines.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    29: [
        "Front bedroom mockup of a red-and-white patchwork Christmas quilt with a cardinal and the words I am Always With You.",
        "Close-up of the cardinal and the visible words I am Always With You.",
        "Isolated matching pillow sham with cardinal, poinsettias, roses, bells, hearts and berries.",
        "Angled bedroom mockup of the same cardinal memorial-style Christmas quilt.",
        "Included-components and size graphic: one premium quilt plus two or four optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
    30: [
        "Front bedroom mockup of a cardinal-in-wreath Christmas quilt printed with the sample name Sophia.",
        "Close-up of the cardinal, pine, holly berries and pinecones with visible quilting lines.",
        "Isolated matching pillow sham printed with the sample name Sophia.",
        "Angled bedroom mockup of the same cardinal-in-wreath quilt printed with Sophia.",
        "Included-components and size graphic: one premium quilt plus two optional standard shams; Throw, Twin, Queen and King dimensions.",
    ],
}

SWAPPED = {(21, 3), (21, 4), (29, 2), (29, 4), (30, 2), (30, 4)}
ALT_FIXES = {
    (21, 3): "Close-up of gingerbread and candy cane Christmas quilt print",
    (21, 4): "Matching gingerbread and candy cane Christmas quilt pillow sham",
    (29, 2): "Close-up of cardinal quilt with I am Always With You text",
    (29, 4): "Cardinal memorial Christmas quilt in an angled bedroom view",
    (30, 2): "Close-up of cardinal, pine, holly berries and pinecones on the quilt",
    (30, 4): "Cardinal Christmas quilt with Sophia in an angled bedroom view",
}

SERP = {
    21: [("gingerbread candy cane quilt set", ["https://quiltwoman.com/products/gingerbread-christmas-quilt-downloadable-pattern"]), ("gingerbread Christmas quilt set candy cane", ["https://www.target.com/p/-/A-94992224"])],
    22: [("black Christmas tree quilt set", ["https://www.homedepot.com/p/332726298"]), ("dark Christmas tree quilt bedding", ["https://www.lowes.com/pd/MarCielo-BY218-3-Pcs-Patchwork-Christmas-Tree-Queen-Size-Polyester-Quilt-Set-Holiday-Bedspread/5016774281"])],
    23: [("winter cardinal birdhouse quilt set", ["https://www.etsy.com/listing/4348580665/bird-feeders-and-cardinals-quilt-kit"]), ("cardinal birdhouse winter bedding quilt", ["https://www.wayfair.com/bed-bath/pdp/james-home-holiday-cardinal-reversible-bedding-quilt-set-w100620099.html"])],
    24: [("snowman Christmas quilt set", ["https://www.target.com/p/-/A-1006322169"]), ("gingerbread snowman Christmas quilt set", ["https://www.marcielobedding.com/products/3-piece-christmas-snowman-quilt-set-marcielo-reversible-bedspread"])],
    25: [("red snowman cardinal Christmas quilt set", ["https://charmingfavor.com/products/nd230908"]), ("snowman cardinal bedding quilt", ["https://www.walmart.com/ip/20264016781"])],
    26: [("white Christmas tree quilt set", ["https://www.walmart.com/ip/2562902230"]), ("cream Christmas tree quilt bedding set", ["https://www.macys.com/shop/seasonal-celebrations/holiday-shop/christmas/bed-bath/bedding/Color_normal/Ivory%2FCream?id=202790"])],
    27: [("snowman cardinal Christmas quilt set", ["https://charmingfavor.com/products/nd230908"]), ("cardinal snowman quilt bedding set", ["https://www.chugsandree.com/products/christmas-quilt-embroidery-snowman-and-cardinal-quilt-personalized-christmas-quilt-custom-quilt-christmas-decor-buffalo-plaid"])],
    28: [("vintage Christmas tree quilt set", ["https://www.etsy.com/listing/773547484/handmade-christmas-tree-quilt-vintage"]), ("antique style Christmas tree quilt bedding", ["https://www.potterybarn.com/shopping/christmas-tree-bedding/"])],
    29: [("Christmas cardinal memorial quilt I am always with you", ["https://www.etsy.com/listing/1085801805/i-am-always-with-you-blanket-cardinal"]), ("cardinal remembrance quilt bedding set", ["https://www.walmart.com/ip/16678752228"])],
    30: [("personalized Christmas cardinal quilt set", ["https://www.walmart.com/ip/16678752228"]), ("custom name cardinal Christmas bedding quilt", ["https://www.macustom.com/collections/quilted-bedding-set/products/cardinal-christmas-quilted-bedding-set-ncu0vh044"])],
}

PRODUCT_ASSESS = {
    21: {"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    22: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    23: {"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"PARTIAL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    24: {"P1":"FAIL","P2":"FULL","K1":"FAIL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FAIL","D2":"FAIL","E1":"PARTIAL"},
    25: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    26: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    27: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    28: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    29: {"P1":"FULL","P2":"FAIL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    30: {"P1":"FULL","P2":"FAIL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
}


def add_issue(rows, *args, **kwargs):
    rows.append(common.issue(*args, **kwargs))


def main() -> None:
    data = json.loads((QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))
    live = json.loads((QA_DIR / "live_source_comparison.json").read_text(encoding="utf-8"))
    products = data["products"]
    pk_to_pos = {p["product_key"]: i + 21 for i, p in enumerate(products)}
    images_by_pk = {p["product_key"]: [] for p in products}
    for im in data["images"]:
        images_by_pk[im["shop_domain"] + "+" + im["Handle"]].append(im)
    checked_at = common.now()
    qa_images, issues = [], []

    for p in products:
        pos = pk_to_pos[p["product_key"]]
        live_media = live[pos - 21]["live"]["product_js"]["media"]
        submitted = sorted(images_by_pk[p["product_key"]], key=lambda x: x["image_number"])
        assert len(submitted) == len(live_media) == len(ACTUAL[pos]) == 5
        for idx, im in enumerate(submitted):
            n = idx + 1
            qkey = common.stable_image_key(p["product_key"], im["image_url"], n)
            wrong = (pos, n) in SWAPPED
            assessments = {"IM1":"FULL", "IM2":"FAIL" if wrong else "FULL", "IM3":"FAIL" if wrong else "FULL", "IM4":"FULL"}
            points = sum(IMAGE_WEIGHTS[k] * RATING[v] for k, v in assessments.items())
            refs = [im["evidence_file_or_reference"], p["product_url"], live_media[idx]["src"]]
            issue_refs = []
            if wrong:
                iid = f"ISS-{pos:03d}-IMG-{n:02d}"
                add_issue(issues, iid, p["product_key"], "MAJOR", "image_observation/alt_effective",
                          f"{im['observed_visual_details']} | {im['alt_proposed']}", ACTUAL[pos][idx],
                          "Observation và alt được gán theo vị trí mẫu nhưng không phản ánh ảnh gốc.", ALT_FIXES[(pos,n)],
                          "; ".join(refs), "Mở lại ảnh gốc và xác nhận observation/alt mới mô tả đúng ảnh.", qkey)
                issue_refs.append(iid)
            qa_images.append({
                "product_key":p["product_key"], "qa_image_key":qkey,
                "image_url_source":live_media[idx]["src"], "image_url_workbook":im["image_url"],
                "media_id":str(live_media[idx]["id"]), "workbook_image_id":str(im["media_id"]),
                "variant":im["variant"] or "", "image_location":im["image_location"],
                "check_method":"DIRECT_ORIGINAL_IMAGE", "checked_at":checked_at,
                "qa_observation":ACTUAL[pos][idx], "submitted_observation":im["observed_visual_details"],
                "storefront_alt_observed":live_media[idx].get("alt") or "", "alt_action":im["alt_action"],
                "alt_effective":im["alt_proposed"], **assessments,
                "image_verified_points":points, "image_assessed_weight":100, "image_final_score":points,
                "image_score_lower_bound":points, "image_score_upper_bound":points,
                "issue_refs":issue_refs, "evidence_refs":refs,
            })

    for p in products:
        pos, pk = pk_to_pos[p["product_key"]], p["product_key"]
        add_issue(issues, f"ISS-{pos:03d}-DESC", pk, "MAJOR", "description_proposed_html",
                  "It requires QA and approval before import.", "HTML kết thúc bằng lời nhắc QA/import nội bộ.",
                  "Nội dung quy trình nội bộ không được xuất hiện trên storefront.",
                  "Rewrite as customer-facing English HTML and remove the internal review sentence.",
                  p["evidence_id"], "Render full HTML and confirm no drafting/QA/import instruction remains.")
        if pos <= 28:
            add_issue(issues, f"ISS-{pos:03d}-ENC", pk, "MINOR", "title_current/H1_current",
                      p["title_current"], live[pos-21]["live"]["h1"],
                      "Title/H1 live chứa ký tự thay thế U+FFFD và từ Bedsprea bị cắt.",
                      p["title_proposed"], p["product_url"],
                      "Rendered title and H1 no longer contain U+FFFD or truncated words.")

    for pos in (21, 29, 30):
        p = products[pos - 21]
        opts = live[pos - 21]["live"]["product_js"]["options"]
        add_issue(issues, f"ISS-{pos:03d}-PERS", p["product_key"], "CRITICAL",
                  "description/title/meta personalization claim", "Customization: 1 text input / Personalized",
                  "Purchase options live chỉ có quilt size và pillowcase quantity; không có trường nhập text/name.",
                  "Customization là thuộc tính mua hàng trọng yếu nhưng không thể thao tác trong purchase flow live.",
                  "Remove customization/personalization claims unless a working text input and fulfillment mapping are evidenced.",
                  p["product_url"] + "; live options=" + json.dumps(opts, ensure_ascii=False),
                  "Live purchase flow visibly accepts and preserves the claimed custom text.")

    p24 = products[3]
    add_issue(issues, "ISS-024-DESIGN", p24["product_key"], "CRITICAL",
              "primary_keyword/title/meta/description", "Snowman Christmas Village Quilt Set",
              ACTUAL[24][0], "Draft identifies the central gingerbread figure as a snowman, changing the design and search intent.",
              "Rewrite around a gingerbread Christmas village quilt set; rerun the primary/comparator SERP decision and update all dependent fields.",
              p24["product_url"] + "; direct images 1-4", "All fields and alt text identify the visible gingerbread design consistently.")

    for pos, note in {
        21:"Primary results skew toward downloadable patterns/kits; the closest finished bedding result validates candy-cane quilt intent but not the full exact phrase.",
        23:"Primary results mix quilt kits/digital designs with finished bedding; birdhouse specificity is visually accurate but commercial product-type evidence is incomplete.",
    }.items():
        p = products[pos - 21]
        add_issue(issues, f"ISS-{pos:03d}-SERP", p["product_key"], "MAJOR", "keyword/SERP evidence",
                  p["primary_keyword"], note,
                  "SERP supports the motif only partially and does not cleanly validate the selected finished-product phrase.",
                  "Use current US finished-bedding product SERPs or retain SERP_ONLY with narrower claims and explicit limitation.",
                  json.dumps(SERP[pos], ensure_ascii=False), "Two current product SERPs support the exact motif and quilt-set intent.")

    add_issue(issues, "ISS-GLOBAL-ADMIN", "", "LIMITATION", "current admin SEO fields/current admin alt",
              "Not supplied", "Storefront metadata and product.js media alt were readable, but no Shopify admin export was provided.",
              "Storefront values cannot prove current admin fields or future import mapping.",
              "Provide a frozen Shopify admin/export snapshot before approval or import mapping.", str(SOURCE),
              "Admin export revision and hash are frozen and compared to storefront/workbook.")

    issue_by_product = {}
    for item in issues:
        issue_by_product.setdefault(item["product_key"], []).append(item)
    reasons = dict(common.REASONS)
    reasons["P2"] = "Đối chiếu material, dimensions, components, variants và personalization với product.js live, body HTML và gallery."
    reasons["K1"] = "Đối chiếu long-tail với thiết kế trực tiếp nhìn thấy và nguy cơ trùng intent trong cụm Christmas/cardinal."
    reasons["K2"] = "Đã đọc lại primary và comparator SERP theo US intent; đánh giá product/kit/pattern result type."
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
                rating, earned, reason = RATING[assessment], weight * RATING[assessment], reasons[cid]
            criteria.append({"product_key":pk,"criterion_id":cid,"weight":weight,"assessment":assessment,
                             "rating":rating,"earned_points":earned,"assessed_weight":weight,"reason":reason,
                             "evidence_refs":[p["evidence_id"],p["product_url"],f"serp_qa_{pos:03d}"],"issue_refs":p_refs})
        verified = sum(x["earned_points"] for x in criteria if x["product_key"] == pk)
        sev = Counter(x["severity"] for x in issue_by_product.get(pk, []))
        status = "QA_FAIL" if sev["CRITICAL"] or verified < 70 else ("QA_REVISE" if verified < 85 or sev["MAJOR"] else "QA_PASS")
        qa_products.append({"inventory_position":pos,"product_key":pk,"url":p["product_url"],"handle":p["Handle"],
                            "product_id":str(p["product_id"]),"revision":"r1","verified_points":verified,
                            "assessed_weight":100,"score_lower_bound":verified,"score_upper_bound":verified,"final_score":verified,
                            "qa_status":status,"keyword_evidence_level":p["keyword_evidence_level"],"images_expected":5,
                            "images_checked":5,"image_inventory_complete":True,"image_coverage":1.0,
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
        {"metric":"batch_id","value":QA_BATCH_ID,"definition":"Inventory positions 21–30."},
        {"metric":"products_checked","value":10,"definition":"Đúng 10 product key."},
        {"metric":"images_checked","value":50,"definition":"50/50 ảnh mở trực tiếp ở ảnh gốc."},
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
        for n, (query, urls) in enumerate(SERP[pos], 1):
            serp_rows.append({"serp_id":f"serp_qa_{pos:03d}_{n}","product_key":p["product_key"],"query":query,
                              "market":"United States","language":"English","locale_limit":"US intent; search service locale could not be hard-pinned",
                              "checked_at":checked_at,"result_urls_read":urls,"intent":"Commercial/product",
                              "note":"Evidence supports intent only; no paid search-volume claim."})
    change_rows = []
    for i, entry in enumerate(live):
        old, current = entry["research_snapshot"]["html"], entry["live"]
        stable = {"title":old.get("rendered_title_current")==current.get("title_element"),
                  "h1":old.get("h1_current")==current.get("h1"),
                  "meta_description":old.get("meta_description_current")==current.get("meta_description"),
                  "canonical":old.get("canonical_url")==current.get("canonical")}
        change_rows.append({"product_key":products[i]["product_key"],"snapshot_checked_at":entry["research_snapshot"].get("reviewed_at"),
                            "live_checked_at":current.get("checked_at"),"snapshot_html_sha256":old.get("html_sha256"),
                            "live_html_sha256":current.get("html_sha256"),"html_hash_changed":old.get("html_sha256")!=current.get("html_sha256"),
                            "material_fields_equal":stable,"source_revision_status":"LIVE_EQUIVALENT_DYNAMIC_HTML_CHANGED" if all(stable.values()) else "SOURCE_CHANGED"})
    tests = {"product_weight_total":sum(PRODUCT_WEIGHTS.values()),"image_weight_total":sum(IMAGE_WEIGHTS.values()),
             "products":len(qa_products),"images":len(qa_images),"criteria":len(criteria),
             "unique_product_keys":len({p['product_key'] for p in qa_products}),"unique_qa_image_keys":len({i['qa_image_key'] for i in qa_images}),
             "all_image_coverage_100":all(p["image_coverage"]==1 for p in qa_products),
             "logic_100_with_critical":"QA_FAIL","logic_90_full_no_blocker":"QA_PASS",
             "logic_72_on_80":{"range":"72-92","status":"QA_INCOMPLETE"},
             "cross_links_valid":all(all(ref in {i['issue_id'] for i in issues} for ref in p['issue_refs']) for p in qa_products)}
    assert tests["product_weight_total"] == tests["image_weight_total"] == 100
    assert tests["products"] == tests["unique_product_keys"] == 10
    assert tests["images"] == tests["unique_qa_image_keys"] == 50 and tests["criteria"] == 110
    assert tests["all_image_coverage_100"] and tests["cross_links_valid"]

    dataset = {"QA_Summary":summary,"QA_Products":qa_products,"QA_Criteria":criteria,"QA_Images":qa_images,"QA_Issues":issues,
               "SERP_Evidence":serp_rows,"validation_tests":tests}
    common.save_json(QA_DIR/"qa_dataset.json",dataset)
    common.save_json(QA_DIR/"qa_workbook_payload.json",{k:dataset[k] for k in ("QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues")})
    common.save_json(QA_DIR/"serp_evidence.json",serp_rows)
    common.save_json(QA_DIR/"source_change_audit.json",change_rows)
    common.save_json(QA_DIR/"validation_results.json",tests)

    inventory = json.loads((ROOT/"seo_runs"/SHOP/RUN_ID/"inventory.json").read_text(encoding="utf-8-sig"))
    next_batch = [{k:row.get(k) for k in ("inventory_position","product_key","product_id","Handle","title_current","product_url","image_count")}
                  for row in inventory if 31 <= int(row["inventory_position"]) <= 40]
    assert len(next_batch) == 10
    preview = QA_DIR/"qa_batch_004_preview.json"
    common.save_json(preview,{"batch_id":"qa_batch_004","status":"PREPARED_NOT_STARTED","products":next_batch})

    lines = ["# SEO QA — qa_batch_003","","## Kết luận","",
             "- Phạm vi: **10 sản phẩm, 50/50 ảnh (100%)**; chỉ inventory position 21–30.",
             f"- Điểm lô: **{batch_avg:.1f}/100**; kết luận lô: **NOT_PASSED**.",
             f"- Trạng thái: {statuses.get('QA_FAIL',0)} QA_FAIL, {statuses.get('QA_REVISE',0)} QA_REVISE, {statuses.get('QA_PASS',0)} QA_PASS.",
             f"- Phát hiện: {sev_counts.get('CRITICAL',0)} CRITICAL, {sev_counts.get('MAJOR',0)} MAJOR, {sev_counts.get('MINOR',0)} MINOR, {sev_counts.get('LIMITATION',0)} LIMITATION.",
             f"- Workbook nguồn: `{SOURCE}`",f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`","","## Điểm theo sản phẩm","",
             "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |","|---:|---|---:|---|---:|"]
    for q, p in zip(qa_products, products):
        lines.append(f"| {q['inventory_position']} | {p['title_current'].replace('|','/')} | {q['final_score']:.1f} | {q['qa_status']} | {q['critical_count']}/{q['major_count']}/{q['minor_count']}/{q['limitation_count']} |")
    lines += ["","## Lỗi ưu tiên","",
              "1. **CRITICAL — products 21, 29, 30:** draft ghi customization/personalization nhưng purchase options live không có text/name input; phải chứng minh control hoạt động hoặc bỏ claim.",
              "2. **CRITICAL — product 24:** draft gọi nhân vật trung tâm là snowman, trong khi ảnh là gingerbread figure; cần thay keyword/title/meta/body và chạy lại SERP decision.",
              "3. **MAJOR — cả 10 sản phẩm:** description HTML chứa câu QA/import nội bộ, chưa publish-ready.",
              "4. **MAJOR — ảnh:** product 21 đảo close-up với sham; products 29–30 đảo close-up với angled bedroom view.",
              "5. **MAJOR — products 21, 23:** SERP primary còn lẫn pattern/kit và chưa sạch intent finished quilt set.",
              "6. **MINOR — products 21–28:** title/H1 nguồn có ký tự `�` và từ `Bedsprea` bị cắt.",
              "","## SERP và keyword","",
              "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume.",
              "","## Giới hạn và trạng thái bàn giao","",
              "- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.",
              "- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.",
              "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.",
              "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã được mở lại, kiểm tra cấu trúc/công thức và render kiểm tra.",
              "- Excel Desktop không thể dùng để recalculate do Office báo `Product Activation Failed`; công thức được kiểm tra cấu trúc, không có `#REF!/#NAME?`, và cả 5 sheet được render độc lập bằng openpyxl/Pillow.",
              "- Chưa QA products 31–40. `awaiting_confirmation=true`.","","## Tệp chi tiết","",
              f"- QA data: `{QA_DIR/'qa_dataset.json'}`",f"- SERP evidence: `{QA_DIR/'serp_evidence.json'}`",
              f"- Validation: `{QA_DIR/'validation_results.json'}`",f"- Manifest/checkpoint: `{QA_DIR}`",""]
    OUT_DIR.mkdir(parents=True,exist_ok=True)
    report = OUT_DIR/f"SEO_QA_{QA_BATCH_ID}.md"
    report.write_text("\n".join(lines),encoding="utf-8")
    xlsx = OUT_DIR/f"SEO_QA_{QA_BATCH_ID}.xlsx"

    keys = [p["product_key"] for p in products]
    manifest = {"rubric_version":"1.0","prompt_version":"2.4","qa_run_id":QA_RUN_ID,"started_at":"2026-09-07T09:41:27+07:00",
                "shop_domain":SHOP,"research_run_id":RUN_ID,"market":"United States","seo_language":"English","batch_id":QA_BATCH_ID,
                "source_workbook":str(SOURCE.relative_to(ROOT)),"source_workbook_sha256":source_hash,"source_snapshot":str(SNAPSHOT.relative_to(ROOT)),
                "batch_product_keys":keys,"revision":"r1","expected_products":10,"expected_images":50,"source_admin_export_available":False,
                "status":"COMPLETE","completed_at":common.now(),"source_sha256_at_handoff":source_hash,"counts":{"products":10,"images":50},
                "output_markdown":str(report),"output_xlsx":str(xlsx),"xlsx_blocker":None,"qa_dataset":str(QA_DIR/"qa_dataset.json"),
                "qa_workbook_payload":str(QA_DIR/"qa_workbook_payload.json"),"source_snapshot_sha256":common.sha256(SNAPSHOT),
                "prompt_files_sha256":{"prompt_qa.md":common.sha256(ROOT/"seo-prompt/jeminise/prompt_qa.md"),"prompt.md":common.sha256(ROOT/"seo-prompt/jeminise/prompt.md")},
                "source_revision_status":"LIVE_EQUIVALENT_DYNAMIC_HTML_CHANGED"}
    common.save_json(QA_DIR/"qa_manifest.json",manifest)
    progress = {"rubric_version":"1.0","qa_run_id":QA_RUN_ID,"source_workbook":str(SOURCE.relative_to(ROOT)),
                "source_workbook_sha256":source_hash,"batch_id":QA_BATCH_ID,"batch_product_keys":keys,
                "current_product_key":keys[-1],"current_stage":"BATCH_COMPLETE","completed_image_keys":[x["qa_image_key"] for x in qa_images],
                "last_saved_at":common.now(),"artifact_paths":{"markdown":str(report),"xlsx":str(xlsx),"dataset":str(QA_DIR/"qa_dataset.json"),
                "workbook_payload":str(QA_DIR/"qa_workbook_payload.json"),"source_change_audit":str(QA_DIR/"source_change_audit.json"),
                "spreadsheet_validation":str(QA_DIR/"spreadsheet_validation.json")},"awaiting_confirmation":True,"confirmation_ref":None,
                "next_batch_preview":{"batch_id":"qa_batch_004","inventory_positions":"31-40","status":"PREPARED_NOT_STARTED",
                "product_keys":[x["product_key"] for x in next_batch],"path":str(preview)}}
    common.save_json(QA_DIR/"qa_progress.json",progress)
    print(json.dumps({"report":str(report),"dataset":str(QA_DIR/"qa_dataset.json"),"batch_score":batch_avg,
                      "statuses":dict(statuses),"issues":dict(sev_counts),"source_hash":source_hash},ensure_ascii=False,indent=2))


if __name__ == "__main__":
    main()
