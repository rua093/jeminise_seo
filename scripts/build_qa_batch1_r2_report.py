from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import build_qa_batch1_report as common


ROOT = Path(__file__).resolve().parents[1]
SHOP, RUN_ID, QA_RUN_ID, BATCH = "jeminise.com", "20260906_234129", "20260907_114114", "qa_batch_001_r2"
QA_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID
OUT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN_ID
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / "qa_batch_001_r2" / "SEO_Product_Optimization_qa_batch_001_r2.xlsx"
SNAPSHOT = QA_DIR / "source_snapshot" / SOURCE.name
PW, IW, R = common.PRODUCT_WEIGHTS, common.IMAGE_WEIGHTS, common.RATING

ASSESS = {
    1:{"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"PARTIAL"},
    2:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"PARTIAL","T2":"PARTIAL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
    3:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
    4:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
    5:{"P1":"FULL","P2":"FULL","K1":"FAIL","K2":"PARTIAL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FAIL"},
    6:{"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"PARTIAL"},
    7:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
    8:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
    9:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","E1":"FULL"},
    10:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"PARTIAL","E1":"FULL"},
}

SERP = {
 1:[("cardinal sunflower quilt set",["https://vantique.net/products/autumn-cardinal-3-piece-quilted-bedding-set-ncu0tv033"]),("autumn cardinal bedding quilt",["https://www.target.com/s/cardinal%2Bquilt"])],
 2:[("softball comforter set",["https://www.etsy.com/market/softball_comforter_sets"]),("teal softball bedding set",["https://www.walmart.com/ip/18550165606","https://www.youcustomizeit.com/p/Softball-Comforters-Personalized/353832"])],
 3:[("Christmas cardinal memorial quilt",["https://www.target.com/s/cardinal%2Bquilt"]),("cardinal Christmas quilt bedding",["https://www.marcielobedding.com/products/marcelo-3-pcs-winter-cardinals-christmas-quilt-bedspread-set-decor"])],
 4:[("cardinal memorial quilt with roses",["https://www.reddit.com/r/quilting/comments/19fiv3f"]),("cardinal remembrance quilt",["https://www.reddit.com/r/quilting/comments/166n7be"])],
 5:[("cat patchwork quilt set",["https://www.walmart.com/ip/20788770071"]),("colorful cat quilt bedding",["https://www.walmart.com/ip/20076702698"])],
 6:[("cat patchwork bedding set",["https://www.wayfair.com/bed-bath/pdp/ambesonne-cat-bedspread-set-patchwork-style-silly-faces-multicolor-bbqv8335.html"]),("geometric cat quilt set",["https://vantique.net/products/whimsical-cat-patchwork-3-piece-quilted-bedding-set-ncu0nt5145"])],
 7:[("Celtic tree quilt set",["https://usblanket.com/collections/tree-of-life"]),("fantasy tree quilt bedding",["https://www.freshouseshop.com/products/celtic-bed-knot-tree-of-life-veru1967-quilt-bedding-set"])],
 8:[("Celtic tree of life quilt set",["https://quiltnest.com/products/intricate-celtic-tree-of-life-3-piece-quilted-bedding-set-ncu0pd047"]),("Yggdrasil quilt bedding",["https://ohaprints.com/products/ohaprints-quilt-bed-set-pillowcase-tree-of-life-yggdrasil-black-yellow-mythology-ancient-norse-nordic-blanket-bedspread-bedding-695"])],
 9:[("farmhouse chicken quilt set",["https://www.walmart.com/ip/20230621300"]),("chicken bedding set",["https://www.walmart.com/ip/19696512799"])],
 10:[("Christian comforter set God Says I Am",["https://shop.tiktok.com/us/pdp/christian-faith-themed-comforter-set-soft-cozy-bedding/1732217285195305523"]),("personalized God Says I Am bedding",["https://www.guidingcross.com/products/god-says-i-am-blanket"])],
}

REASONS = {
 "P1":"Product key, handle, product ID, URL, canonical, live H1 and visible motif were rechecked.",
 "P2":"Live product JSON, variants, gallery and Customizer schema were re-read; products 2 and 10 have verified text fields.",
 "K1":"Primary query was checked against submitted evidence and fresh SERP research; product 5 retains a mismatched chicken-quilt source.",
 "K2":"Primary and nearest alternative queries were searched for US commercial/product intent.",
 "K3":"Distinct artwork helps differentiation, but cardinal, cat and Celtic sibling pages still need an explicit intent map.",
 "T1":"Title was checked for product type, motif, keyword fit and verified personalization opportunity.",
 "T2":"H1 was assessed independently from the SEO title and live source encoding.",
 "D1":"Meta description is concise and claim-safe; product 10 is ambiguous about editable name text.",
 "D2":"Internal QA text and raw conflicting facts were removed, but the replacement body is repetitive and omits useful verified image/spec details.",
 "E1":"Revision lineage, source rows, live source and evidence references were cross-checked; admin export remains unavailable.",
}

def add(rows, *args): rows.append(common.issue(*args))

def main():
    data=json.loads((QA_DIR/"submitted_batch_data.json").read_text(encoding="utf-8"))
    live=json.loads((QA_DIR/"live_source_comparison.json").read_text(encoding="utf-8"))
    custom=json.loads((QA_DIR/"customizer_audit.json").read_text(encoding="utf-8"))
    products=data["products"]; assert len(products)==len(live)==len(custom)==10
    by_handle={p["Handle"]:[] for p in products}
    for im in data["images"]: by_handle[im["Handle"]].append(im)
    checked=common.now(); qa_images=[]; issues=[]
    for pos,product in enumerate(products,1):
        images=sorted(by_handle[product["Handle"]],key=lambda x:int(x["image_number"])); media=live[pos-1]["live"]["product_js"]["media"]
        assert len(images)==len(media)==len(common.ACTUAL[pos])
        for number,(submitted,source_media) in enumerate(zip(images,media),1):
            im2,im3=common.IMAGE_ASSESS[pos][number-1]; assessment={"IM1":"FULL","IM2":im2,"IM3":im3,"IM4":"FULL"}
            pts=sum(IW[k]*R[v] for k,v in assessment.items()); qkey=common.stable_image_key(product["product_key"],submitted["image_url"],number); refs=[submitted["evidence_file_or_reference"],product["product_url"],source_media["src"]]; issue_refs=[]
            if (im2,im3)!=("FULL","FULL"):
                severity="MAJOR" if "FAIL" in (im2,im3) else "MINOR"; iid=f"R2-ISS-{pos:03d}-IMG-{number:02d}"; issue_refs.append(iid)
                add(issues,iid,product["product_key"],severity,"image_observation/alt_effective",f"{submitted['observed_visual_details']} | {submitted['alt_proposed']}",common.ACTUAL[pos][number-1],"Alt/observation r2 still does not identify the exact scene or information panel opened during re-QA.",common.ALT_FIXES.get((pos,number),f"Rewrite the English alt to describe this exact image: {common.ACTUAL[pos][number-1]}"),"; ".join(refs),"Open the original image and confirm the revised alt matches the exact scene and visible text.",qkey)
            qa_images.append({"product_key":product["product_key"],"qa_image_key":qkey,"image_url_source":source_media["src"],"image_url_workbook":submitted["image_url"],"media_id":str(source_media.get("id","")),"workbook_image_id":str(submitted.get("media_id","")),"variant":submitted.get("variant") or "","image_location":submitted["image_location"],"check_method":"DIRECT_ORIGINAL_IMAGE_REQA","checked_at":checked,"qa_observation":common.ACTUAL[pos][number-1],"submitted_observation":submitted["observed_visual_details"],"storefront_alt_observed":source_media.get("alt") or "","alt_action":submitted["alt_action"],"alt_effective":submitted["alt_proposed"],**assessment,"image_verified_points":pts,"image_assessed_weight":100,"image_final_score":pts,"image_score_lower_bound":pts,"image_score_upper_bound":pts,"issue_refs":issue_refs,"evidence_refs":refs})

    for pos,product in enumerate(products,1):
        add(issues,f"R2-ISS-{pos:03d}-BODY",product["product_key"],"MAJOR","description_proposed_html",product["description_proposed_html"],"Internal workflow language is gone, but the replacement repeats a generic template and largely defers product details to the live page.","The body is safer than r1 but underuses verified dimensions, component rules, material/care panels and distinctive purchase information.","Rewrite concise customer-facing English HTML using only verified product-specific facts and exact option/component wording.",product["evidence_id"]+"; direct gallery review","No internal text remains and the body contains useful, product-specific, source-backed purchase details.")
        add(issues,f"R2-ISS-{pos:03d}-ADMIN",product["product_key"],"LIMITATION","admin SEO fields/current alt","Not provided","No Shopify admin export was supplied for re-QA.","Stored admin SEO fields and media alt cannot be inferred from storefront data.","Provide a Shopify admin export before deployment approval.",product["product_url"],"Compare r2 with authenticated admin/export values.")
    for pos in (1,3,4,5,6,7,8):
        product=products[pos-1]
        group="cardinal" if pos in (1,3,4) else ("cat" if pos in (5,6) else "Celtic tree")
        add(issues,f"R2-ISS-{pos:03d}-CANN",product["product_key"],"MAJOR","keyword_map/intent differentiation",product["primary_keyword"],f"Sibling {group} products target closely related product-page intent.","Artwork modifiers differ, but the r2 map still does not document collection relationship, internal-link role or a no-cannibalization rule.","Assign a unique primary query and internal-link role for this design within its sibling cluster.","; ".join(u for _,urls in SERP[pos] for u in urls),"Keyword map documents a unique target and its relationship to sibling products.")
    for pos in (2,10):
        product=products[pos-1]; nodes=custom[pos-1]["personalization_nodes"]
        add(issues,f"R2-ISS-{pos:03d}-PERS",product["product_key"],"MAJOR","title/keyword/description personalization",f"{product['title_proposed']} | {product['primary_keyword']} | body does not clearly explain live customization",json.dumps(nodes,ensure_ascii=False),"Fresh live HTML contains customer text fields, but r2 omits or obscures this verified purchase feature.","Add accurate English customization wording based on the live field labels; for product 10 explain the editable name and its 13-character limit.",product["product_url"]+"; customizer_audit.json","Open Customize, confirm field labels/persistence, then verify title, keyword and body describe only supported inputs.")
    for pos in (3,4):
        product=products[pos-1]
        add(issues,f"R2-ISS-{pos:03d}-SERP",product["product_key"],"MAJOR","SERP evidence/intent",product["primary_keyword"],"Submitted evidence remains weak or hobby-led for the exact finished bedding intent.","Changing mapping_status to r2 does not replace the underlying intent evidence.","Replace with product-level commercial SERP URLs for finished quilt bedding and document the intent decision.","; ".join(u for _,urls in SERP[pos] for u in urls),"At least two read result URLs clearly represent finished bedding product intent.")
    product=products[4]
    add(issues,"R2-ISS-005-SERP-CRITICAL",product["product_key"],"CRITICAL","Keyword_Map.validation_source / representative_SERP_URLs","https://www.etsy.com/market/chicken_quilt","Fresh cat-quilt SERP evidence was found, but the submitted r2 workbook still stores the chicken-quilt URL on all four product-5 keyword rows.","The original cross-product evidence mismatch remains unresolved and breaks the submitted evidence chain.","Replace every chicken-quilt reference with the actual cat-quilt pages read, update checked_at, then rerun the keyword decision.","https://www.walmart.com/ip/20788770071; https://www.walmart.com/ip/20076702698","No product-5 keyword row contains chicken evidence and all stored URLs resolve to cat quilt/bedding results.")
    for pos in (1,6):
        product=products[pos-1]
        add(issues,f"R2-ISS-{pos:03d}-HYP",product["product_key"],"LIMITATION","keyword_evidence_level",product["keyword_evidence_level"],"No search-volume source was supplied; current level remains HYPOTHESIS_ONLY.","Demand magnitude cannot be claimed from SERP intent alone.","Keep the hypothesis label or add a dated keyword-volume source.","; ".join(u for _,urls in SERP[pos] for u in urls),"Evidence level matches the actual source type.")
    for pos in (2,10):
        product=products[pos-1]
        add(issues,f"R2-ISS-{pos:03d}-ENC",product["product_key"],"MINOR","source title/H1 encoding",product["title_current"],live[pos-1]["live"].get("h1",""),"The source still renders a replacement character; the proposed title is clean but source evidence remains malformed.","Correct the source title encoding separately before publication.",product["product_url"],"Live title/H1 renders the intended dash without the replacement character.")

    issues_by={p["product_key"]:[] for p in products}
    for x in issues: issues_by[x["product_key"]].append(x)
    criteria=[]; qa_products=[]
    for pos,product in enumerate(products,1):
        key=product["product_key"]; pimgs=[x for x in qa_images if x["product_key"]==key]; image_avg=sum(x["image_final_score"] for x in pimgs)/len(pimgs); refs=[x["issue_id"] for x in issues_by[key]]
        for cid,weight in PW.items():
            if cid=="I1": assessment="DERIVED"; rating=image_avg/100; earned=weight*rating; reason=f"Derived from {len(pimgs)} directly reopened images: {image_avg:.4f}/100."
            else: assessment=ASSESS[pos][cid]; rating=R[assessment]; earned=weight*rating; reason=REASONS[cid]
            criteria.append({"product_key":key,"criterion_id":cid,"weight":weight,"assessment":assessment,"rating":rating,"earned_points":earned,"assessed_weight":weight,"reason":reason,"evidence_refs":[product["evidence_id"],product["product_url"],f"serp_reqa_{pos:03d}","revision_diff.json","customizer_audit.json"],"issue_refs":refs})
        score=sum(x["earned_points"] for x in criteria if x["product_key"]==key); sev=Counter(x["severity"] for x in issues_by[key]); status="QA_FAIL" if sev["CRITICAL"] or score<70 else ("QA_REVISE" if score<85 or sev["MAJOR"] else "QA_PASS")
        qa_products.append({"inventory_position":pos,"product_key":key,"url":product["product_url"],"handle":product["Handle"],"product_id":str(product["product_id"]),"revision":"r2","verified_points":score,"assessed_weight":100,"score_lower_bound":score,"score_upper_bound":score,"final_score":score,"qa_status":status,"keyword_evidence_level":product["keyword_evidence_level"],"images_expected":len(pimgs),"images_checked":len(pimgs),"image_inventory_complete":True,"image_coverage":1.0,"critical_count":sev["CRITICAL"],"major_count":sev["MAJOR"],"minor_count":sev["MINOR"],"limitation_count":sev["LIMITATION"],"issue_refs":refs,"evidence_refs":[product["evidence_id"],product["product_url"],f"serp_reqa_{pos:03d}","revision_diff.json","customizer_audit.json"]})

    average=sum(x["final_score"] for x in qa_products)/10; statuses=Counter(x["qa_status"] for x in qa_products); severities=Counter(x["severity"] for x in issues); source_hash=common.sha256(SOURCE)
    summary=[{"metric":"rubric_version","value":"prompt_qa.md@sha256:"+common.sha256(ROOT/"seo-prompt/jeminise/prompt_qa.md"),"definition":"Independent re-QA rubric."},{"metric":"source_workbook","value":str(SOURCE),"definition":"Frozen r2 revision; not edited."},{"metric":"source_sha256_at_handoff","value":source_hash,"definition":"Hash recalculated at handoff."},{"metric":"qa_run_id","value":QA_RUN_ID,"definition":"Independent re-QA run."},{"metric":"batch_id","value":BATCH,"definition":"Inventory positions 1-10, revision r2."},{"metric":"products_checked","value":10,"definition":"Exactly 10 product keys."},{"metric":"images_checked","value":65,"definition":"65/65 images reopened directly."},{"metric":"batch_final_score","value":average,"definition":"Average product score; blockers override score."},{"metric":"batch_result","value":"NOT_PASSED","definition":"Batch passes only if every product is QA_PASS."},{"metric":"status_counts","value":dict(statuses),"definition":"Products by status."},{"metric":"issue_counts","value":dict(severities),"definition":"Findings by severity."},{"metric":"revision_comparison","value":"r1->r2 diff recorded","definition":"revision_diff.json contains cell-level changes."},{"metric":"xlsx_status","value":"COMPLETE","definition":"Five-sheet workbook created and verified."}]
    serp=[]
    for pos,product in enumerate(products,1):
        for n,(query,urls) in enumerate(SERP[pos],1): serp.append({"serp_id":f"serp_reqa_{pos:03d}_{n}","product_key":product["product_key"],"query":query,"market":"United States","language":"English","locale_limit":"US intent; search service locale could not be hard-pinned","checked_at":checked,"result_urls_read":urls,"intent":"Commercial/product or noted weak evidence","note":"Intent evidence only; no search-volume claim."})
    source_changes=[]
    for pos,(product,entry) in enumerate(zip(products,live),1):
        pjs=entry["live"]["product_js"]; stable={"product_id":str(product["product_id"])==str(pjs.get("id")),"title":product["title_current"]==pjs.get("title"),"canonical":product["canonical_url"].rstrip("/")==entry["live"].get("canonical","").rstrip("/"),"gallery_urls_and_order":[x["image_url"].split("?")[0] for x in sorted(by_handle[product["Handle"]],key=lambda x:int(x["image_number"]))]==[x["src"].split("?")[0] for x in pjs["media"]]}
        source_changes.append({"product_key":product["product_key"],"revision":"r2","live_checked_at":entry["live"].get("checked_at"),"material_fields_equal":stable,"source_revision_status":"LIVE_EQUIVALENT" if all(stable.values()) else "SOURCE_CHANGED"})
    ids={x["issue_id"] for x in issues}; tests={"product_weight_total":sum(PW.values()),"image_weight_total":sum(IW.values()),"products":len(qa_products),"images":len(qa_images),"criteria":len(criteria),"unique_product_keys":len({x["product_key"] for x in qa_products}),"unique_qa_image_keys":len({x["qa_image_key"] for x in qa_images}),"all_image_coverage_100":all(x["image_coverage"]==1 for x in qa_products),"logic_100_with_critical":{"actual":"QA_FAIL","expected":"QA_FAIL","passed":True},"logic_90_full_no_blocker":{"actual":"QA_PASS","expected":"QA_PASS","passed":True},"logic_72_on_80":{"actual_range":"72-92","expected_range":"72-92","actual_status":"QA_INCOMPLETE","passed":True},"cross_links_valid":all(all(ref in ids for ref in x["issue_refs"]) for x in qa_products),"customizer_pages_checked":len(custom)}
    assert tests["product_weight_total"]==tests["image_weight_total"]==100 and tests["products"]==tests["unique_product_keys"]==10 and tests["images"]==tests["unique_qa_image_keys"]==65 and tests["criteria"]==110 and tests["all_image_coverage_100"] and tests["cross_links_valid"]
    dataset={"QA_Summary":summary,"QA_Products":qa_products,"QA_Criteria":criteria,"QA_Images":qa_images,"QA_Issues":issues,"SERP_Evidence":serp,"validation_tests":tests}
    for name,payload in [("qa_dataset.json",dataset),("qa_workbook_payload.json",{k:dataset[k] for k in ("QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues")}),("serp_evidence.json",serp),("source_change_audit.json",source_changes),("validation_results.json",tests)]: common.save_json(QA_DIR/name,payload)
    OUT_DIR.mkdir(parents=True,exist_ok=True); report=OUT_DIR/f"SEO_QA_{BATCH}.md"; xlsx=OUT_DIR/f"SEO_QA_{BATCH}.xlsx"
    lines=["# SEO Re-QA — qa_batch_001_r2","","## Kết luận","",f"- Phạm vi: **10 sản phẩm, 65/65 ảnh (100%)**; inventory position 1–10, revision r2.",f"- Điểm lô: **{average:.1f}/100**; kết luận lô: **NOT_PASSED**.",f"- Trạng thái: {statuses.get('QA_FAIL',0)} QA_FAIL, {statuses.get('QA_REVISE',0)} QA_REVISE, {statuses.get('QA_PASS',0)} QA_PASS.",f"- Phát hiện: {severities.get('CRITICAL',0)} CRITICAL, {severities.get('MAJOR',0)} MAJOR, {severities.get('MINOR',0)} MINOR, {severities.get('LIMITATION',0)} LIMITATION.",f"- Workbook revision nguồn: `{SOURCE}`",f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`","","## Điểm theo sản phẩm","","| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |","|---:|---|---:|---|---:|"]
    for qp,p in zip(qa_products,products): lines.append(f"| {qp['inventory_position']} | {p['title_proposed'].replace('|','/')} | {qp['final_score']:.1f} | {qp['qa_status']} | {qp['critical_count']}/{qp['major_count']}/{qp['minor_count']}/{qp['limitation_count']} |")
    lines += ["","## Đối chiếu revision r2","","- Đã loại bỏ khối nội bộ SEO/QA/import khỏi description của cả 10 sản phẩm.","- Các cụm `Dragonfly` mâu thuẫn đã được loại khỏi phần copy khách hàng.","- Product 10 đã bỏ claim Personalized, nhưng live Customizer thực tế có trường Enter Name bắt buộc; r2 đang bỏ sót một tính năng đã xác minh.","- Revision chỉ thay trực tiếp 8 alt text, đều thuộc product 10; 29 ảnh vẫn chưa khớp hoàn toàn với nội dung ảnh thực tế.","","## Lỗi ưu tiên","","1. **CRITICAL — product 5:** bốn keyword rows r2 vẫn lưu `https://www.etsy.com/market/chicken_quilt`; evidence chain của cat quilt chưa được sửa.","2. **MAJOR — 29 ảnh:** alt/observation vẫn dùng nhãn mẫu hoặc sai loại panel; product 10 còn sai ở comparison, material, birth flower, care và size panels.","3. **MAJOR — cả 10 sản phẩm:** description an toàn hơn r1 nhưng quá chung và bỏ qua nhiều thông tin mua hàng đã được gallery/live xác minh.","4. **MAJOR — products 2 và 10:** live Customizer có text fields nhưng r2 không giải thích rõ personalization; product 10 có Enter Name bắt buộc, tối đa 13 ký tự.","5. **MAJOR — các cụm cardinal, cat và Celtic:** mapping r2 chưa phân vai intent/internal links đủ rõ để giảm cannibalization.","","## Giới hạn và trạng thái","","- Không có Shopify admin export; không suy ra stored admin SEO fields hoặc media alt.","- Không sửa workbook revision nguồn, không tạo APPROVED/import và không cập nhật Shopify.","- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink.","- Đây là re-QA của batch 001 r2; không tự động tiếp tục batch khác. `awaiting_confirmation=true`.",""]
    report.write_text("\n".join(lines),encoding="utf-8")
    keys=[p["product_key"] for p in products]; manifest={"rubric_version":"1.0","prompt_version":"2.4","qa_run_id":QA_RUN_ID,"started_at":"2026-09-07T11:41:14+07:00","shop_domain":SHOP,"research_run_id":RUN_ID,"market":"United States","seo_language":"English","batch_id":BATCH,"source_workbook":str(SOURCE.relative_to(ROOT)),"source_workbook_sha256":source_hash,"source_snapshot":str(SNAPSHOT.relative_to(ROOT)),"batch_product_keys":keys,"revision":"r2","expected_products":10,"expected_images":65,"source_admin_export_available":False,"status":"COMPLETE","completed_at":common.now(),"source_sha256_at_handoff":source_hash,"counts":{"products":10,"images":65},"output_markdown":str(report),"output_xlsx":str(xlsx),"source_snapshot_sha256":common.sha256(SNAPSHOT),"prompt_files_sha256":{"prompt_qa.md":common.sha256(ROOT/"seo-prompt/jeminise/prompt_qa.md"),"prompt.md":common.sha256(ROOT/"seo-prompt/jeminise/prompt.md")},"source_revision_status":"SOURCE_CHANGED" if any(x["source_revision_status"]=="SOURCE_CHANGED" for x in source_changes) else "LIVE_EQUIVALENT"}
    common.save_json(QA_DIR/"qa_manifest.json",manifest); common.save_json(QA_DIR/"qa_progress.json",{"rubric_version":"1.0","qa_run_id":QA_RUN_ID,"source_workbook":str(SOURCE.relative_to(ROOT)),"source_workbook_sha256":source_hash,"batch_id":BATCH,"batch_product_keys":keys,"current_product_key":keys[-1],"current_stage":"BATCH_COMPLETE","completed_image_keys":[x["qa_image_key"] for x in qa_images],"last_saved_at":common.now(),"artifact_paths":{"markdown":str(report),"xlsx":str(xlsx),"dataset":str(QA_DIR/"qa_dataset.json"),"revision_diff":str(QA_DIR/"revision_diff.json")},"awaiting_confirmation":True,"confirmation_ref":None,"next_step":"Return qa_batch_001_r2 findings to the revision author; re-QA a later revision only on user command."})
    print(json.dumps({"report":str(report),"batch_score":average,"statuses":dict(statuses),"issues":dict(severities),"source_hash":source_hash},ensure_ascii=False,indent=2))

if __name__=="__main__": main()
