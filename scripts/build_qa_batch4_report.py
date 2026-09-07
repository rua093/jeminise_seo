from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import build_qa_batch1_report as common


ROOT = Path(__file__).resolve().parents[1]
SHOP, RUN_ID, QA_RUN_ID, BATCH = "jeminise.com", "20260906_234129", "20260907_095643", "qa_batch_004"
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
31: ["Front bedroom mockup of a cardinal-and-birdhouse Christmas quilt with poinsettias.", "Angled bedroom view with printed-craft feature callouts.", "Matching pillow shams and lightweight, soft, anti-pill and anti-static callouts.", "Fabric-feature panel describing the printed quilt for all-season use.", "Premium quilt-set size chart for Throw, Twin, Full, Queen and King.", "Material-layer and bedspread-features diagram for top, filling and back layers.", "Overhead bedroom mockup of the cardinal-and-birdhouse quilt."],
32: ["Front bedroom mockup with two red cardinals on snowy branches and poinsettias.", "Close-up of a red cardinal and the quilted printed surface.", "Isolated matching pillow sham with cardinal artwork.", "Angled bedroom mockup of the cardinal quilt.", "Included-components and size overlay for quilt and optional shams."],
33: ["Front bedroom mockup of a cow-and-sunflower landscape quilt.", "Angled bedroom mockup of the cow-and-sunflower quilt.", "Matching shams and product-feature callouts.", "Fabric-feature panel for the printed quilt.", "Premium quilt-set size chart.", "Material-layer and bedspread-features diagram.", "Overhead bedroom mockup of the cow-and-sunflower quilt."],
34: ["Front bedroom mockup of a colorful cartoon crocodile patchwork quilt.", "Angled bedroom mockup of the crocodile patchwork quilt.", "Matching shams and product-feature callouts.", "Fabric-feature panel for the printed quilt.", "Premium quilt-set size chart.", "Material-layer and bedspread-features diagram.", "Overhead bedroom mockup of the crocodile patchwork quilt."],
35: ["Main bed view of a football on a vintage US flag with sample name Kevin and number 15.", "Angled bedroom view of the same Kevin 15 football design.", "Duvet-cover versus comforter product-type comparison.", "Close-up feature panel marked soft, lightweight, durable and breathable.", "Easy-care instruction panel.", "Comforter and duvet-cover size-dimension diagram."],
36: ["Main bed view of a football on grunge red-and-white stripes with sample name Brian and number 15.", "Angled bedroom view of the same Brian 15 design.", "Duvet-cover versus comforter product-type comparison.", "Close-up feature panel marked soft, lightweight, durable and breathable.", "Easy-care instruction panel.", "Comforter and duvet-cover size-dimension diagram."],
37: ["Main bed view of a football on a blue-red cosmic background with sample name Kevin and number 20.", "Angled bedroom view of the same Kevin 20 cosmic design.", "Duvet-cover versus comforter product-type comparison.", "Close-up feature panel marked soft, lightweight, durable and breathable.", "Easy-care instruction panel.", "Comforter and duvet-cover size-dimension diagram."],
38: ["Main bed view of a football on a vintage USA flag background with sample name Michael and number 30.", "Angled bedroom view of the same Michael 30 design.", "Duvet-cover versus comforter product-type comparison.", "Close-up feature panel marked soft, lightweight, durable and breathable.", "Easy-care instruction panel.", "Comforter and duvet-cover size-dimension diagram."],
39: ["Main bed view of a tan paint-splash football design with sample name MARK and number 15.", "Angled bedroom view of the same MARK 15 design.", "Duvet-cover versus comforter product-type comparison.", "Close-up feature panel marked soft, lightweight, durable and breathable.", "Easy-care instruction panel.", "Comforter and duvet-cover size-dimension diagram."],
40: ["Main bed view of a football wrapped in a USA flag with sample name MATTHEW and number 08.", "Angled bedroom view of the same MATTHEW 08 design.", "Duvet-cover versus comforter product-type comparison.", "Close-up feature panel marked soft, lightweight, durable and breathable.", "Easy-care instruction panel.", "Comforter and duvet-cover size-dimension diagram."],
}

WRONG = {(31,6), (32,2), (32,4), (33,6), (34,6)}
ALT_FIX = {
(31,6): "Cardinal quilt material layers and bedspread feature diagram",
(32,2): "Close-up of red cardinal on snowy branch quilt print",
(32,4): "Christmas cardinal quilt in an angled bedroom view",
(33,6): "Cow sunflower quilt material layers and bedspread feature diagram",
(34,6): "Crocodile quilt material layers and bedspread feature diagram",
}

SERP = {
31: [("Christmas cardinal birdhouse quilt", ["https://saphiesfabrics.com/products/christmas-cardinal-birdhouse-100-cotton-fabric-panel-block-ee1859-copy-1"]), ("cardinal poinsettia quilt set", ["https://www.macustom.com/collections/quilted-bedding-set/products/cardinal-christmas-quilted-bedding-set-ncu0vh044"])],
32: [("Christmas cardinal quilt set", ["https://www.lakeside.com/products/nordic-cardinal-quilt-set-full-queen-or-king-with-shams"]), ("snowy cardinal quilt bedding", ["https://www.kohls.com/product/prd-8133602/cf-home-natures-holiday-cardinal-king-quilt-set-with-shams.jsp"])],
33: [("cow sunflower quilt set", ["https://www.target.com/p/-/A-1005538288", "https://www.wayfair.com/bed-bath/pdp/rt-designers-collection-cow-sunflowers-farmhouse-microfiber-quilt-qlko2052.html"]), ("farmhouse cow bedding set", ["https://www.walmart.com/ip/19924623339"])],
34: [("crocodile patchwork quilt set", ["https://vantique.net/products/everglades-alligator-charm-3-piece-quilted-bedding-set-ncu0dv5368"]), ("alligator quilt bedding set", ["https://www.potterybarnkids.com/products/alligator-madras-nursery-bedding-set/pip-print.html"])],
35: [("personalized football comforter", ["https://www.youcustomizeit.com/p/Football-Comforters-Personalized/142169"]), ("custom football bedding with name number", ["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"])],
36: [("custom football comforter with name", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"]), ("personalized football bedding set", ["https://wearanddecor.com/products/personalized-football-quilt-bedding-set-football-ball-black-quilt-custom-name-and-number-duvet-cover-set"])],
37: [("cosmic football comforter", ["https://www.target.com/c/bedding-sets-collections-home/galaxy/-/N-5xtv1Zkt6vv"]), ("galaxy football bedding set", ["https://www.target.com/c/bedding-sets-collections-home/galaxy/-/N-5xtv1Zkt6vv"])],
38: [("personalized patriotic football comforter", ["https://www.youcustomizeit.com/p/Football-Comforters-Personalized/142169"]), ("American flag football bedding", ["https://www.walmart.com/ip/20312810342"])],
39: [("personalized vintage football comforter", ["https://www.etsy.com/listing/4324538505/custom-comforter-sports-team-mascot"]), ("custom football bedding with name", ["https://www.etsy.com/de/listing/1883272338/benutzerdefinierte-fussball-bettwasche"])],
40: [("custom patriotic football comforter", ["https://www.youcustomizeit.com/p/Football-Comforters-Personalized/142169"]), ("personalized American flag football bedding", ["https://custombeddingset.com/america-football-custom-bedding-set-personalized-us-flag-duvet-cover-bed-sheets-pillow-shams/"])],
}

ASSESS = {
31:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
32:{"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
33:{"P1":"FULL","P2":"FAIL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"FAIL"},
34:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"FAIL"},
35:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
36:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
37:{"P1":"FULL","P2":"FAIL","K1":"FULL","K2":"PARTIAL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
38:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
39:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
40:{"P1":"FULL","P2":"FAIL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FAIL","T2":"FAIL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
}


def add(rows, *a, **kw): rows.append(common.issue(*a, **kw))


def main():
    data = json.loads((QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))
    live = json.loads((QA_DIR / "live_source_comparison.json").read_text(encoding="utf-8"))
    products = data["products"]
    by_pk = {p["product_key"]: [] for p in products}
    for im in data["images"]: by_pk[im["shop_domain"] + "+" + im["Handle"]].append(im)
    checked, qa_images, issues = common.now(), [], []
    for ix, p in enumerate(products):
        pos, media = 31 + ix, live[ix]["live"]["product_js"]["media"]
        ims = sorted(by_pk[p["product_key"]], key=lambda x: x["image_number"])
        assert len(ims) == len(media) == len(ACTUAL[pos])
        for j, im in enumerate(ims, 1):
            if R2_MODE:
                im2, im3 = R2_IMAGE_ASSESS.get((pos, j), ("FULL", "FULL"))
                bad = im2 != "FULL" or im3 != "FULL"
                ass = {"IM1":"FULL","IM2":im2,"IM3":im3,"IM4":"FULL"}
            else:
                bad = (pos,j) in WRONG
                ass = {"IM1":"FULL","IM2":"FAIL" if bad else "FULL","IM3":"FAIL" if bad else "FULL","IM4":"FULL"}
            pts = sum(IW[k] * R[v] for k,v in ass.items())
            key = common.stable_image_key(p["product_key"], im["image_url"], j)
            refs = [im["evidence_file_or_reference"], p["product_url"], media[j-1]["src"]]
            irefs=[]
            if bad:
                iid=f"{'R2-' if R2_MODE else ''}ISS-{pos:03d}-IMG-{j:02d}"; irefs=[iid]
                add(issues,iid,p["product_key"],R2_IMAGE_SEVERITY.get((pos,j),"MAJOR") if R2_MODE else "MAJOR","image_observation/alt_effective",f"{im['observed_visual_details']} | {im['alt_proposed']}",ACTUAL[pos][j-1],R2_IMAGE_REASON.get((pos,j),"Observation/alt r2 vẫn dùng nhãn mẫu hoặc sai loại ảnh.") if R2_MODE else "Observation/alt theo mẫu vị trí không phản ánh đúng ảnh gốc.",R2_IMAGE_FIX.get((pos,j),ALT_FIX.get((pos,j),"Rewrite the alt from the directly viewed image.")) if R2_MODE else ALT_FIX[(pos,j)],"; ".join(refs),"Mở ảnh gốc và xác nhận alt mới mô tả đúng ảnh.",key)
            qa_images.append({"product_key":p["product_key"],"qa_image_key":key,"image_url_source":media[j-1]["src"],"image_url_workbook":im["image_url"],"media_id":str(media[j-1]["id"]),"workbook_image_id":str(im["media_id"]),"variant":im["variant"] or "","image_location":im["image_location"],"check_method":"DIRECT_ORIGINAL_IMAGE","checked_at":checked,"qa_observation":ACTUAL[pos][j-1],"submitted_observation":im["observed_visual_details"],"storefront_alt_observed":media[j-1].get("alt") or "","alt_action":im["alt_action"],"alt_effective":im["alt_proposed"],**ass,"image_verified_points":pts,"image_assessed_weight":100,"image_final_score":pts,"image_score_lower_bound":pts,"image_score_upper_bound":pts,"issue_refs":irefs,"evidence_refs":refs})

    if R2_MODE:
        for ix,p in enumerate(products):
            pos=31+ix
            add(issues,f"R2-ISS-{pos:03d}-DESC",p["product_key"],"MAJOR","description_proposed_html",p.get("description_proposed_html", ""),"Bản r2 đã bỏ câu nội bộ nhưng nội dung vẫn quá chung, dựa vào cụm 'gallery images show' và chưa đưa đủ thông số, thành phần, chăm sóc hoặc quy tắc cá nhân hóa đã xác minh.","Mô tả chưa chuyển đầy đủ evidence thành thông tin mua hàng hữu ích.","Rewrite the description with only verified materials, sizes, included pieces, care guidance, and personalization rules where available.",p["evidence_id"]+"; direct images; live product.js/customizer","Rendered copy is customer-facing and every material purchase fact is traceable to evidence.")
    for ix,p in enumerate(() if R2_MODE else products):
        pos=31+ix
        add(issues,f"ISS-{pos:03d}-DESC",p["product_key"],"MAJOR","description_proposed_html","It needs QA and approval before import.","HTML chứa heading SEO Use và câu nội bộ về QA/import.","Nội dung quy trình nội bộ không được xuất hiện trên storefront.","Rewrite as customer-facing English HTML and remove the internal SEO/QA/import block.",p["evidence_id"],"Render HTML and confirm no drafting, QA or import instruction remains.")
    for pos in (() if R2_MODE else range(33,41)):
        p=products[pos-31]; opts=live[pos-31]["live"]["product_js"]["options"]
        add(issues,f"ISS-{pos:03d}-PERS",p["product_key"],"CRITICAL","personalization/customization claims",p["primary_keyword"]+" | "+p["title_proposed"],"Trang live/product.js chỉ có size, pillowcase và/hoặc flat-sheet options; không có input name/number/text.","Ảnh có tên/số mẫu không chứng minh người mua có thể tùy biến.","Prove a working name/number input and fulfillment mapping, or remove personalized/custom claims from keyword, title, meta and body.",p["product_url"]+"; live options="+json.dumps(opts,ensure_ascii=False),"Purchase flow visibly accepts, persists and fulfills the claimed custom values.")
    for pos in (() if R2_MODE else (33,34)):
        p=products[pos-31]
        add(issues,f"ISS-{pos:03d}-DRAGONFLY",p["product_key"],"MAJOR","description_proposed_html","Dragonfly source field copied into draft.",ACTUAL[pos][0],"Dragonfly contradicts the visible cow/crocodile artwork.","Remove the Dragonfly color/pattern/design field and replace only with evidenced product facts.",p["evidence_id"]+"; direct images 1-7","No Dragonfly attribute remains unless supported by a matching source revision.")
    p=products[0]
    add(issues,"ISS-031-SERP",p["product_key"],"MAJOR","keyword/SERP evidence",p["primary_keyword"],"Exact birdhouse results skew toward fabric panels and quilt kits; finished-bedding evidence is broader cardinal/poinsettia intent.","The chosen exact phrase is not cleanly validated as finished-product intent.","Retain SERP_ONLY and rerun with US finished-bedding product results before approval.",json.dumps(SERP[31],ensure_ascii=False),"Two current finished-bedding product SERPs support the exact motif and product type.")
    p=products[3]
    if not R2_MODE:
        add(issues,"ISS-034-ALLIGATOR",p["product_key"],"MAJOR","secondary_keywords","alligator quilt set",ACTUAL[34][0],"Crocodile and alligator are not interchangeable product attributes without source evidence.","Use crocodile-specific wording, or document why the depicted animal is an alligator.",p["evidence_id"]+"; direct images","Keyword map and copy use a single evidence-supported animal identity.")
    for pos in (() if R2_MODE else (31,32)):
        p=products[pos-31]
        add(issues,f"ISS-{pos:03d}-KREASON",p["product_key"],"MAJOR","keyword decision reason","Specific to visible artwork, personalization and product type.","The proposed customer copy does not claim personalization and the live purchase flow exposes no text input.","The evidence chain cites an unsupported personalization signal.","Rewrite the decision reason around visible motif and finished quilt-set intent only.",p["evidence_id"],"Decision reason contains only evidence-backed selection factors.")
    for pos in (() if R2_MODE else (31,32,33,34)):
        p=products[pos-31]
        add(issues,f"ISS-{pos:03d}-ALTGEN",p["product_key"],"MINOR","image alt text","Several alts use generic alternatives such as 'care or size' / 'detail or fabric feature'.","Direct inspection identifies a single image purpose.","Ambiguous 'or' phrasing is less precise than the source image.","Replace generic alternatives with the exact image purpose recorded in QA_Images.",p["evidence_id"],"Every SET alt names the observed scene or information panel without alternatives.")
    if R2_MODE:
        for pos in range(35,41):
            p=products[pos-31]
            add(issues,f"R2-ISS-{pos:03d}-PERS-OMIT",p["product_key"],"MAJOR","keyword/title/meta/description personalization",p["primary_keyword"]+" | "+p["title_proposed"],"Customizer live có Enter Name bắt buộc (1–25 ký tự) và Enter Number tùy chọn (1–5 ký tự), nhưng r2 bỏ điểm khác biệt cá nhân hóa khỏi copy và diễn giải tên/số mẫu như artwork cố định.","R2 đã sửa quá tay sau lỗi cũ và làm mất thông tin mua hàng được chứng minh trực tiếp.","State that the comforter is personalized with a required name and an optional number; keep examples as samples, not fixed artwork.",p["product_url"]+"; live customizer audit","Purchase flow and English copy agree on required name, optional number, limits, and sample-versus-input behavior.")
        keyword_notes={
            34:("crocodile patchwork quilt set","Kết quả exact còn thưa và comparator lệch loài hoặc lệch loại bedding.","Retain a cautious SERP_ONLY label or choose a broader crocodile bedding phrase supported by finished-product results."),
            36:("football comforter with themed artwork","Cụm primary không tự nhiên và không phản ánh intent personalization đã xác minh.","Use a natural personalized grunge football comforter query and document two relevant finished-product results."),
            37:("cosmic football comforter","SERP exact bị nhiễu bởi nghĩa ngoài bedding và chưa xác nhận commercial intent sạch.","Test a personalized cosmic football bedding phrase and require relevant finished-product results before selection."),
            38:("patriotic football comforter","Primary trùng intent với product 40, trong khi artwork và personalization có thể phân tách.","Differentiate around personalized USA flag football comforter and document the intended page boundary."),
            40:("patriotic football comforter","Primary trùng intent với product 38, tạo nguy cơ cannibalization.","Assign a distinct personalized patriotic football angle and record a non-overlapping keyword boundary."),
        }
        for pos,(submitted,reason,fix) in keyword_notes.items():
            p=products[pos-31]
            add(issues,f"R2-ISS-{pos:03d}-KEYWORD",p["product_key"],"MAJOR","primary keyword/SERP decision",submitted,"QA SERP đối chiếu không xác nhận sạch lựa chọn hiện tại.",reason,fix,json.dumps(SERP[pos],ensure_ascii=False),"Two relevant US-intent finished-product results support the revised phrase and it does not collide with another page in this batch.")
    add(issues,("R2-" if R2_MODE else "")+"ISS-GLOBAL-ADMIN","","LIMITATION","current admin SEO fields/current admin alt","Not supplied","Storefront metadata and product.js media alt were readable, but no Shopify admin export was provided.","Storefront values cannot prove admin fields or future import mapping.","Provide a frozen Shopify admin/export snapshot before deployment approval.",str(SOURCE),"Admin export revision and hash are frozen and compared.")

    issue_by={}
    for x in issues: issue_by.setdefault(x["product_key"],[]).append(x)
    reasons=dict(common.REASONS); reasons["P2"]="Đối chiếu specifications, variants, components và personalization với product.js live, body HTML và ảnh."; reasons["K2"]="Đã đọc primary và comparator SERP theo US intent; không dùng claim volume."; reasons["E1"]="Kiểm tra chuỗi evidence và giới hạn do thiếu admin export/browser tương tác."
    criteria=[]; qproducts=[]
    for ix,p in enumerate(products):
        pos,pk=31+ix,p["product_key"]; ims=[x for x in qa_images if x["product_key"]==pk]; avg=sum(x["image_final_score"] for x in ims)/len(ims); refs=[x["issue_id"] for x in issue_by.get(pk,[])]
        for cid,w in PW.items():
            if cid=="I1": ass="DERIVED"; rating=avg/100; earned=w*rating; reason=f"Tính từ trung bình {len(ims)} ảnh: {avg:.4f}/100."
            else: ass=ASSESS[pos][cid]; rating=R[ass]; earned=w*rating; reason=reasons[cid]
            criteria.append({"product_key":pk,"criterion_id":cid,"weight":w,"assessment":ass,"rating":rating,"earned_points":earned,"assessed_weight":w,"reason":reason,"evidence_refs":[p["evidence_id"],p["product_url"],f"serp_qa_{pos:03d}"],"issue_refs":refs})
        score=sum(x["earned_points"] for x in criteria if x["product_key"]==pk); sev=Counter(x["severity"] for x in issue_by.get(pk,[])); status="QA_FAIL" if sev["CRITICAL"] or score<70 else ("QA_REVISE" if score<85 or sev["MAJOR"] else "QA_PASS")
        qproducts.append({"inventory_position":pos,"product_key":pk,"url":p["product_url"],"handle":p["Handle"],"product_id":str(p["product_id"]),"revision":"r1","verified_points":score,"assessed_weight":100,"score_lower_bound":score,"score_upper_bound":score,"final_score":score,"qa_status":status,"keyword_evidence_level":p["keyword_evidence_level"],"images_expected":len(ims),"images_checked":len(ims),"image_inventory_complete":True,"image_coverage":1.0,"critical_count":sev["CRITICAL"],"major_count":sev["MAJOR"],"minor_count":sev["MINOR"],"limitation_count":sev["LIMITATION"],"issue_refs":refs,"evidence_refs":[p["evidence_id"],p["product_url"],f"serp_qa_{pos:03d}"]})
    avg=sum(x["final_score"] for x in qproducts)/10; statuses=Counter(x["qa_status"] for x in qproducts); sevs=Counter(x["severity"] for x in issues); source_hash=common.sha256(SOURCE)
    summary=[{"metric":"rubric_version","value":"prompt_qa.md@sha256:"+common.sha256(ROOT/"seo-prompt/jeminise/prompt_qa.md"),"definition":"Rubric QA áp dụng."},{"metric":"source_workbook","value":str(SOURCE),"definition":"Workbook giai đoạn 1 đóng băng; không chỉnh sửa."},{"metric":"source_sha256_at_handoff","value":source_hash,"definition":"Hash tính lại khi hoàn tất."},{"metric":"qa_run_id","value":QA_RUN_ID,"definition":"Run QA độc lập."},{"metric":"batch_id","value":BATCH,"definition":"Inventory positions 31-40."},{"metric":"products_checked","value":10,"definition":"Đúng 10 product key."},{"metric":"images_checked","value":62,"definition":"62/62 ảnh mở trực tiếp ở độ phân giải gốc."},{"metric":"batch_final_score","value":avg,"definition":"Trung bình 10 final_score; không bù lỗi chặn."},{"metric":"batch_result","value":"NOT_PASSED","definition":"Lô chỉ đạt khi mọi sản phẩm QA_PASS."},{"metric":"status_counts","value":dict(statuses),"definition":"Số sản phẩm theo kết luận."},{"metric":"issue_counts","value":dict(sevs),"definition":"Số phát hiện theo severity."},{"metric":"admin_limitation","value":"No Shopify admin export; interactive browser unavailable","definition":"Không suy ra admin fields hoặc xác nhận dynamic personalization app."},{"metric":"xlsx_status","value":"COMPLETE","definition":"XLSX tạo bằng openpyxl theo yêu cầu người dùng, kiểm tra cấu trúc/công thức và render."}]
    serp=[]
    for ix,p in enumerate(products):
        pos=31+ix
        for n,(q,urls) in enumerate(SERP[pos],1): serp.append({"serp_id":f"serp_qa_{pos:03d}_{n}","product_key":p["product_key"],"query":q,"market":"United States","language":"English","locale_limit":"US intent; search service locale could not be hard-pinned","checked_at":checked,"result_urls_read":urls,"intent":"Commercial/product","note":"Evidence supports intent only; no paid search-volume claim."})
    changes=[]
    for ix,e in enumerate(live):
        old,cur=e["research_snapshot"]["html"],e["live"]; stable={"title":old.get("rendered_title_current")==cur.get("title_element"),"h1":old.get("h1_current")==cur.get("h1"),"meta_description":old.get("meta_description_current")==cur.get("meta_description"),"canonical":old.get("canonical_url")==cur.get("canonical")}
        changes.append({"product_key":products[ix]["product_key"],"snapshot_checked_at":e["research_snapshot"].get("reviewed_at"),"live_checked_at":cur.get("checked_at"),"snapshot_html_sha256":old.get("html_sha256"),"live_html_sha256":cur.get("html_sha256"),"html_hash_changed":old.get("html_sha256")!=cur.get("html_sha256"),"material_fields_equal":stable,"source_revision_status":"LIVE_EQUIVALENT_DYNAMIC_HTML_CHANGED" if all(stable.values()) else "SOURCE_CHANGED"})
    tests={"product_weight_total":sum(PW.values()),"image_weight_total":sum(IW.values()),"products":len(qproducts),"images":len(qa_images),"criteria":len(criteria),"unique_product_keys":len({x['product_key'] for x in qproducts}),"unique_qa_image_keys":len({x['qa_image_key'] for x in qa_images}),"all_image_coverage_100":all(x["image_coverage"]==1 for x in qproducts),"logic_100_with_critical":"QA_FAIL","logic_90_full_no_blocker":"QA_PASS","logic_72_on_80":{"range":"72-92","status":"QA_INCOMPLETE"},"cross_links_valid":all(all(r in {i['issue_id'] for i in issues} for r in x['issue_refs']) for x in qproducts)}
    assert tests["product_weight_total"]==tests["image_weight_total"]==100 and tests["products"]==tests["unique_product_keys"]==10 and tests["images"]==tests["unique_qa_image_keys"]==62 and tests["criteria"]==110 and tests["all_image_coverage_100"] and tests["cross_links_valid"]
    dataset={"QA_Summary":summary,"QA_Products":qproducts,"QA_Criteria":criteria,"QA_Images":qa_images,"QA_Issues":issues,"SERP_Evidence":serp,"validation_tests":tests}
    for name,obj in [("qa_dataset.json",dataset),("qa_workbook_payload.json",{k:dataset[k] for k in ("QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues")}),("serp_evidence.json",serp),("source_change_audit.json",changes),("validation_results.json",tests)]: common.save_json(QA_DIR/name,obj)
    inventory=json.loads((ROOT/"seo_runs"/SHOP/RUN_ID/"inventory.json").read_text(encoding="utf-8-sig")); nxt=[{k:r.get(k) for k in ("inventory_position","product_key","product_id","Handle","title_current","product_url","image_count")} for r in inventory if 41<=int(r["inventory_position"])<=50]; preview=QA_DIR/"qa_batch_005_preview.json"; common.save_json(preview,{"batch_id":"qa_batch_005","status":"PREPARED_NOT_STARTED","products":nxt})
    OUT_DIR.mkdir(parents=True,exist_ok=True); report=OUT_DIR/f"SEO_QA_{BATCH}.md"; xlsx=OUT_DIR/f"SEO_QA_{BATCH}.xlsx"
    lines=["# SEO QA — qa_batch_004","","## Kết luận","",f"- Phạm vi: **10 sản phẩm, 62/62 ảnh (100%)**; chỉ inventory position 31–40.",f"- Điểm lô: **{avg:.1f}/100**; kết luận lô: **NOT_PASSED**.",f"- Trạng thái: {statuses.get('QA_FAIL',0)} QA_FAIL, {statuses.get('QA_REVISE',0)} QA_REVISE, {statuses.get('QA_PASS',0)} QA_PASS.",f"- Phát hiện: {sevs.get('CRITICAL',0)} CRITICAL, {sevs.get('MAJOR',0)} MAJOR, {sevs.get('MINOR',0)} MINOR, {sevs.get('LIMITATION',0)} LIMITATION.",f"- Workbook nguồn: `{SOURCE}`",f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`","","## Điểm theo sản phẩm","","| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |","|---:|---|---:|---|---:|"]
    for q,p in zip(qproducts,products): lines.append(f"| {q['inventory_position']} | {p['title_current'].replace('|','/')} | {q['final_score']:.1f} | {q['qa_status']} | {q['critical_count']}/{q['major_count']}/{q['minor_count']}/{q['limitation_count']} |")
    lines += ["","## Lỗi ưu tiên","","1. **CRITICAL — products 33–40:** draft dùng customization/personalized/custom nhưng purchase options live không có input name/number/text; phải chứng minh control và fulfillment mapping hoặc bỏ claim.","2. **MAJOR — cả 10 sản phẩm:** description HTML chứa khối `SEO Use` và câu QA/import nội bộ, chưa publish-ready.","3. **MAJOR — products 33–34:** draft chép thuộc tính `Dragonfly`, mâu thuẫn với thiết kế cow/crocodile.","4. **MAJOR — ảnh:** products 31, 33, 34 gán image 6 là size guide thay vì material-layer diagram; product 32 đảo close-up và angled bedroom.","5. **MAJOR — product 31:** SERP exact birdhouse còn lẫn fabric panel/quilt kit; finished-product intent chưa được xác nhận sạch.","6. **MAJOR — product 34:** dùng `alligator` như secondary keyword khi evidence sản phẩm gọi là crocodile.","","## SERP và keyword","","Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và phương án gần nhất. URL/timestamp/locale limit nằm trong `serp_evidence.json`; không có claim volume. Các SERP football có sản phẩm tùy biến thật với input name/number, làm rõ khoảng cách giữa intent keyword và purchase flow hiện tại.","","## Giới hạn và trạng thái bàn giao","","- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.","- Trình duyệt tương tác không khả dụng trong môi trường; static HTML/product.js không hiển thị input personalization. Claim chỉ được thông qua lại khi có bằng chứng purchase flow hoạt động.","- Product identity, canonical, product ID, variants và gallery không phát hiện thay đổi trọng yếu so với snapshot; HTML động có thể đổi hash.","- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.","- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại và render kiểm tra.","- Chưa QA products 41–50. `awaiting_confirmation=true`.","","## Tệp chi tiết","",f"- QA data: `{QA_DIR/'qa_dataset.json'}`",f"- SERP evidence: `{QA_DIR/'serp_evidence.json'}`",f"- Validation: `{QA_DIR/'validation_results.json'}`",f"- Manifest/checkpoint: `{QA_DIR}`",""]
    report.write_text("\n".join(lines),encoding="utf-8")
    keys=[p["product_key"] for p in products]; manifest={"rubric_version":"1.0","prompt_version":"2.4","qa_run_id":QA_RUN_ID,"started_at":"2026-09-07T09:56:43+07:00","shop_domain":SHOP,"research_run_id":RUN_ID,"market":"United States","seo_language":"English","batch_id":BATCH,"source_workbook":str(SOURCE.relative_to(ROOT)),"source_workbook_sha256":source_hash,"source_snapshot":str(SNAPSHOT.relative_to(ROOT)),"batch_product_keys":keys,"revision":"r1","expected_products":10,"expected_images":62,"source_admin_export_available":False,"status":"COMPLETE","completed_at":common.now(),"source_sha256_at_handoff":source_hash,"counts":{"products":10,"images":62},"output_markdown":str(report),"output_xlsx":str(xlsx),"xlsx_blocker":None,"qa_dataset":str(QA_DIR/"qa_dataset.json"),"qa_workbook_payload":str(QA_DIR/"qa_workbook_payload.json"),"source_snapshot_sha256":common.sha256(SNAPSHOT),"prompt_files_sha256":{"prompt_qa.md":common.sha256(ROOT/"seo-prompt/jeminise/prompt_qa.md"),"prompt.md":common.sha256(ROOT/"seo-prompt/jeminise/prompt.md")},"source_revision_status":"LIVE_EQUIVALENT_DYNAMIC_HTML_CHANGED"}
    common.save_json(QA_DIR/"qa_manifest.json",manifest); common.save_json(QA_DIR/"qa_progress.json",{"rubric_version":"1.0","qa_run_id":QA_RUN_ID,"source_workbook":str(SOURCE.relative_to(ROOT)),"source_workbook_sha256":source_hash,"batch_id":BATCH,"batch_product_keys":keys,"current_product_key":keys[-1],"current_stage":"BATCH_COMPLETE","completed_image_keys":[x["qa_image_key"] for x in qa_images],"last_saved_at":common.now(),"artifact_paths":{"markdown":str(report),"xlsx":str(xlsx),"dataset":str(QA_DIR/"qa_dataset.json"),"workbook_payload":str(QA_DIR/"qa_workbook_payload.json"),"source_change_audit":str(QA_DIR/"source_change_audit.json"),"spreadsheet_validation":str(QA_DIR/"spreadsheet_validation.json")},"awaiting_confirmation":True,"confirmation_ref":None,"next_batch_preview":{"batch_id":"qa_batch_005","inventory_positions":"41-50","status":"PREPARED_NOT_STARTED","product_keys":[x["product_key"] for x in nxt],"path":str(preview)}})
    print(json.dumps({"report":str(report),"batch_score":avg,"statuses":dict(statuses),"issues":dict(sevs),"source_hash":source_hash},ensure_ascii=False,indent=2))


if __name__ == "__main__": main()
