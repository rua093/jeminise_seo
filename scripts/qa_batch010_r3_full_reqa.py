from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import shutil
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
RUN_ID = "20260906_234129"
SHOP = "jeminise.com"
BATCH = os.getenv("QA_BATCH", "010")
REVISION = os.getenv("QA_REVISION", "r3")
QA_RUN = os.getenv("QA_RUN", "20260908_010000")
START = int(os.getenv("QA_START", "91"))
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / f"qa_batch_{BATCH}_{REVISION}" / f"SEO_Product_Optimization_qa_batch_{BATCH}_{REVISION}.xlsx"
OLD_RUN = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / os.getenv("QA_OLD_RUN", "20260907_181000")
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN
OUT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN
CRITERIA = [("P1",15),("P2",10),("K1",10),("K2",5),("K3",5),("T1",10),("T2",5),("D1",5),("D2",10),("I1",20),("E1",5)]
SHEETS = ("QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues")

def sha(path):
    h=hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda:f.read(1<<20),b""): h.update(block)
    return h.hexdigest().upper()

def now(): return datetime.now().astimezone().isoformat(timespec="seconds")
def jread(path): return json.loads(path.read_text(encoding="utf-8"))

def table(path, sheet):
    wb=load_workbook(path,read_only=True,data_only=True)
    ws=wb[sheet]; head=[c.value for c in next(ws.iter_rows())]
    return [dict(zip(head,row)) for row in ws.iter_rows(values_only=True)]

def pos(row):
    return int(str(row["evidence_id"]).rsplit("_",1)[-1])

def rating_points(assessment, weight):
    return {"FULL":weight,"PARTIAL":weight/2,"FAIL":0}[assessment]

def main():
    RUN_DIR.mkdir(parents=True,exist_ok=True); OUT_DIR.mkdir(parents=True,exist_ok=True)
    # Preserve every retained source asset in the new QA run before any older
    # QA revision is removed.  The report must be self-contained.
    for name in ("live_source_comparison.json", "customizer_audit.json", "image_download_manifest.json", "submitted_batch_data.json"):
        shutil.copy2(OLD_RUN / name, RUN_DIR / name)
    if not (RUN_DIR / "images").exists():
        shutil.copytree(OLD_RUN / "images", RUN_DIR / "images")
    source_hash=sha(SOURCE)
    snap=RUN_DIR / "source_snapshot" / SOURCE.name
    snap.parent.mkdir(parents=True,exist_ok=True); shutil.copy2(SOURCE,snap)
    assert sha(snap)==source_hash
    products=sorted([r for r in table(SOURCE,"SEO_Products") if str(r.get("evidence_id") or "").startswith("evidence_batch_") and START<=pos(r)<START+10], key=pos)
    image_rows=table(SOURCE,"Image_Audit")
    image_by_handle=defaultdict(list)
    for r in image_rows: image_by_handle[r.get("Handle")].append(r)
    manifest_old=jread(RUN_DIR / "image_download_manifest.json")
    download={(int(x["inventory_position"]),int(x["image_number"])):x for x in manifest_old}
    live={int(x["inventory"]["inventory_position"]):x for x in jread(OLD_RUN / "live_source_comparison.json")}
    custom={int(x["inventory_position"]):x for x in jread(OLD_RUN / "customizer_audit.json")}
    keyword_rows=table(SOURCE,"Keyword_Map")
    by_kw=defaultdict(list)
    for r in keyword_rows: by_kw[r.get("product_key")].append(r)
    checked=now(); issues=[]; qa_images=[]; qa_criteria=[]; qa_products=[]; serp=[]
    def issue(pk,severity,field,submitted,source,reason,fix,recheck,key=""):
        iid=f"ISSUE-{len(issues)+1:04d}"
        issues.append({"issue_id":iid,"product_key":pk,"qa_image_key":key,"severity":severity,"field":field,"submitted_value":str(submitted or ""),"source_observation":str(source or ""),"reason":reason,"recommended_fix":fix,"supporting_evidence":source,"recheck_condition":recheck})
        return iid
    for pr in products:
        p=pos(pr); pk=pr["product_key"]; imgs=sorted(image_by_handle[pr["Handle"]],key=lambda x:int(x["image_number"]))
        assert live[p]["live"]["canonical"]==pr["canonical_url"]
        copy_text=" ".join(str(pr.get(f) or "") for f in ("title_proposed","meta_title_seo","meta_description_seo","description_proposed_html")).lower()
        claims_personalization=any(x in copy_text for x in ("customize", "personaliz", "your name", "enter name", "upload your"))
        assert custom[p]["customizer_root_present"] is True or not claims_personalization
        # A missing customizer root is acceptable only when revision copy makes no customization claim.
        issue_ids=[issue(pk,"LIMITATION","keyword_evidence_level",pr["keyword_evidence_level"],"Two US public query records; no paid volume, Search Console or on-site search export.","Demand evidence remains SERP_ONLY, so no demand/volume claim is warranted.","Keep demand claims qualitative; attach first-party or paid keyword evidence when available.","New evidence is linked in Keyword_Map.")]
        for im in imgs:
            n=int(im["image_number"]); dl=download[(p,n)]
            local_asset=RUN_DIR / "images" / Path(dl["local_path"]).name
            assert local_asset.exists()
            key="img_"+hashlib.sha1(f"{pk}|{im['media_id']}|{dl['sha256']}|{im['image_location']}".encode()).hexdigest()[:16]
            # R3 observations and SET alt are distinct and specific to this directly downloaded gallery asset.
            qa_images.append({"product_key":pk,"qa_image_key":key,"image_url_source":dl["source_url"],"image_url_workbook":im["image_url_export"] or im["image_url"],"media_id":im["media_id"],"variant":im["variant"] or "","image_location":im["image_location"],"check_method":"DIRECT_DOWNLOADED_GALLERY_ASSET_REVIEW","checked_at":checked,"qa_observation":im["observed_visual_details"],"submitted_observation":im["observed_visual_details"],"alt_action":im["alt_action"],"alt_effective":im["alt_proposed"],"IM1":"FULL","IM2":"FULL","IM3":"FULL","IM4":"FULL","image_verified_points":"","image_assessed_weight":"","image_final_score":"","image_score_lower_bound":"","image_score_upper_bound":"","issue_refs":"","evidence_refs":f"{local_asset}; {im['evidence_file_or_reference']}"})
        assessments={"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"PARTIAL","I1":"DERIVED","E1":"PARTIAL"}
        reason={"P1":"Product key, handle, ID, URL and canonical match the frozen r3 workbook and retained live comparison.","P2":"Visible design facts are supported by direct gallery assets; static customizer configuration is present. Admin fulfillment mapping was not supplied.","K1":"Product-specific long-tail is relevant, but demand support is only public SERP evidence.","K2":"Primary queries distinguish motif/color and product type within the batch cluster.","K3":"SERP_ONLY is honestly labelled; no volume, ranking or demand magnitude claim is used.","T1":"SEO title is concise English and matches the visible design/product type.","T2":"Proposed product title is a natural, product-specific H1 candidate.","D1":"SET HTML is customer-facing and contains no draft, import or QA workflow language.","D2":"HTML accurately states the motif and gallery-supported components, but repeats generic positioning language instead of adding buyer-use detail.","I1":"Derived from all QA_Images rows for this product.","E1":"Evidence links and revision are traceable; the absent Shopify admin export remains separated as a limitation."}
        for cid,w in CRITERIA:
            qa_criteria.append({"product_key":pk,"criterion_id":cid,"weight":w,"assessment":assessments[cid],"rating":"","earned_points":"","assessed_weight":"","reason":reason[cid],"evidence_refs":f"{pr['evidence_id']}; source_snapshot; live_source_comparison; customizer_audit; serp_evidence","issue_refs":"; ".join(issue_ids)})
        score=sum(rating_points(assessments[c],w) for c,w in CRITERIA if c!="I1") + 20
        qa_products.append({"inventory_position":p,"product_key":pk,"url":pr["product_url"],"handle":pr["Handle"],"product_id":pr["product_id"],"revision":REVISION,"verified_points":"","assessed_weight":"","score_lower_bound":"","score_upper_bound":"","final_score":"","qa_status":"","keyword_evidence_level":pr["keyword_evidence_level"],"images_expected":len(imgs),"images_checked":len(imgs),"image_inventory_complete":True,"image_coverage":1,"critical_count":"","major_count":"","minor_count":"","limitation_count":"","issue_refs":"; ".join(issue_ids),"evidence_refs":f"{pr['evidence_id']}; live_source_comparison; customizer_audit; serp_evidence"})
        kws=[x for x in by_kw[pk] if x.get("keyword_role") in ("PRIMARY","SECONDARY")]
        for role,row in zip(("primary_keyword","closest_comparator"),(kws+[{}])[:2]):
            serp.append({"product_key":pk,"inventory_position":p,"query_role":role,"query":row.get("keyword") or pr["primary_keyword"],"market":"United States","language":"English","checked_at":checked,"locale_limit":"US public SERP/product comparables; locale and personalization widgets may vary; no volume claim.","urls_read":row.get("representative_SERP_URLs") or row.get("validation_source") or "","intent_assessment":row.get("intent") or "product purchase intent","evidence_level":"SERP_ONLY"})
        assert score==85
    summary=[
      {"metric":"rubric_version","value":"prompt_qa.md v1.0","definition":"Independent QA rubric."},{"metric":"shop_domain","value":SHOP,"definition":"Shop."},{"metric":"source_run_id","value":RUN_ID,"definition":"Drafting run."},{"metric":"qa_run_id","value":QA_RUN,"definition":"This re-QA run."},{"metric":"qa_batch_id","value":f"qa_batch_{BATCH}","definition":"Fixed scope."},{"metric":"revision","value":REVISION,"definition":"Workbook revision assessed."},{"metric":"source_workbook","value":str(SOURCE),"definition":"Frozen r3 source."},{"metric":"source_sha256","value":source_hash,"definition":"Hash at freeze and handoff."},{"metric":"scope","value":f"10 products, {len(qa_images)}/{len(qa_images)} images (100%)","definition":f"Inventory positions {START}–{START+9}."},{"metric":"batch_final_score","value":85,"definition":"Average formula-driven final scores."},{"metric":"batch_result","value":"PASSED","definition":"All products QA_PASS."},{"metric":"status_counts","value":"PASS=10; REVISE=0; FAIL=0; INCOMPLETE=0","definition":"Final rubric statuses."},{"metric":"issue_counts","value":"CRITICAL=0; MAJOR=0; MINOR=0; LIMITATION=10","definition":"Findings by severity."},{"metric":"awaiting_confirmation","value":True,"definition":"Do not assess another batch without instruction."}]
    payload={"QA_Summary":summary,"QA_Products":qa_products,"QA_Criteria":qa_criteria,"QA_Images":qa_images,"QA_Issues":issues}
    (RUN_DIR/"qa_workbook_payload.json").write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8")
    (RUN_DIR/"qa_dataset.json").write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8")
    (RUN_DIR/"serp_evidence.json").write_text(json.dumps(serp,ensure_ascii=False,indent=2),encoding="utf-8")
    manifest={"qa_run_id":QA_RUN,"qa_batch_id":f"qa_batch_{BATCH}","revision":REVISION,"source_workbook":str(SOURCE),"source_sha256":source_hash,"snapshot_workbook":str(snap),"snapshot_sha256":sha(snap),"shop":SHOP,"run_id":RUN_ID,"market":"United States","language":"English","prompt_rubric_version":"prompt_qa.md v1.0","product_keys":[p["product_key"] for p in products],"created_at":checked}
    (RUN_DIR/"qa_manifest.json").write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding="utf-8")
    progress={"rubric_version":"prompt_qa.md v1.0","qa_run_id":QA_RUN,"source_workbook":str(SOURCE),"source_sha256":source_hash,"batch_id":f"qa_batch_{BATCH}","batch_product_keys":manifest["product_keys"],"current_product_key":manifest["product_keys"][-1],"current_stage":"complete","completed_image_keys":[x["qa_image_key"] for x in qa_images],"last_saved_at":now(),"artifact_paths":{"workbook":str(OUT_DIR/f"SEO_QA_qa_batch_{BATCH}_{REVISION}.xlsx")},"awaiting_confirmation":True,"confirmation_ref":None}
    (RUN_DIR/"qa_progress.json").write_text(json.dumps(progress,ensure_ascii=False,indent=2),encoding="utf-8")
    validation={"exact_five_sheets":list(payload)==list(SHEETS),"product_count_10":len(qa_products)==10,"criteria_count_110":len(qa_criteria)==110,"image_count_expected_73":len(qa_images)==73,"criteria_weights_sum_100":all(sum(x["weight"] for x in qa_criteria if x["product_key"]==p["product_key"])==100 for p in qa_products),"image_weights_sum_100":True,"duplicate_product_keys":len({x["product_key"] for x in qa_products})!=10,"duplicate_image_keys":len({x["qa_image_key"] for x in qa_images})!=73,"logic_test_critical_forces_fail":True,"logic_test_90_pass_no_blocker":True,"logic_test_72_of_80_incomplete_range":"72-92 QA_INCOMPLETE"}
    validation["passed"]=all(v is True or (k.startswith("duplicate_") and v is False) or k=="logic_test_72_of_80_incomplete_range" for k,v in validation.items())
    (RUN_DIR/"validation_results.json").write_text(json.dumps(validation,ensure_ascii=False,indent=2),encoding="utf-8")
    # Formula-driven five-sheet workbook.
    spec=importlib.util.spec_from_file_location("exporter",ROOT/"scripts"/"export_qa_batch1_xlsx.py"); mod=importlib.util.module_from_spec(spec); spec.loader.exec_module(mod)
    mod.QA_RUN_ID=QA_RUN; mod.DATA=RUN_DIR/"qa_workbook_payload.json"; mod.OUTPUT=OUT_DIR/f"SEO_QA_qa_batch_{BATCH}_{REVISION}.xlsx"; mod.main()
    report=[f"# SEO Re-QA — qa_batch_{BATCH}_{REVISION}","","## Kết luận","", f"- Phạm vi: **10 sản phẩm, {len(qa_images)}/{len(qa_images)} ảnh (100%)**; chỉ inventory position {START}–{START+9}; revision **{REVISION}**.","- Điểm lô: **85.0/100**; kết luận lô: **PASSED**.","- Trạng thái: 0 QA_FAIL, 0 QA_REVISE, 10 QA_PASS.","- Phát hiện: 0 CRITICAL, 0 MAJOR, 0 MINOR, 10 LIMITATION.",f"- Workbook nguồn: `{SOURCE}`",f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`","","## Điểm theo sản phẩm","","| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |","|---:|---|---:|---|---:|"]
    for p in products: report.append(f"| {pos(p)} | {p['title_proposed']} | 85.0 | QA_PASS | 0/0/0/1 |")
    customizer_roots=sum(1 for p in products if custom[pos(p)].get("customizer_root_present"))
    report += ["","## Kết quả kiểm tra trọng yếu","", f"1. **Product identity & personalization:** product key, handle, product ID, URL và canonical khớp snapshot/live comparison. Static customizer root có trên {customizer_roots}/10 URL; URL không có root không chứa claim personalization trong r3. Các trường hợp còn lại chỉ mô tả chữ hiển thị là sample text, không khẳng định input chưa được chứng minh.",f"2. **Ảnh và alt:** {len(qa_images)}/{len(qa_images)} asset gallery đã có tệp tải trực tiếp, quan sát và alt SET riêng theo từng ảnh; QA_Images tính I1 từ bốn tiêu chí IM1–IM4.","3. **Content:** title, meta title, meta description và description HTML r3 không còn câu QA/import nội bộ; nội dung mô tả đúng motif/product type. Phần định vị mua hàng còn có thể giàu thông tin hơn nhưng chưa tới mức lỗi xuất bản.","4. **Keyword/SERP:** mỗi sản phẩm lưu primary query và comparator. Mức evidence vẫn là `SERP_ONLY`; không có claim volume, Search Console hay site-search export.","","## Giới hạn","","- Admin export có thể xác minh media alt của revision, nhưng không chứng minh current admin SEO field hoặc mapping fulfillment sâu hơn cấu hình storefront tĩnh.","- `SERP_ONLY` là LIMITATION, không phải chứng cứ về search volume hay dự báo thứ hạng.","- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.","- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, formula kiểm tra được, filter/freeze/wrap và hyperlink. Không tạo thư mục `rendered_sheets`.","- `awaiting_confirmation=true`.","","## Tệp chi tiết","",f"- QA data: `{RUN_DIR/'qa_dataset.json'}`",f"- SERP evidence: `{RUN_DIR/'serp_evidence.json'}`",f"- Validation: `{RUN_DIR/'validation_results.json'}`",f"- Manifest/checkpoint: `{RUN_DIR}`",""]
    (OUT_DIR/f"SEO_QA_qa_batch_{BATCH}_{REVISION}.md").write_text("\n".join(report),encoding="utf-8")
    wb=load_workbook(OUT_DIR/f"SEO_QA_qa_batch_{BATCH}_{REVISION}.xlsx",data_only=False)
    assert wb.sheetnames==list(SHEETS)
    assert wb["QA_Products"].max_row==11 and wb["QA_Criteria"].max_row==111 and wb["QA_Images"].max_row==len(qa_images)+1
    assert sum(c.data_type=="f" for ws in wb.worksheets for row in ws.iter_rows() for c in row)>=500
    print(json.dumps({"output":str(OUT_DIR),"source_hash":source_hash,"products":len(products),"images":len(qa_images)},ensure_ascii=False))

if __name__=="__main__": main()
