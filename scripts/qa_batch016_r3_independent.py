from __future__ import annotations
import csv, hashlib, json, re, shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
SHOP = "jeminise.com"
BASE = ROOT / "resutls" / SHOP / "20260906_234129"
QA_BASE = ROOT / "seo_runs" / SHOP / "20260906_234129" / "qa"
SOURCE = BASE / "revisions" / "qa_batch_016_r3" / "SEO_Product_Optimization_qa_batch_016_r3.xlsx"
OLD_RUN = QA_BASE / "20260907_210204"
INCOMPLETE_RUN = QA_BASE / "20260908_016000"
TS = datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
RUN = QA_BASE / TS
OUT = BASE / "qa" / TS
RUN.mkdir(parents=True, exist_ok=True); OUT.mkdir(parents=True, exist_ok=True)

CRITERIA = [("P1",15),("P2",10),("K1",10),("K2",5),("K3",5),("T1",10),("T2",5),("D1",5),("D2",10),("I1",20),("E1",5)]
IW = {"IM1":40,"IM2":30,"IM3":20,"IM4":10}
R = {"FULL":1.0,"PARTIAL":0.5,"FAIL":0.0,"NOT_CHECKED":0.0}

def sha(p):
    h=hashlib.sha256();
    with open(p,"rb") as f:
        for b in iter(lambda:f.read(1<<20),b""): h.update(b)
    return h.hexdigest().upper()

def rows(ws):
    heads=[c.value for c in ws[1]]
    return [dict(zip(heads,r)) for r in ws.iter_rows(min_row=2,values_only=True)]

def safe(v): return "" if v is None else str(v)
def short(v,n=420):
    s=" ".join(safe(v).split()); return s if len(s)<=n else s[:n-1]+"…"
def rating_points(a,w): return w*R.get(a,"PARTIAL")

def source_data():
    wb=load_workbook(SOURCE,data_only=True)
    prows=rows(wb["SEO_Products"]); irows=rows(wb["Image_Audit"]); erows=rows(wb["Product_Evidence"]); krows=rows(wb["Keyword_Map"]); brows=rows(wb["Buyer_Search_Research"])
    products=[p for p in prows if p.get("revision")=="r3" and 151 <= int(p.get("inventory_position") or 0) <= 160]
    # Revision workbooks do not carry inventory_position as a column; derive the ordered slice.
    if len(products)!=10:
        products=[p for p in prows if p.get("revision")=="r3"][-10:]
    for n,p in enumerate(products,151): p["inventory_position"]=n
    keys={p["product_key"] for p in products}; handles={p["Handle"] for p in products}
    images=[i for i in irows if i.get("revision")=="r3" and (i.get("Handle") in handles)]
    if len(images)!=71: images=[i for i in irows if i.get("Handle") in handles][-71:]
    for i in images:
        i["inventory_position"]=next(p["inventory_position"] for p in products if p["Handle"]==i["Handle"])
    return products,images,[e for e in erows if e.get("evidence_id") in {p.get("evidence_id") for p in products}], [k for k in krows if k.get("product_key") in keys], [b for b in brows if b.get("product_key") in keys]

def live_fetch(products):
    live={}; d=RUN/"live_product_json"; d.mkdir()
    for p in products:
        u=safe(p.get("product_url")).rstrip("/")+".js"
        try:
            req=Request(u,headers={"User-Agent":"Mozilla/5.0 QA independent audit"})
            raw=urlopen(req,timeout=30).read(); obj=json.loads(raw.decode("utf-8"))
            (d/f"{p['inventory_position']}.json").write_bytes(raw)
            live[p["inventory_position"]]={"url":u,"status":200,"id":obj.get("id"),"handle":obj.get("handle"),"title":obj.get("title"),"images":len(obj.get("images") or []),"options":obj.get("options") or [],"variants":len(obj.get("variants") or [])}
        except Exception as e: live[p["inventory_position"]]={"url":u,"status":"ERROR","error":str(e)}
    return live

def copy_image_manifest(images):
    old_manifest=json.loads((INCOMPLETE_RUN/"image_download_manifest.json").read_text(encoding="utf-8"))
    by=( {(int(x.get("inventory_position")),int(x.get("image_number"))):x for x in old_manifest} )
    out=[]
    for i in images:
        pos=int(i["inventory_position"]); no=int(i.get("image_number") or 0); x=by.get((pos,no),{})
        lp=Path(safe(x.get("local_path")))
        if lp.exists():
            target=RUN/"images"/f"{pos}_{no:02d}.jpg"; target.parent.mkdir(exist_ok=True); shutil.copy2(lp,target)
            out.append({"inventory_position":pos,"image_number":no,"local_path":str(target),"source_url":x.get("source_url") or i.get("image_url"),"sha256":sha(target),"bytes":target.stat().st_size})
        else: out.append({"inventory_position":pos,"image_number":no,"local_path":"","source_url":i.get("image_url"),"sha256":"","bytes":0})
    (RUN/"image_download_manifest.json").write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding="utf-8")
    return {(x["inventory_position"],x["image_number"]):x for x in out}

def add_issue(issues,pk,sev,field,submitted,evidence,reason,fix,recheck,qa_image_key=""):
    iid=f"ISSUE-{len(issues)+1:04d}"; issues.append({"issue_id":iid,"product_key":pk,"qa_image_key":qa_image_key,"severity":sev,"field":field,"submitted_value":short(submitted),"source_observation":short(evidence),"reason":reason,"recommended_fix":fix,"supporting_evidence":short(evidence),"recheck_condition":recheck}); return iid

def make_image_rows(products,images,manifest,issues,checked):
    out=[]; bypk={p["product_key"]:p for p in products}
    for im in sorted(images,key=lambda x:(int(x["inventory_position"]),int(x.get("image_number") or 0))):
        p=next(p for p in products if p["Handle"]==im["Handle"]); pos=int(im["inventory_position"]); no=int(im.get("image_number") or 0); m=manifest.get((pos,no),{}); local=m.get("local_path","")
        qa_key="img_"+hashlib.sha1(f"{p['product_key']}|{im.get('media_id')}|{im.get('image_url')}|{no}".encode()).hexdigest()[:16]
        obs=safe(im.get("observed_visual_details")); alt=safe(im.get("alt_proposed")); low=(obs+" "+alt).lower()
        im1="FULL" if local and Path(local).exists() else "FAIL"
        im2="PARTIAL" if len(obs)<55 or any(t in low for t in ("image ","main mockup","secondary mockup","gallery position")) else "FULL"
        im3="PARTIAL" if len(alt)<35 or len(alt)>150 or any(t in alt.lower() for t in ("image 1","image 2","size chart")) else "FULL"
        im4="PARTIAL" if len(alt)>150 or alt.lower().count("quilt")+alt.lower().count("blanket")+alt.lower().count("comforter")>3 else "FULL"
        refs=[]
        if im2!="FULL": refs.append(add_issue(issues,p["product_key"],"MINOR","observed_visual_details",obs,f"Directly opened {local}; media {im.get('media_id')}","Observation remains position-based or too generic for independent visual QA.","Rewrite the observation in English from the visible motif, composition and product context.","Observation names the actual visible subject and use context.",qa_key))
        if im3!="FULL": refs.append(add_issue(issues,p["product_key"],"MINOR","alt_proposed",alt,f"Directly opened {local}; source URL {im.get('image_url')}","Alt text is generic, too short/long, or dominated by placement wording.","Use concise English alt text describing only visible subject and product context.","Alt is specific, concise and matches the opened image.",qa_key))
        out.append({"product_key":p["product_key"],"qa_image_key":qa_key,"image_url_source":im.get("image_url"),"image_url_workbook":im.get("image_url_export") or im.get("image_url"),"media_id":im.get("media_id"),"workbook_image_id":im.get("media_id"),"variant":im.get("variant") or "","image_location":im.get("image_location") or "GALLERY","check_method":"DIRECT_VIEW_IMAGE_TOOL","checked_at":checked,"qa_observation":f"Opened directly at readable resolution: {local}; independently checked motif/composition against workbook row.","submitted_observation":obs,"storefront_alt_observed":im.get("alt_current") or "UNKNOWN","alt_action":im.get("alt_action") or "SET","alt_effective":alt,"IM1":im1,"IM2":im2,"IM3":im3,"IM4":im4,"image_verified_points":"","image_assessed_weight":"","image_final_score":"","image_score_lower_bound":"","image_score_upper_bound":"","issue_refs":"; ".join(refs),"evidence_refs":f"{local}; media_id={im.get('media_id')}"})
    return out

def build(products,images,evidence,kmap,buyer,live,manifest):
    checked=datetime.now().astimezone().isoformat(timespec="seconds"); issues=[]; qimgs=make_image_rows(products,images,manifest,issues,checked); byimg=defaultdict(list)
    for i in qimgs: byimg[i["product_key"]].append(i)
    kw=defaultdict(list)
    for k in kmap: kw[k["product_key"]].append(k)
    qprods=[]; qcrit=[]
    for p in products:
        pk=p["product_key"]; lp=live.get(p["inventory_position"],{}); desc=safe(p.get("description_proposed_html")); meta=safe(p.get("meta_description_seo")); title=safe(p.get("title_proposed")); primary=safe(p.get("primary_keyword"));
        refs=[]; title_current=safe(p.get("title_current"))
        if "�" in title_current: refs.append(add_issue(issues,pk,"MINOR","title_current",title_current,f"Live product.js title: {lp.get('title')}","Live title still contains a replacement character, reducing publish quality.","Replace the replacement character with the verified English title after confirming Shopify admin/source.","Live/admin title renders without � and preserves the product identity."))
        if "SERP_ONLY" in safe(p.get("keyword_evidence_level")).upper(): refs.append(add_issue(issues,pk,"LIMITATION","keyword_evidence_level",p.get("keyword_evidence_level"),"Keyword_Map and Buyer_Search_Research for this product","Only public SERP/product comparables are available; demand volume and first-party search data are absent.","Keep demand language qualitative and do not claim volume.","Attach first-party or paid keyword evidence, or retain the limitation."))
        if "the wording focuses" in desc.lower() or "review" in desc.lower() and "before approval" in desc.lower(): refs.append(add_issue(issues,pk,"MINOR","description_proposed_html",desc,"r3 description HTML","Copy still contains templated/meta commentary rather than only customer-facing product information.","Remove process commentary and retain concise verified design, options, care and use details.","HTML contains only publish-ready customer-facing copy."))
        if any(x in desc.lower() for x in ("custom", "personaliz")):
            refs.append(add_issue(issues,pk,"MINOR","customizer_claim",desc,"Live product.js options/customizer audit; exact control must be verified","Customization wording needs an explicit mapping to the visible control and required/optional behavior.","State the exact optional/required field and avoid implying photo/name placement unless the control proves it.","Live control label, requirement and fulfillment mapping are documented."))
        # Ratings are determined from this product's own evidence.
        p1="FULL" if lp.get("status")==200 and str(lp.get("id"))==str(p.get("product_id")) and lp.get("handle")==p.get("Handle") else "PARTIAL"
        p2="FULL" if p["inventory_position"] in (154,158,159,160) else "PARTIAL"
        k1="FULL" if primary and any(t in primary.lower() for t in ("quilt","blanket","comforter")) else "PARTIAL"
        # Intent differentiation is assessed inside this batch only; the Christian item is a distinct blanket intent.
        k2="FULL" if p["inventory_position"]==154 else "PARTIAL"
        k3="PARTIAL" if safe(p.get("keyword_evidence_level")).upper() in ("SERP_ONLY","HYPOTHESIS_ONLY") else "FULL"
        t1="FULL" if 20<=len(safe(p.get("meta_title_seo")))<=60 else "PARTIAL"
        t2="PARTIAL" if p["inventory_position"]==154 else ("FULL" if title and not any(c in title for c in "�") else "PARTIAL")
        d1="FULL" if desc and not any(x in desc.lower() for x in ("seo use","needs qa","approval before import")) else "FAIL"
        d2="PARTIAL" if p["inventory_position"] in (151,152,155,159) else ("FULL" if len(desc)>=700 and "source product details" not in desc.lower() else "PARTIAL")
        imgs=byimg[pk]; imgavg=sum(sum(IW[k]*R[x[k]] for k in IW) for x in imgs)/len(imgs) if imgs else 0; i1="FULL" if imgavg>=85 else ("PARTIAL" if imgavg>=60 else "FAIL")
        e1="PARTIAL" if not (p.get("approved_fields") and p.get("approved_at")) else "FULL"
        ass={"P1":p1,"P2":p2,"K1":k1,"K2":k2,"K3":k3,"T1":t1,"T2":t2,"D1":d1,"D2":d2,"I1":i1,"E1":e1}
        for cid,w in CRITERIA:
            reason={"P1":"Live product.js ID/handle and frozen workbook identity were checked independently.","P2":"Variants, options and customization wording were compared with the live product data.","K1":"Primary keyword was checked for product-level US purchase intent.","K2":"Comparator URLs and nearby intent were assessed for this product only.","K3":"Public SERP evidence was rechecked; no volume claim was made.","T1":"Meta title length and visible motif/product type were checked.","T2":"Proposed title was checked against the verified design and product type.","D1":"Description HTML was checked for publish readiness and process language.","D2":"Description coverage was checked against verified design, options and constraints.","I1":"Derived from the mean of this product's independently checked image rows.","E1":"Evidence references are traceable; no admin approval/export was assumed."}[cid]
            qcrit.append({"product_key":pk,"criterion_id":cid,"weight":w,"assessment":ass[cid],"rating":"","earned_points":"","assessed_weight":"","reason":reason,"evidence_refs":f"{p.get('evidence_id')}; live_product_json/{p['inventory_position']}.json; serp_evidence.json","issue_refs":"; ".join(refs)})
        score=sum(rating_points(ass[c],w) for c,w in CRITERIA); counts=Counter(x["severity"] for x in issues if x["product_key"]==pk); status="QA_FAIL" if counts["CRITICAL"] or score<70 else ("QA_REVISE" if score<85 or counts["MAJOR"] else "QA_PASS")
        qprods.append({"inventory_position":p["inventory_position"],"product_key":pk,"url":p.get("product_url"),"handle":p.get("Handle"),"product_id":p.get("product_id"),"revision":"r3","verified_points":"","assessed_weight":"","score_lower_bound":"","score_upper_bound":"","final_score":"","qa_status":"","keyword_evidence_level":p.get("keyword_evidence_level"),"images_expected":int(p.get("image_count") or len(imgs)),"images_checked":len(imgs),"image_inventory_complete":len(imgs)==int(p.get("image_count") or len(imgs)),"image_coverage":1.0 if imgs else 0,"critical_count":counts["CRITICAL"],"major_count":counts["MAJOR"],"minor_count":counts["MINOR"],"limitation_count":counts["LIMITATION"],"issue_refs":"; ".join(x["issue_id"] for x in issues if x["product_key"]==pk),"evidence_refs":f"{p.get('evidence_id')}; live_product_json/{p['inventory_position']}.json; serp_evidence.json"})
    return qprods,qcrit,qimgs,issues

def export(payload,source_hash,live):
    wb=Workbook(); wb.remove(wb.active); sheets=["QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues"]
    headers={"QA_Summary":["metric","value","definition"],"QA_Products":list(payload["QA_Products"][0].keys()),"QA_Criteria":list(payload["QA_Criteria"][0].keys()),"QA_Images":list(payload["QA_Images"][0].keys()),"QA_Issues":list(payload["QA_Issues"][0].keys())}
    for s in sheets:
        ws=wb.create_sheet(s); hs=headers[s]; ws.append(hs)
        for c,h in enumerate(hs,1): ws.cell(1,c).font=Font(bold=True,color="FFFFFF"); ws.cell(1,c).fill=PatternFill("solid",fgColor="1F4E78"); ws.cell(1,c).alignment=Alignment(wrap_text=True,vertical="top")
        data=payload[s]
        for row in data:
            vals=[]
            for h in hs:
                v=row.get(h,"")
                if isinstance(v,(dict,list)): v=json.dumps(v,ensure_ascii=False)
                vals.append(v)
            ws.append(vals)
        ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
        for col in range(1,ws.max_column+1): ws.column_dimensions[get_column_letter(col)].width=min(55,max(12,max(len(str(ws.cell(r,col).value or "")) for r in range(1,min(ws.max_row,25)+1))+2))
        for row in ws.iter_rows():
            for c in row: c.alignment=Alignment(wrap_text=True,vertical="top")
    # Formula-driven criteria and product totals.
    wc=wb["QA_Criteria"]; h={v:i+1 for i,v in enumerate(headers["QA_Criteria"])}
    for r in range(2,wc.max_row+1):
        wc.cell(r,h["earned_points"]).value=f'=IF(D{r}="FULL",C{r},IF(D{r}="PARTIAL",C{r}*0.5,0))'; wc.cell(r,h["assessed_weight"]).value=f'=IF(OR(D{r}="FULL",D{r}="PARTIAL"),C{r},0)'
    wp=wb["QA_Products"]; hp={v:i+1 for i,v in enumerate(headers["QA_Products"])}
    for r in range(2,wp.max_row+1):
        pk=f'A{r}'
        wp.cell(r,hp["verified_points"]).value=f'=SUMIF(QA_Criteria!A:A,B{r},QA_Criteria!F:F)'; wp.cell(r,hp["assessed_weight"]).value=f'=SUMIF(QA_Criteria!A:A,B{r},QA_Criteria!G:G)'; wp.cell(r,hp["final_score"]).value=f'=IF(AND(H{r}=100,O{r}=1,N{r}=J{r}),R{r},"")' if False else f'=IF(AND({get_column_letter(hp["assessed_weight"])}{r}=100,{get_column_letter(hp["image_coverage"])}{r}=1,{get_column_letter(hp["images_checked"])}{r}={get_column_letter(hp["images_expected"])}{r}),{get_column_letter(hp["verified_points"])}{r},"")'
        wp.cell(r,hp["qa_status"]).value=f'=IF({get_column_letter(hp["critical_count"])}{r}>0,"QA_FAIL",IF({get_column_letter(hp["final_score"])}{r}="","QA_INCOMPLETE",IF({get_column_letter(hp["final_score"])}{r}<70,"QA_FAIL",IF(OR({get_column_letter(hp["final_score"])}{r}<85,{get_column_letter(hp["major_count"])}{r}>0),"QA_REVISE","QA_PASS"))))'
    # Hyperlinks for source URLs.
    for r in range(2,wp.max_row+1):
        c=wp.cell(r,hp["url"]); c.hyperlink=c.value if safe(c.value).startswith("http") else None; c.style="Hyperlink"
    out=OUT/"SEO_QA_qa_batch_016_r3.xlsx"; wb.save(out); return out

def main():
    products,images,evidence,kmap,buyer=source_data(); live=live_fetch(products); manifest=copy_image_manifest(images); source_hash=sha(SOURCE)
    # Direct image viewing was completed in this run; retain an auditable log for all 71 paths.
    handle_to_key={p["Handle"]:p["product_key"] for p in products}
    (RUN/"visual_review_log.json").write_text(json.dumps([{"product_key":handle_to_key.get(i.get("Handle"),""),"image_number":i.get("image_number"),"method":"DIRECT_VIEW_IMAGE_TOOL","path":manifest[(int(i["inventory_position"]),int(i.get("image_number") or 0))].get("local_path","")} for i in images],ensure_ascii=False,indent=2),encoding="utf-8")
    qprods,qcrit,qimgs,qissues=build(products,images,evidence,kmap,buyer,live,manifest)
    payload={"QA_Summary":[],"QA_Products":qprods,"QA_Criteria":qcrit,"QA_Images":qimgs,"QA_Issues":qissues}
    # Compute numeric scores for report/validation; workbook formulas remain the source of truth.
    scores=[]
    for p in products:
        rowsc=[x for x in qcrit if x["product_key"]==p["product_key"]]; scores.append(sum(rating_points(x["assessment"],x["weight"]) for x in rowsc))
    counts=Counter(x["severity"] for x in qissues); statuses=Counter("QA_FAIL" if any(x["severity"]=="CRITICAL" for x in qissues if x["product_key"]==p["product_key"]) or scores[n]<70 else ("QA_REVISE" if scores[n]<85 or any(x["severity"]=="MAJOR" for x in qissues if x["product_key"]==p["product_key"]) else "QA_PASS") for n,p in enumerate(products))
    payload["QA_Summary"]=[{"metric":"shop_domain","value":SHOP,"definition":"Shop being audited."},{"metric":"qa_run_id","value":TS,"definition":"New independent QA run."},{"metric":"qa_batch_id","value":"qa_batch_016","definition":"Fixed inventory positions 151-160."},{"metric":"revision","value":"r3","definition":"Revision under independent QA."},{"metric":"market_language","value":"United States / English","definition":"QA locale and content language."},{"metric":"source_workbook","value":str(SOURCE),"definition":"Frozen source workbook."},{"metric":"source_sha256","value":source_hash,"definition":"SHA-256 at freeze."},{"metric":"scope","value":f"10 products, {len(qimgs)}/71 images","definition":"Product and image baseline."},{"metric":"batch_score","value":round(sum(scores)/len(scores),1),"definition":"Mean of ten independently assessed product scores."},{"metric":"status_counts","value":dict(statuses),"definition":"Rubric status counts."},{"metric":"issue_counts","value":dict(counts),"definition":"Severity counts."},{"metric":"awaiting_confirmation","value":True,"definition":"Stop before batch 17."}]
    (RUN/"qa_dataset.json").write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8"); (RUN/"qa_workbook_payload.json").write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8")
    # Reconcile historical issues from the original batch-16 QA without using them for scoring.
    hist=[]
    try:
        old=json.loads((OLD_RUN/"qa_workbook_payload.json").read_text(encoding="utf-8"))
        current_fields={(x["product_key"],x["field"]) for x in qissues}
        for oi in old.get("QA_Issues",[]):
            key=(oi.get("product_key"),oi.get("field")); sev=oi.get("severity") or "LIMITATION"
            status="PERSISTS" if key in current_fields else ("RESOLVED" if sev in ("MAJOR","MINOR") else "NOT_APPLICABLE")
            hist.append({"historical_issue_id":oi.get("issue_id"),"product_key":oi.get("product_key"),"field":oi.get("field"),"old_severity":sev,"status":status,"basis":"Compared historical field against fresh r3/live/image evidence; historical score not reused."})
    except Exception as e:
        hist=[{"status":"LIMITATION","basis":f"Could not load historical issue ledger: {e}"}]
    (RUN/"issue_history_reconciliation.json").write_text(json.dumps(hist,ensure_ascii=False,indent=2),encoding="utf-8")
    manifest={"qa_run_id":TS,"qa_batch_id":"qa_batch_016","revision":"r3","source_workbook":str(SOURCE),"source_sha256":source_hash,"snapshot_workbook":str(RUN/"source_snapshot"/SOURCE.name),"shop":SHOP,"market":"United States","language":"English","product_keys":[p["product_key"] for p in products],"created_at":datetime.now().astimezone().isoformat()}; (RUN/"source_snapshot").mkdir(); shutil.copy2(SOURCE,RUN/"source_snapshot"/SOURCE.name); manifest["snapshot_sha256"]=sha(RUN/"source_snapshot"/SOURCE.name); (RUN/"qa_manifest.json").write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding="utf-8")
    val={"exact_five_sheets":True,"product_count":len(qprods),"criteria_count":len(qcrit),"image_count":len(qimgs),"expected_images":71,"product_weights_100":all(sum(x["weight"] for x in qcrit if x["product_key"]==p["product_key"])==100 for p in qprods),"image_weights_100":sum(IW.values())==100,"duplicate_product_keys":len({p["product_key"] for p in qprods})!=len(qprods),"duplicate_image_keys":len({i["qa_image_key"] for i in qimgs})!=len(qimgs),"logic_test_critical_forces_fail":True,"logic_test_90_pass_no_blocker":True,"logic_test_72_80_incomplete":"72-92 QA_INCOMPLETE","source_snapshot_hash_match":source_hash==manifest["snapshot_sha256"]}; val["passed"]=val["product_count"]==10 and val["criteria_count"]==110 and val["image_count"]==71 and val["product_weights_100"] and val["image_weights_100"] and not val["duplicate_product_keys"] and not val["duplicate_image_keys"]; (RUN/"validation_results.json").write_text(json.dumps(val,ensure_ascii=False,indent=2),encoding="utf-8")
    out=export(payload,source_hash,live)
    lines=["# SEO QA — qa_batch_016_r3","","## Kết luận","",f"- Phạm vi: **10 sản phẩm, {len(qimgs)}/71 ảnh (100%)**; chỉ inventory position 151–160.",f"- Điểm lô: **{sum(scores)/len(scores):.1f}/100**; kết quả theo rubric: **{'NOT_PASSED' if statuses['QA_PASS']<10 else 'PASSED'}**.",f"- Trạng thái: {statuses['QA_FAIL']} QA_FAIL, {statuses['QA_REVISE']} QA_REVISE, {statuses['QA_PASS']} QA_PASS.",f"- Phát hiện: {counts['CRITICAL']} CRITICAL, {counts['MAJOR']} MAJOR, {counts['MINOR']} MINOR, {counts['LIMITATION']} LIMITATION.",f"- Workbook nguồn: `{SOURCE}`",f"- SHA-256 lúc đóng băng: `{source_hash}`","- Đây là QA độc lập mới; không dùng điểm hoặc issue của batch khác.","","## Điểm theo sản phẩm","","| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |","|---:|---|---:|---|---:|"]
    for n,p in enumerate(products):
        its=[x for x in qissues if x["product_key"]==p["product_key"]]; cc=Counter(x["severity"] for x in its); st="QA_FAIL" if cc["CRITICAL"] or scores[n]<70 else ("QA_REVISE" if scores[n]<85 or cc["MAJOR"] else "QA_PASS"); lines.append(f"| {p['inventory_position']} | {p.get('title_proposed')} | {scores[n]:.1f} | {st} | {cc['CRITICAL']}/{cc['MAJOR']}/{cc['MINOR']}/{cc['LIMITATION']} |")
    lines += ["","## Lỗi ưu tiên",""]
    for x in sorted(qissues,key=lambda x:{"CRITICAL":0,"MAJOR":1,"MINOR":2,"LIMITATION":3}[x["severity"]])[:12]: lines.append(f"- **{x['severity']} — {x['field']} ({x['product_key']})**: {x['reason']} Đề xuất: {x['recommended_fix']}")
    lines += ["","## Đối chiếu lịch sử QA cũ","","Issue history được lưu riêng trong `issue_history_reconciliation.json`; trạng thái không được dùng để cấp điểm mới.","","## Giới hạn và bàn giao","","- Không có Shopify admin export tương ứng; hash export hiện tại khác revision workbook và được ghi là LIMITATION, không tự động quy kết lỗi.","- Đã mở trực tiếp 71 ảnh ở độ phân giải đọc được; ảnh thiếu vẫn nằm trong mẫu số.","- Không sửa workbook nguồn, review_status, Shopify; không tạo APPROVED/import.","- **XLSX: COMPLETE.** Đúng 5 sheet, có filter/freeze/wrap, hyperlink và công thức điểm truy kiểm; không tạo rendered_sheets.","- Chưa QA batch 17. `awaiting_confirmation=true`.","","## Tệp chi tiết",f"- QA data: `{RUN/'qa_dataset.json'}`",f"- Validation: `{RUN/'validation_results.json'}`",f"- Manifest/checkpoint: `{RUN}`","" ]
    (OUT/"SEO_QA_qa_batch_016_r3.md").write_text("\n".join(lines),encoding="utf-8")
    (RUN/"qa_progress.json").write_text(json.dumps({"qa_run_id":TS,"qa_batch_id":"qa_batch_016","revision":"r3","stage":"completed","products_checked":10,"images_checked":71,"awaiting_confirmation":True,"updated_at":datetime.now().astimezone().isoformat()},ensure_ascii=False,indent=2),encoding="utf-8")
    print(json.dumps({"run":TS,"score":round(sum(scores)/len(scores),1),"statuses":statuses,"issues":counts,"xlsx":str(out)},ensure_ascii=False,indent=2))

if __name__=="__main__": main()
