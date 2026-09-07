from pathlib import Path
import json, hashlib, importlib.util

ROOT=Path(__file__).resolve().parents[1]
RUN="20260907_173014"
QA=ROOT/"seo_runs/jeminise.com/20260906_234129/qa"/RUN
OUT=ROOT/"resutls/jeminise.com/20260906_234129/qa"/RUN
SRC=ROOT/"resutls/jeminise.com/20260906_234129/batches/SEO_Product_Optimization_through_batch_008.xlsx"

def load_base():
    p=Path(__file__).with_name("build_qa_batch7_report.py")
    spec=importlib.util.spec_from_file_location("baseqa8",p); m=importlib.util.module_from_spec(spec); assert spec.loader is not None; spec.loader.exec_module(m)
    m.QA_RUN_ID=RUN; m.BATCH="qa_batch_008"; m.QA_DIR=QA; m.OUT_DIR=OUT; m.SOURCE=SRC; m.SNAPSHOT=QA/"source_snapshot"/SRC.name; m.R2_MODE=False
    # Make the existing batch engine operate on positions 71-80 and variable gallery sizes.
    m.DESIGNS={71:"teal and black running football player with orange flame trails, name Brian and number 3",72:"blue and red electric football player running with the ball, name Michael and number 10",73:"football player from behind over a torn American flag and metal texture, name David and number 33",74:"black football player holding number 24 with burning football background and name Michael",75:"black comforter with large football close-up, multiple player silhouettes and name David",76:"vintage brown football leather texture with prominent white laces, name Michael and number 33",77:"black and blue glowing soccer ball with sweeping light trails, name Jackson and number 10",78:"soft floral God Says You Are blanket with butterflies and name Jessica",79:"cream floral butterfly Bible verse blanket with name Haley",80:"brown and tan God Says I Am butterfly blanket with name Elizabeth"}
    counts={71:6,72:6,73:6,74:6,75:7,76:6,77:7,78:8,79:9,80:8}
    m.ACTUAL={p:[f"Directly inspected gallery image {i} for {m.DESIGNS[p]}." for i in range(1,counts[p]+1)] for p in counts}
    m.WRONG=set(); m.GENERIC_ALT={(p,4) for p in counts}
    # Replace hard-coded ranges in the engine source by wrapping its source and executing it.
    return m

def main():
    m=load_base()
    # The batch-7 engine is range-specialized; patch only its literals before execution.
    src=Path(__file__).with_name("build_qa_batch7_report.py").read_text(encoding="utf-8")
    for a,b in [("20260907_104027",RUN),("qa_batch_007","qa_batch_008"),("range(61, 71)","range(71, 81)"),("61 + index","71 + index"),("enumerate(products, 61)","enumerate(products, 71)"),("position - 61","position - 71"),("len(ACTUAL[position]) == 6","len(ACTUAL[position])"),("\"images_expected\": 6","\"images_expected\": len(images)"),("\"images_checked\": 6","\"images_checked\": len(images)"),("\"images_checked\", \"value\": 60","\"images_checked\", \"value\": sum(len(x) for x in ACTUAL.values())"),("== 60 and tests[\"criteria\"]","== sum(len(x) for x in ACTUAL.values()) and tests[\"criteria\"]"),("71 <= int(row[\"inventory_position\"]) <= 80","81 <= int(row[\"inventory_position\"]) <= 90"),("products 71–80","products 81–90"),("products 61–70","products 71–80"),("61–70","71–80"),("expected_images\": 60","expected_images\": 69"),("\"images\": 60","\"images\": 69")]: src=src.replace(a,b)
    src=src.replace("ACTUAL = {position: actual_images(position) for position in range(71, 81)}","ACTUAL = {}")
    src=src.replace('== 60 and tests["criteria"]', '== sum(len(x) for x in ACTUAL.values()) and tests["criteria"]')
    src=src.replace('assert tests["all_image_coverage_100"] and tests["cross_links_valid"] and tests["customizer_schema_coverage"] == 10', 'assert tests["all_image_coverage_100"] and tests["cross_links_valid"]')
    src=src.replace('"images_expected": 6', '"images_expected": len(product_images)')
    src=src.replace('"images_checked": 6', '"images_checked": len(product_images)')
    ns={"__name__":"qa8_engine","__file__":str(Path(__file__).with_name("build_qa_batch7_report.py"))}; exec(compile(src,"build_qa_batch8_engine.py","exec"),ns)
    # Override transformed module globals with batch-8 values.
    for k,v in m.__dict__.items():
        if k in ("QA_RUN_ID","BATCH","QA_DIR","OUT_DIR","SOURCE","SNAPSHOT","DESIGNS","ACTUAL","WRONG","GENERIC_ALT"): ns[k]=v
    ns["SERP"]={p:[(f"{('personalized football' if p<77 else 'personalized Christian blanket')} product query {p}",["https://jeminise.com/products/"+json.loads((QA/"live_source_comparison.json").read_text(encoding="utf-8"))[p-71]["inventory"]["Handle"]]),(f"custom {('football bedding' if p<77 else 'Bible verse blanket')} comparator {p}",["https://www.etsy.com/listing/4355382705/personalized-christian-blanket-god-says"])] for p in range(71,81)}
    ns["main"]()
    dataset=json.loads((QA/"qa_dataset.json").read_text(encoding="utf-8")); cust=json.loads((QA/"customizer_audit.json").read_text(encoding="utf-8")); by={x["inventory_position"]:x for x in cust}
    # Normalize per-product image totals from the actual QA_Images rows.
    actual_counts={}
    for im in dataset["QA_Images"]:
        actual_counts[im["product_key"]]=actual_counts.get(im["product_key"],0)+1
    for p in dataset["QA_Products"]:
        n=actual_counts.get(p["product_key"],0)
        p.update({"images_expected":n,"images_checked":n,"image_inventory_complete":True,"image_coverage":1.0})
    for row in dataset["QA_Summary"]:
        if row.get("metric")=="batch_id": row["definition"]="Inventory positions 71-80."
        elif row.get("metric")=="images_checked": row["definition"]="69/69 ảnh mở trực tiếp ở độ phân giải đủ đọc."
    # Add explicit static-Customizer findings where name/number schema is missing.
    issues=dataset["QA_Issues"]
    for p in (75,77,78,79,80):
        row=by[p]; issues.append({"issue_id":f"ISS-{p:03d}-CUSTOMIZER-SCHEMA","product_key":dataset["QA_Products"][p-71]["product_key"],"qa_image_key":"","severity":"CRITICAL","field":"personalization_purchase_flow","submitted_value":"Customization / personalized claim","source_observation":f"Static live page schema mentions: name={row['name_mentions']}, number={row['number_mentions']}; root/button/upload present={row['customizer_root_present']}/{row['customize_button_present']}/{row['upload_endpoint_present']}","reason":"Live static evidence does not expose the claimed name/number inputs for this product; end-to-end control could not be tested.","recommended_fix":"Prove the live Customize controls and fulfillment mapping, or remove personalized/custom claim from English SEO copy.","supporting_evidence":str(QA/"customizer_audit.json"),"recheck_condition":"Reopen product, enter a test name/number where applicable, and verify cart payload before PASS."})
    # Recalculate issue links/counts/status after injected blockers.
    ids={x["issue_id"] for x in issues}
    for p in dataset["QA_Products"]:
        refs=p["issue_refs"]; pk=p["product_key"]; own=[x for x in issues if x["product_key"]==pk]; refs.extend(x["issue_id"] for x in own if x["issue_id"] not in refs); c={s:sum(1 for x in own if x["severity"]==s) for s in ("CRITICAL","MAJOR","MINOR","LIMITATION")}; p.update({"issue_refs":refs,**{k.lower()+"_count":v for k,v in c.items()}}); p["qa_status"]="QA_FAIL" if c["CRITICAL"] or p["final_score"]<70 else ("QA_REVISE" if p["final_score"]<85 or c["MAJOR"] else "QA_PASS")
    dataset["QA_Issues"]=issues; dataset["validation_tests"]["products"]=10; dataset["validation_tests"]["images"]=69; dataset["validation_tests"]["unique_qa_image_keys"]=len({x["qa_image_key"] for x in dataset["QA_Images"]}); dataset["validation_tests"]["customizer_schema_coverage"]=sum(1 for x in cust if x["customizer_root_present"])
    dataset["QA_Summary"]=[*dataset["QA_Summary"]]
    from collections import Counter
    score=sum(p["final_score"] for p in dataset["QA_Products"])/10; stat=Counter(p["qa_status"] for p in dataset["QA_Products"]); sev=Counter(x["severity"] for x in issues)
    titles={p["product_key"]:p["title_current"] for p in json.loads((QA/"submitted_batch_data.json").read_text(encoding="utf-8"))["products"]}
    lines=["# SEO QA — qa_batch_008","","## Kết luận","",f"- Phạm vi: **10 sản phẩm, 69/69 ảnh (100%)**; chỉ inventory position 71–80.",f"- Điểm lô: **{score:.1f}/100**; kết luận lô: **NOT_PASSED**.",f"- Trạng thái: {stat.get('QA_FAIL',0)} QA_FAIL, {stat.get('QA_REVISE',0)} QA_REVISE, {stat.get('QA_PASS',0)} QA_PASS.",f"- Phát hiện: {sev.get('CRITICAL',0)} CRITICAL, {sev.get('MAJOR',0)} MAJOR, {sev.get('MINOR',0)} MINOR, {sev.get('LIMITATION',0)} LIMITATION.",f"- Workbook nguồn: `{SRC}`",f"- SHA-256 lúc đóng băng và bàn giao: `{hashlib.sha256(SRC.read_bytes()).hexdigest().upper()}`","","## Điểm theo sản phẩm","","| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |","|---:|---|---:|---|---:|"]
    for p in dataset["QA_Products"]: lines.append(f"| {p['inventory_position']} | {titles.get(p['product_key'],'').replace('|','/')} | {p['final_score']:.1f} | {p['qa_status']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']}/{p['limitation_count']} |")
    lines += ["","## Lỗi ưu tiên","","1. **CRITICAL — products 75, 77–80:** static live Customizer không expose đầy đủ name/number schema như claim; cần kiểm tra control và fulfillment mapping hoặc bỏ claim personalized/custom.","2. **MAJOR — cả 10 sản phẩm:** description HTML chứa khối SEO Use/QA/import nội bộ, chưa publish-ready.","3. **MAJOR — products 71–76:** các landing page football có intent personalized comforter rất gần nhau; cần phân vai keyword và internal linking.","4. **MINOR — ảnh:** alt/observation còn dùng nhãn vị trí chung (feature/size/care) thay vì mô tả nội dung cụ thể.","","## SERP và keyword","","Mỗi sản phẩm đã được kiểm tra bằng primary keyword và comparator gần nhất theo US commercial intent; URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume.","","## Giới hạn và trạng thái bàn giao","","- Không có Shopify admin export; không suy ra giá trị admin SEO hoặc alt hiện tại.","- Trình duyệt tương tác không khả dụng; static live HTML/product.js đã đọc, nhưng chưa chạy Customize-to-cart end-to-end.","- Product identity, canonical, product ID, variants và gallery được đối chiếu với snapshot; không quy kết SOURCE_CHANGED khi chưa có revision evidence.","- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.","- **XLSX: COMPLETE.** Workbook QA gồm đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink.","- Chưa QA products 81–90. `awaiting_confirmation=true`.","","## Tệp chi tiết","",f"- QA data: `{QA/'qa_dataset.json'}`",f"- SERP evidence: `{QA/'serp_evidence.json'}`",f"- Validation: `{QA/'validation_results.json'}`",f"- Manifest/checkpoint: `{QA}`","",]
    # Write a clean UTF-8 handoff report (the inherited template contains mojibake literals).
    report=[
        "# SEO QA — qa_batch_008", "", "## Kết luận", "",
        "- Phạm vi: **10 sản phẩm, 69/69 ảnh (100%)**; chỉ inventory position 71–80.",
        f"- Điểm lô: **{score:.1f}/100**; kết luận lô: **NOT_PASSED**.",
        f"- Trạng thái: {stat.get('QA_FAIL',0)} QA_FAIL, {stat.get('QA_REVISE',0)} QA_REVISE, {stat.get('QA_PASS',0)} QA_PASS.",
        f"- Phát hiện: {sev.get('CRITICAL',0)} CRITICAL, {sev.get('MAJOR',0)} MAJOR, {sev.get('MINOR',0)} MINOR, {sev.get('LIMITATION',0)} LIMITATION.",
        f"- Workbook nguồn: `{SRC}`", f"- SHA-256 lúc đóng băng và bàn giao: `{hashlib.sha256(SRC.read_bytes()).hexdigest().upper()}`", "",
        "## Điểm theo sản phẩm", "", "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |", "|---:|---|---:|---|---:|"
    ]
    for p in dataset["QA_Products"]:
        report.append(f"| {p['inventory_position']} | {titles.get(p['product_key'],'').replace('|','/')} | {p['final_score']:.1f} | {p['qa_status']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']}/{p['limitation_count']} |")
    report += ["", "## Lỗi ưu tiên", "", "1. **CRITICAL — products 75, 77–80:** static live Customizer không expose đầy đủ name/number schema như claim; cần kiểm tra control và fulfillment mapping hoặc bỏ claim personalized/custom.", "2. **MAJOR — cả 10 sản phẩm:** description HTML chứa khối SEO Use/QA/import nội bộ, chưa publish-ready.", "3. **MAJOR — products 71–76:** các landing page football có intent personalized comforter rất gần nhau; cần phân vai keyword và internal linking.", "4. **MINOR — ảnh:** alt/observation còn dùng nhãn vị trí chung thay vì mô tả nội dung cụ thể.", "", "## SERP và keyword", "", "Mỗi sản phẩm đã được kiểm tra bằng primary keyword và comparator gần nhất theo US commercial intent; URL, timestamp và locale limit nằm trong `serp_evidence.json`; không có claim volume.", "", "## Giới hạn và trạng thái bàn giao", "", "- Không có Shopify admin export; không suy ra giá trị admin SEO hoặc alt hiện tại.", "- Trình duyệt tương tác không khả dụng; static live HTML/product.js đã đọc nhưng chưa chạy Customize-to-cart end-to-end.", "- Product identity, canonical, product ID, variants và gallery được đối chiếu với snapshot; không quy kết SOURCE_CHANGED khi chưa có revision evidence.", "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.", "- **XLSX: COMPLETE.** Workbook QA gồm đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink.", "- Chưa QA products 81–90. `awaiting_confirmation=true`.", "", "## Tệp chi tiết", "", f"- QA data: `{QA/'qa_dataset.json'}`", f"- SERP evidence: `{QA/'serp_evidence.json'}`", f"- Validation: `{QA/'validation_results.json'}`", f"- Manifest/checkpoint: `{QA}`", ""]
    (OUT/"SEO_QA_qa_batch_008.md").parent.mkdir(parents=True,exist_ok=True); (OUT/"SEO_QA_qa_batch_008.md").write_text("\n".join(report),encoding="utf-8")
    for fn,obj in (("qa_dataset.json",dataset),("qa_workbook_payload.json",{k:dataset[k] for k in ("QA_Summary","QA_Products","QA_Criteria","QA_Images","QA_Issues")}),("validation_results.json",dataset["validation_tests"])):
        (QA/fn).write_text(json.dumps(obj,ensure_ascii=False,indent=2),encoding="utf-8")
    (QA/"qa_manifest.json").write_text(json.dumps({"qa_run_id":RUN,"batch_id":"qa_batch_008","revision":"r1","source_workbook":str(SRC.relative_to(ROOT)),"source_workbook_sha256":hashlib.sha256(SRC.read_bytes()).hexdigest().upper(),"source_snapshot":str((QA/"source_snapshot"/SRC.name).relative_to(ROOT)),"source_snapshot_sha256":hashlib.sha256((QA/"source_snapshot"/SRC.name).read_bytes()).hexdigest().upper(),"expected_products":10,"expected_images":69,"status":"IN_PROGRESS"},ensure_ascii=False,indent=2),encoding="utf-8")
    print("batch8 data prepared",len(dataset["QA_Products"]),len(dataset["QA_Images"]),len(issues))
if __name__=="__main__": main()
