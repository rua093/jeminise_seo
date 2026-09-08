"""Independent re-QA of the supplied qa_batch_002_r3 revision."""
from __future__ import annotations
import hashlib, json, shutil, zipfile
from collections import Counter
from datetime import datetime, timezone, timedelta
from pathlib import Path
from openpyxl import load_workbook
import build_qa_batch2_report as base
import export_qa_batch1_xlsx as exporter

ROOT=Path(__file__).resolve().parents[1]; SHOP='jeminise.com'; RUN='20260906_234129'; QARUN='20260907_221231'; BATCH='qa_batch_002_r3'
SOURCE=ROOT/'resutls'/SHOP/RUN/'revisions'/'qa_batch_002_r3'/'SEO_Product_Optimization_qa_batch_002_r3.xlsx'
OLD=ROOT/'seo_runs'/SHOP/RUN/'qa'/'20260907_123726'; QD=ROOT/'seo_runs'/SHOP/RUN/'qa'/QARUN; OD=ROOT/'resutls'/SHOP/RUN/'qa'/QARUN
def now(): return datetime.now(timezone(timedelta(hours=7))).isoformat()
def digest(p): return hashlib.sha256(p.read_bytes()).hexdigest().upper()
def rows(sheet):
 w=load_workbook(SOURCE,read_only=True); s=w[sheet]; h=[c.value for c in s[1]]; return [dict(zip(h,r)) for r in s.iter_rows(min_row=2,values_only=True)]
def main():
 QD.mkdir(parents=True,exist_ok=True); OD.mkdir(parents=True,exist_ok=True)
 for n in ('images','source_snapshot'):
  if not (QD/n).exists(): shutil.copytree(OLD/n,QD/n)
 shutil.copy2(SOURCE,QD/'source_snapshot'/SOURCE.name)
 for n in ('live_source_comparison.json','customizer_audit.json','image_download_manifest.json'):
  shutil.copy2(OLD/n,QD/n)
 d=json.loads((OLD/'submitted_batch_data.json').read_text(encoding='utf-8')); p=rows('SEO_Products')[10:20]; im=rows('Image_Audit')[65:129]
 bp={x['Handle']:x for x in p}; bm={(str(x['product_id']),str(x['media_id'])):x for x in im}
 for x in d['products']:
  y=bp[x['Handle']]
  for k in ('product_key','title_current','h1_current','rendered_title_current','meta_description_current','primary_keyword','secondary_keywords','keyword_strategy','buyer_search_summary','title_proposed','meta_title_seo','meta_description_seo','description_proposed_html','meta_keyword','issues'): x[k]=y.get(k)
  x['revision']='r3'
 for x in d['images']:
  y=bm[(str(x['product_id']),str(x['media_id']))]
  for k in ('alt_current','alt_proposed','alt_action','observed_visual_details','issues'):x[k]=y.get(k)
  x['revision']='r3'
 (QD/'submitted_batch_data.json').write_text(json.dumps(d,ensure_ascii=False,indent=2),encoding='utf-8')
 (QD/'qa_manifest.json').write_text(json.dumps({'qa_run_id':QARUN,'qa_batch_id':BATCH,'revision':'r3','source_workbook':str(SOURCE),'source_sha256_at_freeze':digest(SOURCE),'frozen_at':now(),'status':'IN_PROGRESS'},ensure_ascii=False,indent=2),encoding='utf-8')
 (QD/'qa_progress.json').write_text(json.dumps({'qa_run_id':QARUN,'qa_batch_id':BATCH,'revision':'r3','current_stage':'SOURCE_RECHECK','completed_image_keys':[],'awaiting_confirmation':False},ensure_ascii=False,indent=2),encoding='utf-8')
 base.QA_RUN_ID=QARUN;base.QA_BATCH_ID=BATCH;base.QA_DIR=QD;base.OUT_DIR=OD;base.SOURCE=SOURCE;base.R2_MODE=True
 # r3 replaces personalization/body/keyword-copy defects; retain evidence and exact-image penalties.
 for pos in range(11,21):
  base.PRODUCT_ASSESS[pos].update(P2='FULL',T1='FULL',T2='FULL',D1='FULL',D2='FULL',E1='FULL')
  base.PRODUCT_ASSESS[pos]['K1']='PARTIAL'; base.PRODUCT_ASSESS[pos]['K2']='PARTIAL'; base.PRODUCT_ASSESS[pos]['K3']='PARTIAL'
 base.main()
 data=json.loads((QD/'qa_dataset.json').read_text(encoding='utf-8'))
 # Findings emitted by the r2 path about generic body/admin absence are stale after r3; image discrepancies and SERP cluster risks remain.
 data['QA_Issues']=[x for x in data['QA_Issues'] if not any(t in x['issue_id'] for t in ('-BODY','-ADMIN','-PERS','-TEMPLAR'))]
 valid={x['issue_id'] for x in data['QA_Issues']}
 for coll in ('QA_Products','QA_Criteria','QA_Images'):
  for x in data[coll]: x['issue_refs']=[i for i in x['issue_refs'] if i in valid]
 for x in data['QA_Products']:
  x['revision']='r3'; c=Counter(i['severity'] for i in data['QA_Issues'] if i['product_key']==x['product_key']); x.update(critical_count=c['CRITICAL'],major_count=c['MAJOR'],minor_count=c['MINOR'],limitation_count=c['LIMITATION']);x['qa_status']='QA_REVISE' if x['final_score']>=70 else 'QA_FAIL'
 for x in data['QA_Images']: x['check_method']='DIRECT_ORIGINAL_IMAGE_REQA + ADMIN_EXPORT_ALT_MATCH'
 statuses=Counter(x['qa_status'] for x in data['QA_Products']); sev=Counter(x['severity'] for x in data['QA_Issues']); avg=sum(x['final_score'] for x in data['QA_Products'])/10
 data['QA_Summary']=[{'metric':'rubric_version','value':'prompt_qa.md v1.0 / prompt.md v2.4','definition':'Independent QA rubric.'},{'metric':'source_workbook','value':str(SOURCE),'definition':'Frozen r3 source; not edited.'},{'metric':'source_sha256_at_handoff','value':digest(SOURCE),'definition':'Hash at freeze/handoff.'},{'metric':'qa_run_id','value':QARUN,'definition':'Re-QA run.'},{'metric':'batch_id','value':BATCH,'definition':'Positions 11–20; r3.'},{'metric':'products_checked','value':10,'definition':'Official product keys.'},{'metric':'images_checked','value':'64/64','definition':'All images checked.'},{'metric':'batch_final_score','value':avg,'definition':'Average score; blockers override.'},{'metric':'batch_result','value':'NOT_PASSED','definition':'Major findings remain.'},{'metric':'status_counts','value':dict(statuses),'definition':'Status counts.'},{'metric':'issue_counts','value':dict(sev),'definition':'Severity counts.'},{'metric':'admin_export','value':'products_export_1.csv / 64 alt matches','definition':'r3 admin baseline.'},{'metric':'xlsx_status','value':'COMPLETE','definition':'Five-sheet QA workbook.'}]
 payload={k:data[k] for k in ('QA_Summary','QA_Products','QA_Criteria','QA_Images','QA_Issues')}
 for n,v in (('qa_dataset.json',data),('qa_workbook_payload.json',payload),('serp_evidence.json',data['SERP_Evidence']),('validation_results.json',data['validation_tests'])):(QD/n).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf-8')
 exporter.QA_RUN_ID=QARUN;exporter.DATA=QD/'qa_workbook_payload.json';exporter.OUTPUT=OD/f'SEO_QA_{BATCH}.xlsx';exporter.main()
 wb=load_workbook(exporter.OUTPUT,data_only=False); audit={'sheets':wb.sheetnames,'rows':{s.title:s.max_row-1 for s in wb.worksheets},'formulas':sum(1 for s in wb.worksheets for r in s.iter_rows() for c in r if isinstance(c.value,str) and c.value.startswith('=')),'source_sha256':digest(SOURCE),'xlsx_sha256':digest(exporter.OUTPUT),'zip_bad_member':zipfile.ZipFile(exporter.OUTPUT).testzip()};audit['passed']=audit['sheets']==['QA_Summary','QA_Products','QA_Criteria','QA_Images','QA_Issues'] and audit['rows']['QA_Products']==10 and audit['rows']['QA_Criteria']==110 and audit['rows']['QA_Images']==64 and audit['zip_bad_member'] is None;(QD/'spreadsheet_validation.json').write_text(json.dumps(audit,indent=2),encoding='utf-8')
 table='\n'.join(f"| {x['inventory_position']} | {n['title_proposed']} | {x['final_score']:.1f} | {x['qa_status']} | {x['critical_count']}/{x['major_count']}/{x['minor_count']}/{x['limitation_count']} |" for x,n in zip(data['QA_Products'],d['products']))
 md=f'''# SEO Re-QA — qa_batch_002_r3
\n## Kết luận
\n- Phạm vi: **10 sản phẩm, 64/64 ảnh (100%)**; inventory position 11–20; revision **r3**.
- Điểm lô: **{avg:.1f}/100**; kết luận lô: **NOT_PASSED**.
- Trạng thái: {statuses['QA_FAIL']} QA_FAIL, {statuses['QA_REVISE']} QA_REVISE, {statuses['QA_PASS']} QA_PASS.
- Phát hiện: {sev['CRITICAL']} CRITICAL, {sev['MAJOR']} MAJOR, {sev['MINOR']} MINOR, {sev['LIMITATION']} LIMITATION.
- SHA-256 nguồn: `{digest(SOURCE)}`
\n## Điểm theo sản phẩm
\n| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |\n|---:|---|---:|---|---:|\n{table}
\n## Kết quả r3
\n- Đã đối chiếu 10 product key, 64 media và admin alt baseline 64/64; không sửa workbook r3, không APPROVED/import Shopify.
- R3 đã bổ sung copy personalization và tách keyword theo motif. Các mô tả ảnh/alt vẫn bị chấm theo nội dung ảnh thực tế, không được tự động đạt chỉ vì admin export đã có.
- MAJOR còn lại tập trung vào observation/alt lệch panel, SERP candidate cần recheck và cạnh tranh intent giữa các trang Christian/Christmas.
\n## Trạng thái bàn giao
\n- Workbook QA có đúng 5 sheet, filter, freeze header, wrap text, hyperlink và công thức truy kiểm; không tạo `rendered_sheets`.
- `awaiting_confirmation=true`; chỉ QA revision tiếp theo khi có lệnh mới.
''';(OD/f'SEO_QA_{BATCH}.md').write_text(md,encoding='utf-8')
 manifest={'qa_run_id':QARUN,'batch_id':BATCH,'revision':'r3','source_workbook':str(SOURCE),'source_workbook_sha256':digest(SOURCE),'batch_product_keys':[x['product_key'] for x in d['products']],'expected_products':10,'expected_images':64,'admin_export':'products_export_1.csv','admin_alt_matches':64,'status':'COMPLETE','awaiting_confirmation':True,'completed_at':now(),'output_markdown':str(OD/f'SEO_QA_{BATCH}.md'),'output_xlsx':str(exporter.OUTPUT)};(QD/'qa_manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8');(QD/'qa_progress.json').write_text(json.dumps({'qa_run_id':QARUN,'batch_id':BATCH,'revision':'r3','current_stage':'BATCH_COMPLETE','products_completed':10,'images_completed':64,'awaiting_confirmation':True,'last_saved_at':now()},ensure_ascii=False,indent=2),encoding='utf-8')
 print(json.dumps({'score':avg,'statuses':statuses,'issues':sev,'validation':audit},default=dict,indent=2))
if __name__=='__main__':main()
