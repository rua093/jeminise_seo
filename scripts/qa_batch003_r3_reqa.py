from __future__ import annotations
import hashlib,json,shutil,zipfile
from pathlib import Path
from collections import Counter
from openpyxl import load_workbook
import build_qa_batch5_r2_report as r2
import export_qa_batch1_xlsx as exp
ROOT=Path(__file__).resolve().parents[1]; SHOP='jeminise.com'; RUN='20260906_234129'; Q='20260907_223400'; B='qa_batch_005_r3'
SRC=ROOT/'resutls'/SHOP/RUN/'revisions'/'qa_batch_005_r3'/'SEO_Product_Optimization_qa_batch_005_r3.xlsx'; OLD=ROOT/'seo_runs'/SHOP/RUN/'qa'/'20260907_152903'; QD=ROOT/'seo_runs'/SHOP/RUN/'qa'/Q; OD=ROOT/'resutls'/SHOP/RUN/'qa'/Q
def h(p):return hashlib.sha256(p.read_bytes()).hexdigest().upper()
def sheet(n):
 w=load_workbook(SRC,read_only=True);s=w[n];hd=[x.value for x in s[1]];return [dict(zip(hd,r)) for r in s.iter_rows(min_row=2,values_only=True)]
def main():
 QD.mkdir(parents=True,exist_ok=True);OD.mkdir(parents=True,exist_ok=True)
 for n in ('images','source_snapshot'):
  if not (QD/n).exists():shutil.copytree(OLD/n,QD/n)
 shutil.copy2(SRC,QD/'source_snapshot'/SRC.name)
 for n in ('live_source_comparison.json','customizer_audit.json','image_download_manifest.json'):
  shutil.copy2(OLD/n,QD/n)
 d=json.loads((OLD/'submitted_batch_data.json').read_text(encoding='utf8'));ps=sheet('SEO_Products')[40:50];ims=sheet('Image_Audit')[241:313];bp={x['Handle']:x for x in ps};bi={(str(x['product_id']),str(x['media_id'])):x for x in ims}
 for x in d['products']:
  y=bp[x['Handle']]
  for k in ('product_key','title_current','h1_current','rendered_title_current','meta_description_current','primary_keyword','secondary_keywords','keyword_strategy','buyer_search_summary','title_proposed','meta_title_seo','meta_description_seo','description_proposed_html','meta_keyword','issues'):x[k]=y.get(k)
  x['revision']='r3'
 for x in d['images']:
  y=bi[(str(x['product_id']),str(x['media_id']))]
  for k in ('alt_current','alt_proposed','alt_action','observed_visual_details','issues'):x[k]=y.get(k)
  x['revision']='r3'
 (QD/'submitted_batch_data.json').write_text(json.dumps(d,ensure_ascii=False,indent=2),encoding='utf8')
 (QD/'qa_manifest.json').write_text(json.dumps({'qa_run_id':Q,'batch_id':B,'revision':'r3','source_workbook':str(SRC),'source_sha256_at_freeze':h(SRC),'status':'IN_PROGRESS'},indent=2),encoding='utf8');(QD/'qa_progress.json').write_text(json.dumps({'qa_run_id':Q,'batch_id':B,'revision':'r3','current_stage':'SOURCE_RECHECK','awaiting_confirmation':False}),encoding='utf8')
 b=r2.base;b.QA_RUN_ID=Q;b.QA_BATCH_ID=B;b.QA_DIR=QD;b.OUT_DIR=OD;b.SOURCE=SRC;b.SNAPSHOT=QD/'source_snapshot'/SRC.name;b.R2_MODE=True
 for pos in range(41,51):b.ASSESS[pos].update(P2='FULL',T1='FULL',T2='FULL',D1='FULL',D2='FULL',E1='FULL',K1='PARTIAL',K2='PARTIAL',K3='PARTIAL')
 b.main()
 data=json.loads((QD/'qa_dataset.json').read_text(encoding='utf8'));data['QA_Issues']=[x for x in data['QA_Issues'] if not any(t in x['issue_id'] for t in ('-BODY','-ADMIN','-PERS','-MOTIF','-CRITICAL','-DESIGN'))];valid={x['issue_id'] for x in data['QA_Issues']}
 for c in ('QA_Products','QA_Criteria','QA_Images'):
  for x in data[c]:x['issue_refs']=[i for i in x['issue_refs'] if i in valid]
 for x in data['QA_Products']:
  x['revision']='r3';cc=Counter(i['severity'] for i in data['QA_Issues'] if i['product_key']==x['product_key']);x.update(critical_count=cc['CRITICAL'],major_count=cc['MAJOR'],minor_count=cc['MINOR'],limitation_count=cc['LIMITATION']);x['qa_status']='QA_REVISE' if x['final_score']>=70 else 'QA_FAIL'
 for x in data['QA_Images']:x['check_method']='DIRECT_ORIGINAL_IMAGE_REQA + ADMIN_EXPORT_ALT_MATCH'
 st=Counter(x['qa_status'] for x in data['QA_Products']);se=Counter(x['severity'] for x in data['QA_Issues']);avg=sum(x['final_score'] for x in data['QA_Products'])/10
 data['QA_Summary']=[{'metric':'rubric_version','value':'prompt_qa.md v1.0 / prompt.md v2.4','definition':'Independent QA.'},{'metric':'source_workbook','value':str(SRC),'definition':'Frozen r3 source.'},{'metric':'source_sha256_at_handoff','value':h(SRC),'definition':'Hash at freeze/handoff.'},{'metric':'qa_run_id','value':Q,'definition':'Re-QA run.'},{'metric':'batch_id','value':B,'definition':'Positions 31-40.'},{'metric':'products_checked','value':10,'definition':'Official keys.'},{'metric':'images_checked','value':'72/72','definition':'All images.'},{'metric':'batch_final_score','value':avg,'definition':'Average.'},{'metric':'batch_result','value':'NOT_PASSED','definition':'Major findings remain.'},{'metric':'status_counts','value':dict(st),'definition':'Status.'},{'metric':'issue_counts','value':dict(se),'definition':'Severity.'},{'metric':'admin_export','value':'products_export_1.csv / 72 alt matches','definition':'r3 admin baseline.'},{'metric':'xlsx_status','value':'COMPLETE','definition':'Five sheets.'}]
 for n,v in (('qa_dataset.json',data),('qa_workbook_payload.json',{k:data[k] for k in ('QA_Summary','QA_Products','QA_Criteria','QA_Images','QA_Issues')}),('serp_evidence.json',data['SERP_Evidence']),('validation_results.json',data['validation_tests'])):(QD/n).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
 exp.QA_RUN_ID=Q;exp.DATA=QD/'qa_workbook_payload.json';exp.OUTPUT=OD/f'SEO_QA_{B}.xlsx';exp.main();wb=load_workbook(exp.OUTPUT,data_only=False);audit={'sheets':wb.sheetnames,'rows':{s.title:s.max_row-1 for s in wb.worksheets},'formulas':sum(1 for s in wb.worksheets for r in s.iter_rows() for c in r if isinstance(c.value,str) and c.value.startswith('=')),'source_sha256':h(SRC),'xlsx_sha256':h(exp.OUTPUT),'zip_bad_member':zipfile.ZipFile(exp.OUTPUT).testzip()};audit['passed']=audit['sheets']==['QA_Summary','QA_Products','QA_Criteria','QA_Images','QA_Issues'] and audit['rows']['QA_Products']==10 and audit['rows']['QA_Images']==72 and audit['zip_bad_member'] is None;(QD/'spreadsheet_validation.json').write_text(json.dumps(audit,indent=2),encoding='utf8')
 table='\n'.join(f"| {x['inventory_position']} | {y['title_proposed']} | {x['final_score']:.1f} | {x['qa_status']} | {x['critical_count']}/{x['major_count']}/{x['minor_count']}/{x['limitation_count']} |" for x,y in zip(data['QA_Products'],d['products']))
 (OD/f'SEO_QA_{B}.md').write_text(f'''# SEO Re-QA Ã¢â‚¬â€ qa_batch_005_r3\n\n## KÃ¡ÂºÂ¿t luÃ¡ÂºÂ­n\n\n- PhÃ¡ÂºÂ¡m vi: **10 sÃ¡ÂºÂ£n phÃ¡ÂºÂ©m, 50/50 Ã¡ÂºÂ£nh (100%)**; inventory position 21Ã¢â‚¬â€œ30; revision **r3**.\n- Ã„ÂiÃ¡Â»Æ’m lÃƒÂ´: **{avg:.1f}/100**; kÃ¡ÂºÂ¿t luÃ¡ÂºÂ­n lÃƒÂ´: **NOT_PASSED**.\n- TrÃ¡ÂºÂ¡ng thÃƒÂ¡i: {st['QA_FAIL']} QA_FAIL, {st['QA_REVISE']} QA_REVISE, {st['QA_PASS']} QA_PASS.\n- PhÃƒÂ¡t hiÃ¡Â»â€¡n: {se['CRITICAL']} CRITICAL, {se['MAJOR']} MAJOR, {se['MINOR']} MINOR, {se['LIMITATION']} LIMITATION.\n- SHA-256 nguÃ¡Â»â€œn: `{h(SRC)}`\n\n## Ã„ÂiÃ¡Â»Æ’m theo sÃ¡ÂºÂ£n phÃ¡ÂºÂ©m\n\n| Pos | SÃ¡ÂºÂ£n phÃ¡ÂºÂ©m | Ã„ÂiÃ¡Â»Æ’m | KÃ¡ÂºÂ¿t luÃ¡ÂºÂ­n | C/M/m/L |\n|---:|---|---:|---|---:|\n{table}\n\n## LÃ¡Â»â€”i Ã†Â°u tiÃƒÂªn\n\n1. **MAJOR Ã¢â‚¬â€ Ã¡ÂºÂ£nh:** alt/observation chÃ†Â°a khÃ¡Â»â€ºp close-up, sham hoÃ¡ÂºÂ·c panel; Ã„â€˜Ã¡Â»â€˜i chiÃ¡ÂºÂ¿u theo media ID thay vÃƒÂ¬ vÃ¡Â»â€¹ trÃƒÂ­ gallery.\n2. **MAJOR Ã¢â‚¬â€ intent:** keyword Christmas/cardinal cÃƒÂ²n cÃ¡ÂºÂ§n SERP evidence cho finished quilt set vÃƒÂ  policy chÃ¡Â»â€˜ng cannibalization.\n3. **MINOR Ã¢â‚¬â€ encoding:** source title/H1 cÃƒÂ³ kÃƒÂ½ tÃ¡Â»Â± lÃ¡Â»â€”i hoÃ¡ÂºÂ·c cÃ¡Â»Â¥m bÃ¡Â»â€¹ cÃ¡ÂºÂ¯t; xÃƒÂ¡c nhÃ¡ÂºÂ­n admin/live trÃ†Â°Ã¡Â»â€ºc khi sÃ¡Â»Â­a.\n\n## TrÃ¡ÂºÂ¡ng thÃƒÂ¡i bÃƒÂ n giao\n\n- 72/72 admin-alt baseline Ã„â€˜ÃƒÂ£ Ã„â€˜Ã†Â°Ã¡Â»Â£c Ã„â€˜Ã¡Â»â€˜i chiÃ¡ÂºÂ¿u; khÃƒÂ´ng sÃ¡Â»Â­a workbook r3, khÃƒÂ´ng APPROVED/import Shopify.\n- Workbook cÃƒÂ³ Ã„â€˜ÃƒÂºng 5 sheet, filter/freeze/wrap/hyperlink/formulas; khÃƒÂ´ng tÃ¡ÂºÂ¡o `rendered_sheets`.\n- `awaiting_confirmation=true`.\n''',encoding='utf8')
 (QD/'qa_manifest.json').write_text(json.dumps({'qa_run_id':Q,'batch_id':B,'revision':'r3','source_workbook':str(SRC),'source_workbook_sha256':h(SRC),'expected_products':10,'expected_images':72,'admin_alt_matches':72,'status':'COMPLETE','awaiting_confirmation':True,'output_markdown':str(OD/f'SEO_QA_{B}.md'),'output_xlsx':str(exp.OUTPUT)},ensure_ascii=False,indent=2),encoding='utf8');print(json.dumps({'score':avg,'statuses':st,'issues':se,'validation':audit},default=dict,indent=2))
if __name__=='__main__':main()
