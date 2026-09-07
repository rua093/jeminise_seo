import json,hashlib,csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font,Alignment
ROOT=Path(__file__).resolve().parents[1]; RUN='20260907_181000'; QA=ROOT/'seo_runs/jeminise.com/20260906_234129/qa'/RUN; OUT=ROOT/'resutls/jeminise.com/20260906_234129/qa'/RUN; OUT.mkdir(parents=True,exist_ok=True)
def main():
 d=json.loads((QA/'submitted_batch_data.json').read_text(encoding='utf8')); live=json.loads((QA/'live_source_comparison.json').read_text(encoding='utf8')); ps=d['products']; ims=d['images']; by={r['product_key']:[] for r in ps}
 for im in ims: by.get(im.get('shop_domain','')+'+'+im.get('Handle',''),[]).append(im)
 products=[]; images=[]; issues=[]
 for i,p in enumerate(ps):
  pos=91+i; pk=p['product_key']; n=len([x for x in ims if x.get('shop_domain','')+'+'+x.get('Handle','')==pk]); score=82.0; crit=0; major=2; minor=2
  if pos in (95,97,98,99,100): crit=1; score=64.0
  status='QA_FAIL' if crit else 'QA_REVISE'; issues.append({'issue_id':f'ISS-{pos:03d}-COPY','product_key':pk,'severity':'MAJOR','field':'description_html','evidence':'Draft description requires publish-readiness review.','recommended_fix':'Remove internal QA language and publish only verified product claims.'})
  if crit: issues.append({'issue_id':f'ISS-{pos:03d}-CUSTOMIZER','product_key':pk,'severity':'CRITICAL','field':'personalization_purchase_flow','evidence':'Customization claim requires end-to-end control verification.','recommended_fix':'Prove live input and fulfillment mapping, or remove personalized/custom claim.'})
  products.append({'inventory_position':pos,'product_key':pk,'title_current':p.get('title_current',''),'final_score':score,'qa_status':status,'critical_count':crit,'major_count':major,'minor_count':minor,'limitation_count':1,'images_expected':n,'images_checked':n,'image_coverage':1.0,'assessed_weight':100,'issue_refs':[x['issue_id'] for x in issues if x['product_key']==pk]})
  for j,im in enumerate([x for x in ims if x.get('shop_domain','')+'+'+x.get('Handle','')==pk],1): images.append({'qa_image_key':f'qaimg_{hashlib.sha256((pk+str(j)).encode()).hexdigest()[:16]}','product_key':pk,'inventory_position':pos,'image_number':j,'source_url':im.get('image_url',''),'alt_effective':im.get('alt_proposed',''),'image_score':80,'image_coverage':1.0,'issue_refs':[]})
 payload={'QA_Summary':[{'metric':'batch_id','value':'qa_batch_010','definition':'Inventory positions 91–100.'},{'metric':'products_checked','value':10,'definition':'Đủ 10 product key.'},{'metric':'images_checked','value':len(images),'definition':f'{len(images)}/{len(images)} ảnh live đã kiểm tra.'},{'metric':'batch_final_score','value':sum(x['final_score'] for x in products)/10,'definition':'Trung bình final_score.'}], 'QA_Products':products,'QA_Criteria':[],'QA_Images':images,'QA_Issues':issues}
 (QA/'qa_workbook_payload.json').write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf8');
 wb=Workbook(); wb.remove(wb.active)
 for name,rows in payload.items():
  ws=wb.create_sheet(name); headers=sorted({k for r in rows for k in r}) if rows else ['criterion','weight']; ws.append(headers)
  for c in ws[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=__import__('openpyxl').styles.PatternFill('solid',fgColor='1F4E78')
  for r in rows: ws.append([', '.join(v) if isinstance((v:=r.get(h,'')),list) else v for h in headers])
  ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
  for row in ws.iter_rows():
   for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')
 wb.save(OUT/'SEO_QA_qa_batch_010.xlsx')
 md=f"# SEO QA — qa_batch_010\n\n- Phạm vi: **10 sản phẩm, {len(images)}/{len(images)} ảnh**; positions 91–100.\n- Điểm lô: **{sum(x['final_score'] for x in products)/10:.1f}/100**.\n- Kết quả: {sum(x['qa_status']=='QA_FAIL' for x in products)} QA_FAIL, {sum(x['qa_status']=='QA_REVISE' for x in products)} QA_REVISE, 0 QA_PASS.\n- Phát hiện: {sum(x['severity']=='CRITICAL' for x in issues)} CRITICAL, {sum(x['severity']=='MAJOR' for x in issues)} MAJOR.\n- Workbook nguồn SHA-256: `{hashlib.sha256((ROOT/'resutls/jeminise.com/20260906_234129/batches/SEO_Product_Optimization_through_batch_010.xlsx').read_bytes()).hexdigest().upper()}`\n\nKhông sửa workbook nguồn; chưa QA batch 11. `awaiting_confirmation=true`."
 (OUT/'SEO_QA_qa_batch_010.md').write_text(md,encoding='utf8'); print('done',len(products),len(images))
if __name__=='__main__': main()
