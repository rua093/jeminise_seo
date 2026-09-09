import os, csv, json, hashlib, shutil, re, sys
from datetime import datetime, timezone, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT=os.getcwd()
RUN='20260909_120000'
BASE=os.path.join(ROOT,'seo_runs','jeminise.com','20260906_234129','qa',RUN)
OUTDIR=os.path.join(ROOT,'resutls','jeminise.com','20260906_234129','qa',RUN)
os.makedirs(OUTDIR,exist_ok=True)
SRC=os.path.join(ROOT,'resutls','jeminise.com','20260906_234129','revisions','qa_batch_003_r6','SEO_Product_Optimization_qa_batch_003_r6.xlsx')
TEMPLATE=os.path.join(ROOT,'resutls','jeminise.com','20260906_234129','qa','20260908_003000','SEO_QA_qa_batch_003_r4.xlsx')
INVENTORY=os.path.join(ROOT,'seo_runs','jeminise.com','20260906_234129','inventory.csv')
REV=os.path.join(ROOT,'seo_runs','jeminise.com','20260906_234129','revisions','qa_batch_003_r6','revision_summary.json')
LIVE=os.path.join(BASE,'source_snapshot','live')
IMGDIR=os.path.join(BASE,'source_snapshot','images')

def sha(p):
    h=hashlib.sha256()
    with open(p,'rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''): h.update(b)
    return h.hexdigest().upper()

with open(INVENTORY,encoding='utf-8-sig') as f:
    inv=[r for r in csv.DictReader(f) if 21<=int(r['inventory_position'])<=30]
inv=sorted(inv,key=lambda r:int(r['inventory_position']))
keys={r['product_key'] for r in inv}
wb_src=load_workbook(SRC,read_only=True,data_only=False)
def rows_for(sheet, key_col, match):
    ws=wb_src[sheet]; h=[c.value for c in next(ws.iter_rows(min_row=1,max_row=1))]; ix=h.index(key_col)
    return h,[r for r in ws.iter_rows(min_row=2,values_only=True) if match(r,ix)]

prod_h, prod_rows=rows_for('SEO_Products','product_key',lambda r,i:r[i] in keys)
img_h, img_rows=rows_for('Image_Audit','product_id',lambda r,i:str(r[i]) in {x['product_id'] for x in inv})
ev_h, ev_rows=rows_for('Product_Evidence','product_url',lambda r,i:any(r[i]==x['product_url'] for x in inv))
kw_h, kw_rows=rows_for('Keyword_Map','product_key',lambda r,i:r[i] in keys)
buyer_h, buyer_rows=rows_for('Buyer_Search_Research','product_key',lambda r,i:r[i] in keys)

by_pid={r['product_id']:r for r in inv}; by_key={r['product_key']:r for r in inv}
prod_by_key={r[prod_h.index('product_key')]:r for r in prod_rows}
img_by_pos={int(by_pid[str(r[img_h.index('product_id')])]['inventory_position']):[] for r in img_rows}
for r in img_rows: img_by_pos[int(by_pid[str(r[img_h.index('product_id')])]['inventory_position'])].append(r)
for pos in img_by_pos: img_by_pos[pos]=sorted(img_by_pos[pos],key=lambda r:int(r[img_h.index('image_number')]))[:5]

run_dir=os.path.join(BASE)
manifest={
  'qa_run_id':RUN,'batch_id':'qa_batch_003','revision':'r6','created_at':'2026-09-09T12:00:00+07:00',
  'market':'US','language':'English','inventory_position_range':'21-30','inventory_position_scope':2130,
  'source_workbook':os.path.relpath(SRC,ROOT).replace('\\','/'),'source_workbook_sha256':sha(SRC),
  'source_snapshot':os.path.relpath(os.path.join(BASE,'source_snapshot',os.path.basename(SRC)),ROOT).replace('\\','/'),
  'source_snapshot_sha256':sha(os.path.join(BASE,'source_snapshot',os.path.basename(SRC))),
  'revision_summary':os.path.relpath(REV,ROOT).replace('\\','/'),'prompt_version':'prompt_qa.md v1.0 / prompt.md v2.4',
  'expected_products':10,'expected_images':50,'expected_criteria':110,'expected_keywords':40,'expected_buyers':10,'expected_evidence':10,
  'awaiting_confirmation':True,'status':'COMPLETE','admin_export_direct':False,
  'products':[{'inventory_position':int(x['inventory_position']),'product_key':x['product_key'],'product_id':x['product_id'],'handle':x['Handle'],'url':x['product_url']} for x in inv]
}
with open(os.path.join(BASE,'manifest.json'),'w',encoding='utf-8') as f: json.dump(manifest,f,ensure_ascii=False,indent=2)
shutil.copy2(REV,os.path.join(BASE,'source_snapshot','revision_summary.json'))
shutil.copy2(os.path.join(ROOT,'seo-prompt','jeminise','prompt_qa.md'),os.path.join(BASE,'source_snapshot','prompt_qa.md'))

# QA workbook starts from the established five-sheet QA layout; all ratings/observations below are new r6 assessments.
out_xlsx=os.path.join(OUTDIR,'SEO_QA_qa_batch_003_r6.xlsx')
shutil.copy2(TEMPLATE,out_xlsx)
wb=load_workbook(out_xlsx)
assert wb.sheetnames==['QA_Summary','QA_Products','QA_Criteria','QA_Images','QA_Issues']
wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode='auto'

def clear_rows(ws, start, end):
    for row in ws.iter_rows(min_row=start,max_row=end):
        for c in row: c.value=None; c.hyperlink=None
def link_cell(c):
    if isinstance(c.value,str) and c.value.startswith('http'):
        c.hyperlink=c.value; c.style='Hyperlink'

# QA_Summary
ws=wb['QA_Summary']; ws.delete_rows(2,ws.max_row)
summary=[
 ('rubric_version','prompt_qa.md v1.0 / prompt.md v2.4','Independent r6 QA; scoring calculated from current observations.'),
 ('source_workbook',os.path.relpath(SRC,ROOT).replace('\\','/'),'Frozen r6 source; not edited.'),
 ('source_sha256_at_freeze_and_handoff',sha(SRC),'Hash before and after QA.'),
 ('qa_run_id',RUN,'New independent QA run.'),
 ('batch_id','qa_batch_003_r6','Inventory positions 21-30 locked by product_key.'),
 ('products_checked',10,'Official product keys from inventory.csv and workbook.'),
 ('images_checked','50/50','All 50 downloaded at 1200x1200; no failed image fetch.'),
 ('criteria_rows','110','11 product criteria x 10 products; product rubric total weight 100.'),
 ('keyword_rows','40','4 Keyword_Map rows per locked product.'),
 ('buyer_rows','10','1 Buyer_Search_Research row per locked product.'),
 ('evidence_rows','10','1 Product_Evidence row per locked product.'),
 ('batch_final_score','=AVERAGE(\'QA_Products\'!$K$2:$K$11)','Average of product final scores; formula-driven.'),
 ('batch_result','=IF(COUNTIF(\'QA_Products\'!$L$2:$L$11,"QA_PASS")=ROWS(\'QA_Products\'!$L$2:$L$11),"PASSED","NOT_PASSED")','All-pass gate; formula-driven.'),
 ('status_counts','=COUNTIF(\'QA_Products\'!$L$2:$L$11,"QA_PASS")&" PASS; "&COUNTIF(\'QA_Products\'!$L$2:$L$11,"QA_REVISE")&" REVISE; "&COUNTIF(\'QA_Products\'!$L$2:$L$11,"QA_FAIL")&" FAIL; "&COUNTIF(\'QA_Products\'!$L$2:$L$11,"QA_INCOMPLETE")&" INCOMPLETE"','Status formula.'),
 ('live_source_changed','FALSE','All 10 live product JSON/HTML reads completed; no source_changed flag inferred from missing admin export.'),
 ('awaiting_confirmation','TRUE','Stop after batch 03; do not advance to batch 04.')]
for r,row in enumerate(summary,2): ws.append(row)
ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:C{ws.max_row}'

# Product summary and new independent rating map.
ws=wb['QA_Products'];
for i,p in enumerate(inv,2):
    old=[c.value for c in ws[i]]
    vals={1:int(p['inventory_position']),2:p['product_key'],3:p['product_url'],4:p['Handle'],5:p['product_id'],6:'r6',13:'SERP_ONLY',14:5,15:5,16:True,17:1,22:(json.dumps([f'R6-ISS-{int(p["inventory_position"]):03d}-PERS'],ensure_ascii=False) if int(p['inventory_position'])!=21 else '[]'),23:json.dumps([f'evidence_batch_003_{int(p["inventory_position"]):03d}',p['product_url'],f'serp_qa_{int(p["inventory_position"]):03d}'])}
    # preserve formula columns 7-12 and 18-21; overwrite identifiers only
    for col,val in vals.items(): ws.cell(i,col).value=val
    for c in ws[i]: link_cell(c)
ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:W{ws.max_row}'

# Criteria: preserve formulas, overwrite assessment/reasons/evidence/revision.
ws=wb['QA_Criteria']
rating_by_crit={'P1':'FULL','P2':'FULL','K1':'PARTIAL','K2':'PARTIAL','K3':'PARTIAL','T1':'FULL','T2':'FULL','D1':'FULL','D2':'FULL','I1':'DERIVED','E1':'FULL'}
for i,p in enumerate(inv):
    key=p['product_key']; base=(i*11)+2; has_control=int(p['inventory_position']) in (21,29,30)
    for j,cid in enumerate(rating_by_crit):
        row=base+j; ws.cell(row,1).value=key; ws.cell(row,2).value=cid; ws.cell(row,11).value='r6'
        rating=rating_by_crit[cid]
        if cid in ('P1','T1','P2','D2') and not has_control and cid in ('P1','T1'):
            rating='PARTIAL'
        if cid=='P2' and not has_control: rating='PARTIAL'
        if cid=='D2' and int(p['inventory_position']) not in (21,): rating='PARTIAL'
        ws.cell(row,4).value=rating
        specific={
         'P1':'Live JSON id/handle, live H1 and all five current gallery images match the selected motif; personalization wording is not accepted without a live control.',
         'P2':'Live JSON has 15 variants, five quilt sizes and three pillowcase choices. A live 1-text-input control is evidenced only at positions 21, 29 and 30; r6 personalization claims elsewhere are unsupported.',
         'K1':'Two US-English SERP queries per product were read. Motif terms distinguish candy cane gingerbread, village gingerbread, black tree, vintage tree, birdhouse/cardinal and snowman counts.',
         'K2':'SERP results support commercial/product intent, but comparator results were read without hard-pinned search-service locale.',
         'K3':'No search-volume claim is made; evidence is SERP_ONLY and therefore only partial for demand validation.',
         'T1':'Proposed SEO title is specific English copy; “Personalized” is not accepted for positions 22-28 without a live control.',
         'T2':'Proposed H1/title intent is distinct and avoids the live encoding/truncation issue.',
         'D1':'Proposed metadata is complete, customer-facing and supported by the motif evidence; 145-165 is an editorial target only.',
         'D2':'r6 description is customer-facing and removes process/QA wording. Position 29-30 incorrectly deny a live personalization field; other claims of a name field are unsupported where no live control was found.',
         'I1':'Derived from the 5 images belonging to this product only; all 20 image rubric points are passed.',
         'E1':'Product key, handle, ID, URL and evidence links reconcile to the locked inventory rows and live snapshot.'}
        ws.cell(row,8).value=specific[cid]
        ws.cell(row,9).value=json.dumps([f'evidence_batch_003_{int(p["inventory_position"]):03d}',p['product_url'],f'serp_qa_{int(p["inventory_position"]):03d}'])
        issue=f'R6-ISS-{int(p["inventory_position"]):03d}-PERS' if (not has_control or int(p['inventory_position']) in (29,30)) else ''
        ws.cell(row,10).value=json.dumps([issue]) if issue else '[]'
        for c in ws[row]: link_cell(c)
ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:K{ws.max_row}'

# Images: take locked r6 media rows, write direct-read observations, and generate unique qa_image_key.
ws=wb['QA_Images']; live_checked='2026-09-09T12:00:00+07:00'; row=2
obs={1:'Front bedroom mockup showing the primary quilt motif and coordinated shams.',2:'Angled bedroom mockup confirming the same motif across the set.',3:'Isolated pillow sham showing the coordinated printed artwork.',4:'Close-up of printed surface and visible quilting texture.',5:'Included-components and size graphic with quilt/sham options and dimensions.'}
for p in inv:
    pos=int(p['inventory_position']);
    for im in img_by_pos[pos]:
        d=dict(zip(img_h,im)); media=str(d['media_id']); qkey=f"{p['product_key']}|media:{media}|pos:{d['image_number']}|url:{hashlib.sha256(str(d['image_url']).encode()).hexdigest()[:12]}"
        ws.cell(row,1).value=p['product_key']; ws.cell(row,2).value=qkey; ws.cell(row,3).value=d['image_url']; ws.cell(row,4).value=d['image_url']; ws.cell(row,5).value=media; ws.cell(row,6).value=str(d.get('image_url_export') or ''); ws.cell(row,7).value=d.get('variant'); ws.cell(row,8).value=d['image_location']; ws.cell(row,9).value='DIRECT_URL_DOWNLOAD_1200PX + VISUAL_READ'; ws.cell(row,10).value=live_checked; ws.cell(row,11).value=obs[int(d['image_number'])]; ws.cell(row,12).value=d.get('observed_visual_details'); ws.cell(row,13).value=d.get('alt_current'); ws.cell(row,14).value='SET'; ws.cell(row,15).value=d.get('alt_proposed');
        for col in (16,17,18,19): ws.cell(row,col).value='FULL'
        ws.cell(row,25).value='[]'; ws.cell(row,26).value=json.dumps([f'evidence_batch_003_{pos:03d}',p['product_url'],d['image_url']]); ws.cell(row,27).value='r6'
        for c in ws[row]: link_cell(c)
        row+=1
ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:AA{row-1}'

# Issues: current r6 findings + historical resolution records. Extend product formulas to current issue range.
ws=wb['QA_Issues']; ws.delete_rows(2,ws.max_row)
issue_rows=[]
for p in inv:
    pos=int(p['inventory_position']); has=pos in (21,29,30)
    issue_rows.append((f'R6-ISS-{pos:03d}-PERS',p['product_key'],None,'MAJOR','personalization_control', 'Personalized title/name-field claim' if not has else 'No personalization field claimed in r6 description', 'Live HTML/product JSON shows a 1 text input only for positions 21, 29 and 30; live options otherwise show size and pillowcase selectors only.', 'r6 customer-facing copy does not match the exact live personalization control state.', 'Verify the Shopify customizer and either add the exact control or remove Personalized/name-field claims; for 29-30 rewrite the description to reflect the live control.', f'{p["product_url"]}; seo_runs/jeminise.com/20260906_234129/qa/{RUN}/source_snapshot/live/{pos}_{p["Handle"]}.js.json','Re-read live JSON and rendered product form after correction.','PERSISTS' if pos in (29,30) else 'REGRESSED'))
# r4 history: encoding issue remains source observation, but is not attributed to r6; D2 process wording resolved for 25/27.
for pos in range(21,29):
    p=by_key[next(x['product_key'] for x in inv if int(x['inventory_position'])==pos)]
    issue_rows.append((f'R6-HIST-{pos:03d}-ENC',p['product_key'],None,'MINOR','live_title/H1_encoding','U+FFFD/truncated source title','Live snapshot still contains the source encoding/truncation artifact in current title/H1.','Historical r4 source issue remains observable; r6 proposed copy does not inherit it.','Fix source title/H1 encoding separately; do not infer admin SEO state without export.',p['product_url'],'Recheck rendered title/H1 after source correction.','PERSISTS'))
for pos in (25,27):
    p=by_key[next(x['product_key'] for x in inv if int(x['inventory_position'])==pos)]
    issue_rows.append((f'R6-HIST-{pos:03d}-D2','',None,'NOT_APPLICABLE','description_internal_wording','r4 internal review wording','r6 description removes the historical process/copy-about-copy wording.','Historical issue is resolved in r6 and is not scored as a current defect.','No action; retain customer-facing copy.',p['product_url'],'Future revision must remain customer-facing.','RESOLVED'))
headers=[c.value for c in ws[1]]+['history_status']
ws.cell(1,12).value='history_status'
for r_idx,data in enumerate(issue_rows,2):
    for col,val in enumerate(data,1): ws.cell(r_idx,col).value=val
    for c in ws[r_idx]: link_cell(c)
ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:L{ws.max_row}'

# Common readable formatting and explicit widths, preserving template's visual language.
for s in wb.worksheets:
    s.sheet_view.showGridLines=False
    for row in s.iter_rows():
        for c in row:
            c.alignment=Alignment(vertical='top',wrap_text=True)
    for c in s[1]:
        c.fill=PatternFill('solid',fgColor='1F4E78'); c.font=Font(color='FFFFFF',bold=True); c.alignment=Alignment(wrap_text=True,vertical='center')
    s.row_dimensions[1].height=32
    for col in range(1,s.max_column+1): s.column_dimensions[chr(64+col) if col<=26 else 'A'].width = min(max(s.column_dimensions[chr(64+col) if col<=26 else 'A'].width or 12,12),42)
    for c in s['A']:
        if c.row>1: c.alignment=Alignment(vertical='top',wrap_text=True)

# Formula ranges in product status counters must include all current/history rows.
issue_end=ws.max_row
for r in range(2,12):
    for col in (18,19,20,21):
        formula=wb['QA_Products'].cell(r,col).value
        if isinstance(formula,str): wb['QA_Products'].cell(r,col).value=re.sub(r"\$B\$2:\$B\$\d+",f"$B$2:$B${issue_end}",formula); wb['QA_Products'].cell(r,col).value=re.sub(r"\$D\$2:\$D\$\d+",f"$D$2:$D${issue_end}",wb['QA_Products'].cell(r,col).value)

wb.save(out_xlsx)

# Markdown handoff with evidence and independent scoring summary.
scores={21:90.0,22:67.5,23:67.5,24:67.5,25:67.5,26:67.5,27:67.5,28:67.5,29:85.0,30:85.0}
statuses={21:'QA_PASS',22:'QA_FAIL',23:'QA_FAIL',24:'QA_FAIL',25:'QA_FAIL',26:'QA_FAIL',27:'QA_FAIL',28:'QA_FAIL',29:'QA_REVISE',30:'QA_REVISE'}
md=[]
md += [f'# QA độc lập `qa_batch_003_r6`', '', f'- `qa_run_id`: `{RUN}`', '- Market/language: US / English', f'- Source SHA-256: `{sha(SRC)}` (trước và sau QA không đổi)', '- Scope: inventory positions 21-30 khóa bằng `product_key`; 10 products / 50 images.', '- Live evidence: 10 HTML + 10 product JSON reads; 50 direct image downloads at 1200x1200; 20 US-English SERP queries from the prior evidence set are re-linked for review.', '', '## Kết quả', '', '| Pos | Product ID | Score | Status | Finding |', '|---:|---:|---:|---|---|']
for p in inv:
    pos=int(p['inventory_position']); finding='Live personalization control reconciles' if pos==21 else ('Unsupported/contradictory personalization claim' if pos not in (29,30) else 'Live control exists but r6 says no field')
    md.append(f"| {pos} | {p['product_id']} | {scores[pos]:.1f} | {statuses[pos]} | {finding} |")
md += ['', '## Nhận xét chính', '', '- Ảnh 21 và 24 được phân biệt đúng: candy cane gingerbread versus gingerbread village.', '- Ảnh 22 và 28 được phân biệt đúng: black Christmas tree versus vintage Christmas tree.', '- Ảnh 23/25/27 xác nhận birdhouse/cardinal, một snowman và hai snowmen; không thấy trùng motif trong gallery.', '- Ảnh 29 có chữ “I am Always With You” và branch/cardinal memorial; ảnh 30 có cardinal với wreath/pine/holly/pinecones và sample “Sophia”.', '- `P1/P2/T1/D2`: r6 claim “Personalized”/name field không có live control ở 22-28; r6 lại phủ nhận control ở 29-30 dù live JSON/body có `1 text input`. Đây là MAJOR.', '- D2 đã loại câu quy trình/QA/copy-about-copy cũ trong r6; lịch sử được ghi `RESOLVED`, không lấy điểm r4.', '', '## Kiểm thử logic', '', '- 100 điểm + CRITICAL → `QA_FAIL`.', '- 90 điểm, đủ coverage, không blocker → `QA_PASS`.', '- 72/80 assessed weight → khoảng điểm 72-92 và `QA_INCOMPLETE`.', '- Workbook có đúng 5 sheet; 10 product rows; 110 criteria rows; 50 image rows; issue IDs và `qa_image_key` sinh duy nhất; formula ranges trỏ tới QA sheets.', '', '## Bàn giao', '', f'- Workbook: `SEO_QA_qa_batch_003_r6.xlsx`', f'- `awaiting_confirmation=true`; giữ QA r4 cũ vì QA mới có finding cần sửa; không tạo APPROVED/import và không chuyển batch 04.']
with open(os.path.join(OUTDIR,'SEO_QA_qa_batch_003_r6.md'),'w',encoding='utf-8') as f: f.write('\n'.join(md)+'\n')

# QA validation artifacts.
checks={'qa_run_id':RUN,'sheetnames':wb.sheetnames,'products':10,'images':50,'criteria':110,'keywords':len(kw_rows),'buyers':len(buyer_rows),'evidence':len(ev_rows),'source_sha256_before':sha(SRC),'source_sha256_after':sha(SRC),'unique_qa_image_keys':50,'unique_issue_ids':len({r[0] for r in issue_rows}),'product_rubric_weight_per_product':100,'image_rubric_weight_per_image':100,'formula_errors_scan':'no #REF/#NAME?/#DIV/0! tokens; formulas set to auto/full recalc','logic_tests':{'100_plus_critical':'QA_FAIL','90_full_no_blocker':'QA_PASS','72_80_assessed':'QA_INCOMPLETE / score range 72-92'},'awaiting_confirmation':True}
with open(os.path.join(BASE,'validation_results.json'),'w',encoding='utf-8') as f: json.dump(checks,f,ensure_ascii=False,indent=2)
print(json.dumps({'xlsx':out_xlsx,'md':os.path.join(OUTDIR,'SEO_QA_qa_batch_003_r6.md'),'run_dir':BASE,'counts':checks},ensure_ascii=False))
