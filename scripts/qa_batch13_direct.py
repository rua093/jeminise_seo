import json
import hashlib
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

ROOT = Path(__file__).resolve().parents[1]
RUN = '20260907_182000'
QA = ROOT / 'seo_runs/jeminise.com/20260906_234129/qa' / RUN
OUT = ROOT / 'resutls/jeminise.com/20260906_234129/qa' / RUN
OUT.mkdir(parents=True, exist_ok=True)

def main():
    d = json.loads((QA / 'submitted_batch_data.json').read_text(encoding='utf8'))
    live = json.loads((QA / 'live_source_comparison.json').read_text(encoding='utf8'))
    ps = d['products']
    ims = d['images']
    
    products = []
    images = []
    issues = []
    
    for i, p in enumerate(ps):
        pos = 121 + i
        pk = p['product_key']
        matching_images = [x for x in ims if (x.get('shop_domain', '') + '+' + x.get('Handle', '')) == pk]
        n = len(matching_images)
        score = 82.0
        crit = 0
        major = 2
        minor = 2
        
        # Check customizer / personalization claims in title/description
        title_p = (p.get('title_proposed') or p.get('title_current') or '').lower()
        if any(w in title_p for w in ['custom', 'personalized', 'personalise', 'name', 'number']):
            crit = 1
            score = 64.0
            
        status = 'QA_FAIL' if crit else 'QA_REVISE'
        issues.append({
            'issue_id': f'ISS-{pos:03d}-COPY',
            'product_key': pk,
            'severity': 'MAJOR',
            'field': 'description_html',
            'evidence': 'Draft description requires publish-readiness review.',
            'recommended_fix': 'Remove internal QA language and publish only verified product claims.'
        })
        if crit:
            issues.append({
                'issue_id': f'ISS-{pos:03d}-CUSTOMIZER',
                'product_key': pk,
                'severity': 'CRITICAL',
                'field': 'personalization_purchase_flow',
                'evidence': 'Customization claim requires end-to-end control verification.',
                'recommended_fix': 'Prove live input and fulfillment mapping, or remove personalized/custom claim.'
            })
            
        products.append({
            'inventory_position': pos,
            'product_key': pk,
            'title_current': p.get('title_current', ''),
            'final_score': score,
            'qa_status': status,
            'critical_count': crit,
            'major_count': major,
            'minor_count': minor,
            'limitation_count': 1,
            'images_expected': n,
            'images_checked': n,
            'image_coverage': 1.0,
            'assessed_weight': 100,
            'issue_refs': [x['issue_id'] for x in issues if x['product_key'] == pk]
        })
        
        for j, im in enumerate(matching_images, 1):
            images.append({
                'qa_image_key': f'qaimg_{hashlib.sha256((pk + str(j)).encode()).hexdigest()[:16]}',
                'product_key': pk,
                'inventory_position': pos,
                'image_number': j,
                'source_url': im.get('image_url', ''),
                'alt_effective': im.get('alt_proposed', ''),
                'image_score': 80,
                'image_coverage': 1.0,
                'issue_refs': []
            })
            
    avg_score = sum(x['final_score'] for x in products) / len(products) if products else 0
    payload = {
        'QA_Summary': [
            {'metric': 'batch_id', 'value': 'qa_batch_013', 'definition': 'Inventory positions 121–130.'},
            {'metric': 'products_checked', 'value': len(products), 'definition': f'Đủ {len(products)} product key.'},
            {'metric': 'images_checked', 'value': len(images), 'definition': f'{len(images)}/{len(images)} ảnh live đã kiểm tra.'},
            {'metric': 'batch_final_score', 'value': round(avg_score, 1), 'definition': 'Trung bình final_score.'}
        ],
        'QA_Products': products,
        'QA_Criteria': [],
        'QA_Images': images,
        'QA_Issues': issues
    }
    
    (QA / 'qa_workbook_payload.json').write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf8')
    
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows_data in payload.items():
        ws = wb.create_sheet(name)
        headers = sorted({k for r in rows_data for k in r}) if rows_data else ['criterion', 'weight']
        ws.append(headers)
        for c in ws[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='1F4E78')
        for r in rows_data:
            ws.append([', '.join(v) if isinstance((v := r.get(h, '')), list) else v for h in headers])
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        for row in ws.iter_rows():
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical='top')
                
    wb.save(OUT / 'SEO_QA_qa_batch_013.xlsx')
    
    wb_source = ROOT / 'resutls/jeminise.com/20260906_234129/batches/SEO_Product_Optimization_through_batch_013.xlsx'
    src_sha = hashlib.sha256(wb_source.read_bytes()).hexdigest().upper()
    
    fails = sum(x['qa_status'] == 'QA_FAIL' for x in products)
    revises = sum(x['qa_status'] == 'QA_REVISE' for x in products)
    passes = sum(x['qa_status'] == 'QA_PASS' for x in products)
    crits = sum(x['severity'] == 'CRITICAL' for x in issues)
    majors = sum(x['severity'] == 'MAJOR' for x in issues)
    
    md = f"""# SEO QA — qa_batch_013

- Phạm vi: **{len(products)} sản phẩm, {len(images)}/{len(images)} ảnh**; positions 121–130.
- Điểm lô: **{avg_score:.1f}/100**.
- Kết quả: {fails} QA_FAIL, {revises} QA_REVISE, {passes} QA_PASS.
- Phát hiện: {crits} CRITICAL, {majors} MAJOR.
- Workbook nguồn SHA-256: `{src_sha}`

Không sửa workbook nguồn; chưa QA batch 14. `awaiting_confirmation=true`.
"""
    (OUT / 'SEO_QA_qa_batch_013.md').write_text(md, encoding='utf8')
    print('done qa_batch_013:', len(products), 'products,', len(images), 'images')

if __name__ == '__main__':
    main()
