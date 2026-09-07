import csv, json, importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('c', Path(__file__).with_name('collect_qa_batch2_sources.py'))
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

ROOT = m.ROOT
RUN_ID = '20260907_204400'
QA = ROOT / f'seo_runs/jeminise.com/20260906_234129/qa/{RUN_ID}'
WB = ROOT / 'resutls/jeminise.com/20260906_234129/batches/SEO_Product_Optimization_through_batch_014.xlsx'

def main():
    QA.mkdir(parents=True, exist_ok=True)
    (QA / 'source_snapshot').mkdir(exist_ok=True)
    (QA / 'source_snapshot' / WB.name).write_bytes(WB.read_bytes())
    
    rows = list(csv.DictReader((ROOT / 'seo_runs/jeminise.com/20260906_234129/inventory.csv').open(encoding='utf-8-sig')))[130:140]
    keys = {r['product_key'] for r in rows}
    handles = {r['Handle'] for r in rows}
    urls = {r['product_url'] for r in rows}
    
    book = m.load_workbook(WB, read_only=True, data_only=True)
    sub = {
        'products': [r for r in m.rows(book, 'SEO_Products') if r['product_key'] in keys],
        'images': [r for r in m.rows(book, 'Image_Audit') if r['Handle'] in handles],
        'product_evidence': [r for r in m.rows(book, 'Product_Evidence') if r['product_url'] in urls],
        'keyword_map': [r for r in m.rows(book, 'Keyword_Map') if r['product_key'] in keys],
        'buyer_search_research': [r for r in m.rows(book, 'Buyer_Search_Research') if r['product_key'] in keys]
    }
    (QA / 'submitted_batch_data.json').write_text(json.dumps(sub, ensure_ascii=False, indent=2), encoding='utf8')
    
    live = []
    for r in rows:
        try:
            x = m.fetch_live(r['product_url'])
            live.append({'inventory': r, 'live': x, 'fetch_error': ''})
        except Exception as e:
            live.append({'inventory': r, 'live': {}, 'fetch_error': str(e)})
            
    (QA / 'live_source_comparison.json').write_text(json.dumps(live, ensure_ascii=False, indent=2), encoding='utf8')
    print('batch14 sources collected: products =', len(sub['products']), ', images =', len(sub['images']))

if __name__ == '__main__':
    main()
