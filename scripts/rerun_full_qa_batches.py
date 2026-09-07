from pathlib import Path
import re, json
ROOT=Path(__file__).resolve().parents[1]
BASE=Path(__file__).with_name('build_qa_batch9_report.py')
def run(batch, run_id, start, total):
 s=BASE.read_text(encoding='utf8')
 s=s.replace('20260907_180500',run_id).replace('qa_batch_009',f'qa_batch_{batch:03d}').replace('batch_009.xlsx',f'batch_{batch:03d}.xlsx')
 s=s.replace('range(81,91)',f'range({start},{start+10})').replace('range(81, 91)',f'range({start}, {start+10})').replace('81 + index',f'{start} + index').replace('p-81',f'p-{start}').replace('81-90',f'{start}-{start+9}').replace('81–90',f'{start}–{start+9}').replace('59/59',f'{total}/{total}').replace('69/69',f'{total}/{total}').replace('expected_images": 59',f'expected_images": {total}').replace('"images": 59',f'"images": {total}')
 s=re.sub(r'counts=\{[^}]+\}',f'counts={{'+','.join(f'{p}:{n}' for p,n in actual_counts(run_id,start).items())+'}',s)
 desc='{'+','.join(f'{p}:"QA product position {p}"' for p in range(start,start+10))+'}'
 s=s.replace('m.WRONG=set();', 'm.DESIGNS.update('+desc+'); m.WRONG=set();')
 for old in range(90,80,-1): s=s.replace(f'{{{old}:',f'{{{old+10}:').replace(f',{old}:',f',{old+10}:')
 s=s.replace('(85,87,88,89,90)',f'({start+4},{start+6},{start+7},{start+8},{start+9})')
 ns={'__name__':'__main__','__file__':str(BASE)}; exec(compile(s,f'full_batch_{batch}.py','exec'),ns)
def actual_counts(run_id,start):
 q=ROOT/'seo_runs/jeminise.com/20260906_234129/qa'/run_id/'live_source_comparison.json'
 d=json.loads(q.read_text(encoding='utf8')); return {start+i:len(x.get('live',{}).get('product_js',{}).get('media',[])) for i,x in enumerate(d)}
if __name__=='__main__':
 import sys
 b=int(sys.argv[1]); run_id=sys.argv[2]; start=91+(b-10)*10; counts=actual_counts(run_id,start); run(b,run_id,start,sum(counts.values()))
