from pathlib import Path
from datetime import datetime
import json, hashlib, html, re, requests

ROOT=Path(__file__).resolve().parents[1]
CONFIG={14:('20260907_204400',131)}
HEAD={'User-Agent':'Mozilla/5.0 independent SEO QA/1.0'}

def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest().upper()

def main():
 for batch,(run,start) in CONFIG.items():
  qa=ROOT/'seo_runs/jeminise.com/20260906_234129/qa'/run
  source=ROOT/'resutls/jeminise.com/20260906_234129/batches'/f'SEO_Product_Optimization_through_batch_{batch:03d}.xlsx'
  snap=qa/'source_snapshot'; snap.mkdir(parents=True,exist_ok=True); frozen=snap/source.name
  if not frozen.exists() or sha(frozen)!=sha(source): frozen.write_bytes(source.read_bytes())
  live=json.loads((qa/'live_source_comparison.json').read_text(encoding='utf8')); imgdir=qa/'images'; imgdir.mkdir(exist_ok=True)
  image_manifest=[]; custom=[]
  for idx,rec in enumerate(live):
   pos=start+idx; inv=rec['inventory']; url=inv['product_url']; page=requests.get(url,headers=HEAD,timeout=45).text
   match=re.search(r'data-amzcustom-root[^>]*data-config="([^"]+)"',page); cfg=json.loads(html.unescape(match.group(1))) if match else {}; text=json.dumps(cfg,ensure_ascii=False)
   custom.append({'inventory_position':pos,'product_key':inv['product_key'],'url':url,'checked_at':datetime.now().astimezone().isoformat(),'customizer_root_present':bool(match),'customize_button_present':'amzcustom-open' in page,'upload_endpoint_present':'/apps/amazon-customizer/upload' in page,'surface_count':len(cfg.get('surfaces',[])),'name_mentions':len(re.findall('name',text,re.I)),'number_mentions':len(re.findall('number',text,re.I)),'text_or_input_mentions':len(re.findall('text|input',text,re.I)),'config_excerpt':text[:2000]})
   media=rec.get('live',{}).get('product_js',{}).get('media',[])
   for num,item in enumerate(media,1):
    src=item['src']; ext=Path(src.split('?',1)[0]).suffix or '.jpg'; target=imgdir/f'{pos:03d}_{num:02d}{ext}'
    if not target.exists():
     response=requests.get(src,headers=HEAD,timeout=60); response.raise_for_status(); target.write_bytes(response.content)
    image_manifest.append({'inventory_position':pos,'product_key':inv['product_key'],'image_number':num,'media_id':str(item.get('id','')),'source_url':src,'local_path':str(target),'sha256':sha(target),'bytes':target.stat().st_size,'downloaded_at':datetime.now().astimezone().isoformat()})
  (qa/'customizer_audit.json').write_text(json.dumps(custom,ensure_ascii=False,indent=2),encoding='utf8')
  (qa/'image_download_manifest.json').write_text(json.dumps(image_manifest,ensure_ascii=False,indent=2),encoding='utf8')
  print(batch,len(live),len(image_manifest),sha(source))
if __name__=='__main__': main()
