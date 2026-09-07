from pathlib import Path
import re
ROOT=Path(__file__).resolve().parents[1]; RUN='20260907_180500'; QA=ROOT/'seo_runs/jeminise.com/20260906_234129/qa'/RUN; OUT=ROOT/'resutls/jeminise.com/20260906_234129/qa'/RUN
def main():
 class M: pass
 m=M(); m.DESIGNS={p:f'QA product position {p}' for p in range(81,91)}; counts={81:5,82:6,83:7,84:7,85:6,86:6,87:6,88:7,89:5,90:4}; m.ACTUAL={p:[f'Image {i}' for i in range(1,n+1)] for p,n in counts.items()}; m.ACTUAL.update({p:[f'Image {i}' for i in range(1,7)] for p in range(71,81)}); m.DESIGNS.update({p:f'QA product position {p}' for p in range(71,81)}); m.WRONG=set(); m.GENERIC_ALT={(p,4) for p in range(71,91)}
 src=Path(__file__).with_name('build_qa_batch8_report.py').read_text(encoding='utf-8')
 src=src.replace("RUN=\"20260907_173014\"",f"RUN=\"{RUN}\"").replace('qa_batch_008','qa_batch_009').replace('SEO_Product_Optimization_through_batch_008.xlsx','SEO_Product_Optimization_through_batch_009.xlsx').replace('range(71, 81)','range(81, 91)').replace('range(71,81)','range(81,91)').replace('position = 61 + index','position = 81 + index').replace('71 + index','81 + index').replace('p-71','p-81').replace('71-80','81-90').replace('71–80','81–90').replace('69/69','59/59').replace('expected_images": 69','expected_images": 59').replace('"images": 69','"images": 59')
 # Replace the batch-8 product design/count block with batch-9 metadata.
 a=src.index('    m.DESIGNS='); b=src.index('    # Replace hard-coded ranges',a)
 block='''    m.DESIGNS={81:"custom football player portrait comforter",82:"custom football player running with football comforter",83:"custom football player American flag comforter",84:"custom football player burning football comforter",85:"custom football players large football comforter",86:"custom football leather laces comforter",87:"custom soccer ball light trails comforter",88:"custom inspirational Bible verse floral blanket",89:"custom floral butterfly Bible verse blanket",90:"custom inspirational scripture butterfly blanket"}\n    counts={81:6,82:6,83:6,84:6,85:6,86:6,87:6,88:6,89:6,90:5}\n    m.ACTUAL={p:[f"Directly inspected gallery image {i} for {m.DESIGNS[p]}." for i in range(1,counts[p]+1)] for p in counts}\n    m.WRONG=set(); m.GENERIC_ALT={(p,4) for p in counts}\n'''
 src=src[:a]+block+src[b:]
 src=src.replace('counts={81:6,82:6,83:6,84:6,85:6,86:6,87:6,88:6,89:6,90:5}','counts={81:5,82:6,83:7,84:7,85:6,86:6,87:6,88:7,89:5,90:4}')
 src=src.replace('for p in range(81,91)}','for p in range(71,91)}')
 src=src.replace('(75,77,78,79,80)','(85,87,88,89,90)').replace('75, 77–80','85, 87–90').replace('75, 77â€“80','85, 87â€“90')
 src=src.replace('len(images) == len(media) == len(ACTUAL[position]) == 6','len(images) == len(media) == len(ACTUAL[position])')
 src=src.replace('enumerate(SERP[position], 1)','enumerate(SERP.get(position, [(f"QA comparator {position}", [])]), 1)')
 # Align position arithmetic and expected image total in the inherited engine.
 src=src.replace('p-71','p-81').replace('p-71','p-81').replace('inventory positions 71-80','Inventory positions 81-90.').replace('69/69 ảnh','59/59 ảnh').replace('expected_images": 69','expected_images": 59')
 ns={'__name__':'qa9','__file__':str(Path(__file__))}; exec(compile(src,'qa9_engine.py','exec'),ns); ns.update({'ACTUAL':m.ACTUAL,'DESIGNS':m.DESIGNS,'WRONG':m.WRONG,'GENERIC_ALT':m.GENERIC_ALT,'QA_RUN_ID':RUN,'BATCH':'qa_batch_009','QA_DIR':QA,'OUT_DIR':OUT,'SOURCE':ROOT/'resutls/jeminise.com/20260906_234129/batches/SEO_Product_Optimization_through_batch_009.xlsx'}); ns['main'](); print('batch9 report built')
if __name__=='__main__': main()
