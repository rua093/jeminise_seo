from pathlib import Path
def main():
 src=Path(__file__).with_name('finalize_qa_batch8.py').read_text(encoding='utf-8').replace('20260907_173014','20260907_180500').replace('qa_batch_008','qa_batch_009').replace('batch_008.xlsx','batch_009.xlsx').replace('71<=int(r["inventory_position"])<=80','81<=int(r["inventory_position"])<=90').replace('71<=int(row["inventory_position"])<=80','81<=int(row["inventory_position"])<=90').replace('image_count_69','image_count_59').replace('len(ims)==69','len(ims)==59').replace('==69','==59')
 exec(compile(src,'finalize9_engine.py','exec'),{'__name__':'__main__','__file__':str(Path(__file__))})
if __name__=='__main__': main()
