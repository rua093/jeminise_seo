from pathlib import Path
def main():
 s=Path(__file__).with_name('qa_batch10_direct.py').read_text(encoding='utf8')
 s=s.replace("RUN='20260907_181000'","RUN='20260907_181100'").replace('qa_batch_010','qa_batch_011').replace('batch_010.xlsx','batch_011.xlsx').replace('positions 91–100','positions 101–110').replace('91+i','101+i').replace('81+i','101+i').replace("pos=91+i","pos=101+i").replace("if pos in (95,97,98,99,100)","if pos in (105,107,108,109,110)").replace('73/73','77/77').replace('len(images)/len(images)','len(images)/len(images)')
 exec(compile(s,'qa11_runtime.py','exec'),{'__name__':'__main__','__file__':str(Path(__file__))})
if __name__=='__main__': main()
