import importlib.util
from pathlib import Path
spec=importlib.util.spec_from_file_location('d',Path(__file__).with_name('download_qa_batch8_images.py')); m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
m.QA=m.ROOT/'seo_runs/jeminise.com/20260906_234129/qa/20260907_180500'; m.IMAGE_DIR=m.QA/'images'
if __name__=='__main__': m.main()
