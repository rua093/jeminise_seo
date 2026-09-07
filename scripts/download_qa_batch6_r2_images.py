import importlib.util
from pathlib import Path

BASE = Path(__file__).with_name("download_qa_batch6_images.py")
spec = importlib.util.spec_from_file_location("base6img", BASE)
base = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(base)
base.QA_DIR = base.ROOT / "seo_runs/jeminise.com/20260906_234129/qa/20260907_161639"
base.IMAGE_DIR = base.QA_DIR / "images"
base.main()
