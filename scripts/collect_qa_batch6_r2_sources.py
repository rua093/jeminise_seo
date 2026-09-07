import importlib.util
from pathlib import Path

BASE = Path(__file__).with_name("collect_qa_batch6_sources.py")
spec = importlib.util.spec_from_file_location("base6", BASE)
base = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(base)

base.module.QA_RUN_ID = "20260907_161639"
base.module.QA_DIR = base.module.ROOT / "seo_runs" / base.module.SHOP / base.module.RUN_ID / "qa" / base.module.QA_RUN_ID
base.module.WORKBOOK = base.module.QA_DIR / "source_snapshot" / "SEO_Product_Optimization_qa_batch_006_r2.xlsx"
base.main()
