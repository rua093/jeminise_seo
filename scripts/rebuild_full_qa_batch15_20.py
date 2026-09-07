import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location("builder", Path(__file__).with_name("rebuild_full_qa_batch18_19.py"))
module = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(module)
module.CONFIG = {
    "015": {"qa_run": "20260907_214000", "source": "SEO_Product_Optimization_through_batch_015.xlsx", "positions": range(141, 151), "expected_images": 0},
    "020": {"qa_run": "20260907_214100", "source": "SEO_Product_Optimization_through_batch_020.xlsx", "positions": range(191, 201), "expected_images": 0},
}

if __name__ == "__main__":
    module.main()
