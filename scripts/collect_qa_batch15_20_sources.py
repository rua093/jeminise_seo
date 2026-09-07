import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location("collector", Path(__file__).with_name("collect_qa_batch18_19_sources.py"))
module = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(module)
module.CONFIG = {
    15: ("20260907_214000", 140, "SEO_Product_Optimization_through_batch_015.xlsx"),
    20: ("20260907_214100", 190, "SEO_Product_Optimization_through_batch_020.xlsx"),
}

if __name__ == "__main__":
    module.main()
