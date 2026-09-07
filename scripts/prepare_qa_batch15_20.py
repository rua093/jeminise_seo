import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location("preparer", Path(__file__).with_name("prepare_full_qa_10_13.py"))
module = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(module)
module.CONFIG = {15: ("20260907_214000", 141), 20: ("20260907_214100", 191)}

if __name__ == "__main__":
    module.main()
