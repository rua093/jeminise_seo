import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location("preparer", Path(__file__).with_name("prepare_full_qa_10_13.py"))
module = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(module)
module.CONFIG = {16: ("20260907_210204", 151)}

if __name__ == "__main__":
    module.main()
