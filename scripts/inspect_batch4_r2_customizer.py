from pathlib import Path
import inspect_batch7_customizer as base

ROOT = Path(__file__).resolve().parents[1]
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / "20260907_144539"
base.START_POSITION = 31

if __name__ == "__main__":
    base.main()
