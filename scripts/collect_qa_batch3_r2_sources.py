from pathlib import Path
import collect_qa_batch3_sources as base

ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_140123"
base.module.QA_RUN_ID = RUN
base.module.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.module.WORKBOOK = base.module.QA_DIR / "source_snapshot" / "SEO_Product_Optimization_qa_batch_003_r2.xlsx"

if __name__ == "__main__":
    base.main()
