from pathlib import Path

import collect_qa_batch1_sources as base


ROOT = Path(__file__).resolve().parents[1]
QA_RUN_ID = "20260907_114114"
QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / QA_RUN_ID

base.QA_RUN_ID = QA_RUN_ID
base.QA_DIR = QA_DIR
base.WORKBOOK = QA_DIR / "source_snapshot" / "SEO_Product_Optimization_qa_batch_001_r2.xlsx"


if __name__ == "__main__":
    base.main()
