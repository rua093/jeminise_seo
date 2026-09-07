from pathlib import Path

import download_qa_batch4_images as base


ROOT = Path(__file__).resolve().parents[1]
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / "20260907_114114"
base.IMAGE_DIR = base.QA_DIR / "images"
base.QA_RUN_ID = "20260907_114114"
base.BATCH_ID = "qa_batch_001_r2"
base.SOURCE_WORKBOOK_NAME = "SEO_Product_Optimization_qa_batch_001_r2.xlsx"


if __name__ == "__main__":
    base.main()
