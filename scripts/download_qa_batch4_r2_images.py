from pathlib import Path
import download_qa_batch4_images as base

ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_144539"
base.QA_RUN_ID = RUN
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.IMAGE_DIR = base.QA_DIR / "images"
base.BATCH_ID = "qa_batch_004_r2"
base.SOURCE_WORKBOOK_NAME = "SEO_Product_Optimization_qa_batch_004_r2.xlsx"

if __name__ == "__main__":
    base.main()
