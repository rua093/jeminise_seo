from pathlib import Path
import download_qa_batch5_images as base


ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_152903"
base.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.IMAGE_DIR = base.QA_DIR / "images"


if __name__ == "__main__":
    base.main()
