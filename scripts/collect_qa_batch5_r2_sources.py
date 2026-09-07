from pathlib import Path
import hashlib
import collect_qa_batch5_sources as base


ROOT = Path(__file__).resolve().parents[1]
RUN = "20260907_152903"
SOURCE = ROOT / "resutls" / "jeminise.com" / "20260906_234129" / "revisions" / "qa_batch_005_r2" / "SEO_Product_Optimization_qa_batch_005_r2.xlsx"
base.module.QA_RUN_ID = RUN
base.module.QA_DIR = ROOT / "seo_runs" / "jeminise.com" / "20260906_234129" / "qa" / RUN
base.module.WORKBOOK = base.module.QA_DIR / "source_snapshot" / SOURCE.name


if __name__ == "__main__":
    base.module.WORKBOOK.parent.mkdir(parents=True, exist_ok=True)
    if not base.module.WORKBOOK.exists():
        base.module.WORKBOOK.write_bytes(SOURCE.read_bytes())
    assert hashlib.sha256(SOURCE.read_bytes()).digest() == hashlib.sha256(base.module.WORKBOOK.read_bytes()).digest()
    base.main()
