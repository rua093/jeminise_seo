import importlib.util
import json
import time
from pathlib import Path


spec = importlib.util.spec_from_file_location("collector17", Path(__file__).with_name("collect_qa_batch17_sources.py"))
collector = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(collector)

ROOT = collector.ROOT
QA = ROOT / "seo_runs/jeminise.com/20260906_234129/qa/20260907_213200"


def main():
    records = json.loads((QA / "live_source_comparison.json").read_text(encoding="utf-8"))
    for record in records:
        if not record.get("fetch_error"):
            continue
        time.sleep(3)
        try:
            record["live"] = collector.module.m.fetch_live(record["inventory"]["product_url"])
            record["fetch_error"] = ""
        except Exception as exc:
            record["fetch_error"] = str(exc)
        print(record["inventory"]["inventory_position"], "ok" if not record["fetch_error"] else record["fetch_error"][:80])
        # Persist each retry so a time-limited command resumes instead of repeating completed requests.
        (QA / "live_source_comparison.json").write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
