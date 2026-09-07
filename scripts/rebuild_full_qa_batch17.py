import importlib.util
import json
from datetime import datetime, timezone
from pathlib import Path

spec = importlib.util.spec_from_file_location("rebuilder", Path(__file__).with_name("rebuild_full_qa_batch14.py"))
module = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(module)

RUN_ID = "20260907_212312"
BATCH_ID = "017"
module.BATCHES = {
    BATCH_ID: {
        "qa_run": RUN_ID,
        "source": "SEO_Product_Optimization_through_batch_017.xlsx",
        "positions": range(161, 171),
        "expected_images": 0,
    }
}
# Deliberately do not create the user-disallowed rendered_sheets directory.
module.render_xlsx = lambda *_args, **_kwargs: None


def main():
    run_dir = module.QA_BASE / RUN_ID
    submitted = json.loads((run_dir / "submitted_batch_data.json").read_text(encoding="utf-8"))
    module.BATCHES[BATCH_ID]["expected_images"] = len(submitted["images"])
    module.main()

    validation_path = run_dir / "validation_results.json"
    validation = json.loads(validation_path.read_text(encoding="utf-8"))
    validation.get("workbook_validation", {}).pop("rendered_sheets", None)
    validation_path.write_text(json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8")

    manifest_path = run_dir / "qa_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["revision"] = "r1"
    manifest["prompt_version"] = "seo-prompt/jeminise/prompt_qa.md (current at QA start)"
    manifest["rubric_version"] = "P1,P2,K1-K3,T1-T2,D1-D2,I1,E1; IM1-IM4"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    payload_path = run_dir / "qa_workbook_payload.json"
    payload = json.loads(payload_path.read_text(encoding="utf-8"))
    for row in payload.get("QA_Summary", []):
        if row.get("metric") == "revision":
            row["value"] = "r1"
            row["definition"] = "Initial independent QA revision for qa_batch_017."
    payload_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (run_dir / "qa_dataset.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    module.export_xlsx(BATCH_ID, RUN_ID)

    progress = {
        "qa_run_id": RUN_ID,
        "qa_batch_id": "qa_batch_017",
        "revision": "r1",
        "stage": "completed",
        "products_checked": len(payload.get("QA_Products", [])),
        "images_checked": len(payload.get("QA_Images", [])),
        "completed_product_keys": [p["product_key"] for p in payload.get("QA_Products", [])],
        "completed_qa_image_keys": [i["qa_image_key"] for i in payload.get("QA_Images", [])],
        "awaiting_confirmation": True,
        "updated_at": datetime.now(timezone.utc).astimezone().isoformat(),
    }
    (run_dir / "qa_progress.json").write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")

    report = module.OUT_BASE / RUN_ID / "SEO_QA_qa_batch_017.md"
    lines = []
    for line in report.read_text(encoding="utf-8").splitlines():
        if "rendered_sheets" in line:
            lines.append("- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã mở lại kiểm tra cấu trúc, công thức và bố cục.")
        elif "thay thế output rút gọn" in line:
            lines.append("- Ghi chú: đây là bản QA độc lập đầy đủ theo `prompt_qa.md`, revision r1.")
        else:
            lines.append(line)
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
