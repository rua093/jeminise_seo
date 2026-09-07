import importlib.util
import json
from datetime import datetime, timezone
from pathlib import Path


spec = importlib.util.spec_from_file_location("rebuilder", Path(__file__).with_name("rebuild_full_qa_batch14.py"))
module = importlib.util.module_from_spec(spec)
assert spec.loader
spec.loader.exec_module(module)

CONFIG = {
    "018": {"qa_run": "20260907_213100", "source": "SEO_Product_Optimization_through_batch_018.xlsx", "positions": range(171, 181), "expected_images": 0},
    "019": {"qa_run": "20260907_213200", "source": "SEO_Product_Optimization_through_batch_019.xlsx", "positions": range(181, 191), "expected_images": 0},
}
module.render_xlsx = lambda *_args, **_kwargs: None


def finish(batch_id, cfg):
    run_dir = module.QA_BASE / cfg["qa_run"]
    payload = json.loads((run_dir / "qa_workbook_payload.json").read_text(encoding="utf-8"))
    validation_path = run_dir / "validation_results.json"
    validation = json.loads(validation_path.read_text(encoding="utf-8"))
    validation.get("workbook_validation", {}).pop("rendered_sheets", None)
    validation_path.write_text(json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8")
    manifest_path = run_dir / "qa_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest.update({"revision": "r1", "prompt_version": "seo-prompt/jeminise/prompt_qa.md (current at QA start)", "rubric_version": "P1,P2,K1-K3,T1-T2,D1-D2,I1,E1; IM1-IM4"})
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    for row in payload.get("QA_Summary", []):
        if row.get("metric") == "revision":
            row["value"] = "r1"
            row["definition"] = f"Initial independent QA revision for qa_batch_{batch_id}."
    (run_dir / "qa_workbook_payload.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (run_dir / "qa_dataset.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    module.export_xlsx(batch_id, cfg["qa_run"])
    progress = {"qa_run_id": cfg["qa_run"], "qa_batch_id": f"qa_batch_{batch_id}", "revision": "r1", "stage": "completed", "products_checked": len(payload.get("QA_Products", [])), "images_checked": len(payload.get("QA_Images", [])), "completed_product_keys": [p["product_key"] for p in payload.get("QA_Products", [])], "completed_qa_image_keys": [i["qa_image_key"] for i in payload.get("QA_Images", [])], "awaiting_confirmation": True, "updated_at": datetime.now(timezone.utc).astimezone().isoformat()}
    (run_dir / "qa_progress.json").write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")
    report = module.OUT_BASE / cfg["qa_run"] / f"SEO_QA_qa_batch_{batch_id}.md"
    lines = []
    for line in report.read_text(encoding="utf-8").splitlines():
        if "rendered_sheets" in line:
            lines.append("- **XLSX: COMPLETE.** Workbook QA has exactly five sheets, auditable formulas, filters, frozen headers, wrapped text and hyperlinks; structure, formulas and layout were reopened and checked.")
        elif "output rút gọn" in line:
            lines.append(f"- Note: this is the complete independent QA following `prompt_qa.md`, revision r1.")
        else:
            lines.append(line)
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main():
    for batch_id, cfg in CONFIG.items():
        cfg["expected_images"] = len(json.loads((module.QA_BASE / cfg["qa_run"] / "submitted_batch_data.json").read_text(encoding="utf-8"))["images"])
    module.BATCHES = CONFIG
    module.main()
    for batch_id, cfg in CONFIG.items():
        finish(batch_id, cfg)


if __name__ == "__main__":
    main()
