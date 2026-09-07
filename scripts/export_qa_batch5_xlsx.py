from pathlib import Path
import json

import export_qa_batch1_xlsx as exporter
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
RUN_ID, QA_RUN_ID = "20260906_234129", "20260907_101050"
exporter.DATA = ROOT / "seo_runs" / "jeminise.com" / RUN_ID / "qa" / QA_RUN_ID / "qa_workbook_payload.json"
exporter.OUTPUT = ROOT / "resutls" / "jeminise.com" / RUN_ID / "qa" / QA_RUN_ID / "SEO_QA_qa_batch_005.xlsx"


if __name__ == "__main__":
    exporter.main()
    wb = load_workbook(exporter.OUTPUT, data_only=False, read_only=False)
    formulas = [c.value for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.data_type == "f"]
    hyperlinks = [c.hyperlink.target for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.hyperlink]
    audit = {
        "sheet_names": wb.sheetnames,
        "exact_five_sheets": wb.sheetnames == list(exporter.SHEETS),
        "rows": {ws.title: ws.max_row - 1 for ws in wb.worksheets},
        "formula_count": len(formulas),
        "formula_ref_errors": [f for f in formulas if "#REF!" in f or "#NAME?" in f],
        "hyperlink_count": len(hyperlinks),
        "filters_present": all(bool(ws.auto_filter.ref) for ws in wb.worksheets),
        "freeze_headers": all(ws.freeze_panes == "A2" for ws in wb.worksheets),
        "wrap_text": all(ws.cell(2, 1).alignment.wrap_text for ws in wb.worksheets),
        "rendered_sheets_expected": list(exporter.SHEETS),
        "excel_com_note": "Excel Desktop recalculation was not used; formulas were structurally audited and the workbook was rendered independently with openpyxl/Pillow.",
    }
    assert audit["exact_five_sheets"] and not audit["formula_ref_errors"]
    assert audit["filters_present"] and audit["freeze_headers"] and audit["wrap_text"]
    assert audit["rows"] == {"QA_Summary":13, "QA_Products":10, "QA_Criteria":110, "QA_Images":72, "QA_Issues":71}
    path = ROOT / "seo_runs" / "jeminise.com" / RUN_ID / "qa" / QA_RUN_ID / "spreadsheet_validation.json"
    path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(audit, ensure_ascii=False, indent=2))
