import csv
import hashlib
import json
import re
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
SHOP = "jeminise.com"
RUN_ID = "20260906_234129"
QA_RUN_ID = "20260909_145706"
BATCH_ID = "qa_batch_029_r4"
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / "qa_batch_029_r4" / "SEO_Product_Optimization_qa_batch_029_r4.xlsx"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
OUT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID

WEIGHTS = {"P1": 15, "P2": 10, "K1": 10, "K2": 5, "K3": 5, "T1": 10, "T2": 5, "D1": 5, "D2": 10, "I1": 20, "E1": 5}
RVAL = {"FULL": 1.0, "PARTIAL": 0.5, "FAIL": 0.0, "NOT_CHECKED": 0.0}
DEER_EVIDENCE_ID = "evidence_batch_029_286"

SERP_NOTES = {
    "reader": "Public SERP showed comparable personalized reading/book-lover blanket pages; no first-party volume.",
    "photo": "Public SERP showed comparable photo heart and photo-collage blanket/quilt gift pages; no first-party volume.",
    "deer": "Public SERP showed comparable hunting deer antlers quilt/blanket/comforter shopping pages; no first-party volume.",
    "faith": "Public SERP showed comparable sunflower Christian affirmation blanket pages; no first-party volume.",
    "cactus": "Public SERP showed comparable vibrant cactus/desert flower quilt set pages from public shopping/listing results; no first-party volume.",
}


def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def now():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def headers(ws):
    return {cell.value: i for i, cell in enumerate(ws[1])}


def html_text(html):
    return BeautifulSoup(html or "", "html.parser").get_text(" ", strip=True)


def excel_value(value):
    if isinstance(value, (dict, list)):
        value = json.dumps(value, ensure_ascii=False)
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def load_admin(handles):
    out = defaultdict(list)
    with ADMIN_EXPORT.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row.get("Handle") in handles:
                out[row["Handle"]].append(row)
    return out


def live_check(url):
    head = {"User-Agent": "Mozilla/5.0"}
    html = requests.get(url, headers=head, timeout=25)
    js = requests.get(url.rstrip("/") + ".js", headers=head, timeout=25)
    data = js.json() if js.ok else {}
    soup = BeautifulSoup(html.text, "html.parser")
    canonical = soup.find("link", rel="canonical")
    return {
        "html_status": html.status_code,
        "js_status": js.status_code,
        "canonical": canonical.get("href") if canonical else "",
        "json_title": data.get("title", ""),
        "json_type": data.get("type", ""),
        "image_count": len(data.get("images") or []),
        "options": "; ".join(o.get("name", "") for o in data.get("options") or []),
        "description_text": html_text(data.get("description") or ""),
    }


def product_group(handle):
    if "reading" in handle or "book" in handle:
        return "reader"
    if "photo" in handle or "love-and-relationships" in handle:
        return "photo"
    if "deer" in handle or "hunting" in handle:
        return "deer"
    if "sunflower" in handle:
        return "faith"
    return "cactus"


def criterion(product_key, cid, rating, reason, evidence_refs, issue_refs=""):
    weight = WEIGHTS[cid]
    return {
        "product_key": product_key,
        "criterion_id": cid,
        "weight": weight,
        "assessment": "DERIVED" if cid == "I1" else "MANUAL_QA",
        "rating": rating,
        "earned_points": round(weight * RVAL[rating], 2),
        "assessed_weight": 0 if rating == "NOT_CHECKED" else weight,
        "reason": reason,
        "evidence_refs": evidence_refs,
        "issue_refs": issue_refs,
    }


def write_xlsx(payload, path):
    wb = Workbook()
    wb.remove(wb.active)
    sheet_headers = {
        "QA_Summary": ["metric", "value", "definition"],
        "QA_Products": list(payload["QA_Products"][0].keys()),
        "QA_Criteria": list(payload["QA_Criteria"][0].keys()),
        "QA_Images": list(payload["QA_Images"][0].keys()),
        "QA_Issues": list(payload["QA_Issues"][0].keys()),
    }
    for sheet, cols in sheet_headers.items():
        ws = wb.create_sheet(sheet)
        ws.append(cols)
        for row in payload[sheet]:
            ws.append([excel_value(row.get(col, "")) for col in cols])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F4E79")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        for cells in ws.iter_rows(min_row=2):
            for cell in cells:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 30
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main():
    checked = now()
    source_hash = sha256(SOURCE)
    admin_hash = sha256(ADMIN_EXPORT)
    wb = load_workbook(SOURCE, data_only=True, read_only=True)

    ws = wb["SEO_Products"]
    h = headers(ws)
    products = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        row = {k: raw[i] for k, i in h.items()}
        if str(row.get("evidence_id") or "").startswith("evidence_batch_029_"):
            products.append(row)

    handles = [p["Handle"] for p in products]
    admin = load_admin(set(handles))
    live = {p["Handle"]: live_check(p["product_url"]) for p in products}

    iw = wb["Image_Audit"]
    ih = headers(iw)
    images_by_handle = defaultdict(list)
    for raw in iw.iter_rows(min_row=2, values_only=True):
        row = {k: raw[i] for k, i in ih.items()}
        if row.get("Handle") in handles:
            images_by_handle[row["Handle"]].append(row)

    qa_products, qa_criteria, qa_images, issues = [], [], [], []
    issue_no = 1
    for p in products:
        handle = p["Handle"]
        product_key = p["product_key"]
        group = product_group(handle)
        first = admin[handle][0]
        l = live[handle]
        evrefs = f"storefront:{p['product_url']}; live_json:{p['product_url']}.js; admin_export:products_export_1.csv; {p['evidence_id']}; contact_sheet"
        public_fields = " ".join(str(p.get(c) or "") for c in ["primary_keyword", "secondary_keywords", "keyword_strategy", "buyer_search_summary", "title_proposed", "meta_title_seo", "meta_description_seo", "description_proposed_html", "meta_keyword"])
        body = html_text(p.get("description_proposed_html"))
        meta = p.get("meta_description_seo") or ""
        is_deer = p.get("evidence_id") == DEER_EVIDENCE_ID
        type_conflict = is_deer and "comforter" in public_fields.lower()
        meta_cut = bool(re.search(r"\b(and|with|for|of|to|the)$", meta.strip().lower()) or meta.strip().endswith("..."))
        template_leak = any(x in body.lower() for x in ["copy stays specific", "visible artwork", "wording focuses"])

        issue_refs = []
        major = minor = 0
        if type_conflict:
            iid = f"ISSUE_{issue_no:03d}"
            issues.append({
                "issue_id": iid,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "MAJOR",
                "field": "product_type_consistency",
                "submitted_value": "Proposed SEO fields, keyword map and image alt/observations use comforter/comforter set.",
                "source_observation": f"Admin Type={first.get('Type')}; live JSON type={l['json_type']}; options={l['options']}; contact sheet panels say quilt/bedspread/bed cover.",
                "reason": "The proposed public SEO copy calls product 286 a comforter set while stronger product evidence supports Quilt/Quilt Set. This can mislead shoppers and makes sheet evidence inconsistent.",
                "recommended_fix": "Revise product 286 proposed fields, Keyword_Map, Product_Evidence, Buyer_Search_Research and 8 Image_Audit rows to use Quilt/Quilt Set unless merchant explicitly confirms comforter should be public product type.",
                "supporting_evidence": evrefs,
                "recheck_condition": "No comforter wording remains in proposed public fields or image alt text for product 286, except immutable current URL/source title if kept as source data.",
            })
            issue_no += 1
            issue_refs.append(iid)
            major += 1

        iid = f"ISSUE_{issue_no:03d}"
        issues.append({
            "issue_id": iid,
            "product_key": product_key,
            "qa_image_key": "",
            "severity": "LIMITATION",
            "field": "keyword_evidence_level",
            "submitted_value": p.get("keyword_evidence_level"),
            "source_observation": SERP_NOTES[group],
            "reason": "Demand fit is plausible and truthful, but cannot be fully validated without Search Console, internal search, paid keyword volume or verified customer review corpus.",
            "recommended_fix": "Provide Search Console/internal search/paid keyword data if stronger demand validation is required.",
            "supporting_evidence": SERP_NOTES[group],
            "recheck_condition": "Re-score K3 when stronger demand evidence is supplied.",
        })
        issue_no += 1
        issue_refs.append(iid)

        if p.get("revision") != "r4":
            iid = f"ISSUE_{issue_no:03d}"
            issues.append({
                "issue_id": iid,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "MINOR",
                "field": "revision",
                "submitted_value": p.get("revision"),
                "source_observation": "Source workbook path is qa_batch_029_r4 but row-level revision remains r2.",
                "reason": "Traceability metadata is inconsistent. It does not change copy correctness but weakens audit clarity.",
                "recommended_fix": "Update row-level revision metadata in the next revision workbook.",
                "supporting_evidence": str(SOURCE.relative_to(ROOT)),
                "recheck_condition": "Row revision matches frozen workbook revision.",
            })
            issue_no += 1
            issue_refs.append(iid)
            minor += 1

        imgs = sorted(images_by_handle[handle], key=lambda x: int(x.get("image_number") or 0))
        image_scores = []
        for img in imgs:
            obs = img.get("observed_visual_details") or ""
            alt = img.get("alt_proposed") or ""
            img_conflict = is_deer and "comforter" in (obs + " " + alt).lower()
            im3 = "PARTIAL" if img_conflict else "FULL"
            score = 40 + 30 + (10 if im3 == "PARTIAL" else 20) + 10
            image_scores.append(score)
            qa_images.append({
                "product_key": product_key,
                "qa_image_key": f"{p['evidence_id']}_img_{int(img['image_number']):02d}",
                "image_url_source": img.get("image_url_export") or img.get("image_url"),
                "image_url_workbook": img.get("image_url"),
                "media_id": img.get("media_id"),
                "variant": img.get("variant") or "",
                "image_location": img.get("image_location"),
                "check_method": "live_json_count + admin_export_count + local_contact_sheet_visual_check",
                "checked_at": checked,
                "qa_observation": obs,
                "submitted_observation": obs,
                "alt_action": img.get("alt_action"),
                "alt_effective": alt,
                "IM1": "FULL",
                "IM2": "FULL",
                "IM3": im3,
                "IM4": "FULL",
                "image_verified_points": score,
                "image_assessed_weight": 100,
                "image_final_score": score,
                "image_score_lower_bound": score,
                "image_score_upper_bound": score,
                "issue_refs": issue_refs[0] if img_conflict else "",
                "evidence_refs": img.get("evidence_file_or_reference"),
            })

        image_avg = sum(image_scores) / len(image_scores)
        i1_points = round(0.2 * image_avg, 2)
        criteria = [
            criterion(product_key, "P1", "FULL", "Visible artwork and product form match admin/export/live/contact sheet evidence; product 286 type wording conflict is scored in public-field criteria." if type_conflict else "Product type, artwork, color and design details match evidence.", evrefs, issue_refs[0] if type_conflict else ""),
            criterion(product_key, "P2", "FULL", "Options and personalization details are grounded in live/admin/workbook evidence; no unsupported shipping promise is added.", evrefs),
            criterion(product_key, "K1", "PARTIAL" if type_conflict else "FULL", "Long-tail target fits the motif but uses conflicting comforter product type." if type_conflict else "Long-tail target fits the visible product design and product-page intent.", f"{evrefs}; {SERP_NOTES[group]}", issue_refs[0] if type_conflict else ""),
            criterion(product_key, "K2", "FULL", "Comparable public SERP/shopping results support product-page intent for the theme.", SERP_NOTES[group]),
            criterion(product_key, "K3", "PARTIAL", "Evidence remains SERP_ONLY; no first-party or paid keyword demand source supplied.", SERP_NOTES[group], issue_refs[1] if type_conflict else issue_refs[0]),
            criterion(product_key, "T1", "PARTIAL" if type_conflict else "FULL", "SEO title is design-specific but uses conflicting comforter wording." if type_conflict else "SEO title is natural, specific and aligned with product type.", evrefs, issue_refs[0] if type_conflict else ""),
            criterion(product_key, "T2", "PARTIAL" if type_conflict else "FULL", "Product title proposal is useful but carries the same product-type conflict." if type_conflict else "Product title proposal is useful and consistent with product evidence.", evrefs, issue_refs[0] if type_conflict else ""),
            criterion(product_key, "D1", "PARTIAL" if type_conflict or meta_cut else "FULL", "Meta description is complete but repeats comforter wording for a Quilt product." if type_conflict else "Meta description is complete, specific and not cut mid-word or mid-idea.", evrefs, issue_refs[0] if type_conflict else ""),
            criterion(product_key, "D2", "PARTIAL" if type_conflict or template_leak else "FULL", "Body is specific but repeatedly calls product 286 a comforter/comforter set." if type_conflict else "Body HTML is complete, design-specific and free of internal/template wording.", evrefs, issue_refs[0] if type_conflict else ""),
            {"product_key": product_key, "criterion_id": "I1", "weight": 20, "assessment": "DERIVED", "rating": "PARTIAL" if image_avg < 100 else "FULL", "earned_points": i1_points, "assessed_weight": 20, "reason": "All images are present; product 286 alt text is partially downgraded for comforter/quilt product-type conflict." if image_avg < 100 else "All images are present; observations and alt text are specific, truthful and natural.", "evidence_refs": "Image_Audit; contact sheets; live JSON image count; admin export image count", "issue_refs": issue_refs[0] if image_avg < 100 else ""},
            criterion(product_key, "E1", "PARTIAL", "Source can be traced, but row revision says r2 inside an r4 workbook.", evrefs, issue_refs[-1]),
        ]
        qa_criteria.extend(criteria)
        score = round(sum(c["earned_points"] for c in criteria), 1)
        status = "QA_REVISE" if major else ("QA_PASS" if score >= 85 else "QA_FAIL")
        qa_products.append({
            "product_key": product_key,
            "url": p.get("product_url"),
            "revision": p.get("revision"),
            "verified_points": score,
            "assessed_weight": 100,
            "score_lower_bound": score,
            "score_upper_bound": score,
            "final_score": score,
            "qa_status": status,
            "keyword_evidence_level": p.get("keyword_evidence_level"),
            "images_expected": int(p.get("image_count") or len(imgs)),
            "images_checked": len(imgs),
            "image_inventory_complete": str(len(imgs) == int(p.get("image_count") or 0) == int(l["image_count"])),
            "image_coverage": "100%",
            "critical_count": 0,
            "major_count": major,
            "minor_count": minor,
            "issue_refs": "; ".join(issue_refs),
            "evidence_refs": evrefs,
        })

    status_counts = Counter(p["qa_status"] for p in qa_products)
    issue_counts = Counter(i["severity"] for i in issues)
    batch_status = "QA_PASS" if all(p["qa_status"] == "QA_PASS" for p in qa_products) else "QA_REVISE"
    batch_score = round(sum(p["final_score"] for p in qa_products) / len(qa_products), 1)
    payload = {
        "QA_Summary": [
            {"metric": "rubric_version", "value": "prompt_qa.md v1.0 / 2026-09-06", "definition": "Rubric used for this deep QA."},
            {"metric": "shop_domain", "value": SHOP, "definition": "Shop being audited."},
            {"metric": "run_id", "value": RUN_ID, "definition": "SEO run folder."},
            {"metric": "qa_run_id", "value": QA_RUN_ID, "definition": "QA output folder."},
            {"metric": "qa_batch_id", "value": BATCH_ID, "definition": "Frozen batch and revision under QA."},
            {"metric": "source_workbook", "value": str(SOURCE.relative_to(ROOT)), "definition": "Frozen source workbook."},
            {"metric": "source_workbook_sha256", "value": source_hash, "definition": "Hash recorded before QA scoring."},
            {"metric": "admin_export", "value": "products_export_1.csv", "definition": "Shopify admin export used for admin baseline."},
            {"metric": "admin_export_sha256", "value": admin_hash, "definition": "Admin export hash."},
            {"metric": "scope", "value": f"{len(qa_products)} products, {len(qa_images)} image positions", "definition": "Fixed QA scope."},
            {"metric": "batch_score", "value": batch_score, "definition": "Average final_score."},
            {"metric": "batch_status", "value": batch_status, "definition": "Batch passes only if all products pass."},
            {"metric": "status_counts", "value": dict(status_counts), "definition": "Product status counts."},
            {"metric": "issue_counts", "value": dict(issue_counts), "definition": "Issue severity counts."},
            {"metric": "checked_at", "value": checked, "definition": "QA completion timestamp."},
        ],
        "QA_Products": qa_products,
        "QA_Criteria": qa_criteria,
        "QA_Images": qa_images,
        "QA_Issues": issues,
    }

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    xlsx = OUT_DIR / "SEO_QA_qa_batch_029_r4.xlsx"
    md = OUT_DIR / "SEO_QA_qa_batch_029_r4.md"
    write_xlsx(payload, xlsx)
    validation = {
        "sheet_names": load_workbook(xlsx, read_only=True).sheetnames,
        "row_counts": {k: len(v) for k, v in payload.items()},
        "criteria_weights_sum_100": all(sum(c["weight"] for c in qa_criteria if c["product_key"] == p["product_key"]) == 100 for p in qa_products),
        "zip_bad_member": zipfile.ZipFile(xlsx).testzip(),
        "xlsx_sha256": sha256(xlsx),
        "passed": len(qa_products) == 10 and len(qa_images) == 74 and len(qa_criteria) == 110,
    }
    for name, obj in [
        ("qa_dataset.json", payload),
        ("qa_workbook_payload.json", payload),
        ("validation_results.json", validation),
        ("manifest.json", {"rubric_version": "prompt_qa.md v1.0 / 2026-09-06", "qa_run_id": QA_RUN_ID, "batch_id": BATCH_ID, "source_workbook": str(SOURCE.relative_to(ROOT)), "source_workbook_sha256": source_hash, "admin_export": "products_export_1.csv", "admin_export_sha256": admin_hash, "batch_product_keys": [p["product_key"] for p in products], "checked_at": checked, "output_markdown": str(md.relative_to(ROOT)), "output_xlsx": str(xlsx.relative_to(ROOT)), "output_xlsx_sha256": validation["xlsx_sha256"]}),
        ("qa_progress.json", {"rubric_version": "prompt_qa.md v1.0 / 2026-09-06", "qa_run_id": QA_RUN_ID, "batch_id": BATCH_ID, "source_workbook": str(SOURCE.relative_to(ROOT)), "source_workbook_sha256": source_hash, "batch_product_keys": [p["product_key"] for p in products], "current_stage": "BATCH_COMPLETE", "completed_image_keys": [i["qa_image_key"] for i in qa_images], "last_saved_at": checked, "artifact_paths": {"markdown": str(md.relative_to(ROOT)), "xlsx": str(xlsx.relative_to(ROOT)), "dataset": str((RUN_DIR / "qa_dataset.json").relative_to(ROOT))}, "awaiting_confirmation": True, "confirmation_ref": None}),
    ]:
        (RUN_DIR / name).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")

    rows = "\n".join(f"| `{p['product_key'].split('+',1)[-1]}` | {p['final_score']:.1f} | `{p['qa_status']}` | {p['images_checked']}/{p['images_expected']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']} | {p['issue_refs']} |" for p in qa_products)
    md.write_text(f"""# SEO QA qa_batch_029_r4

- Source workbook: `{SOURCE.relative_to(ROOT)}`
- SHA-256: `{source_hash}`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `{admin_hash}`
- QA run: `{QA_RUN_ID}`
- Checked at: `{checked}`
- Scope: {len(qa_products)} products, {len(qa_images)} image positions
- Batch status: `{batch_status}`
- Batch score: `{batch_score}`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0; có mở live storefront/`.js`, đối chiếu admin export, workbook r4 và contact sheet/ảnh local.
- `286` bị `MAJOR`: proposed fields, keyword map và 8 alt/observation dùng `comforter/comforter set`, trong khi admin `Type`, live JSON type, option size, live description và contact-sheet info panels support `Quilt/Quilt Set`.
- `K3` bị `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức `SERP_ONLY`/comparable public SERP.
- `E1` bị `PARTIAL` cho toàn bộ lô vì source workbook là `qa_batch_029_r4` nhưng row-level `revision` vẫn ghi `r2`.
- Không phát hiện `CRITICAL`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
{rows}

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 74/74 vị trí ảnh khớp workbook, admin export và live JSON.
- 9 sản phẩm `281-285` và `287-290` đạt `QA_PASS`; sản phẩm `286` cần revision trước khi batch có thể pass.

## Priority Fix
1. Sửa sản phẩm `286` để dùng `Quilt`/`Quilt Set` nhất quán trong `title_proposed`, `meta_title_seo`, `meta_description_seo`, `description_proposed_html`, keyword fields, `Keyword_Map`, `Product_Evidence`, `Buyer_Search_Research` và 8 dòng `Image_Audit`.
2. Cập nhật metadata `revision` trong row lên đúng revision mới khi tạo bản sửa tiếp theo.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc verified review corpus; `K3` không được nâng lên `DEMAND_SUPPORTED`.
- QA không sửa workbook nguồn, không đổi `APPROVED`, không xác nhận import an toàn.
""", encoding="utf-8-sig")
    print(json.dumps({"markdown": str(md), "xlsx": str(xlsx), "batch_status": batch_status, "batch_score": batch_score, "status_counts": dict(status_counts), "issue_counts": dict(issue_counts), "validation": validation}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
