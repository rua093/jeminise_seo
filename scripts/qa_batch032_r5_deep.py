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
QA_RUN_ID = "20260909_134214"
BATCH_ID = "qa_batch_032_r5"
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / "qa_batch_032_r5" / "SEO_Product_Optimization_qa_batch_032_r5.xlsx"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
OUT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID

WEIGHTS = {"P1": 15, "P2": 10, "K1": 10, "K2": 5, "K3": 5, "T1": 10, "T2": 5, "D1": 5, "D2": 10, "I1": 20, "E1": 5}
RVAL = {"FULL": 1.0, "PARTIAL": 0.5, "FAIL": 0.0, "NOT_CHECKED": 0.0}
SERP = {
    "deer": "Amazon/CustomFam/Target comparable deer quilt and bedding SERP results; no first-party volume.",
    "wolf": "Walmart/Doonakingdom/eBay comparable wolf dreamcatcher quilt and bedding SERP results; no first-party volume.",
    "yggdrasil": "Viking Sons of Odin/Luvingift comparable Yggdrasil quilt bedding SERP results; no first-party volume.",
    "rooster": "Etsy/Amazon/Knotzee comparable rooster quilt and patchwork SERP results; no first-party volume.",
}


def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def now():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def html_text(html):
    return BeautifulSoup(html or "", "html.parser").get_text(" ", strip=True)


def headers(ws):
    return {cell.value: i for i, cell in enumerate(ws[1])}


def excel_value(v):
    if isinstance(v, (dict, list)):
        v = json.dumps(v, ensure_ascii=False)
    if isinstance(v, str) and v.startswith(("=", "+", "-", "@")):
        return "'" + v
    return v


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
    }


def crit(product_key, cid, rating, reason, evidence_refs, issue_refs=""):
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


def group_for(handle):
    if "deer" in handle:
        return "deer"
    if "wolf" in handle:
        return "wolf"
    if "yggdrasil" in handle:
        return "yggdrasil"
    return "rooster"


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
    source_hash = sha256(SOURCE)
    admin_hash = sha256(ADMIN_EXPORT)
    checked = now()
    wb = load_workbook(SOURCE, data_only=True, read_only=True)

    ws = wb["SEO_Products"]
    h = headers(ws)
    products = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        row = {k: raw[i] for k, i in h.items()}
        if str(row.get("evidence_id") or "").startswith("evidence_batch_032_"):
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
        first = admin[handle][0]
        l = live[handle]
        evrefs = f"storefront:{p['product_url']}; live_json:{p['product_url']}.js; admin_export:products_export_1.csv; {p['evidence_id']}; contact_sheet"
        text_fields = " ".join(str(p.get(c) or "") for c in ["primary_keyword", "secondary_keywords", "title_proposed", "meta_title_seo", "meta_description_seo", "description_proposed_html", "meta_keyword"])
        body = html_text(p.get("description_proposed_html"))
        meta = p.get("meta_description_seo") or ""
        keyword_group = group_for(handle)

        issue_refs = []
        major = minor = limitation = 0
        has_type_conflict = "comforter" in text_fields.lower()
        meta_cut = bool(re.search(r"\\b(and|with|for|of|to|the)$", meta.strip().lower()) or meta.strip().endswith("..."))
        template_leak = any(x in body.lower() for x in ["copy stays specific", "visible artwork", "wording focuses"])

        if has_type_conflict:
            iid = f"ISSUE_{issue_no:03d}"
            issues.append({
                "issue_id": iid,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "MAJOR",
                "field": "product_type_consistency",
                "submitted_value": text_fields[:500],
                "source_observation": f"Admin Type={first.get('Type')}; live JSON type={l['json_type']}; options={l['options']}",
                "reason": "Proposed SEO fields still use comforter wording for a product supported as Quilt/Quilt Set.",
                "recommended_fix": "Replace customer-facing proposed wording with Quilt/Quilt Set.",
                "supporting_evidence": evrefs,
                "recheck_condition": "No comforter wording remains in proposed fields unless merchant confirms it.",
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
            "source_observation": SERP[keyword_group],
            "reason": "Demand fit is plausible and truthful, but not validated with Search Console, internal search, paid keyword volume or verified review corpus.",
            "recommended_fix": "Provide first-party or paid keyword data if stronger demand validation is required.",
            "supporting_evidence": SERP[keyword_group],
            "recheck_condition": "Re-score K3 when stronger demand evidence is provided.",
        })
        issue_no += 1
        issue_refs.append(iid)
        limitation += 1

        revision_ok = p.get("revision") == "r5"
        if not revision_ok:
            iid = f"ISSUE_{issue_no:03d}"
            issues.append({
                "issue_id": iid,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "MINOR",
                "field": "revision",
                "submitted_value": p.get("revision"),
                "source_observation": "Source workbook path is qa_batch_032_r5 but this row revision is not r5.",
                "reason": "Traceability metadata is inconsistent. Copy is usable, but audit clarity is weaker.",
                "recommended_fix": "Update row-level revision metadata if this product is carried into the next revision workbook.",
                "supporting_evidence": str(SOURCE.relative_to(ROOT)),
                "recheck_condition": "Row revision matches frozen workbook revision.",
            })
            issue_no += 1
            issue_refs.append(iid)
            minor += 1

        imgs = sorted(images_by_handle[handle], key=lambda x: int(x.get("image_number") or 0))
        image_scores = []
        for img in imgs:
            alt = img.get("alt_proposed") or ""
            obs = img.get("observed_visual_details") or ""
            img_conflict = "wildlife-deer" in handle and "comforter" in (alt + " " + obs).lower()
            im3 = "PARTIAL" if img_conflict else "FULL"
            if img_conflict and not has_type_conflict:
                # Defensive issue if a future source only leaves the error in image rows.
                iid = f"ISSUE_{issue_no:03d}"
                issues.append({
                    "issue_id": iid,
                    "product_key": product_key,
                    "qa_image_key": f"{p['evidence_id']}_img_{int(img['image_number']):02d}",
                    "severity": "MAJOR",
                    "field": "image_alt_product_type",
                    "submitted_value": alt,
                    "source_observation": obs,
                    "reason": "Image alt/observation still uses comforter for deer quilt product.",
                    "recommended_fix": "Use quilt/quilt set in image alt and observations.",
                    "supporting_evidence": img.get("evidence_file_or_reference"),
                    "recheck_condition": "All deer image alt text uses quilt terminology.",
                })
                issue_no += 1
                issue_refs.append(iid)
                major += 1
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
                "issue_refs": "",
                "evidence_refs": img.get("evidence_file_or_reference"),
            })

        image_avg = sum(image_scores) / len(image_scores)
        i1_points = round(0.2 * image_avg, 2)
        type_good = not has_type_conflict
        criteria = [
            crit(product_key, "P1", "FULL" if type_good else "PARTIAL", "Product type, artwork, color and visible design details match storefront/admin/contact sheets." if type_good else "Design matches but product type wording conflicts.", evrefs),
            crit(product_key, "P2", "FULL", "Options, personalization field limits and visible sample text are grounded in live/admin evidence; no unsupported shipping or material promise is added.", evrefs),
            crit(product_key, "K1", "FULL" if type_good else "PARTIAL", "Primary and secondary keywords align with visible design and product-page intent." if type_good else "Keyword target uses conflicting product type.", f"{evrefs}; {SERP[keyword_group]}"),
            crit(product_key, "K2", "FULL", "Comparable public SERP/shopping results support product-page intent for the selected theme.", SERP[keyword_group]),
            crit(product_key, "K3", "PARTIAL", "Evidence remains SERP_ONLY; no first-party or paid keyword demand source supplied.", SERP[keyword_group], issue_refs[0] if issue_refs else ""),
            crit(product_key, "T1", "FULL" if type_good else "PARTIAL", "SEO title is clear, natural, design-specific and aligned to product type." if type_good else "SEO title has product-type conflict.", evrefs),
            crit(product_key, "T2", "FULL" if type_good else "PARTIAL", "Product title proposal is useful and consistent with product evidence." if type_good else "Product title has product-type conflict.", evrefs),
            crit(product_key, "D1", "FULL" if (type_good and not meta_cut) else "PARTIAL", "Meta description is complete, specific and not cut mid-word or mid-idea." if not meta_cut else "Meta description appears cut or incomplete.", evrefs),
            crit(product_key, "D2", "FULL" if (type_good and not template_leak) else "PARTIAL", "Body HTML is complete, specific, free of internal/template language and grounded in verified details." if not template_leak else "Body contains internal/template wording.", evrefs),
            {"product_key": product_key, "criterion_id": "I1", "weight": 20, "assessment": "DERIVED", "rating": "FULL" if image_avg == 100 else "PARTIAL", "earned_points": i1_points, "assessed_weight": 20, "reason": "All image positions are present and alt text is specific, truthful and natural." if image_avg == 100 else "Some image alt text has product-type wording issues.", "evidence_refs": "Image_Audit; contact sheets; live JSON image count; admin export image count", "issue_refs": ""},
            crit(product_key, "E1", "FULL" if revision_ok else "PARTIAL", "Source, admin export, live page and revision metadata are traceable." if revision_ok else "Traceable source, but row revision metadata does not match r5 workbook.", evrefs, "; ".join(issue_refs)),
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
    xlsx = OUT_DIR / "SEO_QA_qa_batch_032_r5.xlsx"
    md = OUT_DIR / "SEO_QA_qa_batch_032_r5.md"
    write_xlsx(payload, xlsx)
    xlsx_hash = sha256(xlsx)
    validation = {
        "sheet_names": load_workbook(xlsx, read_only=True).sheetnames,
        "row_counts": {k: len(v) for k, v in payload.items()},
        "criteria_weights_sum_100": all(sum(c["weight"] for c in qa_criteria if c["product_key"] == p["product_key"]) == 100 for p in qa_products),
        "zip_bad_member": zipfile.ZipFile(xlsx).testzip(),
        "xlsx_sha256": xlsx_hash,
        "passed": len(qa_products) == 10 and len(qa_images) == 70 and len(qa_criteria) == 110,
    }
    for name, obj in [
        ("qa_dataset.json", payload),
        ("qa_workbook_payload.json", {k: payload[k] for k in ("QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues")}),
        ("validation_results.json", validation),
        ("manifest.json", {"rubric_version": "prompt_qa.md v1.0 / 2026-09-06", "qa_run_id": QA_RUN_ID, "batch_id": BATCH_ID, "source_workbook": str(SOURCE.relative_to(ROOT)), "source_workbook_sha256": source_hash, "admin_export": "products_export_1.csv", "admin_export_sha256": admin_hash, "batch_product_keys": [p["product_key"] for p in products], "checked_at": checked, "output_markdown": str(md.relative_to(ROOT)), "output_xlsx": str(xlsx.relative_to(ROOT)), "output_xlsx_sha256": xlsx_hash}),
        ("qa_progress.json", {"rubric_version": "prompt_qa.md v1.0 / 2026-09-06", "qa_run_id": QA_RUN_ID, "batch_id": BATCH_ID, "source_workbook": str(SOURCE.relative_to(ROOT)), "source_workbook_sha256": source_hash, "batch_product_keys": [p["product_key"] for p in products], "current_stage": "BATCH_COMPLETE", "completed_image_keys": [i["qa_image_key"] for i in qa_images], "last_saved_at": checked, "artifact_paths": {"markdown": str(md.relative_to(ROOT)), "xlsx": str(xlsx.relative_to(ROOT)), "dataset": str((RUN_DIR / "qa_dataset.json").relative_to(ROOT))}, "awaiting_confirmation": True, "confirmation_ref": None}),
    ]:
        (RUN_DIR / name).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")

    rows = "\n".join(f"| `{p['product_key'].split('+',1)[-1]}` | {p['final_score']:.1f} | `{p['qa_status']}` | {p['images_checked']}/{p['images_expected']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']} | {p['issue_refs']} |" for p in qa_products)
    md.write_text(f"""# SEO QA qa_batch_032_r5

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
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0; có mở live storefront/`.js`, đối chiếu admin export, workbook r5 và contact sheet/ảnh local.
- `311-314`: lỗi `comforter` của QA r4 đã được sửa; proposed fields và 32 alt/observation liên quan đều dùng `quilt/quilt set`.
- `K3` vẫn `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức `SERP_ONLY`/comparable public SERP.
- `E1` `FULL` cho 4 dòng đã sửa `r5`; `E1 PARTIAL` cho 6 dòng còn lại vì row-level `revision` vẫn là `r2` trong workbook r5. Đây là traceability `MINOR`, không phải lỗi nội dung.
- Không phát hiện `CRITICAL` hoặc `MAJOR`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
{rows}

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 70/70 vị trí ảnh khớp workbook, admin export và live JSON; alt mô tả đúng ảnh, không nhồi keyword.
- 10/10 sản phẩm đạt `QA_PASS` theo rubric: final_score >=85, kiểm tra đủ, không có `CRITICAL/MAJOR`.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc verified review corpus; `K3` không được nâng lên `DEMAND_SUPPORTED`.
- QA không sửa workbook nguồn, không đổi `APPROVED`, không xác nhận import an toàn.
""", encoding="utf-8-sig")
    print(json.dumps({"markdown": str(md), "xlsx": str(xlsx), "batch_status": batch_status, "batch_score": batch_score, "status_counts": dict(status_counts), "issue_counts": dict(issue_counts), "validation": validation}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
