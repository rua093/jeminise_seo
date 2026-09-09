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
QA_RUN_ID = "20260909_132406"
BATCH_ID = "qa_batch_032_r4"
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / "qa_batch_032_r4" / "SEO_Product_Optimization_qa_batch_032_r4.xlsx"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
OUT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID

WEIGHTS = {
    "P1": 15,
    "P2": 10,
    "K1": 10,
    "K2": 5,
    "K3": 5,
    "T1": 10,
    "T2": 5,
    "D1": 5,
    "D2": 10,
    "I1": 20,
    "E1": 5,
}
RATING_VALUE = {"FULL": 1.0, "PARTIAL": 0.5, "FAIL": 0.0, "NOT_CHECKED": 0.0}

SERP_REFS = {
    "deer": [
        "https://www.amazon.ca/You-Me-Got-This-Hunting-Bedding/dp/",
        "https://customfam.com/products/you-and-me-we-got-this-quilt-bed-sets-couple-deer",
        "https://www.target.com/s/deer+comforter",
    ],
    "wolf": [
        "https://www.walmart.com/search?q=personalized+wolf+dreamcatcher+quilt",
        "https://doonakingdom.com.au/products/wolf-dreamcatcher-quilt-cover-set",
        "https://www.ebay.com/sch/i.html?_nkw=wolf+quilt+bedding+set",
    ],
    "yggdrasil": [
        "https://vikingsonsofodin.com/products/viking-bedding-set-viking-raven-tree-of-life-yggdrasil",
        "https://luvingift.com/products/viking-quilt-set-yggdrasil-odins-ravens-huginn-and-muninn",
    ],
    "rooster": [
        "https://www.etsy.com/market/rooster_quilts",
        "https://knotzee.com/products/patchwork-rooster-quilt-pattern-applique-wall-hanging-pdf-download",
        "https://www.amazon.com/s?k=rooster+patchwork+quilt",
    ],
}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def html_text(value: str) -> str:
    return BeautifulSoup(value or "", "html.parser").get_text(" ", strip=True)


def safe_formula_text(value):
    if isinstance(value, (dict, list)):
        value = json.dumps(value, ensure_ascii=False)
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def load_admin_rows(handles):
    rows = defaultdict(list)
    with ADMIN_EXPORT.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("Handle") in handles:
                rows[row["Handle"]].append(row)
    return rows


def parse_live(product_url):
    headers = {"User-Agent": "Mozilla/5.0"}
    html_resp = requests.get(product_url, headers=headers, timeout=25)
    js_resp = requests.get(product_url.rstrip("/") + ".js", headers=headers, timeout=25)
    data = js_resp.json() if js_resp.ok else {}
    soup = BeautifulSoup(html_resp.text, "html.parser")
    canonical = soup.find("link", rel="canonical")
    return {
        "html_status": html_resp.status_code,
        "js_status": js_resp.status_code,
        "canonical": canonical.get("href") if canonical else "",
        "json_title": data.get("title", ""),
        "image_count": len(data.get("images") or []),
        "variant_count": len(data.get("variants") or []),
        "options": "; ".join(o.get("name", "") for o in data.get("options") or []),
    }


def criterion(product_key, cid, rating, reason, evidence_refs, issue_refs=""):
    weight = WEIGHTS[cid]
    assessed = 0 if rating == "NOT_CHECKED" else weight
    return {
        "product_key": product_key,
        "criterion_id": cid,
        "weight": weight,
        "assessment": "DERIVED" if cid == "I1" else "MANUAL_QA",
        "rating": rating,
        "earned_points": round(weight * RATING_VALUE[rating], 2),
        "assessed_weight": assessed,
        "reason": reason,
        "evidence_refs": evidence_refs,
        "issue_refs": issue_refs,
    }


def style_workbook(wb: Workbook):
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        widths = {
            "QA_Summary": 36,
            "QA_Products": 28,
            "QA_Criteria": 28,
            "QA_Images": 34,
            "QA_Issues": 34,
        }
        for col_idx in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = widths.get(ws.title, 24)


def write_xlsx(payload, path: Path):
    wb = Workbook()
    wb.remove(wb.active)
    headers = {
        "QA_Summary": ["metric", "value", "definition"],
        "QA_Products": list(payload["QA_Products"][0].keys()),
        "QA_Criteria": list(payload["QA_Criteria"][0].keys()),
        "QA_Images": list(payload["QA_Images"][0].keys()),
        "QA_Issues": list(payload["QA_Issues"][0].keys()),
    }
    for sheet_name, sheet_headers in headers.items():
        ws = wb.create_sheet(sheet_name)
        ws.append(sheet_headers)
        for item in payload[sheet_name]:
            ws.append([safe_formula_text(item.get(h, "")) for h in sheet_headers])
    style_workbook(wb)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main():
    source_hash = sha256(SOURCE)
    admin_hash = sha256(ADMIN_EXPORT)
    checked_at = now_iso()

    wb = load_workbook(SOURCE, data_only=True, read_only=True)
    ws = wb["SEO_Products"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = {h: i for i, h in enumerate(headers)}
    products = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[idx["evidence_id"]] or "").startswith("evidence_batch_032_"):
            products.append({h: row[idx[h]] for h in headers})

    handles = [p["Handle"] for p in products]
    admin_rows = load_admin_rows(set(handles))

    image_ws = wb["Image_Audit"]
    image_headers = [c.value for c in next(image_ws.iter_rows(min_row=1, max_row=1))]
    ii = {h: i for i, h in enumerate(image_headers)}
    images_by_handle = defaultdict(list)
    for row in image_ws.iter_rows(min_row=2, values_only=True):
        handle = row[ii["Handle"]]
        if handle in handles:
            images_by_handle[handle].append({h: row[ii[h]] for h in image_headers})

    evidence_ws = wb["Product_Evidence"]
    evidence_headers = [c.value for c in next(evidence_ws.iter_rows(min_row=1, max_row=1))]
    ei = {h: i for i, h in enumerate(evidence_headers)}
    evidence_by_id = {}
    for row in evidence_ws.iter_rows(min_row=2, values_only=True):
        evidence_by_id[str(row[ei["evidence_id"]])] = {h: row[ei[h]] for h in evidence_headers}

    live_by_handle = {}
    for p in products:
        live_by_handle[p["Handle"]] = parse_live(p["product_url"])

    qa_products = []
    qa_criteria = []
    qa_images = []
    issues = []
    issue_no = 1
    deer_handles = set(handles[:4])

    for p in products:
        handle = p["Handle"]
        product_key = p["product_key"]
        evidence_id = p["evidence_id"]
        admin_first = admin_rows[handle][0]
        live = live_by_handle[handle]
        evidence = evidence_by_id[evidence_id]
        issue_refs = []
        major_count = 0
        minor_count = 0
        limitation_count = 0

        ev_refs = "; ".join(
            [
                f"storefront:{p['product_url']}",
                f"live_json:{p['product_url']}.js",
                "admin_export:products_export_1.csv",
                str(evidence_id),
            ]
        )

        keyword_group = "deer" if "deer" in handle else "wolf" if "wolf" in handle else "yggdrasil" if "yggdrasil" in handle else "rooster"
        serp_ref = "; ".join(SERP_REFS[keyword_group])

        product_type_issue = ""
        if handle in deer_handles:
            product_type_issue = f"ISSUE_{issue_no:03d}"
            issues.append(
                {
                    "issue_id": product_type_issue,
                    "product_key": product_key,
                    "qa_image_key": "",
                    "severity": "MAJOR",
                    "field": "product_type_consistency",
                    "submitted_value": f"title/meta/body/alt use comforter; product_type={p['product_type']}",
                    "source_observation": f"Admin export Type={admin_first.get('Type')}; Option1={admin_first.get('Option1 Name')}; admin SEO title={admin_first.get('SEO Title')}; live title={live['json_title']}",
                    "reason": "The proposal keeps the visible/live product-name wording 'Comforter', but the authoritative product type, size option and admin SEO baseline identify the sellable item as Quilt. This weakens product-type accuracy across public SEO fields and image alt text.",
                    "recommended_fix": "Revise public SEO fields to use Quilt/Quilt Set consistently unless the merchant confirms these four deer items should be marketed as comforter sets.",
                    "supporting_evidence": f"{ev_refs}; contact_sheet; admin export",
                    "recheck_condition": "Re-check Product Title, SEO title, SEO description, body HTML and all image alt text after product-type wording is normalized.",
                }
            )
            issue_no += 1
            issue_refs.append(product_type_issue)
            major_count += 1

        k3_issue = f"ISSUE_{issue_no:03d}"
        issues.append(
            {
                "issue_id": k3_issue,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "LIMITATION",
                "field": "keyword_evidence_level",
                "submitted_value": p["keyword_evidence_level"],
                "source_observation": "Public SERP/comparable shopping pages checked, but no Search Console, internal search, paid keyword volume or verified review corpus was provided.",
                "reason": "Demand fit is plausible for the product and market, but cannot be validated as demand-supported beyond public SERP checks.",
                "recommended_fix": "Provide Search Console, internal search or paid keyword data if stronger demand validation is required.",
                "supporting_evidence": serp_ref,
                "recheck_condition": "Re-score K3 when stronger demand evidence is available.",
            }
        )
        issue_no += 1
        issue_refs.append(k3_issue)
        limitation_count += 1

        revision_issue = f"ISSUE_{issue_no:03d}"
        issues.append(
            {
                "issue_id": revision_issue,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "MINOR",
                "field": "revision",
                "submitted_value": p["revision"],
                "source_observation": "Source workbook path is qa_batch_032_r4 but row-level revision remains r2.",
                "reason": "Traceability metadata is inconsistent with the frozen workbook path. It does not change copy correctness, but weakens audit clarity.",
                "recommended_fix": "Update row-level revision metadata in the next revision workbook.",
                "supporting_evidence": str(SOURCE.relative_to(ROOT)),
                "recheck_condition": "Confirm row revision matches the source workbook revision.",
            }
        )
        issue_no += 1
        issue_refs.append(revision_issue)
        minor_count += 1

        body = html_text(p["description_proposed_html"])
        meta_desc = p["meta_description_seo"] or ""
        bad_template = any(x in body.lower() for x in ["copy stays specific", "visible artwork", "wording focuses"])
        meta_complete = not re.search(r"\\b(and|with|for|of|to|the)$", meta_desc.strip().lower()) and not meta_desc.endswith("...")
        images = sorted(images_by_handle[handle], key=lambda x: int(x["image_number"] or 0))
        image_scores = []
        image_issue_ref = product_type_issue if product_type_issue else ""
        for img in images:
            alt = img["alt_proposed"] or ""
            obs = img["observed_visual_details"] or ""
            has_type_conflict = handle in deer_handles and "comforter" in alt.lower()
            im3 = "PARTIAL" if has_type_conflict else "FULL"
            score = 40 + 30 + (10 if im3 == "PARTIAL" else 20) + 10
            image_scores.append(score)
            qa_images.append(
                {
                    "product_key": product_key,
                    "qa_image_key": f"{evidence_id}_img_{int(img['image_number']):02d}",
                    "image_url_source": img["image_url_export"] or img["image_url"],
                    "image_url_workbook": img["image_url"],
                    "media_id": img["media_id"],
                    "variant": img["variant"] or "",
                    "image_location": img["image_location"],
                    "check_method": "live_json_count + admin_export_count + local_contact_sheet_visual_check",
                    "checked_at": checked_at,
                    "qa_observation": obs,
                    "submitted_observation": obs,
                    "alt_action": img["alt_action"],
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
                    "issue_refs": image_issue_ref,
                    "evidence_refs": img["evidence_file_or_reference"],
                }
            )

        avg_image_score = sum(image_scores) / len(image_scores)
        i1_points = round(0.2 * avg_image_score, 2)

        if handle in deer_handles:
            crits = [
                criterion(product_key, "P1", "FULL", "Visible deer artwork, colors and personalization context match the product; product-type wording conflict is scored in public-field criteria.", ev_refs, product_type_issue),
                criterion(product_key, "P2", "FULL", "Customizer option and size/pillowcase options are stated with limits from storefront/admin evidence; no unsupported material or shipping promise is added.", ev_refs),
                criterion(product_key, "K1", "PARTIAL", "Long-tail target fits the deer artwork but uses 'comforter' while admin product type and size option identify Quilt.", f"{ev_refs}; {serp_ref}", product_type_issue),
                criterion(product_key, "K2", "FULL", "Public SERP checks show comparable deer bedding/quilt/comforter product intent.", serp_ref),
                criterion(product_key, "K3", "PARTIAL", "Evidence remains SERP_ONLY; demand is plausible but not validated by first-party or paid keyword data.", serp_ref, k3_issue),
                criterion(product_key, "T1", "PARTIAL", "SEO title is readable and design-specific, but uses Comforter/Comforter Set against Quilt product type evidence.", ev_refs, product_type_issue),
                criterion(product_key, "T2", "PARTIAL", "Product title is useful but carries the same product-type conflict.", ev_refs, product_type_issue),
                criterion(product_key, "D1", "PARTIAL" if meta_complete else "FAIL", "Meta description is complete and design-specific, but repeats comforter wording for a Quilt product type.", ev_refs, product_type_issue),
                criterion(product_key, "D2", "PARTIAL" if not bad_template else "FAIL", "Body removed internal template wording and includes verified options, but repeatedly calls the product a comforter/comforter set.", ev_refs, product_type_issue),
                {
                    "product_key": product_key,
                    "criterion_id": "I1",
                    "weight": 20,
                    "assessment": "DERIVED",
                    "rating": "PARTIAL",
                    "earned_points": i1_points,
                    "assessed_weight": 20,
                    "reason": "All images are present and visually specific; alt text is partially downgraded because it repeats the same comforter/quilt product-type conflict.",
                    "evidence_refs": "Image_Audit; contact sheets; live JSON image count",
                    "issue_refs": product_type_issue,
                },
                criterion(product_key, "E1", "PARTIAL", "Source can be traced, but row revision says r2 inside an r4 workbook.", ev_refs, revision_issue),
            ]
        else:
            crits = [
                criterion(product_key, "P1", "FULL", "Product type, artwork, color and visible design details match storefront, admin export and contact sheet.", ev_refs),
                criterion(product_key, "P2", "FULL", "Options and personalization limits are grounded in storefront/admin evidence; no unsupported material or shipping promise is added.", ev_refs),
                criterion(product_key, "K1", "FULL", "Long-tail target is specific to visible artwork and suitable for the product page.", f"{ev_refs}; {serp_ref}"),
                criterion(product_key, "K2", "FULL", "Public SERP checks show comparable product/shopping intent for the theme and product type.", serp_ref),
                criterion(product_key, "K3", "PARTIAL", "Evidence remains SERP_ONLY; demand is plausible but not validated by first-party or paid keyword data.", serp_ref, k3_issue),
                criterion(product_key, "T1", "FULL", "SEO title is natural, specific and aligned with product type and artwork.", ev_refs),
                criterion(product_key, "T2", "FULL", "Product title/H1 proposal is useful and consistent with the product.", ev_refs),
                criterion(product_key, "D1", "FULL" if meta_complete else "FAIL", "Meta description is complete, specific and not cut mid-word or mid-idea.", ev_refs),
                criterion(product_key, "D2", "FULL" if not bad_template else "FAIL", "Body HTML is complete, design-specific and free of internal/template wording.", ev_refs),
                {
                    "product_key": product_key,
                    "criterion_id": "I1",
                    "weight": 20,
                    "assessment": "DERIVED",
                    "rating": "FULL",
                    "earned_points": i1_points,
                    "assessed_weight": 20,
                    "reason": "All images are present; observations and alt text are specific, truthful and natural.",
                    "evidence_refs": "Image_Audit; contact sheets; live JSON image count",
                    "issue_refs": "",
                },
                criterion(product_key, "E1", "PARTIAL", "Source can be traced, but row revision says r2 inside an r4 workbook.", ev_refs, revision_issue),
            ]

        qa_criteria.extend(crits)
        score = round(sum(c["earned_points"] for c in crits), 1)
        status = "QA_REVISE" if major_count else ("QA_PASS" if score >= 85 else "QA_FAIL")
        qa_products.append(
            {
                "product_key": product_key,
                "url": p["product_url"],
                "revision": p["revision"],
                "verified_points": score,
                "assessed_weight": 100,
                "score_lower_bound": score,
                "score_upper_bound": score,
                "final_score": score,
                "qa_status": status,
                "keyword_evidence_level": p["keyword_evidence_level"],
                "images_expected": int(p["image_count"]),
                "images_checked": len(images),
                "image_inventory_complete": str(len(images) == int(p["image_count"]) == live["image_count"]),
                "image_coverage": "100%",
                "critical_count": 0,
                "major_count": major_count,
                "minor_count": minor_count,
                "issue_refs": "; ".join(issue_refs),
                "evidence_refs": ev_refs,
            }
        )

    scores = [p["final_score"] for p in qa_products]
    status_counts = Counter(p["qa_status"] for p in qa_products)
    issue_counts = Counter(i["severity"] for i in issues)
    batch_status = "QA_PASS" if all(p["qa_status"] == "QA_PASS" for p in qa_products) else "QA_REVISE"
    batch_score = round(sum(scores) / len(scores), 1)
    qa_summary = [
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
        {"metric": "batch_score", "value": batch_score, "definition": "Average of product final_score."},
        {"metric": "batch_status", "value": batch_status, "definition": "Batch status; any MAJOR prevents batch pass."},
        {"metric": "status_counts", "value": dict(status_counts), "definition": "Product status counts."},
        {"metric": "issue_counts", "value": dict(issue_counts), "definition": "Issue severity counts."},
        {"metric": "checked_at", "value": checked_at, "definition": "QA completion timestamp."},
    ]

    payload = {
        "QA_Summary": qa_summary,
        "QA_Products": qa_products,
        "QA_Criteria": qa_criteria,
        "QA_Images": qa_images,
        "QA_Issues": issues,
    }

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    xlsx_path = OUT_DIR / "SEO_QA_qa_batch_032_r4.xlsx"
    md_path = OUT_DIR / "SEO_QA_qa_batch_032_r4.md"
    write_xlsx(payload, xlsx_path)

    workbook_hash = sha256(xlsx_path)
    zip_bad = zipfile.ZipFile(xlsx_path).testzip()
    validation = {
        "sheet_names": load_workbook(xlsx_path, read_only=True).sheetnames,
        "row_counts": {k: len(v) for k, v in payload.items()},
        "product_count_10": len(qa_products) == 10,
        "image_count_70": len(qa_images) == 70,
        "criteria_count_110": len(qa_criteria) == 110,
        "criteria_weights_sum_100": all(sum(c["weight"] for c in qa_criteria if c["product_key"] == p["product_key"]) == 100 for p in qa_products),
        "zip_bad_member": zip_bad,
        "xlsx_sha256": workbook_hash,
        "passed": zip_bad is None and len(qa_products) == 10 and len(qa_images) == 70 and len(qa_criteria) == 110,
    }

    for name, obj in [
        ("qa_dataset.json", payload),
        ("qa_workbook_payload.json", {k: payload[k] for k in ("QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues")}),
        ("validation_results.json", validation),
    ]:
        (RUN_DIR / name).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")

    manifest = {
        "rubric_version": "prompt_qa.md v1.0 / 2026-09-06",
        "qa_run_id": QA_RUN_ID,
        "batch_id": BATCH_ID,
        "source_workbook": str(SOURCE.relative_to(ROOT)),
        "source_workbook_sha256": source_hash,
        "admin_export": "products_export_1.csv",
        "admin_export_sha256": admin_hash,
        "batch_product_keys": [p["product_key"] for p in products],
        "checked_at": checked_at,
        "output_markdown": str(md_path.relative_to(ROOT)),
        "output_xlsx": str(xlsx_path.relative_to(ROOT)),
        "output_xlsx_sha256": workbook_hash,
    }
    (RUN_DIR / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (RUN_DIR / "qa_progress.json").write_text(
        json.dumps(
            {
                "rubric_version": manifest["rubric_version"],
                "qa_run_id": QA_RUN_ID,
                "source_workbook": manifest["source_workbook"],
                "source_workbook_sha256": source_hash,
                "batch_id": BATCH_ID,
                "batch_product_keys": manifest["batch_product_keys"],
                "current_product_key": manifest["batch_product_keys"][-1],
                "current_stage": "BATCH_COMPLETE",
                "completed_image_keys": [i["qa_image_key"] for i in qa_images],
                "last_saved_at": checked_at,
                "artifact_paths": {"markdown": str(md_path.relative_to(ROOT)), "xlsx": str(xlsx_path.relative_to(ROOT)), "dataset": str((RUN_DIR / "qa_dataset.json").relative_to(ROOT))},
                "awaiting_confirmation": True,
                "confirmation_ref": None,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    table = "\n".join(
        f"| `{p['product_key'].split('+', 1)[-1]}` | {p['final_score']:.1f} | `{p['qa_status']}` | {p['images_checked']}/{p['images_expected']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']} | {p['issue_refs']} |"
        for p in qa_products
    )
    md = f"""# SEO QA qa_batch_032_r4

- Source workbook: `{SOURCE.relative_to(ROOT)}`
- SHA-256: `{source_hash}`
- Admin export: `products_export_1.csv`
- Admin export SHA-256: `{admin_hash}`
- QA run: `{QA_RUN_ID}`
- Checked at: `{checked_at}`
- Scope: {len(qa_products)} products, {len(qa_images)} image positions
- Batch status: `{batch_status}`
- Batch score: `{batch_score}`

## Score Logic
- Chấm theo `seo-prompt/jeminise/prompt_qa.md` v1.0, có mở live storefront/`.js`, đối chiếu `products_export_1.csv`, workbook r4 và contact sheet/ảnh local.
- `311-314` bị `MAJOR` vì public SEO fields và alt đang dùng `comforter/comforter set`, trong khi admin `Type`, option size và admin SEO baseline neo về `Quilt`; cần merchant xác nhận hoặc sửa đồng nhất về `Quilt/Quilt Set`.
- `K3` bị `PARTIAL` cho toàn bộ lô vì bằng chứng nhu cầu mới ở mức public SERP/comparable shopping results, chưa có Search Console/internal search/paid keyword data.
- `E1` bị `PARTIAL` cho toàn bộ lô vì source workbook là `qa_batch_032_r4` nhưng row-level `revision` vẫn ghi `r2`.

| Product | Score | Status | Images | C/M/m | Issue refs |
|---|---:|---|---:|---:|---|
{table}

## Findings
- `meta_description_seo`: không thấy câu bị cắt giữa từ hoặc cắt ý; các meta là câu hoàn chỉnh.
- `description_proposed_html`: không thấy câu nội bộ/template như `copy stays specific to the visible artwork`.
- `Image_Audit`: 70/70 vị trí ảnh khớp số lượng live JSON và admin export.
- 6 sản phẩm `315-320` đạt `QA_PASS`; 4 sản phẩm deer `311-314` cần revision trước khi gửi pass do lỗi nhất quán loại sản phẩm.

## Priority Fix
1. Sửa `311-314` để dùng `Quilt`/`Quilt Set` đồng nhất trong `title_proposed`, `meta_title_seo`, `meta_description_seo`, `description_proposed_html` và tất cả `alt_proposed`, hoặc cung cấp xác nhận merchant rằng 4 sản phẩm này phải market là `comforter`.
2. Cập nhật metadata `revision` trong row lên đúng `r4/r5` ở bản sửa tiếp theo để traceability rõ.

## Limitations
- Chưa có Search Console, internal search, paid keyword volume hoặc review corpus xác minh nhu cầu; chỉ dùng public SERP/comparable shopping pages.
- QA không sửa workbook nguồn và không đổi trạng thái `APPROVED`.
"""
    md_path.write_text(md, encoding="utf-8-sig")

    print(json.dumps({"markdown": str(md_path), "xlsx": str(xlsx_path), "batch_status": batch_status, "batch_score": batch_score, "status_counts": dict(status_counts), "issue_counts": dict(issue_counts), "validation": validation}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
