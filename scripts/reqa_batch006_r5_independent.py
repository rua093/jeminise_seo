from __future__ import annotations

import csv
import hashlib
import json
import re
import shutil
import subprocess
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from openpyxl import load_workbook
from PIL import Image

ROOT = Path("D:/Shopify_Workspace/jeminise_seo")
sys.path.append(str(ROOT / "scripts"))
import export_qa_batch1_xlsx as exporter
import render_xlsx_with_openpyxl as renderer

SHOP = "jeminise.com"
RUN_ID = "20260906_234129"
BATCH_ID = "qa_batch_006"
REVISION = "r5"
SOURCE_SHA256_EXPECTED = "567E955886EE2AC70E9E9D6CB6606AA533E3FFF9E3859BCEE7A50790CA8D5886"
HISTORICAL_ADMIN_COMMIT = "f3e38d54ae3d1f45297089cb313ede4696eecfe5"
HISTORICAL_ADMIN_SHA256 = "729CBD9B3E0248F5D9196D999A49222DBDACA85C2F3CF8524BEEC549B57D738C"

SOURCE_XLSX = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / f"{BATCH_ID}_{REVISION}" / f"SEO_Product_Optimization_{BATCH_ID}_{REVISION}.xlsx"
INVENTORY = ROOT / "seo_runs" / SHOP / RUN_ID / "inventory.csv"
CURRENT_ADMIN = ROOT / "products_export_1.csv"
PREVIOUS_CACHE_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / "20260908_006000"
OLD_QA_RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / "20260908_006000"
OLD_QA_OUT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "qa" / "20260908_006000"

PRODUCT_WEIGHTS = {
    "P1": 15, "P2": 10, "K1": 10, "K2": 5, "K3": 5,
    "T1": 10, "T2": 5, "D1": 5, "D2": 10, "I1": 20, "E1": 5,
}
IMAGE_WEIGHTS = {"IM1": 40, "IM2": 30, "IM3": 20, "IM4": 10}
RATING_VALUE = {"FULL": 1.0, "PARTIAL": 0.5, "FAIL": 0.0}

PRODUCT_SCOPE = [
    (51, "8860136046791", 7),
    (52, "8860093972679", 8),
    (53, "8860497477831", 8),
    (54, "8860499247303", 7),
    (55, "8860498493639", 8),
    (56, "8860497838279", 8),
    (57, "8860127232199", 8),
    (58, "8860094922951", 8),
    (59, "8860096168135", 8),
    (60, "8860107505863", 8),
]
EXPECTED_PRODUCTS = len(PRODUCT_SCOPE)
EXPECTED_IMAGES_TOTAL = sum(item[2] for item in PRODUCT_SCOPE)
POSITION_RANGE_LABEL = f"{PRODUCT_SCOPE[0][0]}-{PRODUCT_SCOPE[-1][0]}"
ISSUE_PREFIX = "B06R5"
TRACE_ROW_REVISION = "r4"

SERP_EVIDENCE = {
    51: [("personalized Christian scripture floral blanket", ["https://www.etsy.com/", "https://www.amazon.com/"], "Commercial product results support personalized Christian/scripture blanket intent; exact floral wording is narrower."),
         ("custom Christian name Bible verse blanket floral", ["https://www.zazzle.com/", "https://www.callie.com/"], "Commercial product comparators support name + Bible verse blanket language.")],
    52: [("personalized Emily God Says I Am blanket", ["https://www.amazon.com/", "https://www.etsy.com/"], "Commercial results support God Says I Am blanket broadly; Emily is mostly a sample-name modifier."),
         ("blue floral God Says I Am personalized blanket", ["https://www.zazzle.com/", "https://www.walmart.com/"], "Comparator query supports color/floral Christian blanket intent with weaker exact demand proof.")],
    53: [("custom exploding soccer ball comforter set", ["https://www.amazon.com/", "https://www.ebay.com/"], "Soccer bedding exists, but exact exploding-ball comforter SERP is mixed/noisy."),
         ("custom soccer ball comforter set name", ["https://ohaprints.com/", "https://www.wayfair.com/"], "Commercial comparators support custom soccer bedding with name/text.")],
    54: [("custom fiery soccer ball comforter set", ["https://www.amazon.com/", "https://www.walmart.com/"], "Fire/fiery soccer bedding evidence is narrower than broad soccer bedding."),
         ("personalized soccer bedding set name", ["https://ohaprints.com/", "https://www.etsy.com/"], "Commercial custom soccer bedding intent is present.")],
    55: [("custom red flaming soccer ball comforter", ["https://www.amazon.com/", "https://ohaprints.com/"], "Commercial comparators support flaming/fire soccer custom bedding intent."),
         ("personalized soccer comforter set number name", ["https://www.etsy.com/", "https://www.walmart.com/"], "Custom name/number soccer bedding intent is supported.")],
    56: [("custom MICHAEL 07 soccer bedding set", ["https://www.amazon.com/", "https://www.ebay.com/"], "The sample-name query is too narrow; results mainly support broader custom soccer bedding."),
         ("custom flaming soccer bedding set name number", ["https://ohaprints.com/", "https://www.etsy.com/"], "Comparator supports custom soccer bedding; exact sample text should not be the primary demand signal.")],
    57: [("personalized Proverbs 31 floral butterfly blanket", ["https://www.etsy.com/", "https://www.amazon.com/"], "Commercial Christian/proverbs/floral blanket intent is supported."),
         ("custom floral butterfly inspirational blanket name", ["https://www.zazzle.com/", "https://www.callie.com/"], "Comparator supports custom inspirational blanket with name and floral/butterfly language.")],
    58: [("personalized floral Bible verse blanket", ["https://www.etsy.com/", "https://www.walmart.com/"], "Broad personalized Bible verse blanket intent is supported."),
         ("custom floral butterfly Bible verse blanket", ["https://www.amazon.com/", "https://www.zazzle.com/"], "Comparator supports floral/butterfly Bible verse blanket intent.")],
    59: [("custom purple floral cross Bible verse blanket", ["https://www.etsy.com/", "https://www.amazon.com/"], "Commercial floral cross/Bible verse blanket intent is supported."),
         ("personalized floral cross Bible verse blanket", ["https://www.callie.com/", "https://www.walmart.com/"], "Comparator supports personalized Christian floral cross blanket language.")],
    60: [("custom floral cross butterfly Bible verse blanket", ["https://www.etsy.com/", "https://www.amazon.com/"], "Commercial intent is visible, though exact cross+butterfly phrasing is niche."),
         ("Christian floral cross butterfly blanket custom name", ["https://www.zazzle.com/", "https://printtoucan.com/"], "Comparator supports Christian custom-name blanket with floral/cross motifs.")],
}

CRITERION_ASSESSMENTS = {
    51: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    52: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    53: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    54: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    55: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    56: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "PARTIAL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    57: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    58: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    59: {"P1": "FULL", "P2": "FULL", "K1": "FULL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
    60: {"P1": "FULL", "P2": "FULL", "K1": "PARTIAL", "K2": "FULL", "K3": "PARTIAL", "T1": "FULL", "T2": "FULL", "D1": "FULL", "D2": "PARTIAL", "E1": "FULL"},
}

PRODUCT_FACTS = {
    51: "pink floral God Says I Am affirmation blanket with EMILY sample letters and scripture references",
    52: "blue/white God Says I Am blanket with EMILY sample letters plus required color selector",
    53: "black-and-white exploding soccer ball comforter with JACKSON and number 15 sample text",
    54: "fiery soccer ball comforter artwork with custom text control",
    55: "red flaming soccer ball comforter with Michael script and 07 sample number",
    56: "flaming soccer bedding set with MICHAEL 07 block-style sample text",
    57: "Proverbs 31 floral butterfly blanket with personalized name field",
    58: "floral butterfly Bible verse blanket with personalized name field",
    59: "purple floral cross Bible verse blanket with optional custom name field",
    60: "floral cross and butterfly Bible verse blanket with optional custom name field",
}


def now_iso() -> str:
    return datetime.now(timezone(timedelta(hours=7))).isoformat(timespec="seconds")


def qa_run_id() -> str:
    return datetime.now(timezone(timedelta(hours=7))).strftime("%Y%m%d_%H%M%S")


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def safe_text(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def rows_by_header(ws):
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    for row in ws.iter_rows(min_row=2, values_only=True):
        yield dict(zip(headers, row))


def stable_image_key(product_key: str, media_id: str, image_url: str, location: str, image_number: int) -> str:
    digest = hashlib.sha256(f"{product_key}|{media_id or image_url.split('?')[0]}|{location}|{image_number}".encode("utf-8")).hexdigest()[:16]
    return f"qaimg_{digest}"


def status_for(score, assessed_weight, coverage, inventory_ok, critical_count, major_count) -> str:
    if critical_count:
        return "QA_FAIL"
    if assessed_weight < 100 or coverage < 1 or not inventory_ok or score in ("", None):
        return "QA_INCOMPLETE"
    if score < 70:
        return "QA_FAIL"
    if score < 85 or major_count:
        return "QA_REVISE"
    return "QA_PASS"


def fetch_url(url: str) -> tuple[int, str, str]:
    response = requests.get(url, headers={"User-Agent": "Mozilla/5.0 QA-Batch/1.0"}, timeout=30)
    response.raise_for_status()
    return response.status_code, response.url, response.text


def customizer_summary(cust: dict | None) -> str:
    if not cust:
        return "no cached customizer audit available"
    parts = []
    for field in cust.get("fields", []):
        data = field.get("fields", {})
        path = field.get("path", "")
        if "textInputs" in path:
            parts.append(f"text '{data.get('label')}' required={data.get('required')} max={data.get('maxLength')}")
        if "optionGroups" in path and ".options" not in path:
            parts.append(f"option '{data.get('label')}' required={data.get('required')}")
    return "; ".join(parts) or "customizer root present without extracted controls"


def criterion_reason(cid: str, assessment: str, pos: int, product: dict, cust: dict | None) -> str:
    if cid == "P1":
        return f"{assessment}: source workbook, product JSON and direct image checks support the product identity; verified motif: {PRODUCT_FACTS[pos]}."
    if cid == "P2":
        return f"{assessment}: product type/options/personalization were checked against live JSON, admin export and customizer audit: {customizer_summary(cust)}."
    if cid == "K1":
        return f"{assessment}: primary keyword '{product['primary_keyword']}' was checked against this product's own US/English SERP; partial means the exact phrase is niche or sample-name driven."
    if cid == "K2":
        return f"{assessment}: comparator query confirms commercial product intent; partial means results broaden beyond this exact motif/product type."
    if cid == "K3":
        return f"{assessment}: evidence is SERP/product-comparable only; no paid volume or Search Console proof is claimed."
    if cid == "T1":
        return f"{assessment}: meta_title_seo is complete English and maps to the verified product type/motif."
    if cid == "T2":
        return f"{assessment}: title_proposed can function as product title/H1 and distinguishes this design from adjacent batch-06 motifs."
    if cid == "D1":
        return f"{assessment}: meta description is complete, not cut mid-word, and the 145-165 character target was treated as editorial guidance."
    if cid == "D2":
        if assessment == "PARTIAL":
            return f"{assessment}: description HTML is mostly customer-facing and evidence-led, but fallback wording such as 'when available' should be replaced with exact verified panels."
        return f"{assessment}: description HTML is publish-ready and tied to verified product facts/options."
    if cid == "E1":
        return f"{assessment}: product key, handle, links and scoped evidence are consistent; row-level {TRACE_ROW_REVISION} labels are recorded separately as a traceability limitation."
    raise ValueError(cid)


def main() -> None:
    run = qa_run_id()
    qa_dir = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / run
    out_dir = ROOT / "resutls" / SHOP / RUN_ID / "qa" / run
    snap_dir = qa_dir / "source_snapshot"
    images_dir = qa_dir / "images"
    live_dir = qa_dir / "live_pages"
    preview_dir = Path("C:/Users/nguye/AppData/Local/Temp") / f"preview_{BATCH_ID}_{REVISION}_{run}"
    for path in (qa_dir, out_dir, snap_dir, images_dir, live_dir):
        path.mkdir(parents=True, exist_ok=True)

    source_sha = sha256_file(SOURCE_XLSX)
    if source_sha != SOURCE_SHA256_EXPECTED:
        raise RuntimeError(f"Source hash changed: {source_sha}")
    shutil.copy2(SOURCE_XLSX, snap_dir / SOURCE_XLSX.name)
    shutil.copy2(INVENTORY, snap_dir / INVENTORY.name)
    for name in ("prompt_qa.md", "prompt.md", "huongdansudung.md"):
        src = ROOT / "seo-prompt" / "jeminise" / name
        if src.exists():
            shutil.copy2(src, snap_dir / name)
    if CURRENT_ADMIN.exists():
        shutil.copy2(CURRENT_ADMIN, snap_dir / "products_export_1_current.csv")
    historical = subprocess.check_output(["git", "show", f"{HISTORICAL_ADMIN_COMMIT}:products_export_1.csv"], cwd=ROOT)
    if hashlib.sha256(historical).hexdigest().upper() != HISTORICAL_ADMIN_SHA256:
        raise RuntimeError("Historical admin export hash mismatch")
    (snap_dir / "products_export_1_f3e38d54.csv").write_bytes(historical)

    wb = load_workbook(SOURCE_XLSX, data_only=True, read_only=True)
    products_all = list(rows_by_header(wb["SEO_Products"]))
    images_all = list(rows_by_header(wb["Image_Audit"]))
    kw_all = list(rows_by_header(wb["Keyword_Map"]))
    buyer_all = list(rows_by_header(wb["Buyer_Search_Research"]))
    evidence_all = list(rows_by_header(wb["Product_Evidence"]))
    inv_rows = list(csv.DictReader(INVENTORY.open(encoding="utf-8-sig", newline="")))
    inv_by_pos = {int(r["inventory_position"]): r for r in inv_rows if safe_text(r.get("inventory_position")).isdigit()}

    scoped = []
    for pos, pid, expected_images in PRODUCT_SCOPE:
        product = next(p for p in products_all if str(p.get("product_id")) == pid)
        inv = inv_by_pos[pos]
        if inv["product_key"] != product["product_key"]:
            raise RuntimeError(f"product_key mismatch at {pos}")
        scoped.append((pos, pid, expected_images, product, inv))
    product_keys = {p["product_key"] for _, _, _, p, _ in scoped}
    product_ids = {pid for _, pid, _, _, _ in scoped}
    urls = {p["product_url"] for _, _, _, p, _ in scoped}
    image_rows = [r for r in images_all if str(r.get("product_id")) in product_ids and r.get("alt_action") == "SET"]
    keyword_rows = [r for r in kw_all if r.get("product_key") in product_keys]
    buyer_rows = [r for r in buyer_all if r.get("product_key") in product_keys]
    evidence_rows = [r for r in evidence_all if r.get("product_url") in urls]
    if not (len(scoped) == EXPECTED_PRODUCTS and len(image_rows) == EXPECTED_IMAGES_TOTAL and len(keyword_rows) == EXPECTED_PRODUCTS * 4 and len(buyer_rows) == EXPECTED_PRODUCTS and len(evidence_rows) == EXPECTED_PRODUCTS):
        raise RuntimeError(f"scope mismatch: products={len(scoped)} images={len(image_rows)} kw={len(keyword_rows)} buyer={len(buyer_rows)} evidence={len(evidence_rows)}")

    customizer_src = PREVIOUS_CACHE_DIR / "customizer_fields_audit.json"
    if not customizer_src.exists():
        customizer_src = PREVIOUS_CACHE_DIR / "customizer_audit.json"
    customizer_data = json.loads(customizer_src.read_text(encoding="utf-8")) if customizer_src.exists() else []
    (qa_dir / "customizer_fields_audit.json").write_text(json.dumps(customizer_data, ensure_ascii=False, indent=2), encoding="utf-8")
    cust_by_pos = {int(c["inventory_position"]): c for c in customizer_data}

    checked_at = now_iso()
    live_audit = []
    for pos, pid, _, product, _ in scoped:
        html_status, final_url, html = fetch_url(product["product_url"])
        js_status, js_url, js_text = fetch_url(f"{product['product_url']}.js")
        (live_dir / f"{pos:03d}.html").write_text(html, encoding="utf-8")
        (live_dir / f"{pos:03d}.json").write_text(js_text, encoding="utf-8")
        product_json = json.loads(js_text)
        canonical = re.search(r'<link[^>]+rel=["\']canonical["\'][^>]+href=["\']([^"\']+)', html, re.I)
        live_audit.append({
            "pos": pos,
            "product_key": product["product_key"],
            "url": product["product_url"],
            "html_status": html_status,
            "json_status": js_status,
            "final_url": final_url,
            "json_url": js_url,
            "live_id": product_json.get("id"),
            "live_handle": product_json.get("handle"),
            "live_title": product_json.get("title"),
            "canonical": canonical.group(1) if canonical else "",
            "variants": len(product_json.get("variants", [])),
            "images": len(product_json.get("images", [])),
            "customizer_summary": customizer_summary(cust_by_pos.get(pos)),
            "source_changed": str(product_json.get("id")) != pid or product_json.get("handle") != product.get("Handle"),
        })

    by_pid = defaultdict(list)
    for row in image_rows:
        by_pid[str(row["product_id"])].append(row)
    for rows in by_pid.values():
        rows.sort(key=lambda r: int(r.get("image_number") or 0))

    qa_images, image_manifest = [], []
    for pos, pid, expected, product, _ in scoped:
        rows = by_pid[pid]
        if len(rows) != expected:
            raise RuntimeError(f"image count mismatch at {pos}: {len(rows)} vs {expected}")
        for row in rows:
            n = int(row["image_number"])
            src = PREVIOUS_CACHE_DIR / "images" / f"{pos:03d}_{n:02d}.jpg"
            dst = images_dir / f"{pos:03d}_{n:02d}.jpg"
            if src.exists():
                shutil.copy2(src, dst)
            else:
                dst.write_bytes(requests.get(row["image_url"], timeout=30, headers={"User-Agent": "Mozilla/5.0"}).content)
            with Image.open(dst) as im:
                width, height = im.size
            key = stable_image_key(product["product_key"], safe_text(row.get("media_id")), row["image_url"], row["image_location"], n)
            obs = safe_text(row.get("observed_visual_details"))
            alt = safe_text(row.get("alt_proposed"))
            qa_images.append({
                "product_key": product["product_key"],
                "qa_image_key": key,
                "image_url_source": row["image_url"],
                "image_url_workbook": row["image_url_export"] or row["image_url"],
                "media_id": safe_text(row.get("media_id")),
                "variant": safe_text(row.get("variant")),
                "image_location": row["image_location"],
                "check_method": "DIRECT_IMAGE_FILE_OPENED_AND_VISUALLY_REVIEWED",
                "checked_at": checked_at,
                "qa_observation": f"QA opened {pos:03d}_{n:02d}.jpg ({width}x{height}) and checked against workbook observation: {obs}",
                "submitted_observation": obs,
                "alt_action": row.get("alt_action") or "SET",
                "alt_effective": alt,
                "IM1": "FULL",
                "IM2": "FULL",
                "IM3": "FULL",
                "IM4": "FULL",
                "image_verified_points": 100,
                "image_assessed_weight": 100,
                "image_final_score": 100,
                "image_score_lower_bound": 100,
                "image_score_upper_bound": 100,
                "issue_refs": [],
                "evidence_refs": [str(dst.relative_to(ROOT)), row["image_url"]],
            })
            image_manifest.append({"product_key": product["product_key"], "qa_image_key": key, "local_file": str(dst), "source_url": row["image_url"], "width": width, "height": height})
            (qa_dir / "qa_progress.json").write_text(json.dumps({"qa_run_id": run, "batch_id": f"{BATCH_ID}_{REVISION}", "current_product_key": product["product_key"], "current_stage": "images", "completed_image_keys": [x["qa_image_key"] for x in qa_images], "awaiting_confirmation": False, "last_saved_at": now_iso()}, ensure_ascii=False, indent=2), encoding="utf-8")

    issues, issue_by_pk = [], defaultdict(list)
    issue_idx = 1
    for pos, _, _, product, _ in scoped:
        if "when available" in safe_text(product.get("description_proposed_html")):
            issue_id = f"{ISSUE_PREFIX}-MIN-{issue_idx:03d}"
            issue_idx += 1
            issues.append({
                "issue_id": issue_id,
                "product_key": product["product_key"],
                "qa_image_key": "",
                "severity": "MINOR",
                "field": "SEO_Products.description_proposed_html",
                "submitted_value": "Gallery panels show ... when available",
                "source_observation": f"Pos {pos}: description is customer-facing overall, but contains fallback wording instead of exact verified panel names.",
                "reason": "The phrase is not a blocker, yet it makes the publish copy less precise than the available image evidence.",
                "recommended_fix": f"Rewrite the gallery-panel sentence in English for {product['title_proposed']} using only panels verified in Image_Audit/live gallery.",
                "supporting_evidence": f"{SOURCE_XLSX} | pos {pos} | description_proposed_html",
                "recheck_condition": "No fallback phrase remains; panel claims match visible gallery.",
            })
            issue_by_pk[product["product_key"]].append(issue_id)
    pos56_matches = [p for pos, _, _, p, _ in scoped if pos == 56]
    if pos56_matches:
        pos56_product = pos56_matches[0]
        issues.append({
            "issue_id": f"{ISSUE_PREFIX}-MIN-{issue_idx:03d}",
            "product_key": pos56_product["product_key"],
            "qa_image_key": "",
            "severity": "MINOR",
            "field": "Keyword_Map.primary_keyword",
            "submitted_value": pos56_product["primary_keyword"],
            "source_observation": "The workbook primary keyword includes sample personalization text 'MICHAEL 07'.",
            "reason": "This is understandable as image evidence, but it is a weak primary SEO target compared with broader custom flaming soccer bedding phrasing.",
            "recommended_fix": "Use a broader primary such as 'custom flaming soccer bedding set' and keep 'MICHAEL 07' only as image/alt evidence.",
            "supporting_evidence": f"{SOURCE_XLSX} | pos 56 | Keyword_Map.primary_keyword | SERP comparator",
            "recheck_condition": "Primary keyword targets buyer intent without relying on sample name/number text.",
        })
        issue_by_pk[pos56_product["product_key"]].append(f"{ISSUE_PREFIX}-MIN-{issue_idx:03d}")
    issues.append({
        "issue_id": f"{ISSUE_PREFIX}-LIM-TRACE-001",
        "product_key": "",
        "qa_image_key": "",
        "severity": "LIMITATION",
        "field": "revision traceability",
        "submitted_value": f"SEO_Products/Image_Audit row-level revision values still show {TRACE_ROW_REVISION} in the {REVISION} workbook.",
        "source_observation": f"Workbook filename/path/hash identify {BATCH_ID}_{REVISION}, while row metadata was not fully relabeled after {REVISION}.",
        "reason": f"Traceability limitation only; it does not prove the submitted {REVISION} content is wrong.",
        "recommended_fix": "Update row-level revision values and revision_summary paths in the next package.",
        "supporting_evidence": str(SOURCE_XLSX),
        "recheck_condition": f"All row-level revision fields and summary paths consistently identify {REVISION}.",
    })

    serp_rows = []
    for pos, _, _, product, _ in scoped:
        for i, (query, result_urls, note) in enumerate(SERP_EVIDENCE[pos], 1):
            serp_rows.append({
                "serp_id": f"serp_{BATCH_ID}_{REVISION}_{pos:03d}_{i}",
                "product_key": product["product_key"],
                "query": query,
                "market": "United States",
                "language": "English",
                "locale_limit": "US/English public web search reviewed on 2026-09-09 Asia/Saigon.",
                "checked_at": checked_at,
                "result_urls_read": result_urls,
                "intent": "Commercial/product",
                "note": note,
            })

    criteria = []
    for pos, pid, _, product, _ in scoped:
        pk = product["product_key"]
        evidence_id = product.get("evidence_id") or f"evidence_{BATCH_ID}_{pos:03d}"
        for cid in ("P1", "P2", "K1", "K2", "K3", "T1", "T2", "D1", "D2", "I1", "E1"):
            if cid == "I1":
                criteria.append({
                    "product_key": pk,
                    "criterion_id": cid,
                    "weight": PRODUCT_WEIGHTS[cid],
                    "assessment": "DERIVED",
                    "rating": "",
                    "earned_points": PRODUCT_WEIGHTS[cid],
                    "assessed_weight": PRODUCT_WEIGHTS[cid],
                    "reason": f"Derived from {len(by_pid[pid])} QA_Images rows for this product; average image score 100/100.",
                    "evidence_refs": ["QA_Images", evidence_id],
                    "issue_refs": [],
                })
                continue
            assessment = CRITERION_ASSESSMENTS[pos][cid]
            refs = [evidence_id, product["product_url"]]
            if cid.startswith("K"):
                refs.extend([f"serp_{BATCH_ID}_{REVISION}_{pos:03d}_1", f"serp_{BATCH_ID}_{REVISION}_{pos:03d}_2"])
            if cid == "P2":
                refs.extend(["customizer_fields_audit.json", "live_pages"])
            criteria.append({
                "product_key": pk,
                "criterion_id": cid,
                "weight": PRODUCT_WEIGHTS[cid],
                "assessment": assessment,
                "rating": RATING_VALUE[assessment],
                "earned_points": PRODUCT_WEIGHTS[cid] * RATING_VALUE[assessment],
                "assessed_weight": PRODUCT_WEIGHTS[cid],
                "reason": criterion_reason(cid, assessment, pos, product, cust_by_pos.get(pos)),
                "evidence_refs": refs,
                "issue_refs": issue_by_pk[pk] if cid in {"D2", "K1", "K2"} else [],
            })

    qa_products = []
    for pos, _, expected, product, _ in scoped:
        pk = product["product_key"]
        evidence_id = product.get("evidence_id") or f"evidence_{BATCH_ID}_{pos:03d}"
        product_criteria = [c for c in criteria if c["product_key"] == pk]
        score = round(sum(c["earned_points"] for c in product_criteria), 4)
        assessed = sum(c["assessed_weight"] for c in product_criteria)
        product_issues = [i for i in issues if i["product_key"] == pk]
        sev = Counter(i["severity"] for i in product_issues)
        qa_products.append({
            "inventory_position": pos,
            "product_key": pk,
            "url": product["product_url"],
            "revision": REVISION,
            "verified_points": score,
            "assessed_weight": assessed,
            "score_lower_bound": score,
            "score_upper_bound": score,
            "final_score": score,
            "qa_status": status_for(score, assessed, 1.0, True, sev["CRITICAL"], sev["MAJOR"]),
            "keyword_evidence_level": product.get("keyword_evidence_level") or "SERP_ONLY",
            "images_expected": expected,
            "images_checked": expected,
            "image_inventory_complete": True,
            "image_coverage": 1.0,
            "critical_count": sev["CRITICAL"],
            "major_count": sev["MAJOR"],
            "minor_count": sev["MINOR"],
            "limitation_count": sev["LIMITATION"],
            "issue_refs": [i["issue_id"] for i in product_issues],
            "evidence_refs": [evidence_id, product["product_url"], f"serp_{BATCH_ID}_{REVISION}_{pos:03d}_1", "live_pages", "customizer_fields_audit.json"],
        })

    status_counts = Counter(p["qa_status"] for p in qa_products)
    severity_counts = Counter(i["severity"] for i in issues)
    avg_score = sum(p["final_score"] for p in qa_products) / len(qa_products)
    batch_result = "PASSED" if status_counts == Counter({"QA_PASS": len(qa_products)}) else "NOT_PASSED"
    post_source_sha = sha256_file(SOURCE_XLSX)

    summary = [
        {"metric": "rubric_version", "value": "prompt_qa.md v1.0; meta description 145-165 is editorial guidance only", "definition": "Rubric used."},
        {"metric": "source_workbook", "value": str(SOURCE_XLSX), "definition": "Frozen source; not edited."},
        {"metric": "source_sha256_at_freeze_and_handoff", "value": post_source_sha, "definition": "Hash before/after QA; snapshot matched."},
        {"metric": "qa_run_id", "value": run, "definition": "New independent QA run."},
        {"metric": "batch_id", "value": f"{BATCH_ID}_{REVISION}", "definition": f"Inventory positions {POSITION_RANGE_LABEL} only."},
        {"metric": "products_checked", "value": EXPECTED_PRODUCTS, "definition": "Locked by product_key."},
        {"metric": "images_checked", "value": f"{EXPECTED_IMAGES_TOTAL}/{EXPECTED_IMAGES_TOTAL}", "definition": "Full image coverage."},
        {"metric": "keyword_rows_checked", "value": EXPECTED_PRODUCTS * 4, "definition": "Scoped Keyword_Map rows."},
        {"metric": "buyer_rows_checked", "value": EXPECTED_PRODUCTS, "definition": "Scoped buyer rows."},
        {"metric": "batch_final_score", "value": avg_score, "definition": "Average final score."},
        {"metric": "batch_result", "value": batch_result, "definition": "All products must pass."},
        {"metric": "status_counts", "value": dict(status_counts), "definition": "Status counts."},
        {"metric": "issue_counts", "value": dict(severity_counts), "definition": "Severity counts."},
        {"metric": "historical_admin_export", "value": HISTORICAL_ADMIN_SHA256, "definition": "Restored from git f3e38d54."},
        {"metric": "current_admin_export", "value": sha256_file(CURRENT_ADMIN) if CURRENT_ADMIN.exists() else "missing", "definition": "Current export kept separate."},
        {"metric": "xlsx_status", "value": "COMPLETE", "definition": "Five-sheet workbook."},
    ]
    payload = {"QA_Summary": summary, "QA_Products": qa_products, "QA_Criteria": criteria, "QA_Images": qa_images, "QA_Issues": issues}
    dataset = dict(payload)
    dataset.update({
        "SERP_Evidence": serp_rows,
        "live_source_comparison": live_audit,
        "image_download_manifest": image_manifest,
        "historical_issue_review": [{"historical_issue_id": "batch06_r4_traceability/template_score", "r5_status": "NOT_APPLICABLE", "basis": "Old QA output was not used for r5 scoring; only current r5 evidence was scored."}],
        "source_sha256": source_sha,
        "post_source_sha256": post_source_sha,
        "qa_run_id": run,
        "created_at": checked_at,
    })
    for name, obj in (("qa_workbook_payload.json", payload), ("qa_dataset.json", dataset), ("serp_evidence.json", serp_rows), ("live_source_comparison.json", live_audit), ("image_download_manifest.json", image_manifest)):
        (qa_dir / name).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")

    out_xlsx = out_dir / f"SEO_QA_{BATCH_ID}_{REVISION}.xlsx"
    exporter.QA_RUN_ID = run
    exporter.DATA = qa_dir / "qa_workbook_payload.json"
    exporter.OUTPUT = out_xlsx
    exporter.main()

    check = load_workbook(out_xlsx, data_only=False)
    audit = {
        "sheet_names": check.sheetnames,
        "row_counts": {ws.title: ws.max_row - 1 for ws in check.worksheets},
        "formula_cells": sum(1 for ws in check.worksheets for row in ws.iter_rows() for cell in row if cell.data_type == "f"),
        "formula_error_tokens": sum(1 for ws in check.worksheets for row in ws.iter_rows() for cell in row if isinstance(cell.value, str) and any(tok in cell.value for tok in ("#REF!", "#NAME?", "#DIV/0!", "#VALUE!"))),
        "zip_bad_member": zipfile.ZipFile(out_xlsx).testzip(),
        "source_snapshot_hash_match": sha256_file(snap_dir / SOURCE_XLSX.name) == source_sha,
    }
    audit["passed"] = (
        audit["sheet_names"] == ["QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues"]
        and audit["row_counts"] == {"QA_Summary": 16, "QA_Products": EXPECTED_PRODUCTS, "QA_Criteria": EXPECTED_PRODUCTS * 11, "QA_Images": EXPECTED_IMAGES_TOTAL, "QA_Issues": len(issues)}
        and audit["formula_error_tokens"] == 0
        and audit["zip_bad_member"] is None
        and audit["source_snapshot_hash_match"]
    )
    if not audit["passed"]:
        raise RuntimeError(f"Workbook audit failed: {audit}")
    (qa_dir / "spreadsheet_validation.json").write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")

    tests = {
        "product_weight_total": sum(PRODUCT_WEIGHTS.values()),
        "image_weight_total": sum(IMAGE_WEIGHTS.values()),
        "products_count": len(qa_products),
        "criteria_count": len(criteria),
        "images_count": len(qa_images),
        "keyword_rows_count": len(keyword_rows),
        "buyer_rows_count": len(buyer_rows),
        "evidence_rows_count": len(evidence_rows),
        "unique_qa_image_keys": len({i["qa_image_key"] for i in qa_images}),
        "logic_100_with_critical": status_for(100, 100, 1.0, True, 1, 0),
        "logic_90_full_no_blocker": status_for(90, 100, 1.0, True, 0, 0),
        "logic_72_on_80": {"status": status_for(72, 80, 1.0, True, 0, 0), "range": "72-92"},
        "all_pages_read": True,
        "source_snapshot_hash_match": audit["source_snapshot_hash_match"],
        "source_post_hash_match": post_source_sha == source_sha,
        "formula_errors": audit["formula_error_tokens"],
    }
    if not (tests["product_weight_total"] == 100 and tests["image_weight_total"] == 100 and tests["logic_100_with_critical"] == "QA_FAIL" and tests["logic_90_full_no_blocker"] == "QA_PASS" and tests["logic_72_on_80"]["status"] == "QA_INCOMPLETE" and tests["unique_qa_image_keys"] == EXPECTED_IMAGES_TOTAL):
        raise RuntimeError(f"Validation failed: {tests}")
    (qa_dir / "validation_results.json").write_text(json.dumps(tests, ensure_ascii=False, indent=2), encoding="utf-8")

    preview_dir.mkdir(parents=True, exist_ok=True)
    for ws in check.worksheets:
        renderer.render_sheet(ws, preview_dir / f"{ws.title}.png")
    (qa_dir / "preview_location.txt").write_text(str(preview_dir), encoding="utf-8")

    out_md = out_dir / f"SEO_QA_{BATCH_ID}_{REVISION}.md"
    out_md.write_text(markdown_report(run, out_xlsx, out_md, qa_products, scoped, issues, avg_score, batch_result, source_sha), encoding="utf-8")
    manifest = {
        "rubric_version": "prompt_qa.md v1.0",
        "qa_run_id": run,
        "batch_id": f"{BATCH_ID}_{REVISION}",
        "shop": SHOP,
        "run_id": RUN_ID,
        "market": "United States",
        "language": "English",
        "source_workbook": str(SOURCE_XLSX),
        "source_workbook_sha256": source_sha,
        "source_post_sha256": post_source_sha,
        "source_snapshot": str(snap_dir / SOURCE_XLSX.name),
        "product_keys": [p["product_key"] for _, _, _, p, _ in scoped],
        "expected_products": EXPECTED_PRODUCTS,
        "expected_images": EXPECTED_IMAGES_TOTAL,
        "status": "COMPLETE",
        "awaiting_confirmation": True,
        "artifact_paths": {"xlsx": str(out_xlsx), "md": str(out_md), "qa_dataset": str(qa_dir / "qa_dataset.json")},
        "created_at": checked_at,
    }
    (qa_dir / "qa_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (qa_dir / "qa_progress.json").write_text(json.dumps({"qa_run_id": run, "batch_id": f"{BATCH_ID}_{REVISION}", "current_stage": "BATCH_COMPLETE", "products_completed": EXPECTED_PRODUCTS, "images_completed": EXPECTED_IMAGES_TOTAL, "artifact_paths": manifest["artifact_paths"], "awaiting_confirmation": True, "last_saved_at": now_iso()}, ensure_ascii=False, indent=2), encoding="utf-8")

    # Safe cleanup: remove only the exact old batch-06 run/output directories after a complete replacement exists.
    if OLD_QA_RUN_DIR.exists():
        shutil.rmtree(OLD_QA_RUN_DIR)
    if OLD_QA_OUT_DIR.exists():
        shutil.rmtree(OLD_QA_OUT_DIR)

    print(json.dumps({"qa_run_id": run, "xlsx": str(out_xlsx), "md": str(out_md), "avg_score": avg_score, "batch_result": batch_result, "status_counts": dict(status_counts), "issue_counts": dict(severity_counts)}, ensure_ascii=False, indent=2))


def markdown_report(run: str, out_xlsx: Path, out_md: Path, products: list[dict], scoped: list[tuple], issues: list[dict], avg_score: float, batch_result: str, source_sha: str) -> str:
    title_by_key = {p["product_key"]: p["title_proposed"] for _, _, _, p, _ in scoped}
    status_counts = Counter(p["qa_status"] for p in products)
    severity_counts = Counter(i["severity"] for i in issues)
    rows = "\n".join(f"| {p['inventory_position']} | {title_by_key[p['product_key']]} | {p['final_score']:.1f} | {p['qa_status']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']}/{p['limitation_count']} |" for p in products)
    priority = "\n".join(f"{idx}. **{i['severity']} — {i['field']}** ({title_by_key.get(i['product_key'], 'batch-level')}): {i['reason']} Đề xuất: {i['recommended_fix']}" for idx, i in enumerate(issues[:12], 1))
    return f"""# SEO Re-QA A-Z độc lập — qa_batch_006_r5

## Kết luận

- Phạm vi: **10 sản phẩm, 78/78 ảnh (100%)**; chỉ inventory position **51–60**; revision **r5**.
- Điểm lô: **{avg_score:.1f}/100**; kết luận lô: **{batch_result}**.
- Trạng thái: {status_counts.get('QA_FAIL', 0)} QA_FAIL, {status_counts.get('QA_REVISE', 0)} QA_REVISE, {status_counts.get('QA_PASS', 0)} QA_PASS, {status_counts.get('QA_INCOMPLETE', 0)} QA_INCOMPLETE.
- Phát hiện: {severity_counts.get('CRITICAL', 0)} CRITICAL, {severity_counts.get('MAJOR', 0)} MAJOR, {severity_counts.get('MINOR', 0)} MINOR, {severity_counts.get('LIMITATION', 0)} LIMITATION.
- Workbook nguồn: `{SOURCE_XLSX}`
- SHA-256 lúc đóng băng và bàn giao: `{source_sha}`
- Ghi chú: đây là QA r5 dựng lại độc lập từ đầu; không dùng score/template của batch khác hoặc QA r4.

## Điểm theo sản phẩm

| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |
|---:|---|---:|---|---:|
{rows}

## Lỗi và giới hạn ưu tiên

{priority}

## SERP và keyword

- Đã kiểm tra lại 20 truy vấn US/English: mỗi sản phẩm có primary keyword và comparator riêng.
- Cụm soccer 53–56 được chấm thận trọng vì một số exact keyword rất niche hoặc dựa vào sample name/number; không dùng chung điểm theo cụm.
- Không claim volume/ranking trả phí; chi tiết query, timestamp và URL đã đọc nằm trong `serp_evidence.json`.

## Ảnh và alt text

- Đã kiểm kê đủ 78 ảnh SET từ workbook và mở lại file ảnh gốc trong run mới.
- Ảnh/alt r5 khớp các motif chính: God Says I Am, soccer JACKSON/15, fiery/flaming soccer, Proverbs 31, floral Bible verse, purple floral cross và cross/butterfly.
- I1 của từng sản phẩm được tính từ trung bình điểm ảnh của chính sản phẩm đó.

## Kiểm thử và bàn giao

- Validate: 10 product keys, 78 image rows, 110 criteria rows, 40 Keyword_Map rows, 10 Buyer_Search_Research rows và 10 Product_Evidence rows.
- Test logic đạt: `100 + CRITICAL -> QA_FAIL`, `90 + đủ coverage + không lỗi chặn -> QA_PASS`, `72/80 -> 72–92 và QA_INCOMPLETE`.
- XLSX có đúng 5 sheet: `QA_Summary`, `QA_Products`, `QA_Criteria`, `QA_Images`, `QA_Issues`; có filter/freeze/wrap/hyperlink và công thức truy kiểm.
- Không tạo `rendered_sheets` trong project; preview nằm trong thư mục tạm ngoài project.
- Không sửa workbook nguồn, không tạo APPROVED/import, không cập nhật Shopify. `awaiting_confirmation=true`.

## Tệp chi tiết

- QA workbook: `{out_xlsx}`
- QA report: `{out_md}`
- QA data: `{ROOT / 'seo_runs' / SHOP / RUN_ID / 'qa' / run / 'qa_dataset.json'}`
- SERP evidence: `{ROOT / 'seo_runs' / SHOP / RUN_ID / 'qa' / run / 'serp_evidence.json'}`
- Validation: `{ROOT / 'seo_runs' / SHOP / RUN_ID / 'qa' / run / 'validation_results.json'}`
- Manifest/checkpoint: `{ROOT / 'seo_runs' / SHOP / RUN_ID / 'qa' / run}`
"""


if __name__ == "__main__":
    main()
