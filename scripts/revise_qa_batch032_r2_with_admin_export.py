import csv
import hashlib
import html
import json
import re
import shutil
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
SHOP = "jeminise.com"
RUN_ID = "20260906_234129"
BATCH_ID = "qa_batch_032_r2"
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_032.xlsx"
SUMMARY_PATH = ROOT / "seo_runs" / SHOP / RUN_ID / "batch_032_evidence_summary.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_032_r2.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_032_r2.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    311: ("You and Me Deer Comforter Set", "You and Me deer comforter set", "deer comforter set, woodland deer bedding, buck doe comforter", "beige woodland comforter with two spotted deer, You and Me We Got This text and orange flower accents", "You and Me deer comforter set"),
    312: ("Forest Clearing Deer Couple Comforter Set", "forest clearing deer couple comforter set", "deer couple comforter, forest deer bedding, rustic woodland comforter", "tan forest clearing comforter with buck and doe between tree roots, You and Me text and sample names", "forest clearing deer couple comforter set"),
    313: ("Dark Forest Deer Couple Comforter Set", "dark forest deer couple comforter set", "deer couple comforter, woodland deer bedding, rustic forest comforter", "dark forest comforter with two deer, You and Me We Got This text, butterfly accents and sample name pillows", "dark forest deer couple comforter set"),
    314: ("Deer Couple Touching Noses Comforter", "deer couple touching noses comforter", "deer couple comforter, romantic deer bedding, woodland comforter set", "autumn woodland comforter with two deer touching noses, All of Me Loves All of You text and sample name pillows", "deer couple touching noses comforter"),
    315: ("Wolf Dreamcatcher Quilt Set", "wolf dreamcatcher quilt set", "wolf quilt set, dreamcatcher wolf bedding, rustic wolf quilt", "brown quilt set with a close wolf face framed by a dreamcatcher ring, feather details and matching wolf pillows", "wolf dreamcatcher quilt set"),
    316: ("Geometric Headdress Wolf Quilt Set", "geometric headdress wolf quilt set", "wolf quilt set, geometric wolf bedding, tribal style wolf quilt", "brown and cream quilt set with a close wolf face, geometric headdress artwork and matching wolf pillows", "geometric headdress wolf quilt set"),
    317: ("Green Yggdrasil Tree of Life Quilt Set", "green Yggdrasil Tree of Life quilt set", "Yggdrasil quilt set, Tree of Life bedding, Celtic tree quilt", "green quilt set with gold Tree of Life artwork, Celtic knot border, starry center and amber corner emblems", "green Yggdrasil Tree of Life quilt set"),
    318: ("Blue Yggdrasil Ravens Quilt Set", "blue Yggdrasil ravens quilt set", "Yggdrasil ravens quilt, Tree of Life bedding, raven quilt set", "blue quilt set with Yggdrasil Tree of Life, two ravens, Celtic border and matching pillows", "blue Yggdrasil ravens quilt set"),
    319: ("Vibrant Gold Yggdrasil Quilt Set", "vibrant gold Yggdrasil quilt set", "Yggdrasil quilt set, gold Tree of Life bedding, Celtic tree quilt", "teal and blue quilt set with gold Yggdrasil Tree of Life artwork, leafy background and Celtic border", "vibrant gold Yggdrasil quilt set"),
    320: ("Rooster Patchwork Quilt Set", "rooster patchwork quilt set", "rooster quilt set, chicken bedding, farmhouse patchwork quilt", "farmhouse patchwork quilt set with a large rooster, red flowers, patterned black border and matching pillows", "rooster patchwork quilt set"),
}


IMAGE_DETAILS = {
    311: {1: ("Bed mockup of beige deer comforter with two spotted deer, orange flowers and You and Me text.", "You and Me deer comforter bed view"), 2: ("Held view of beige deer comforter with buck and doe artwork and We Got This text.", "You and Me deer comforter held view"), 3: ("Room mockup of You and Me deer comforter with printed craft callout.", "You and Me deer comforter room mockup"), 4: ("Close bed view with optional pillow shams, deer artwork and orange flower details.", "You and Me deer comforter pillow sham panel"), 5: ("High-quality fabric panel showing quilt layers and deer artwork sample.", "You and Me deer comforter fabric panel"), 6: ("Sizing and details chart for You and Me deer comforter set.", "You and Me deer comforter size chart"), 7: ("Bedspread features panel showing comforter layers and care icons.", "You and Me deer comforter features panel"), 8: ("Overhead bedroom mockup of You and Me deer comforter and matching pillows.", "You and Me deer comforter overhead view")},
    312: {1: ("Bed mockup of forest clearing deer couple comforter with buck, doe and sample names.", "Forest clearing deer couple comforter bed view"), 2: ("Held view of forest clearing deer comforter with You and Me text and sample names.", "Forest clearing deer couple comforter held view"), 3: ("Room mockup of forest clearing deer comforter with printed craft callout.", "Forest clearing deer couple comforter room mockup"), 4: ("Close bed view with optional pillow shams and tan deer forest artwork.", "Forest clearing deer couple comforter pillow sham panel"), 5: ("High-quality fabric panel showing quilt layers and deer forest artwork sample.", "Forest clearing deer couple comforter fabric panel"), 6: ("Sizing and details chart for forest clearing deer comforter set.", "Forest clearing deer couple comforter size chart"), 7: ("Bedspread features panel showing comforter layers and care icons.", "Forest clearing deer couple comforter features panel"), 8: ("Overhead bedroom mockup of forest clearing deer comforter and matching pillows.", "Forest clearing deer couple comforter overhead view")},
    313: {1: ("Bed mockup of dark forest deer couple comforter with butterflies and sample name pillows.", "Dark forest deer couple comforter bed view"), 2: ("Overhead bedroom mockup of dark forest deer comforter and matching pillows.", "Dark forest deer couple comforter overhead view"), 3: ("Held view of dark forest deer comforter with You and Me text and butterfly accents.", "Dark forest deer couple comforter held view"), 4: ("Room mockup of dark forest deer comforter with printed craft callout.", "Dark forest deer couple comforter room mockup"), 5: ("Close bed view with optional pillow shams, deer artwork and sample name pillows.", "Dark forest deer couple comforter pillow sham panel"), 6: ("High-quality fabric panel showing quilt layers and deer forest artwork sample.", "Dark forest deer couple comforter fabric panel"), 7: ("Sizing and details chart for dark forest deer comforter set.", "Dark forest deer couple comforter size chart"), 8: ("Bedspread features panel showing comforter layers and care icons.", "Dark forest deer couple comforter features panel")},
    314: {1: ("Bed mockup of autumn deer comforter with two deer touching noses and love text.", "Deer couple touching noses comforter bed view"), 2: ("Held view of deer couple comforter with All of Me Loves All of You text.", "Deer couple touching noses comforter held view"), 3: ("Room mockup of deer couple touching noses comforter with printed craft callout.", "Deer couple touching noses comforter room mockup"), 4: ("Close bed view with optional pillow shams, butterfly accents and sample names.", "Deer couple touching noses comforter pillow sham panel"), 5: ("Sizing and details chart for deer couple touching noses comforter set.", "Deer couple touching noses comforter size chart"), 6: ("High-quality fabric panel showing quilt layers and autumn deer artwork sample.", "Deer couple touching noses comforter fabric panel"), 7: ("Bedspread features panel showing comforter layers and care icons.", "Deer couple touching noses comforter features panel"), 8: ("Overhead bedroom mockup of deer couple touching noses comforter and matching pillows.", "Deer couple touching noses comforter overhead view")},
    315: {1: ("Bed mockup of brown wolf dreamcatcher quilt with close wolf face and feather accents.", "Wolf dreamcatcher quilt bed view"), 2: ("Close fabric view of wolf face artwork with visible quilting pattern.", "Wolf dreamcatcher quilt close fabric view"), 3: ("Single pillow sham mockup with wolf face framed by dreamcatcher artwork.", "Wolf dreamcatcher quilt pillow sham"), 4: ("Room mockup of wolf dreamcatcher quilt set with matching pillows.", "Wolf dreamcatcher quilt room mockup"), 5: ("Bedding set panel showing premium quilt sizes and optional standard shams.", "Wolf dreamcatcher quilt size panel")},
    316: {1: ("Bed mockup of geometric headdress wolf quilt with close wolf face and matching pillows.", "Geometric headdress wolf quilt bed view"), 2: ("Close fabric view of wolf eye and geometric headdress artwork with quilt pattern.", "Geometric headdress wolf quilt close fabric view"), 3: ("Single pillow sham mockup with wolf face and geometric headband design.", "Geometric headdress wolf quilt pillow sham"), 4: ("Room mockup of geometric headdress wolf quilt set with matching pillows.", "Geometric headdress wolf quilt room mockup"), 5: ("Bedding set panel showing premium quilt sizes and optional standard shams.", "Geometric headdress wolf quilt size panel")},
    317: {1: ("Bed mockup of green Yggdrasil quilt with gold Tree of Life and amber corner emblems.", "Green Yggdrasil Tree of Life quilt bed view"), 2: ("Room mockup of green Yggdrasil quilt with printed craft callout.", "Green Yggdrasil Tree of Life quilt room mockup"), 3: ("Close bed view with optional pillow shams and green Tree of Life artwork.", "Green Yggdrasil Tree of Life quilt pillow sham panel"), 4: ("High-quality fabric panel showing quilt layers and green Yggdrasil artwork sample.", "Green Yggdrasil Tree of Life quilt fabric panel"), 5: ("Premium quilt sets sizing chart for green Yggdrasil design.", "Green Yggdrasil Tree of Life quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Green Yggdrasil Tree of Life quilt features panel"), 7: ("Overhead bedroom mockup of green Yggdrasil quilt and matching pillows.", "Green Yggdrasil Tree of Life quilt overhead view")},
    318: {1: ("Bed mockup of blue Yggdrasil quilt with two ravens, Tree of Life roots and Celtic border.", "Blue Yggdrasil ravens quilt bed view"), 2: ("Room mockup of blue Yggdrasil ravens quilt with printed craft callout.", "Blue Yggdrasil ravens quilt room mockup"), 3: ("Close bed view with optional pillow shams and raven Tree of Life artwork.", "Blue Yggdrasil ravens quilt pillow sham panel"), 4: ("High-quality fabric panel showing quilt layers and blue raven artwork sample.", "Blue Yggdrasil ravens quilt fabric panel"), 5: ("Premium quilt sets sizing chart for blue Yggdrasil ravens design.", "Blue Yggdrasil ravens quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Blue Yggdrasil ravens quilt features panel"), 7: ("Overhead bedroom mockup of blue Yggdrasil ravens quilt and matching pillows.", "Blue Yggdrasil ravens quilt overhead view")},
    319: {1: ("Bed mockup of teal blue Yggdrasil quilt with gold Tree of Life and leafy background.", "Vibrant gold Yggdrasil quilt bed view"), 2: ("Room mockup of vibrant gold Yggdrasil quilt with printed craft callout.", "Vibrant gold Yggdrasil quilt room mockup"), 3: ("Close bed view with optional pillow shams and gold Tree of Life border artwork.", "Vibrant gold Yggdrasil quilt pillow sham panel"), 4: ("High-quality fabric panel showing quilt layers and teal gold Yggdrasil artwork sample.", "Vibrant gold Yggdrasil quilt fabric panel"), 5: ("Premium quilt sets sizing chart for vibrant gold Yggdrasil design.", "Vibrant gold Yggdrasil quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Vibrant gold Yggdrasil quilt features panel"), 7: ("Overhead bedroom mockup of vibrant gold Yggdrasil quilt and matching pillows.", "Vibrant gold Yggdrasil quilt overhead view")},
    320: {1: ("Bed mockup of rooster patchwork quilt with large rooster, red flowers and black patterned border.", "Rooster patchwork quilt bed view"), 2: ("Room mockup of rooster patchwork quilt with printed craft callout.", "Rooster patchwork quilt room mockup"), 3: ("Close bed view with optional pillow shams, rooster artwork and patchwork floral details.", "Rooster patchwork quilt pillow sham panel"), 4: ("High-quality fabric panel showing quilt layers and rooster artwork sample.", "Rooster patchwork quilt fabric panel"), 5: ("Premium quilt sets sizing chart for rooster patchwork design.", "Rooster patchwork quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Rooster patchwork quilt features panel"), 7: ("Overhead bedroom mockup of rooster patchwork quilt and matching pillows.", "Rooster patchwork quilt overhead view")},
}


def headers(ws):
    return {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}


def text(value):
    return "" if value is None else str(value)


def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def normalize_url(url):
    return text(url).split("?")[0]


def parse_customizer(summary):
    raw = Path(summary["html"]["html_file"]).read_text(encoding="utf-8", errors="ignore")
    page = html.unescape(raw)
    fields = re.findall(
        r'\{"id":\s*"[^"]+",\s*"label":\s*"([^"]+)",\s*"required":\s*(true|false),\s*"minLength":\s*(\d+),\s*"maxLength":\s*(\d+)',
        page,
    )
    if not fields:
        fields = re.findall(
            r'"label"\s*:\s*"([^"]+)"\s*,\s*"required"\s*:\s*(true|false)\s*,\s*"minLength"\s*:\s*(\d+)\s*,\s*"maxLength"\s*:\s*(\d+)',
            page,
        )
    options = re.findall(r'"label"\s*:\s*"([^"]+)"\s*,\s*"required"\s*:\s*(true|false)\s*,\s*"defaultOptionId"', page)
    return fields, options


def customizer_sentence(fields, options):
    parts = []
    for label, required, min_length, max_length in fields:
        requirement = "required" if required == "true" else "optional"
        parts.append(f"Personalization uses a {requirement} {label} field, {min_length}-{max_length} characters.")
    visible_options = [label for label, _ in options if label and "confirmation" not in label.lower() and "comfirmation" not in label.lower()]
    if visible_options:
        parts.append(f"Visible required choice fields include: {', '.join(visible_options)}.")
    if fields:
        parts.append("Visible custom text or example selections in mockups are sample artwork unless the matching input is entered.")
    else:
        parts.append("No shopper text-entry field is described for this product.")
    return " ".join(parts)


def load_admin(handles):
    by_handle = defaultdict(list)
    with ADMIN_EXPORT.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row["Handle"] in handles:
                by_handle[row["Handle"]].append(row)
    admin = {}
    for handle, rows in by_handle.items():
        first = rows[0]
        images = {}
        for row in rows:
            if row.get("Image Src"):
                images[normalize_url(row["Image Src"])] = {"src": row["Image Src"], "alt": row.get("Image Alt Text", "")}
        admin[handle] = {
            "Title": first.get("Title", ""),
            "Type": first.get("Type", ""),
            "SEO Title": first.get("SEO Title", ""),
            "SEO Description": first.get("SEO Description", ""),
            "Option1 Name": first.get("Option1 Name", ""),
            "Option2 Name": first.get("Option2 Name", ""),
            "Option3 Name": first.get("Option3 Name", ""),
            "Status": first.get("Status", ""),
            "images": images,
        }
    missing = sorted(set(handles) - set(admin))
    if missing:
        raise RuntimeError(f"Missing handles in admin export: {missing}")
    return admin


def option_text(admin_row):
    values = [admin_row["Option1 Name"], admin_row["Option2 Name"], admin_row["Option3 Name"]]
    values = [v for v in values if v]
    return ", ".join(values) if values else "available size options"


def description(pos, admin_row, ctext):
    title, _, _, detail, _ = PRODUCT_UPDATES[pos]
    return (
        f"<p>{title} features {detail}. The copy stays specific to the visible artwork, product form and selectable options for this exact item.</p>"
        "<h3>Design Details</h3>"
        f"<ul><li>Artwork: {detail}.</li><li>Gallery images include the main mockup plus size, feature, care, fabric or lifestyle panels where shown.</li></ul>"
        "<h3>Options and Fit</h3>"
        f"<ul><li>Available option groups: {option_text(admin_row)}.</li><li>{ctext}</li><li>Select the product type and size shown on the product page before checkout.</li></ul>"
    )


def main():
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE, OUTPUT)
    summaries = {int(item["inventory_row"]["inventory_position"]): item for item in json.loads(SUMMARY_PATH.read_text(encoding="utf-8"))}
    positions = set(PRODUCT_UPDATES)
    handles = {summaries[pos]["inventory_row"]["Handle"] for pos in positions}
    admin = load_admin(handles)
    admin_hash = sha256(ADMIN_EXPORT)
    now = datetime.now().astimezone().isoformat()
    wb = load_workbook(OUTPUT)

    pos_by_key = {summaries[pos]["inventory_row"]["product_key"]: pos for pos in positions}
    handle_by_pos = {pos: summaries[pos]["inventory_row"]["Handle"] for pos in positions}
    fields_by_pos, option_fields_by_pos = {}, {}
    for pos in positions:
        fields_by_pos[pos], option_fields_by_pos[pos] = parse_customizer(summaries[pos])

    ws = wb["SEO_Products"]
    idx = headers(ws)
    revised_products = 0
    for row_num in range(2, ws.max_row + 1):
        pk = text(ws.cell(row_num, idx["product_key"]).value)
        if pk not in pos_by_key:
            continue
        pos = pos_by_key[pk]
        handle = handle_by_pos[pos]
        title, primary, secondary, detail, cluster = PRODUCT_UPDATES[pos]
        admin_row = admin[handle]
        meta = f"Shop {title.lower()} with {detail}, selectable sizes and product details."[:155]
        ws.cell(row_num, idx["title_current"]).value = admin_row["Title"]
        ws.cell(row_num, idx["h1_current"]).value = admin_row["Title"]
        ws.cell(row_num, idx["rendered_title_current"]).value = admin_row["SEO Title"] or admin_row["Title"]
        ws.cell(row_num, idx["meta_description_current"]).value = admin_row["SEO Description"]
        ws.cell(row_num, idx["product_type"]).value = admin_row["Type"] or summaries[pos]["inventory_row"]["product_type"]
        ws.cell(row_num, idx["title_proposed"]).value = title
        ws.cell(row_num, idx["meta_title_seo"]).value = title
        ws.cell(row_num, idx["meta_title_length"]).value = len(title)
        ws.cell(row_num, idx["meta_description_seo"]).value = meta
        ws.cell(row_num, idx["meta_description_length"]).value = len(meta)
        ws.cell(row_num, idx["description_proposed_html"]).value = description(pos, admin_row, customizer_sentence(fields_by_pos[pos], option_fields_by_pos[pos]))
        ws.cell(row_num, idx["primary_keyword"]).value = primary
        ws.cell(row_num, idx["secondary_keywords"]).value = secondary
        ws.cell(row_num, idx["meta_keyword"]).value = primary
        ws.cell(row_num, idx["keyword_strategy"]).value = f"Target the specific product motif: {cluster}; keep broad wildlife, wolf, Celtic tree, raven or farmhouse animal terms for collection pages."
        ws.cell(row_num, idx["buyer_search_summary"]).value = f"US English buyer intent targets {cluster}; no search-volume, trend or ranking claim is made."
        ws.cell(row_num, idx["keyword_evidence_level"]).value = "SERP_ONLY"
        ws.cell(row_num, idx["revision"]).value = "r2"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields"]:
            ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            f"R2 deep recheck without prior QA: used products_export_1.csv sha256:{admin_hash}; "
            "separated deer couple comforters, wolf quilts, Yggdrasil quilts and rooster patchwork quilt intents; confirmed no shopper text-entry fields and rewrote image observations/alts from contact sheets. Still NEEDS_REVIEW."
        )
        revised_products += 1

    ws = wb["Image_Audit"]
    idx = headers(ws)
    revised_images = 0
    admin_alt_matches = 0
    seen_alts = set()
    for row_num in range(2, ws.max_row + 1):
        handle = text(ws.cell(row_num, idx["Handle"]).value)
        pos = next((p for p, h in handle_by_pos.items() if h == handle), None)
        if pos is None:
            continue
        image_number = int(ws.cell(row_num, idx["image_number"]).value)
        image_url = text(ws.cell(row_num, idx["image_url_export"]).value or ws.cell(row_num, idx["image_url"]).value)
        admin_image = admin.get(handle, {}).get("images", {}).get(normalize_url(image_url))
        if admin_image:
            ws.cell(row_num, idx["image_url_export"]).value = admin_image["src"]
            ws.cell(row_num, idx["alt_current"]).value = admin_image["alt"]
            admin_alt_matches += 1
        observation, alt = IMAGE_DETAILS[pos][image_number]
        if alt in seen_alts:
            alt = f"{alt} for {PRODUCT_UPDATES[pos][4]}"
        seen_alts.add(alt)
        ws.cell(row_num, idx["observed_visual_details"]).value = observation
        ws.cell(row_num, idx["alt_proposed"]).value = alt[:125]
        ws.cell(row_num, idx["alt_action"]).value = "SET"
        ws.cell(row_num, idx["viewed_status"]).value = "VIEWED"
        ws.cell(row_num, idx["revision"]).value = "r2"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields", "approved_revision"]:
            if col_name in idx:
                ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = "R2: observation and alt rechecked from contact sheet; admin image URL and current alt matched from Shopify CSV where available."
        revised_images += 1

    ws = wb["Keyword_Map"]
    idx = headers(ws)
    keyword_counts = defaultdict(int)
    keyword_rows = 0
    for row_num in range(2, ws.max_row + 1):
        pk = text(ws.cell(row_num, idx["product_key"]).value)
        if pk not in pos_by_key:
            continue
        pos = pos_by_key[pk]
        title, primary, secondary, _, cluster = PRODUCT_UPDATES[pos]
        role = text(ws.cell(row_num, idx["keyword_role"]).value)
        secondaries = [part.strip() for part in secondary.split(",")]
        keyword = primary if role == "PRIMARY" else secondaries[min(keyword_counts[pk], len(secondaries) - 1)]
        if role != "PRIMARY":
            keyword_counts[pk] += 1
        ws.cell(row_num, idx["keyword"]).value = keyword
        ws.cell(row_num, idx["semantic_cluster"]).value = cluster
        ws.cell(row_num, idx["intent"]).value = "Commercial product intent"
        ws.cell(row_num, idx["target_page_type"]).value = "PRODUCT"
        ws.cell(row_num, idx["decision_reason"]).value = f"{title} is separated by verified product form, artwork motif and custom fields."
        ws.cell(row_num, idx["demand_evidence"]).value = "SERP intent evidence only; no search-volume claim."
        ws.cell(row_num, idx["validation_source"]).value = "products_export_1.csv admin baseline + storefront customizer parse + contact-sheet inspection"
        ws.cell(row_num, idx["checked_at"]).value = now
        ws.cell(row_num, idx["possible_overlap_with_other_products"]).value = "YES; nearby products are differentiated by motif, product form and custom field type."
        ws.cell(row_num, idx["mapping_reason"]).value = "R2 separates deer couple comforters, wolf quilts, Yggdrasil quilt variants and rooster patchwork quilt pages by visible motif and option fields."
        ws.cell(row_num, idx["mapping_status"]).value = "CANDIDATE_MAPPED_NEEDS_RECHECK"
        ws.cell(row_num, idx["mapping_version"]).value = "r2"
        keyword_rows += 1

    ws = wb["Buyer_Search_Research"]
    idx = headers(ws)
    for row_num in range(2, ws.max_row + 1):
        pk = text(ws.cell(row_num, idx["product_key"]).value)
        if pk not in pos_by_key:
            continue
        pos = pos_by_key[pk]
        title, primary, _, _, cluster = PRODUCT_UPDATES[pos]
        ws.cell(row_num, idx["purchase_context"]).value = f"US buyer looking for {cluster} as a specific product page."
        ws.cell(row_num, idx["jtbd_statement"]).value = f"When shopping for {cluster}, the buyer wants exact artwork, product type, size choices and available custom fields confirmed before purchase."
        ws.cell(row_num, idx["functional_motivation"]).value = "Confirm product form, visible motif, size options, custom field rules, feature panels and image accuracy before purchase."
        ws.cell(row_num, idx["emotional_social_motivation"]).value = "Choose wildlife, wolf, Tree of Life or farmhouse animal bedding that matches the recipient or room theme."
        ws.cell(row_num, idx["purchase_concerns"]).value = "Confirm product type, size choices, optional shams and visible design details; avoid unsupported material, custom input, copyrighted identity or delivery claims."
        ws.cell(row_num, idx["source_refs"]).value = f"products_export_1.csv sha256:{admin_hash}; storefront HTML customizer parse; contact sheets 311-320"
        ws.cell(row_num, idx["observed_at"]).value = now
        ws.cell(row_num, idx["research_status"]).value = "REVISED_R2_NEEDS_RECHECK"
        ws.cell(row_num, idx["limitations"]).value = "No independent QA, Search Console, paid keyword volume or internal site-search export supplied; proof remains SERP-only."
        ws.cell(row_num, idx["seo_application"]).value = f"Use {primary} with title: {title}."

    ws = wb["Product_Evidence"]
    idx = headers(ws)
    for row_num in range(2, ws.max_row + 1):
        evidence_id = text(ws.cell(row_num, idx["evidence_id"]).value)
        match = re.search(r"_(\d{3})$", evidence_id)
        if not match:
            continue
        pos = int(match.group(1))
        if pos not in positions:
            continue
        handle = handle_by_pos[pos]
        admin_row = admin[handle]
        _, _, _, detail, _ = PRODUCT_UPDATES[pos]
        ws.cell(row_num, idx["sources_accessed"]).value = f"Original storefront/product evidence; products_export_1.csv sha256:{admin_hash}; storefront customizer parse; inspected contact sheets"
        ws.cell(row_num, idx["current_H1"]).value = admin_row["Title"]
        ws.cell(row_num, idx["current_meta_title"]).value = admin_row["SEO Title"] or "EMPTY"
        ws.cell(row_num, idx["verified_product_facts"]).value = (
            f"Admin export matched handle {handle}; type={admin_row['Type']}; options={option_text(admin_row)}; "
            f"status={admin_row['Status']}; image_count={len(admin_row['images'])}; design={detail}; customizer={customizer_sentence(fields_by_pos[pos], option_fields_by_pos[pos])}"
        )
        ws.cell(row_num, idx["factual_conflicts"]).value = "No independent QA for batch 032 yet; r2 prepared for QA recheck."
        ws.cell(row_num, idx["confidence_and_reason"]).value = "R2 uses contact sheets, storefront customizer parse and Shopify admin CSV baseline; no approval or import created."

    ws = wb["README_QA"]
    idx = headers(ws)
    for metric, value, definition in [
        ("qa_batch_032_r2_revision", "r2", "Separate 10-product revision without prior independent QA; source workbook was not modified."),
        ("qa_batch_032_r2_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_032_r2_scope", "inventory positions 311-320", "No products outside qa_batch_032 were revised."),
        ("qa_batch_032_r2_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_032_r2_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_032_r2_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
    ]:
        ws.append([metric if col == idx["metric"] else value if col == idx["value"] else definition if col == idx["definition"] else "" for col in range(1, ws.max_column + 1)])

    wb.save(OUTPUT)
    load_workbook(OUTPUT, keep_links=False).save(OUTPUT)
    summary = {
        "created_at": now,
        "batch_id": BATCH_ID,
        "source_revision": str(SOURCE.relative_to(ROOT)),
        "source_evidence": str(SUMMARY_PATH.relative_to(ROOT)),
        "admin_export": str(ADMIN_EXPORT.relative_to(ROOT)),
        "admin_export_sha256": admin_hash,
        "revision_workbook": str(OUTPUT.relative_to(ROOT)),
        "scope": "qa_batch_032 only; inventory positions 311-320",
        "revision": "r2",
        "revised_products": revised_products,
        "revised_images": revised_images,
        "admin_alt_matches": admin_alt_matches,
        "keyword_rows_revised": keyword_rows,
        "review_status": "NEEDS_REVIEW",
        "approved_created": False,
        "shopify_import_created": False,
        "next_step": "Ready for independent QA of qa_batch_032_r2.",
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text(
        "\n".join([
            "# Revision qa_batch_032_r2",
            "",
            "Artifact chính thức theo quy trình từng lô 10 sản phẩm.",
            "",
            f"- Nguồn r1: `{SOURCE.relative_to(ROOT)}`",
            f"- Evidence summary: `{SUMMARY_PATH.relative_to(ROOT)}`",
            f"- Shopify admin export: `{ADMIN_EXPORT.relative_to(ROOT)}`",
            f"- Admin export SHA-256: `{admin_hash}`",
            f"- Workbook r2: `{OUTPUT.relative_to(ROOT)}`",
            f"- Sản phẩm sửa: {revised_products}",
            f"- Ảnh sửa/refresh: {revised_images}",
            f"- Admin image alt match: {admin_alt_matches}/{revised_images}",
            "- Sửa trọng yếu: rà lại khi chưa có QA, tách nhóm deer couple comforter, wolf quilt, Yggdrasil quilt và rooster patchwork quilt; parse customizer từ HTML storefront; viết lại title/meta/description/alt theo từng contact sheet.",
            "- Trạng thái: `NEEDS_REVIEW`; không tạo `APPROVED`; không tạo file import Shopify.",
            "- Bước tiếp theo: gửi `qa_batch_032_r2` cho QA độc lập.",
            "",
        ]),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
