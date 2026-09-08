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
BATCH_ID = "qa_batch_022_r2"
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_022.xlsx"
SUMMARY_PATH = ROOT / "seo_runs" / SHOP / RUN_ID / "batch_022_evidence_summary.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_022_r2.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_022_r2.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    211: ("Basketball Close-Up Number Blanket", "basketball close-up number blanket", "basketball number blanket, basketball throw blanket, basketball player blanket", "black blanket with oversized close-up basketball, large jersey number and vertical sample name", "basketball close-up number blanket"),
    212: ("Custom Basketball Court Comforter", "custom basketball court comforter", "basketball court comforter, personalized basketball bedding, custom basketball comforter", "black and gray basketball comforter with court perspective, large ball texture, sample name and jersey number", "custom name number basketball court comforter"),
    213: ("Custom Basketball Cracked Wall Comforter", "custom basketball cracked wall comforter", "basketball cracked wall comforter, personalized basketball bedding, custom sports bedding", "white and gray cracked-wall basketball comforter with an orange ball breaking through and vertical sample name", "custom name number basketball cracked wall comforter"),
    214: ("Custom Rainbow Basketball Net Comforter", "custom rainbow basketball net comforter", "basketball net comforter, personalized basketball comforter, basketball hoop bedding", "bright rainbow fire basketball comforter with glowing hoop, net, sample name and jersey number", "custom name number rainbow basketball net comforter"),
    215: ("Custom Basketball Fire Water Blanket", "custom basketball fire water blanket", "personalized basketball blanket, custom basketball throw, basketball gift blanket", "blanket with basketball, fire, water splash, lightning, hoop background, vertical sample name and jersey number", "custom basketball fire water blanket with optional name number"),
    216: ("Custom Flaming Basketball Comforter", "custom flaming basketball comforter", "flaming basketball comforter, personalized basketball bedding, custom sports bedding", "black comforter with a flaming basketball flying across a smoky court, vertical sample name and jersey number", "custom name number flaming basketball comforter"),
    217: ("Custom Light Burst Basketball Comforter", "custom light burst basketball comforter", "glowing basketball comforter, personalized basketball bedding, basketball number comforter", "gold and orange basketball comforter with glowing light burst, sample name and jersey number", "custom name number light burst basketball comforter"),
    218: ("Custom Basketball Court Blanket", "custom basketball court blanket", "personalized basketball blanket, custom basketball throw, basketball player gift", "stadium court blanket with a hand holding a basketball, sample name arched over the ball and jersey number", "custom basketball court blanket with optional name number"),
    219: ("Custom Basketball Player Blanket", "custom basketball player blanket", "personalized basketball blanket, custom player blanket, basketball name blanket", "black blanket with player arm holding a basketball, vertical sample name and jersey number on the ball", "custom basketball player blanket with optional name number"),
    220: ("Custom Orange Basketball Hoop Comforter", "custom orange basketball hoop comforter", "orange basketball hoop comforter, personalized basketball bedding, custom sports bedding", "black and orange basketball comforter with glowing hoop, sample name and jersey number on pillow shams", "custom name number orange basketball hoop comforter"),
}


IMAGE_DETAILS = {
    211: {1: ("Blanket mockup with close-up basketball, large sample number and vertical sample name.", "Basketball close-up number blanket"), 2: ("Basketball number blanket displayed on sofa with dark ball artwork.", "Basketball number blanket on sofa"), 3: ("Folded basketball blanket showing close-up ball print and sample number.", "Basketball number blanket folded view"), 4: ("Basketball blanket size or style panel.", "Basketball number blanket size panel"), 5: ("Close-up fabric panel for basketball number blanket.", "Basketball number blanket fabric panel"), 6: ("Lifestyle use image for basketball blanket.", "Basketball blanket lifestyle image"), 7: ("Additional basketball blanket mockup with close-up ball artwork.", "Basketball number blanket mockup"), 8: ("Basketball blanket feature panel.", "Basketball blanket feature panel"), 9: ("Product shape illustration for basketball number blanket.", "Basketball number blanket shape view")},
    212: {1: ("Bedroom mockup of basketball court perspective comforter with sample name and number.", "Custom basketball court comforter"), 2: ("Second basketball court bedding mockup with duvet fold.", "Basketball court comforter mockup"), 3: ("Feature view of basketball court bedding with comfort icons.", "Basketball court bedding feature view"), 4: ("Microfiber panel with basketball court print samples.", "Basketball court microfiber panel"), 5: ("Easy care panel for basketball court comforter.", "Basketball court easy care panel"), 6: ("Bedding type comparison panel for basketball bedding.", "Basketball court bedding type panel"), 7: ("Size dimension panel for basketball court comforter.", "Basketball court comforter size panel")},
    213: {1: ("Bedroom mockup of cracked wall basketball comforter with orange ball and vertical sample name.", "Custom basketball cracked wall comforter"), 2: ("Second cracked wall basketball bedding mockup with duvet fold.", "Cracked wall basketball bedding mockup"), 3: ("Feature view of cracked wall basketball bedding with comfort icons.", "Cracked wall basketball feature view"), 4: ("Microfiber panel showing cracked wall basketball print samples.", "Cracked wall basketball microfiber panel"), 5: ("Easy care panel for basketball bedding.", "Basketball bedding easy care panel"), 6: ("Bedding type comparison panel for basketball bedding.", "Basketball bedding type panel"), 7: ("Size dimension panel for basketball comforter.", "Basketball comforter size panel")},
    214: {1: ("Bedroom mockup of rainbow basketball net comforter with glowing hoop and sample name.", "Custom rainbow basketball net comforter"), 2: ("High-density weaving panel comparing fabric samples.", "High-density weaving feature panel"), 3: ("Machine washable care panel with washing machine and laundry basket.", "Machine washable bedding care panel"), 4: ("Bottom zippered closure panel with white zipper close-up.", "Bottom zippered closure panel"), 5: ("Rainbow basketball net comforter product mockup with bright court artwork.", "Rainbow basketball net comforter mockup")},
    215: {1: ("Blanket mockup with basketball, fire, water splash and vertical sample name.", "Custom basketball fire water blanket"), 2: ("Basketball fire water blanket draped on sofa.", "Basketball fire water blanket on sofa"), 3: ("Folded basketball blanket showing fire and water splash artwork.", "Basketball fire water blanket folded view"), 4: ("Basketball blanket size or style panel.", "Basketball fire water blanket size panel"), 5: ("Close-up fabric panel for basketball fire water blanket.", "Basketball fire water blanket fabric panel"), 6: ("Lifestyle use image for basketball blanket.", "Basketball blanket lifestyle panel"), 7: ("Additional basketball blanket mockup with fire and water artwork.", "Basketball fire water blanket mockup"), 8: ("Basketball blanket feature panel.", "Basketball fire water blanket feature panel")},
    216: {1: ("Bedroom mockup of flaming basketball comforter with flying ball, sample name and number.", "Custom flaming basketball comforter"), 2: ("Second flaming basketball bedding mockup with duvet fold.", "Flaming basketball bedding mockup"), 3: ("Feature view of flaming basketball bedding with comfort icons.", "Flaming basketball feature view"), 4: ("Microfiber panel showing flaming basketball print samples.", "Flaming basketball microfiber panel"), 5: ("Easy care panel for flaming basketball bedding.", "Flaming basketball easy care panel"), 6: ("Bedding type comparison panel for basketball bedding.", "Flaming basketball bedding type panel"), 7: ("Size dimension panel for flaming basketball comforter.", "Flaming basketball size panel")},
    217: {1: ("Bedroom mockup of light burst basketball comforter with sample name and number.", "Custom light burst basketball comforter"), 2: ("Second light burst basketball bedding mockup with duvet fold.", "Light burst basketball bedding mockup"), 3: ("Feature view of glowing basketball bedding with comfort icons.", "Light burst basketball feature view"), 4: ("Microfiber panel showing light burst basketball print samples.", "Light burst basketball microfiber panel"), 5: ("Easy care panel for light burst basketball bedding.", "Light burst basketball easy care panel"), 6: ("Bedding type comparison panel for basketball bedding.", "Light burst basketball bedding type panel"), 7: ("Size dimension panel for light burst basketball comforter.", "Light burst basketball size panel")},
    218: {1: ("Court blanket mockup with hand holding basketball, arched sample name and jersey number.", "Custom basketball court blanket"), 2: ("Basketball court blanket draped on sofa.", "Basketball court blanket on sofa"), 3: ("Folded basketball court blanket showing hand and ball artwork.", "Basketball court blanket folded view"), 4: ("Basketball blanket size or style panel.", "Basketball court blanket size panel"), 5: ("Close-up fabric panel for basketball court blanket.", "Basketball court blanket fabric panel"), 6: ("Lifestyle use image for basketball blanket.", "Basketball blanket lifestyle panel"), 7: ("Additional basketball court blanket mockup.", "Basketball court blanket mockup"), 8: ("Basketball blanket feature panel.", "Basketball court blanket feature panel")},
    219: {1: ("Blanket mockup with player arm holding basketball, sample name and jersey number.", "Custom basketball player blanket"), 2: ("Basketball player blanket displayed on sofa.", "Basketball player blanket on sofa"), 3: ("Folded basketball player blanket showing arm and ball artwork.", "Basketball player blanket folded view"), 4: ("Basketball blanket size or style panel.", "Basketball player blanket size panel"), 5: ("Close-up fabric panel for basketball player blanket.", "Basketball player blanket fabric panel"), 6: ("Lifestyle use image for basketball blanket.", "Basketball blanket lifestyle image"), 7: ("Additional basketball player blanket mockup.", "Basketball player blanket mockup"), 8: ("Basketball blanket feature panel.", "Basketball player blanket feature panel")},
    220: {1: ("Bedroom mockup of orange basketball hoop comforter with sample name and number on shams.", "Custom orange basketball hoop comforter"), 2: ("Second orange basketball hoop bedding mockup with duvet fold.", "Orange basketball bedding mockup"), 3: ("Feature view of orange basketball hoop bedding with comfort icons.", "Orange basketball feature view"), 4: ("Microfiber panel showing orange basketball print samples.", "Orange basketball microfiber panel"), 5: ("Easy care panel for orange basketball bedding.", "Orange basketball easy care panel"), 6: ("Bedding type comparison panel for basketball bedding.", "Orange basketball bedding type panel"), 7: ("Size dimension panel for orange basketball comforter.", "Orange basketball size panel")},
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


def load_summaries():
    summaries = json.loads(SUMMARY_PATH.read_text(encoding="utf-8"))
    return {int(item["inventory_row"]["inventory_position"]): item for item in summaries}


def parse_customizer(summary):
    html_file = Path(summary["html"]["html_file"])
    raw = html_file.read_text(encoding="utf-8", errors="ignore")
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
    useful_options = [label for label, _ in options if label and "confirmation" not in label.lower()]
    if useful_options:
        parts.append(f"Additional visible customization option: {', '.join(useful_options)}.")
    labels = " ".join(label for label, *_ in fields).lower()
    if "number" in labels:
        parts.append("Names and numbers visible in mockups are sample artwork unless the matching input is entered.")
    elif "name" in labels:
        parts.append("Visible names in mockups are sample artwork unless the matching input is entered.")
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
                images[normalize_url(row["Image Src"])] = {
                    "src": row["Image Src"],
                    "position": row.get("Image Position", ""),
                    "alt": row.get("Image Alt Text", ""),
                }
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
        f"<p>{title} features {detail}. "
        "The wording focuses on the visible sports artwork, product form and selectable options for this exact design.</p>"
        "<h3>Design Details</h3>"
        f"<ul><li>Artwork: {detail}.</li>"
        "<li>Gallery images show the main bedding or blanket mockup plus feature, care, bedding-type, size or lifestyle panels where present.</li></ul>"
        "<h3>Options and Fit</h3>"
        f"<ul><li>Available option groups: {option_text(admin_row)}.</li>"
        f"<li>{ctext}</li>"
        "<li>Select the product type and size shown on the product page before checkout.</li></ul>"
    )


def main():
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE, OUTPUT)
    summaries = load_summaries()
    batch_positions = set(PRODUCT_UPDATES)
    handles = {summaries[pos]["inventory_row"]["Handle"] for pos in batch_positions}
    admin = load_admin(handles)
    admin_hash = sha256(ADMIN_EXPORT)
    now = datetime.now().astimezone().isoformat()

    wb = load_workbook(OUTPUT)
    pos_by_key = {summaries[pos]["inventory_row"]["product_key"]: pos for pos in batch_positions}
    handle_by_pos = {pos: summaries[pos]["inventory_row"]["Handle"] for pos in batch_positions}
    fields_by_pos = {}
    option_fields_by_pos = {}
    for pos in batch_positions:
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
        ctext = customizer_sentence(fields_by_pos[pos], option_fields_by_pos[pos])
        if fields_by_pos[pos]:
            meta = f"Shop {title.lower()} with {detail}, selectable sizes and verified custom text fields."
        else:
            meta = f"Shop {title.lower()} with {detail}, visible basketball artwork and selectable size options."
        ws.cell(row_num, idx["title_current"]).value = admin_row["Title"]
        ws.cell(row_num, idx["h1_current"]).value = admin_row["Title"]
        ws.cell(row_num, idx["rendered_title_current"]).value = admin_row["SEO Title"] or admin_row["Title"]
        ws.cell(row_num, idx["meta_description_current"]).value = admin_row["SEO Description"]
        ws.cell(row_num, idx["product_type"]).value = admin_row["Type"] or summaries[pos]["inventory_row"]["product_type"]
        ws.cell(row_num, idx["title_proposed"]).value = title
        ws.cell(row_num, idx["meta_title_seo"]).value = title
        ws.cell(row_num, idx["meta_title_length"]).value = len(title)
        # 145–165 characters is an editorial target, never a truncation boundary.
        ws.cell(row_num, idx["meta_description_seo"]).value = meta
        ws.cell(row_num, idx["meta_description_length"]).value = len(meta)
        ws.cell(row_num, idx["description_proposed_html"]).value = description(pos, admin_row, ctext)
        ws.cell(row_num, idx["primary_keyword"]).value = primary
        ws.cell(row_num, idx["secondary_keywords"]).value = secondary
        ws.cell(row_num, idx["meta_keyword"]).value = primary
        ws.cell(row_num, idx["keyword_strategy"]).value = (
            f"Target the specific product motif: {cluster}; keep broader personalized sports bedding terms for collections."
        )
        ws.cell(row_num, idx["buyer_search_summary"]).value = (
            f"US English buyer intent targets {cluster}; no search-volume, trend or ranking claim is made."
        )
        ws.cell(row_num, idx["keyword_evidence_level"]).value = "SERP_ONLY"
        ws.cell(row_num, idx["revision"]).value = "r2"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields"]:
            ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            f"R2 deep recheck without prior QA: used products_export_1.csv sha256:{admin_hash}; "
            "removed internal workflow copy, refreshed publishable title values, parsed storefront custom fields, "
            "split baseball and basketball intents by visible motif, and rewrote image observations/alts from contact sheets. Still NEEDS_REVIEW."
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
        ws.cell(row_num, idx["issues"]).value = (
            "R2: observation and alt rechecked from contact sheet; admin image URL and current alt matched from Shopify CSV where available."
        )
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
        ws.cell(row_num, idx["decision_reason"]).value = f"{title} is separated by verified sport, layout, product form and custom fields."
        ws.cell(row_num, idx["demand_evidence"]).value = "SERP intent evidence only; no search-volume claim."
        ws.cell(row_num, idx["validation_source"]).value = "products_export_1.csv admin baseline + storefront customizer parse + contact-sheet inspection"
        ws.cell(row_num, idx["checked_at"]).value = now
        ws.cell(row_num, idx["possible_overlap_with_other_products"]).value = "YES; nearby baseball/basketball designs are differentiated by motif."
        ws.cell(row_num, idx["mapping_reason"]).value = "R2 separates nearby pages by visible motif, sport, layout and verified custom fields."
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
        ws.cell(row_num, idx["jtbd_statement"]).value = (
            f"When shopping for {cluster}, the buyer wants the page to confirm exact artwork, product type, size choices, "
            "and the available custom fields before purchase when those fields are captured."
        )
        ws.cell(row_num, idx["functional_motivation"]).value = (
            "Confirm product form, sport motif, visible layout, custom text fields, care or size panels and image accuracy before purchase."
        )
        ws.cell(row_num, idx["emotional_social_motivation"]).value = (
            "Choose a personalized baseball or basketball bedding item that matches a preferred player, court, flag, flame or ball design."
        )
        ws.cell(row_num, idx["purchase_concerns"]).value = (
            "Confirm spelling and number where fields are available; avoid unsupported team, player, league, material or delivery claims."
        )
        ws.cell(row_num, idx["source_refs"]).value = f"products_export_1.csv sha256:{admin_hash}; storefront HTML customizer parse; contact sheets 211-220"
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
        if pos not in batch_positions:
            continue
        handle = handle_by_pos[pos]
        admin_row = admin[handle]
        _, _, _, detail, _ = PRODUCT_UPDATES[pos]
        ctext = customizer_sentence(fields_by_pos[pos], option_fields_by_pos[pos])
        ws.cell(row_num, idx["sources_accessed"]).value = (
            f"Original storefront/product evidence; products_export_1.csv sha256:{admin_hash}; storefront customizer parse; inspected contact sheets"
        )
        ws.cell(row_num, idx["current_H1"]).value = admin_row["Title"]
        ws.cell(row_num, idx["current_meta_title"]).value = admin_row["SEO Title"] or "EMPTY"
        ws.cell(row_num, idx["verified_product_facts"]).value = (
            f"Admin export matched handle {handle}; type={admin_row['Type']}; options={option_text(admin_row)}; "
            f"status={admin_row['Status']}; image_count={len(admin_row['images'])}; design={detail}; customizer={ctext}"
        )
        ws.cell(row_num, idx["factual_conflicts"]).value = "No independent QA for batch 022 yet; r2 prepared for QA recheck."
        ws.cell(row_num, idx["confidence_and_reason"]).value = (
            "R2 uses contact sheets, storefront customizer parse and Shopify admin CSV baseline; no approval or import created."
        )

    ws = wb["README_QA"]
    idx = headers(ws)
    for metric, value, definition in [
        ("qa_batch_022_r2_revision", "r2", "Separate 10-product revision without prior independent QA; source workbook was not modified."),
        ("qa_batch_022_r2_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_022_r2_scope", "inventory positions 211-220", "No products outside qa_batch_022 were revised."),
        ("qa_batch_022_r2_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_022_r2_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_022_r2_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
    ]:
        ws.append([
            metric if col == idx["metric"] else value if col == idx["value"] else definition if col == idx["definition"] else ""
            for col in range(1, ws.max_column + 1)
        ])

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
        "scope": "qa_batch_022 only; inventory positions 211-220",
        "revision": "r2",
        "revised_products": revised_products,
        "revised_images": revised_images,
        "admin_alt_matches": admin_alt_matches,
        "keyword_rows_revised": keyword_rows,
        "review_status": "NEEDS_REVIEW",
        "approved_created": False,
        "shopify_import_created": False,
        "next_step": "Ready for independent QA of qa_batch_022_r2.",
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text(
        "\n".join([
            "# Revision qa_batch_022_r2",
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
            "- Sửa trọng yếu: rà lại khi chưa có QA, bỏ nội dung nội bộ, giữ title/H1 publishable, parse customizer từ HTML storefront, tách intent baseball/basketball theo motif nhìn thấy và viết lại mô tả/alt theo contact sheets.",
            "- Trạng thái: `NEEDS_REVIEW`; không tạo `APPROVED`; không tạo file import Shopify.",
            "- Bước tiếp theo: gửi `qa_batch_022_r2` cho QA độc lập.",
            "",
        ]),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
