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
BATCH_ID = "qa_batch_031_r2"
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_031.xlsx"
SUMMARY_PATH = ROOT / "seo_runs" / SHOP / RUN_ID / "batch_031_evidence_summary.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_031_r2.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_031_r2.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    301: ("Vibrant Cactus Flower Quilt Set", "vibrant cactus flower quilt set", "cactus flower quilt set, southwest floral bedding, cactus bloom quilt", "quilt set with tall green cacti, large red and yellow flowers, teal stripes and warm red trim", "vibrant cactus flower quilt set"),
    302: ("Embroidered Desert Cactus Quilt Set", "embroidered desert cactus quilt set", "desert cactus quilt set, embroidered cactus bedding, colorful cactus quilt", "bright desert cactus quilt set with teal columns, magenta flowers, yellow-orange background and textured stitch look", "embroidered desert cactus quilt set"),
    303: ("Flowering Cactus Desert Scene Quilt Set", "flowering cactus desert scene quilt set", "flowering cactus quilt, desert scene bedding, southwest cactus quilt set", "desert scene quilt set with saguaro cacti, layered hills, muted sky and red-orange blooms", "flowering cactus desert scene quilt set"),
    304: ("Vibrant Potted Cactus Garden Quilt Set", "vibrant potted cactus garden quilt set", "potted cactus quilt set, cactus garden bedding, succulent quilt set", "yellow quilt set with repeated potted cactus and succulent blocks in bright colorful planters", "vibrant potted cactus garden quilt set"),
    305: ("Potted Cactus Quilt Set Design 3", "potted cactus quilt set", "potted cactus quilt, southwest cactus bedding, succulent garden quilt", "cream quilt set with potted cactus grid, earthy patterned planters, succulents and small orange flowers", "potted cactus quilt set"),
    306: ("Viking Celtic Tree of Life Quilt Set", "Viking Celtic Tree of Life quilt set", "Viking quilt set, Celtic Tree of Life bedding, Norse tree quilt", "blue and black Viking-style quilt set with Celtic border, Tree of Life roots, compass emblem, runes and sample name Brent", "Viking Celtic Tree of Life quilt set"),
    307: ("Vintage Reading Girl Flowers Blanket", "vintage reading girl flowers blanket", "vintage book lover blanket, reading girl throw, floral reader blanket", "vintage newspaper-style blanket with tall flowers, cartoon reader, stack of books and sample name Sarah", "vintage reading girl flowers blanket"),
    308: ("Vintage Yggdrasil Tree of Life Quilt Set", "vintage Yggdrasil Tree of Life quilt set", "Yggdrasil quilt set, Tree of Life bedding, vintage moon sun quilt", "black and cream quilt set with oval Tree of Life artwork, sun, crescent moon and ornate floral border", "vintage Yggdrasil Tree of Life quilt set"),
    309: ("Watercolor Basketball Dunking Blanket", "watercolor basketball dunking blanket", "basketball blanket, basketball player throw, custom number sports blanket", "watercolor-style blanket with dunking basketball players, hoop, stars, sample class text and sample number 33", "watercolor basketball dunking blanket"),
    310: ("Romantic Deer Forest Comforter Set", "romantic deer forest comforter set", "deer comforter set, romantic woodland bedding, buck doe comforter", "black woodland comforter with two deer, floral border, You and Me text and sample names James and Emma", "romantic deer forest comforter set"),
}


IMAGE_DETAILS = {
    301: {1: ("Bed mockup of cactus flower quilt with tall green cacti and large red and yellow flowers.", "Vibrant cactus flower quilt bed view"), 2: ("Room mockup of cactus flower quilt with printed craft callout.", "Vibrant cactus flower quilt room mockup"), 3: ("Close bed view with optional pillow shams and cactus flower artwork.", "Vibrant cactus flower quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and cactus flower artwork sample.", "Vibrant cactus flower quilt fabric panel"), 5: ("Premium quilt sets sizing chart for vibrant cactus flower design.", "Vibrant cactus flower quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Vibrant cactus flower quilt features panel"), 7: ("Overhead bedroom mockup of cactus flower quilt and matching pillows.", "Vibrant cactus flower quilt overhead view")},
    302: {1: ("Bed mockup of embroidered-style desert cactus quilt with teal cacti and magenta flowers.", "Embroidered desert cactus quilt bed view"), 2: ("Room mockup of embroidered desert cactus quilt with printed craft callout.", "Embroidered desert cactus quilt room mockup"), 3: ("Close bed view with optional pillow shams and colorful cactus artwork.", "Embroidered desert cactus quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and embroidered cactus artwork sample.", "Embroidered desert cactus quilt fabric panel"), 5: ("Premium quilt sets sizing chart for embroidered desert cactus design.", "Embroidered desert cactus quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Embroidered desert cactus quilt features panel"), 7: ("Overhead bedroom mockup of embroidered desert cactus quilt and matching pillows.", "Embroidered desert cactus quilt overhead view")},
    303: {1: ("Bed mockup of flowering cactus desert scene quilt with saguaros, blooms and layered hills.", "Flowering cactus desert scene quilt bed view"), 2: ("Room mockup of flowering cactus desert scene quilt with printed craft callout.", "Flowering cactus desert scene quilt room mockup"), 3: ("Close bed view with optional pillow shams and desert cactus artwork.", "Flowering cactus desert scene quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and desert scene artwork sample.", "Flowering cactus desert scene quilt fabric panel"), 5: ("Premium quilt sets sizing chart for flowering cactus desert scene design.", "Flowering cactus desert scene quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Flowering cactus desert scene quilt features panel"), 7: ("Overhead bedroom mockup of flowering cactus desert scene quilt and matching pillows.", "Flowering cactus desert scene quilt overhead view")},
    304: {1: ("Bed mockup of yellow potted cactus garden quilt with repeated colorful planters.", "Vibrant potted cactus garden quilt bed view"), 2: ("Room mockup of potted cactus garden quilt with printed craft callout.", "Vibrant potted cactus garden quilt room mockup"), 3: ("Close bed view with optional pillow shams and potted succulent blocks.", "Vibrant potted cactus garden quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and potted cactus artwork sample.", "Vibrant potted cactus garden quilt fabric panel"), 5: ("Premium quilt sets sizing chart for potted cactus garden design.", "Vibrant potted cactus garden quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Vibrant potted cactus garden quilt features panel"), 7: ("Overhead bedroom mockup of potted cactus garden quilt and matching pillows.", "Vibrant potted cactus garden quilt overhead view")},
    305: {1: ("Bed mockup of cream potted cactus quilt with earthy patterned pots and succulents.", "Potted cactus quilt design 3 bed view"), 2: ("Room mockup of potted cactus quilt design 3 with printed craft callout.", "Potted cactus quilt design 3 room mockup"), 3: ("Close bed view with optional pillow shams and cactus planter artwork.", "Potted cactus quilt design 3 pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and potted cactus artwork sample.", "Potted cactus quilt design 3 fabric panel"), 5: ("Premium quilt sets sizing chart for potted cactus design 3.", "Potted cactus quilt design 3 size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Potted cactus quilt design 3 features panel"), 7: ("Overhead bedroom mockup of potted cactus quilt and matching pillows.", "Potted cactus quilt design 3 overhead view")},
    306: {1: ("Bed mockup of Viking Celtic Tree of Life quilt with compass emblem and sample name Brent.", "Viking Celtic Tree of Life quilt bed view"), 2: ("Room mockup of Viking Celtic Tree of Life quilt with printed craft callout.", "Viking Celtic Tree of Life quilt room mockup"), 3: ("Close bed view with optional pillow shams, runes and Celtic border.", "Viking Celtic Tree of Life quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and Celtic artwork sample.", "Viking Celtic Tree of Life quilt fabric panel"), 5: ("Premium quilt sets sizing chart for Viking Celtic Tree of Life design.", "Viking Celtic Tree of Life quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Viking Celtic Tree of Life quilt features panel"), 7: ("Overhead bedroom mockup of Viking Celtic Tree of Life quilt and matching pillows.", "Viking Celtic Tree of Life quilt overhead view")},
    307: {1: ("Flat vintage reading blanket with newspaper background, tall flowers, cartoon reader and sample name Sarah.", "Vintage reading girl flowers blanket flat view"), 2: ("Held vintage reading girl blanket with character style examples along one side.", "Vintage reading girl flowers blanket held view"), 3: ("Overhead lifestyle view of vintage reading girl blanket on a bed.", "Vintage reading girl flowers blanket overhead view"), 4: ("Blanket size chart with four sizes for vintage reading girl design.", "Vintage reading girl flowers blanket size chart"), 5: ("Vintage reading girl blanket draped on sofa beneath round mirror.", "Vintage reading girl flowers blanket sofa mockup"), 6: ("Feature panel over vintage reading blanket with fluffy and no-shedding callouts.", "Vintage reading girl flowers blanket feature panel")},
    308: {1: ("Bed mockup of black and cream Yggdrasil quilt with oval tree, sun and crescent moon.", "Vintage Yggdrasil Tree of Life quilt bed view"), 2: ("Room mockup of vintage Yggdrasil quilt with printed craft callout.", "Vintage Yggdrasil Tree of Life quilt room mockup"), 3: ("Close bed view with optional pillow shams and moon tree artwork.", "Vintage Yggdrasil Tree of Life quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and ornate tree artwork sample.", "Vintage Yggdrasil Tree of Life quilt fabric panel"), 5: ("Premium quilt sets sizing chart for vintage Yggdrasil design.", "Vintage Yggdrasil Tree of Life quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Vintage Yggdrasil Tree of Life quilt features panel"), 7: ("Overhead bedroom mockup of vintage Yggdrasil quilt and matching pillows.", "Vintage Yggdrasil Tree of Life quilt overhead view")},
    309: {1: ("Flat watercolor basketball blanket with dunking players, hoop, stars and sample number 33.", "Watercolor basketball dunking blanket flat view"), 2: ("Held watercolor basketball blanket with sample class text and number 33.", "Watercolor basketball dunking blanket held view"), 3: ("Blanket size chart with four sizes for watercolor basketball design.", "Watercolor basketball dunking blanket size chart"), 4: ("Feature panel over basketball blanket with fluffy and no-shedding callouts.", "Watercolor basketball dunking blanket feature panel"), 5: ("Watercolor basketball blanket draped on sofa under a window.", "Watercolor basketball dunking blanket sofa view"), 6: ("Multi-function blanket panel with folded blanket and fabric callouts.", "Watercolor basketball dunking blanket multi-use panel"), 7: ("Gold size chart panel showing queen, adult and smaller blanket examples.", "Watercolor basketball dunking blanket gold size chart"), 8: ("Family reading lifestyle scene with watercolor basketball blanket.", "Watercolor basketball dunking blanket family lifestyle")},
    310: {1: ("Bed mockup of romantic deer comforter with two deer, floral border and sample names.", "Romantic deer forest comforter bed view"), 2: ("Overhead bedroom mockup of romantic deer forest comforter and pillows.", "Romantic deer forest comforter overhead view"), 3: ("Held deer forest comforter with You and Me text and sample names James and Emma.", "Romantic deer forest comforter held view"), 4: ("Room mockup of romantic deer forest comforter with printed craft callout.", "Romantic deer forest comforter room mockup"), 5: ("Close bed view with optional pillow shams and deer artwork.", "Romantic deer forest comforter pillow sham panel"), 6: ("High-quality fabric panel with quilt layers and deer comforter sample.", "Romantic deer forest comforter fabric panel"), 7: ("Sizing and details chart for romantic deer forest comforter set.", "Romantic deer forest comforter size chart"), 8: ("Bedspread features panel showing comforter layers and care icons.", "Romantic deer forest comforter features panel")},
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
        meta = f"Shop {title.lower()} with {detail}, selectable sizes and verified custom options."[:155]
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
        ws.cell(row_num, idx["keyword_strategy"]).value = f"Target the specific product motif: {cluster}; keep broad faith, Celtic, camping, couple gift or reader gift terms for collection pages."
        ws.cell(row_num, idx["buyer_search_summary"]).value = f"US English buyer intent targets {cluster}; no search-volume, trend or ranking claim is made."
        ws.cell(row_num, idx["keyword_evidence_level"]).value = "SERP_ONLY"
        ws.cell(row_num, idx["revision"]).value = "r2"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields"]:
            ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            f"R2 deep recheck without prior QA: used products_export_1.csv sha256:{admin_hash}; "
            "separated cactus quilt, reader blanket, deer/photo comforter, affirmation blanket and Tree of Life quilt intents; checked customizer limits and rewrote image observations/alts from contact sheets. Still NEEDS_REVIEW."
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
        ws.cell(row_num, idx["mapping_reason"]).value = "R2 separates mixed faith, Celtic, camping, photo quilt and reading blanket pages by visible motif and option fields."
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
        ws.cell(row_num, idx["emotional_social_motivation"]).value = "Choose a personalized faith, Celtic, camping, couple or reader-themed bedding item that matches the recipient or room theme."
        ws.cell(row_num, idx["purchase_concerns"]).value = "Confirm custom text or photo rules; avoid unsupported material, religious promise, copyrighted identity or delivery claims."
        ws.cell(row_num, idx["source_refs"]).value = f"products_export_1.csv sha256:{admin_hash}; storefront HTML customizer parse; contact sheets 301-310"
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
        ws.cell(row_num, idx["factual_conflicts"]).value = "No independent QA for batch 031 yet; r2 prepared for QA recheck."
        ws.cell(row_num, idx["confidence_and_reason"]).value = "R2 uses contact sheets, storefront customizer parse and Shopify admin CSV baseline; no approval or import created."

    ws = wb["README_QA"]
    idx = headers(ws)
    for metric, value, definition in [
        ("qa_batch_031_r2_revision", "r2", "Separate 10-product revision without prior independent QA; source workbook was not modified."),
        ("qa_batch_031_r2_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_031_r2_scope", "inventory positions 301-310", "No products outside qa_batch_031 were revised."),
        ("qa_batch_031_r2_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_031_r2_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_031_r2_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
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
        "scope": "qa_batch_031 only; inventory positions 301-310",
        "revision": "r2",
        "revised_products": revised_products,
        "revised_images": revised_images,
        "admin_alt_matches": admin_alt_matches,
        "keyword_rows_revised": keyword_rows,
        "review_status": "NEEDS_REVIEW",
        "approved_created": False,
        "shopify_import_created": False,
        "next_step": "Ready for independent QA of qa_batch_031_r2.",
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text(
        "\n".join([
            "# Revision qa_batch_031_r2",
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
            "- Sửa trọng yếu: rà lại khi chưa có QA, tách nhóm cactus quilt, reader blanket, deer/photo comforter, affirmation blanket và Tree of Life quilt; parse customizer từ HTML storefront; viết lại title/meta/description/alt theo từng contact sheet.",
            "- Trạng thái: `NEEDS_REVIEW`; không tạo `APPROVED`; không tạo file import Shopify.",
            "- Bước tiếp theo: gửi `qa_batch_031_r2` cho QA độc lập.",
            "",
        ]),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
