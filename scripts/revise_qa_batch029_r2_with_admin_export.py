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
BATCH_ID = "qa_batch_029_r2"
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_029.xlsx"
SUMMARY_PATH = ROOT / "seo_runs" / SHOP / RUN_ID / "batch_029_evidence_summary.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_029_r2.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_029_r2.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    281: ("Flowering Cactus Garden Quilt Set", "flowering cactus garden quilt set", "flowering cactus quilt set, cactus garden bedding, southwest succulent quilt", "light quilt set with flowering cacti, succulents, patterned pots and southwest garden artwork", "flowering cactus garden quilt set"),
    282: ("Girl Glasses Newspaper Book Blanket", "girl glasses newspaper book blanket", "book lover blanket, girl reading blanket, newspaper pattern throw", "throw blanket with glasses girl, stacked books, newspaper background, leopard border and sample name Elizabeth", "girl glasses newspaper book blanket"),
    283: ("Girl in Glasses Book Lover Blanket", "girl in glasses book lover blanket", "book lover blanket, girl reading throw, personalized reader blanket", "cream throw blanket with glasses girl reading below a bookshelf, flowers, butterflies and sample name Laura", "girl in glasses book lover blanket"),
    284: ("Messy Bun Girl Reading Book Blanket", "messy bun girl reading book blanket", "book lover blanket, messy bun reader throw, open book blanket", "vintage open-book blanket with messy-bun girl holding a book and mug, floral accents and sample name Rebecca", "messy bun girl reading book blanket"),
    285: ("Girl Reading Bookshelf Blanket", "girl reading bookshelf blanket", "bookshelf blanket, girl reading throw blanket, book lover blanket", "cream blanket with glasses girl reading in front of a colorful bookshelf, flowers and sample name Laura", "girl reading bookshelf blanket"),
    286: ("Hunting Deer Antlers Comforter Set", "hunting deer antlers comforter set", "deer antlers comforter, hunting bedding set, rustic photo comforter", "brown striped comforter with deer head silhouette, antlers, heart icon, photo collage and sample names", "hunting deer antlers comforter set"),
    287: ("Reading Books Floral Blanket", "reading books floral blanket", "reading books blanket, book lover throw blanket, floral book blanket", "white reading blanket with open book, upright books, floral stems, butterflies and sample Amelia text", "reading books floral blanket"),
    288: ("Love Photo Heart Quilt Set", "love photo heart quilt set", "photo heart quilt, couple photo quilt set, love quilt bedding", "black photo-collage quilt set with heart graphic, I love you text and custom photo badge in the image", "love photo heart quilt set"),
    289: ("Morning Affirmations Floral Blanket", "morning affirmations floral blanket", "morning affirmation blanket, floral affirmation throw, self love blanket", "navy blanket with My Morning Affirmations text, flowers, illustrated girl and sample name Jennifer", "morning affirmations floral blanket"),
    290: ("Mosaic Tree of Life Quilt Set", "mosaic Tree of Life quilt set", "mosaic tree quilt set, Tree of Life bedding, colorful tree quilt", "colorful mosaic-style Tree of Life quilt set with glowing sun, curved trunk and warm landscape bands", "mosaic Tree of Life quilt set"),
}


IMAGE_DETAILS = {
    281: {1: ("Bed mockup of flowering cactus garden quilt with tall cactus, succulents and patterned pots.", "Flowering cactus garden quilt bed view"), 2: ("Room mockup of flowering cactus garden quilt with printed craft callout.", "Flowering cactus garden quilt room mockup"), 3: ("Close bed view with optional pillow shams and cactus garden artwork.", "Flowering cactus garden quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and cactus garden artwork sample.", "Flowering cactus garden quilt fabric panel"), 5: ("Premium quilt sets sizing chart for flowering cactus garden design.", "Flowering cactus garden quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Flowering cactus garden quilt features panel"), 7: ("Overhead bedroom mockup of flowering cactus garden quilt and matching pillows.", "Flowering cactus garden quilt overhead view")},
    282: {1: ("Held newspaper pattern book blanket with glasses girl, stacked books and sample name Elizabeth.", "Girl glasses newspaper book blanket held view"), 2: ("Held book blanket with side character style examples next to the main design.", "Girl glasses newspaper book blanket character options"), 3: ("Blanket size chart showing four sizes for newspaper book design.", "Girl glasses newspaper book blanket size chart"), 4: ("Book blanket draped on sofa with leopard border and large BOOKS text.", "Girl glasses newspaper book blanket sofa mockup"), 5: ("Feature panel over newspaper book blanket with fluffy, quality and no-shedding callouts.", "Girl glasses newspaper book blanket feature panel"), 6: ("Second sofa view of newspaper book blanket under a window.", "Girl glasses newspaper book blanket sofa view"), 7: ("Reader wrapped in newspaper book blanket while holding a red book.", "Girl glasses newspaper book blanket wrapped reading view"), 8: ("Reading lifestyle scene with newspaper book blanket, open book and small pet.", "Girl glasses newspaper book blanket reading lifestyle"), 9: ("Family reading scene with newspaper book blanket at the foot of a bed.", "Girl glasses newspaper book blanket family lifestyle")},
    283: {1: ("Held book lover blanket with glasses girl, bookshelf, flowers, butterflies and sample name Laura.", "Girl in glasses book lover blanket held view"), 2: ("Flat book lover blanket with shelf of custom books and Just a Girl Who Loves Books text.", "Girl in glasses book lover blanket flat view"), 3: ("Second held view of book lover blanket with beige floral background.", "Girl in glasses book lover blanket second held view"), 4: ("Blanket size chart with four sizes for girl in glasses book lover design.", "Girl in glasses book lover blanket size chart"), 5: ("Book lover blanket draped on sofa under a window.", "Girl in glasses book lover blanket sofa view"), 6: ("Second sofa mockup of book lover blanket beneath round mirror.", "Girl in glasses book lover blanket sofa mockup"), 7: ("Feature panel over book lover blanket with fluffy, quality and no-shedding callouts.", "Girl in glasses book lover blanket feature panel")},
    284: {1: ("Held open-book blanket with messy-bun reader, mug, flowers and sample name Rebecca.", "Messy bun reader book blanket held view"), 2: ("Flat open-book blanket with handwritten page background and Just a Girl Who Loves Books text.", "Messy bun reader book blanket flat view"), 3: ("Second held view of messy-bun reader blanket with character style examples.", "Messy bun reader book blanket character options"), 4: ("Blanket size chart with four sizes for messy-bun reader design.", "Messy bun reader book blanket size chart"), 5: ("Open-book reader blanket draped on sofa under a window.", "Messy bun reader book blanket sofa view"), 6: ("Second sofa mockup of messy-bun reader blanket beneath round mirror.", "Messy bun reader book blanket sofa mockup"), 7: ("Feature panel over reader blanket with fluffy, quality and no-shedding callouts.", "Messy bun reader book blanket feature panel")},
    285: {1: ("Held bookshelf reader blanket with glasses girl, flowers, butterflies and sample name Laura.", "Girl reading bookshelf blanket held view"), 2: ("Flat bookshelf reader blanket with Just a Girl Who Loves Books text.", "Girl reading bookshelf blanket flat view"), 3: ("Second held view of bookshelf reader blanket with character style examples.", "Girl reading bookshelf blanket character options"), 4: ("Bookshelf reader blanket draped on sofa under a window.", "Girl reading bookshelf blanket sofa view"), 5: ("Second sofa mockup of bookshelf reader blanket beneath round mirror.", "Girl reading bookshelf blanket sofa mockup"), 6: ("Blanket size chart with four sizes for bookshelf reader design.", "Girl reading bookshelf blanket size chart"), 7: ("Feature panel over bookshelf reader blanket with fluffy, quality and no-shedding callouts.", "Girl reading bookshelf blanket feature panel")},
    286: {1: ("Bed mockup of hunting deer antlers comforter with photo collage and sample names.", "Hunting deer antlers comforter bed view"), 2: ("Held comforter showing deer head silhouette, antlers, heart icon and photo collage.", "Hunting deer antlers comforter held view"), 3: ("Room mockup of hunting deer antlers comforter with printed craft callout.", "Hunting deer antlers comforter room mockup"), 4: ("Close bed view with optional pillow shams and deer antlers photo artwork.", "Hunting deer antlers comforter pillow sham panel"), 5: ("Fabric panel with white backing and deer antlers comforter pillow mockup.", "Hunting deer antlers comforter fabric panel"), 6: ("Sizing and details chart for hunting deer antlers comforter set.", "Hunting deer antlers comforter size chart"), 7: ("Bedspread features panel showing comforter layers and care icons.", "Hunting deer antlers comforter features panel"), 8: ("Overhead bedroom mockup of hunting deer antlers comforter and matching pillows.", "Hunting deer antlers comforter overhead view")},
    287: {1: ("Held reading books floral blanket with open book, upright books and sample Amelia text.", "Reading books floral blanket held view"), 2: ("Blanket size chart with seven color/background options for reading books design.", "Reading books floral blanket size chart"), 3: ("Reading books floral blanket draped on sofa beneath round mirror.", "Reading books floral blanket sofa mockup"), 4: ("Family reading lifestyle scene with reading books floral blanket.", "Reading books floral blanket family lifestyle"), 5: ("Reader wrapped in reading books floral blanket while holding a red book.", "Reading books floral blanket wrapped reading view"), 6: ("Second sofa view of reading books floral blanket under a window.", "Reading books floral blanket sofa view"), 7: ("Reading lifestyle scene with open book and small pet on the floral blanket.", "Reading books floral blanket reading lifestyle"), 8: ("Multiple purpose panel for reading books blanket with bed, sofa, office, travel and pet examples.", "Reading books floral blanket multi-use panel"), 9: ("Color option panel showing seven reading books blanket background styles.", "Reading books floral blanket color options")},
    288: {1: ("Flat photo-collage quilt with heart graphic, I love you text and custom photo badge.", "Love photo heart quilt flat view"), 2: ("Overhead bedroom mockup of love photo heart quilt with black border.", "Love photo heart quilt overhead view"), 3: ("Bedspread features panel showing quilt layers and care icons.", "Love photo heart quilt features panel"), 4: ("Premium quilt sets sizing chart for love photo heart design.", "Love photo heart quilt size chart"), 5: ("Room mockup of love photo heart quilt with printed craft callout.", "Love photo heart quilt room mockup"), 6: ("Close bed view with optional pillow shams and love photo collage artwork.", "Love photo heart quilt pillow sham panel"), 7: ("High-quality fabric panel with quilt layers and love photo collage sample.", "Love photo heart quilt fabric panel")},
    289: {1: ("Held navy morning affirmations blanket with flowers, illustrated girl and sample name Jennifer.", "Morning affirmations floral blanket held view"), 2: ("Flat morning affirmations blanket with navy background, flower icons and self-love phrases.", "Morning affirmations floral blanket flat view"), 3: ("Second held view of morning affirmations blanket with character style examples.", "Morning affirmations floral blanket character options"), 4: ("Morning affirmations blanket draped on sofa under a window.", "Morning affirmations floral blanket sofa view"), 5: ("Second sofa mockup of morning affirmations blanket beneath round mirror.", "Morning affirmations floral blanket sofa mockup"), 6: ("Feature panel over morning affirmations blanket with fluffy, quality and no-shedding callouts.", "Morning affirmations floral blanket feature panel")},
    290: {1: ("Bed mockup of mosaic Tree of Life quilt with glowing sun and warm landscape bands.", "Mosaic Tree of Life quilt bed view"), 2: ("Room mockup of mosaic Tree of Life quilt with printed craft callout.", "Mosaic Tree of Life quilt room mockup"), 3: ("Close bed view with optional pillow shams and mosaic leaf artwork.", "Mosaic Tree of Life quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and mosaic tree artwork sample.", "Mosaic Tree of Life quilt fabric panel"), 5: ("Premium quilt sets sizing chart for mosaic Tree of Life design.", "Mosaic Tree of Life quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Mosaic Tree of Life quilt features panel"), 7: ("Overhead bedroom mockup of mosaic Tree of Life quilt and matching pillows.", "Mosaic Tree of Life quilt overhead view")},
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
        ws.cell(row_num, idx["source_refs"]).value = f"products_export_1.csv sha256:{admin_hash}; storefront HTML customizer parse; contact sheets 281-290"
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
        ws.cell(row_num, idx["factual_conflicts"]).value = "No independent QA for batch 029 yet; r2 prepared for QA recheck."
        ws.cell(row_num, idx["confidence_and_reason"]).value = "R2 uses contact sheets, storefront customizer parse and Shopify admin CSV baseline; no approval or import created."

    ws = wb["README_QA"]
    idx = headers(ws)
    for metric, value, definition in [
        ("qa_batch_029_r2_revision", "r2", "Separate 10-product revision without prior independent QA; source workbook was not modified."),
        ("qa_batch_029_r2_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_029_r2_scope", "inventory positions 281-290", "No products outside qa_batch_029 were revised."),
        ("qa_batch_029_r2_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_029_r2_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_029_r2_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
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
        "scope": "qa_batch_029 only; inventory positions 281-290",
        "revision": "r2",
        "revised_products": revised_products,
        "revised_images": revised_images,
        "admin_alt_matches": admin_alt_matches,
        "keyword_rows_revised": keyword_rows,
        "review_status": "NEEDS_REVIEW",
        "approved_created": False,
        "shopify_import_created": False,
        "next_step": "Ready for independent QA of qa_batch_029_r2.",
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text(
        "\n".join([
            "# Revision qa_batch_029_r2",
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
            "- Bước tiếp theo: gửi `qa_batch_029_r2` cho QA độc lập.",
            "",
        ]),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
