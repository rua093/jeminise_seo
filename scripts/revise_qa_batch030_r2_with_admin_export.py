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
BATCH_ID = "qa_batch_030_r2"
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_030.xlsx"
SUMMARY_PATH = ROOT / "seo_runs" / SHOP / RUN_ID / "batch_030_evidence_summary.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_030_r2.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_030_r2.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    291: ("Reading Girl Book Flowers Blanket", "reading girl book flowers blanket", "book lover blanket, reading girl throw blanket, floral book blanket", "beige throw blanket with girl sitting on an open book, flowers, notebook pattern and sample name Rebecca", "reading girl book flowers blanket"),
    292: ("Reading Girl Heart Bookshelf Blanket", "reading girl heart bookshelf blanket", "bookshelf blanket, book lover throw blanket, girl reading blanket", "newspaper-style throw blanket with heart bookshelf, stacked book titles, butterflies and sample name Elizabeth", "reading girl heart bookshelf blanket"),
    293: ("Best Mom Ever Photo Quilt Set", "Best Mom Ever photo quilt set", "mom photo quilt, Mother's Day quilt set, custom family photo quilt", "black-and-white photo collage quilt set with Best Mom Ever text, daisies and custom photo badge in the image", "Best Mom Ever photo quilt set"),
    294: ("Seasonal Bookworms Sophia Blanket", "seasonal bookworms blanket", "seasonal bookworm blanket, reader blanket, library book lover throw", "library-themed blanket with Seasonal Bookworms text, illustrated reader, plants and sample name Sophia", "seasonal bookworms blanket"),
    295: ("Southwestern Wolf Head Comforter Set", "southwestern wolf head comforter set", "wolf comforter set, southwestern bedding, geometric wolf quilt", "southwestern-style comforter with large wolf head portrait, geometric border and matching wolf pillows", "southwestern wolf head comforter set"),
    296: ("Sunflower Christian Bible Verse Blanket", "sunflower Christian Bible verse blanket", "Christian Bible verse blanket, sunflower affirmation blanket, God says you are blanket", "yellow sunflower blanket with silhouette profile, God Says You Are scripture words and sample name Jessica", "sunflower Christian Bible verse blanket"),
    297: ("Sunflower God Says You Are Blanket", "sunflower God Says You Are blanket", "God Says You Are blanket, sunflower Christian blanket, Bible verse throw blanket", "rustic sunflower blanket with woodgrain silhouette, scripture affirmations and sample name Sophia", "sunflower God Says You Are blanket"),
    298: ("Sunflower Self Affirmations Blanket", "sunflower self affirmations blanket", "self affirmation blanket, sunflower affirmation throw, personalized encouragement blanket", "cream blanket with My Self-Affirmations text, sunflowers, butterflies, portrait sample and sample name Jennifer", "sunflower self affirmations blanket"),
    299: ("Teal Argyle Reading Girl Blanket", "teal argyle reading girl blanket", "reading girl blanket, book lover throw, teal book blanket", "teal argyle throw blanket with reader seen from behind, dragonflies, books and And She Read Happily Ever After text", "teal argyle reading girl blanket"),
    300: ("Vibrant Cactus Desert Flowers Quilt Set", "vibrant cactus desert flowers quilt set", "cactus flower quilt set, desert flowers bedding, succulent quilt set", "dark quilt set with vibrant cactus blooms, succulent forms, patterned pots and desert flower artwork", "vibrant cactus desert flowers quilt set"),
}


IMAGE_DETAILS = {
    291: {1: ("Held beige reading blanket with girl on an open book, flowers and sample name Rebecca.", "Reading girl book flowers blanket held view"), 2: ("Sofa mockup of reading girl blanket with books and mug nearby.", "Reading girl book flowers blanket sofa mockup"), 3: ("Detail collage showing notebook print, red script text and fleece texture.", "Reading girl book flowers blanket detail collage"), 4: ("Reading lifestyle scene with open book and small pet on the blanket.", "Reading girl book flowers blanket reading lifestyle"), 5: ("Feature panel showing white fleece edge and machine washable, soft and warm callouts.", "Reading girl book flowers blanket feature panel"), 6: ("Close-up reading scene showing red script text and notebook background.", "Reading girl book flowers blanket close-up"), 7: ("Size and use panel for reading girl book flowers blanket.", "Reading girl book flowers blanket size panel")},
    292: {1: ("Held reading blanket with heart bookshelf, stacked book titles and sample name Elizabeth.", "Reading girl heart bookshelf blanket held view"), 2: ("Blanket size chart with four sizes for heart bookshelf reader design.", "Reading girl heart bookshelf blanket size chart"), 3: ("Sofa mockup of heart bookshelf blanket beneath round mirror.", "Reading girl heart bookshelf blanket sofa mockup"), 4: ("Feature panel over newspaper-style blanket with fluffy and no-shedding callouts.", "Reading girl heart bookshelf blanket feature panel"), 5: ("Reader wrapped in heart bookshelf blanket while holding a red book.", "Reading girl heart bookshelf blanket wrapped reading view"), 6: ("Reading lifestyle scene with open book and small pet on the blanket.", "Reading girl heart bookshelf blanket reading lifestyle"), 7: ("Second sofa view of heart bookshelf blanket under a window.", "Reading girl heart bookshelf blanket sofa view"), 8: ("Multiple purpose panel for heart bookshelf blanket with bed, sofa, office, travel and pet examples.", "Reading girl heart bookshelf blanket multi-use panel"), 9: ("Second held view of heart bookshelf blanket with large top book-lover text.", "Reading girl heart bookshelf blanket second held view")},
    293: {1: ("Flat Best Mom Ever photo collage quilt with daisies and custom photo badge.", "Best Mom Ever photo quilt flat view"), 2: ("Premium quilt sets sizing chart for Best Mom Ever photo collage design.", "Best Mom Ever photo quilt size chart"), 3: ("Bedspread features panel showing quilt layers and care icons.", "Best Mom Ever photo quilt features panel"), 4: ("Overhead bedroom mockup of Best Mom Ever photo quilt with black border.", "Best Mom Ever photo quilt overhead view"), 5: ("High-quality fabric panel with quilt layers and Best Mom Ever photo collage sample.", "Best Mom Ever photo quilt fabric panel"), 6: ("Room mockup of Best Mom Ever photo quilt with printed craft callout.", "Best Mom Ever photo quilt room mockup"), 7: ("Close bed view with optional pillow shams and Best Mom Ever photo collage artwork.", "Best Mom Ever photo quilt pillow sham panel")},
    294: {1: ("Held Seasonal Bookworms blanket with library shelves, illustrated reader and sample name Sophia.", "Seasonal Bookworms blanket held view"), 2: ("Blanket size chart with four sizes for Seasonal Bookworms design.", "Seasonal Bookworms blanket size chart"), 3: ("Seasonal Bookworms blanket draped on sofa beneath round mirror.", "Seasonal Bookworms blanket sofa mockup"), 4: ("Feature panel over library blanket with fluffy and no-shedding callouts.", "Seasonal Bookworms blanket feature panel"), 5: ("Multiple purpose panel for Seasonal Bookworms blanket with bed, sofa, office, travel and pet examples.", "Seasonal Bookworms blanket multi-use panel"), 6: ("Second held view of Seasonal Bookworms blanket with library background.", "Seasonal Bookworms blanket second held view"), 7: ("Reader wrapped in Seasonal Bookworms blanket while holding a red book.", "Seasonal Bookworms blanket wrapped reading view"), 8: ("Reading lifestyle scene with open book and small pet on Seasonal Bookworms blanket.", "Seasonal Bookworms blanket reading lifestyle"), 9: ("Sofa view of Seasonal Bookworms blanket under a window.", "Seasonal Bookworms blanket sofa view")},
    295: {1: ("Bed mockup of southwestern wolf head comforter with geometric border and matching pillows.", "Southwestern wolf head comforter bed view"), 2: ("Close-up of wolf face artwork with golden eyes and quilted texture.", "Southwestern wolf head comforter close-up"), 3: ("Single pillow sham mockup with wolf head and southwestern border.", "Southwestern wolf head pillow sham"), 4: ("Room mockup of southwestern wolf head comforter on bed.", "Southwestern wolf head comforter room mockup"), 5: ("Size and set panel listing one premium quilt and optional standard shams.", "Southwestern wolf head comforter size panel")},
    296: {1: ("Held sunflower Christian blanket with silhouette profile, scripture words and sample name Jessica.", "Sunflower Christian Bible verse blanket held view"), 2: ("Second held view of sunflower Christian blanket with profile style examples.", "Sunflower Christian Bible verse blanket style options"), 3: ("Flat sunflower Bible verse blanket with God Says You Are text.", "Sunflower Christian Bible verse blanket flat view"), 4: ("Sofa mockup of sunflower Christian blanket with scripture words.", "Sunflower Christian Bible verse blanket sofa mockup"), 5: ("Second sofa view of sunflower Christian blanket beneath round mirror.", "Sunflower Christian Bible verse blanket sofa view"), 6: ("Feature panel over sunflower Bible verse blanket with fluffy and no-shedding callouts.", "Sunflower Christian Bible verse blanket feature panel")},
    297: {1: ("Held rustic sunflower blanket with woodgrain silhouette, scripture affirmations and sample name Sophia.", "Sunflower God Says You Are blanket held view"), 2: ("Feature panel showing white fleece edge with machine washable and stitching callouts.", "Sunflower God Says You Are blanket feature panel"), 3: ("Close-up reading scene showing woodgrain silhouette and scripture words.", "Sunflower God Says You Are blanket reading close-up"), 4: ("Size and use panel for sunflower God Says You Are blanket.", "Sunflower God Says You Are blanket size panel"), 5: ("Detail collage showing sunflower artwork, silhouette text and fleece texture.", "Sunflower God Says You Are blanket detail collage"), 6: ("Sunflower God Says You Are blanket draped over an armchair.", "Sunflower God Says You Are blanket chair mockup"), 7: ("Reading lifestyle scene with sunflower God Says You Are blanket and open book.", "Sunflower God Says You Are blanket reading lifestyle")},
    298: {1: ("Held self-affirmations blanket with sunflowers, butterflies, portrait sample and sample name Jennifer.", "Sunflower self affirmations blanket held view"), 2: ("Flat self-affirmations blanket with cream background and sunflower border.", "Sunflower self affirmations blanket flat view"), 3: ("Self-affirmations blanket draped on sofa under a window.", "Sunflower self affirmations blanket sofa view"), 4: ("Second sofa mockup of self-affirmations blanket beneath round mirror.", "Sunflower self affirmations blanket sofa mockup"), 5: ("Feature panel over self-affirmations blanket with fluffy and no-shedding callouts.", "Sunflower self affirmations blanket feature panel"), 6: ("Photo guideline panel showing recommended and discouraged portrait examples.", "Sunflower self affirmations blanket photo guidelines")},
    299: {1: ("Held teal argyle reading blanket with back-view reader, books, dragonflies and sample name Laura.", "Teal argyle reading girl blanket held view"), 2: ("Second held view of teal reading blanket with hair style examples.", "Teal argyle reading girl blanket style options"), 3: ("Overhead lifestyle view of teal reading blanket on a bed.", "Teal argyle reading girl blanket overhead view"), 4: ("Blanket size chart with four sizes for teal reading girl design.", "Teal argyle reading girl blanket size chart"), 5: ("Reader wrapped in teal argyle blanket with large And She Read text.", "Teal argyle reading girl blanket wrapped view"), 6: ("Teal argyle reading blanket draped on sofa beneath round mirror.", "Teal argyle reading girl blanket sofa mockup"), 7: ("Feature panel over teal reading blanket with fluffy and no-shedding callouts.", "Teal argyle reading girl blanket feature panel")},
    300: {1: ("Bed mockup of dark cactus flower quilt with vibrant blooms and patterned pots.", "Vibrant cactus desert flowers quilt bed view"), 2: ("Room mockup of vibrant cactus desert flowers quilt with printed craft callout.", "Vibrant cactus desert flowers quilt room mockup"), 3: ("Close bed view with optional pillow shams and cactus flower artwork.", "Vibrant cactus desert flowers quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and cactus flower artwork sample.", "Vibrant cactus desert flowers quilt fabric panel"), 5: ("Premium quilt sets sizing chart for vibrant cactus desert flowers design.", "Vibrant cactus desert flowers quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Vibrant cactus desert flowers quilt features panel"), 7: ("Overhead bedroom mockup of vibrant cactus desert flowers quilt and matching pillows.", "Vibrant cactus desert flowers quilt overhead view")},
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
        ws.cell(row_num, idx["source_refs"]).value = f"products_export_1.csv sha256:{admin_hash}; storefront HTML customizer parse; contact sheets 291-300"
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
        ws.cell(row_num, idx["factual_conflicts"]).value = "No independent QA for batch 030 yet; r2 prepared for QA recheck."
        ws.cell(row_num, idx["confidence_and_reason"]).value = "R2 uses contact sheets, storefront customizer parse and Shopify admin CSV baseline; no approval or import created."

    ws = wb["README_QA"]
    idx = headers(ws)
    for metric, value, definition in [
        ("qa_batch_030_r2_revision", "r2", "Separate 10-product revision without prior independent QA; source workbook was not modified."),
        ("qa_batch_030_r2_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_030_r2_scope", "inventory positions 291-300", "No products outside qa_batch_030 were revised."),
        ("qa_batch_030_r2_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_030_r2_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_030_r2_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
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
        "scope": "qa_batch_030 only; inventory positions 291-300",
        "revision": "r2",
        "revised_products": revised_products,
        "revised_images": revised_images,
        "admin_alt_matches": admin_alt_matches,
        "keyword_rows_revised": keyword_rows,
        "review_status": "NEEDS_REVIEW",
        "approved_created": False,
        "shopify_import_created": False,
        "next_step": "Ready for independent QA of qa_batch_030_r2.",
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text(
        "\n".join([
            "# Revision qa_batch_030_r2",
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
            "- Bước tiếp theo: gửi `qa_batch_030_r2` cho QA độc lập.",
            "",
        ]),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
