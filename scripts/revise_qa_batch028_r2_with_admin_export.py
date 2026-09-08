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
BATCH_ID = "qa_batch_028_r2"
SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_028.xlsx"
SUMMARY_PATH = ROOT / "seo_runs" / SHOP / RUN_ID / "batch_028_evidence_summary.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_028_r2.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_028_r2.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    271: ("Personalized Deer Couple Heart Comforter Set", "personalized deer couple comforter set", "deer couple comforter set, buck doe bedding, rustic couple comforter", "rustic woodland comforter set with buck and doe inside a heart frame plus sample couple text", "personalized deer couple comforter set"),
    272: ("Personalized Buck and Doe Comforter Set", "personalized buck and doe comforter set", "buck and doe comforter, deer couple bedding, rustic hunting couple bedding", "blue and white deer silhouette comforter with buck and doe artwork, forest border and sample names", "personalized buck and doe comforter set"),
    273: ("Personalized Deer Heart Outline Comforter", "personalized deer heart outline comforter", "deer heart comforter, buck doe comforter set, rustic couple bedding", "brown and white comforter with deer silhouettes forming a heart outline and sample couple names", "personalized deer heart outline comforter"),
    274: ("Desert Cactus Sunset Quilt Set", "desert cactus sunset quilt set", "desert cactus quilt set, southwest sunset bedding, cactus quilt bedding", "southwest desert quilt set with orange sun, layered hills, cacti and warm striped colors", "desert cactus sunset quilt set"),
    275: ("Desert Sunset Cactus Floral Quilt Set", "desert sunset cactus floral quilt set", "pink desert cactus quilt, southwest floral bedding, cactus sunset quilt set", "bright pink and orange desert quilt set with flowering cactus, layered sunset and bird silhouettes", "desert sunset cactus floral quilt set"),
    276: ("Floral Cross Christian Affirmation Blanket", "floral cross Christian affirmation blanket", "Christian affirmation blanket, floral cross blanket, Bible verse throw blanket", "white throw blanket with floral cross, I am chosen affirmation text and sample name Elizabeth", "floral cross Christian affirmation blanket"),
    277: ("God Says I Am Floral Cross Blanket", "God Says I Am floral cross blanket", "God Says I Am blanket, Christian verse blanket, floral cross throw blanket", "white floral cross blanket with God Says I Am scripture words, sample portrait and sample name Elizabeth", "God Says I Am floral cross blanket"),
    278: ("Floral Cross Personalized Text Blanket", "floral cross personalized text blanket", "floral cross blanket, personalized text blanket, Christian affirmation throw", "white floral cross blanket with I am affirmation text and sample name Elizabeth", "floral cross personalized text blanket"),
    279: ("Floral Line Art Book Lover Blanket", "floral line art book lover blanket", "book lover blanket, girl reading blanket, personalized reader throw", "cream blanket with line-art girl holding a book, butterflies, flowers and sample name Jessica", "floral line art book lover blanket"),
    280: ("Flowering Cactus Succulent Quilt Set", "flowering cactus succulent quilt set", "cactus succulent quilt set, desert floral bedding, flowering cactus quilt", "dark quilt set with flowering cacti, succulents, colorful pots and garden-style desert artwork", "flowering cactus succulent quilt set"),
}


IMAGE_DETAILS = {
    271: {1: ("Bed mockup of rustic deer couple comforter with buck and doe inside a heart frame.", "Deer couple heart comforter bed view"), 2: ("Held comforter showing buck, doe, heart border, sample couple quote and sample names.", "Deer couple heart comforter held view"), 3: ("Room mockup of deer couple heart comforter with printed craft callout.", "Deer couple heart comforter room mockup"), 4: ("Close bed view with optional pillow shams and deer heart artwork.", "Deer couple heart comforter pillow sham panel"), 5: ("Fabric panel with white backing and deer couple pillow mockup.", "Deer couple heart comforter fabric panel"), 6: ("Sizing and details chart for deer couple heart comforter set.", "Deer couple heart comforter size chart"), 7: ("Bedspread features panel showing comforter layers and care icons.", "Deer couple heart comforter features panel"), 8: ("Overhead bedroom mockup of deer couple heart comforter and matching pillows.", "Deer couple heart comforter overhead view")},
    272: {1: ("Bed mockup of blue and white buck and doe comforter with forest border and sample names.", "Buck and doe comforter bed view"), 2: ("Held comforter showing deer silhouettes, You and Me We Got This text and sample names.", "Buck and doe comforter held view"), 3: ("Room mockup of buck and doe comforter with printed craft callout.", "Buck and doe comforter room mockup"), 4: ("Close bed view with Her Buck and His Doe pillow shams plus optional sham label.", "Buck and doe comforter pillow sham panel"), 5: ("Fabric panel with white backing and buck and doe pillow mockup.", "Buck and doe comforter fabric panel"), 6: ("Sizing and details chart for buck and doe comforter set.", "Buck and doe comforter size chart"), 7: ("Bedspread features panel showing comforter layers and care icons.", "Buck and doe comforter features panel"), 8: ("Overhead bedroom mockup of buck and doe comforter and matching pillows.", "Buck and doe comforter overhead view")},
    273: {1: ("Bed mockup of deer heart outline comforter with brown camouflage background and sample names.", "Deer heart outline comforter bed view"), 2: ("Held comforter showing deer silhouettes, heart outline and She keeps me wild text.", "Deer heart outline comforter held view"), 3: ("Room mockup of deer heart outline comforter with printed craft callout.", "Deer heart outline comforter room mockup"), 4: ("Close bed view with optional pillow shams and deer heart outline artwork.", "Deer heart outline comforter pillow sham panel"), 5: ("Fabric panel with white backing and deer heart pillow mockup.", "Deer heart outline comforter fabric panel"), 6: ("Sizing and details chart for deer heart outline comforter set.", "Deer heart outline comforter size chart"), 7: ("Bedspread features panel showing comforter layers and care icons.", "Deer heart outline comforter features panel"), 8: ("Overhead bedroom mockup of deer heart outline comforter and matching pillows.", "Deer heart outline comforter overhead view")},
    274: {1: ("Bed mockup of southwest desert cactus quilt with large orange sun and layered hills.", "Desert cactus sunset quilt bed view"), 2: ("Room mockup of desert cactus sunset quilt with printed craft callout.", "Desert cactus sunset quilt room mockup"), 3: ("Close bed view with optional pillow shams and cactus sunset artwork.", "Desert cactus sunset quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and desert cactus artwork sample.", "Desert cactus sunset quilt fabric panel"), 5: ("Premium quilt sets sizing chart for desert cactus sunset design.", "Desert cactus sunset quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Desert cactus sunset quilt features panel"), 7: ("Overhead bedroom mockup of desert cactus sunset quilt and matching pillows.", "Desert cactus sunset quilt overhead view")},
    275: {1: ("Bed mockup of pink desert cactus floral quilt with sunset bands and bird silhouettes.", "Desert cactus floral quilt bed view"), 2: ("Room mockup of desert cactus floral quilt with printed craft callout.", "Desert cactus floral quilt room mockup"), 3: ("Close bed view with optional pillow shams and flowering cactus artwork.", "Desert cactus floral quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and pink sunset cactus artwork sample.", "Desert cactus floral quilt fabric panel"), 5: ("Premium quilt sets sizing chart for desert cactus floral design.", "Desert cactus floral quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Desert cactus floral quilt features panel"), 7: ("Overhead bedroom mockup of desert cactus floral quilt and matching pillows.", "Desert cactus floral quilt overhead view")},
    276: {1: ("Person holds white floral cross Christian affirmation blanket with I am and sample name Elizabeth.", "Floral cross affirmation blanket held view"), 2: ("Flat blanket view with floral cross, I am text, scripture affirmations and sample name.", "Floral cross affirmation blanket flat view"), 3: ("Floral cross affirmation blanket draped on sofa under a window.", "Floral cross affirmation blanket sofa view"), 4: ("Second sofa mockup of floral cross affirmation blanket beneath round mirror.", "Floral cross affirmation blanket sofa mockup"), 5: ("Feature panel over floral cross blanket with fluffy, quality and no-shedding callouts.", "Floral cross affirmation blanket feature panel")},
    277: {1: ("Person holds God Says I Am floral cross blanket with sample portrait, scripture words and sample name.", "God Says I Am floral cross blanket held view"), 2: ("Feature panel showing white fleece edge with machine washable and stitching callouts.", "God Says I Am floral cross blanket feature panel"), 3: ("Close-up reading scene showing floral cross blanket near an open book.", "God Says I Am floral cross blanket reading close-up"), 4: ("Size and use panel for God Says I Am floral cross blanket.", "God Says I Am floral cross blanket size panel"), 5: ("Detail collage showing printed portrait, floral cross, text and white fleece texture.", "God Says I Am floral cross blanket detail collage"), 6: ("God Says I Am floral cross blanket draped over an armchair.", "God Says I Am floral cross blanket chair mockup"), 7: ("Lifestyle reading scene with God Says I Am floral cross blanket.", "God Says I Am floral cross blanket reading lifestyle"), 8: ("Sofa mockup of God Says I Am floral cross blanket with books and mug.", "God Says I Am floral cross blanket sofa mockup")},
    278: {1: ("Person holds white floral cross blanket with I am text, affirmations and sample name Elizabeth.", "Floral cross personalized text blanket held view"), 2: ("Flat blanket view with floral cross, I am wording and scripture affirmation list.", "Floral cross personalized text blanket flat view"), 3: ("Family reading lifestyle scene with floral cross affirmation blanket.", "Floral cross personalized text blanket family lifestyle"), 4: ("Floral cross personalized text blanket draped on sofa under a window.", "Floral cross personalized text blanket sofa view"), 5: ("Second sofa mockup of floral cross personalized text blanket beneath round mirror.", "Floral cross personalized text blanket sofa mockup"), 6: ("Feature panel over floral cross blanket with fluffy, quality and no-shedding callouts.", "Floral cross personalized text blanket feature panel")},
    279: {1: ("Person holds cream book lover blanket with line-art girl, butterflies, flowers and sample name Jessica.", "Floral line art book lover blanket held view"), 2: ("Flat book lover blanket with Just a Girl Who Loves Books text and line-art reader.", "Floral line art book lover blanket flat view"), 3: ("Blanket size chart with four sizes for floral line-art book lover design.", "Floral line art book lover blanket size chart"), 4: ("Book lover blanket draped on sofa with reader illustration and large bottom text.", "Floral line art book lover blanket sofa view"), 5: ("Second sofa mockup of floral line-art book lover blanket beneath round mirror.", "Floral line art book lover blanket sofa mockup"), 6: ("Feature panel over book lover blanket with fluffy, quality and no-shedding callouts.", "Floral line art book lover blanket feature panel")},
    280: {1: ("Bed mockup of flowering cactus succulent quilt with colorful pots and dark background.", "Flowering cactus succulent quilt bed view"), 2: ("Room mockup of flowering cactus succulent quilt with printed craft callout.", "Flowering cactus succulent quilt room mockup"), 3: ("Close bed view with optional pillow shams and cactus succulent artwork.", "Flowering cactus succulent quilt pillow sham panel"), 4: ("High-quality fabric panel with quilt layers and cactus succulent artwork sample.", "Flowering cactus succulent quilt fabric panel"), 5: ("Premium quilt sets sizing chart for flowering cactus succulent design.", "Flowering cactus succulent quilt size chart"), 6: ("Bedspread features panel showing microfiber layers and care icons.", "Flowering cactus succulent quilt features panel"), 7: ("Overhead bedroom mockup of flowering cactus succulent quilt and matching pillows.", "Flowering cactus succulent quilt overhead view")},
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
            "separated deer couple comforter, cactus quilt, Christian blanket and book lover blanket intents; checked customizer limits and rewrote image observations/alts from contact sheets. Still NEEDS_REVIEW."
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
        ws.cell(row_num, idx["source_refs"]).value = f"products_export_1.csv sha256:{admin_hash}; storefront HTML customizer parse; contact sheets 271-280"
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
        ws.cell(row_num, idx["factual_conflicts"]).value = "No independent QA for batch 028 yet; r2 prepared for QA recheck."
        ws.cell(row_num, idx["confidence_and_reason"]).value = "R2 uses contact sheets, storefront customizer parse and Shopify admin CSV baseline; no approval or import created."

    ws = wb["README_QA"]
    idx = headers(ws)
    for metric, value, definition in [
        ("qa_batch_028_r2_revision", "r2", "Separate 10-product revision without prior independent QA; source workbook was not modified."),
        ("qa_batch_028_r2_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_028_r2_scope", "inventory positions 271-280", "No products outside qa_batch_028 were revised."),
        ("qa_batch_028_r2_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_028_r2_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_028_r2_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
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
        "scope": "qa_batch_028 only; inventory positions 271-280",
        "revision": "r2",
        "revised_products": revised_products,
        "revised_images": revised_images,
        "admin_alt_matches": admin_alt_matches,
        "keyword_rows_revised": keyword_rows,
        "review_status": "NEEDS_REVIEW",
        "approved_created": False,
        "shopify_import_created": False,
        "next_step": "Ready for independent QA of qa_batch_028_r2.",
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text(
        "\n".join([
            "# Revision qa_batch_028_r2",
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
            "- Sửa trọng yếu: rà lại khi chưa có QA, tách nhóm deer couple comforter, cactus quilt, Christian blanket và book lover blanket; parse customizer từ HTML storefront; viết lại title/meta/description/alt theo từng contact sheet.",
            "- Trạng thái: `NEEDS_REVIEW`; không tạo `APPROVED`; không tạo file import Shopify.",
            "- Bước tiếp theo: gửi `qa_batch_028_r2` cho QA độc lập.",
            "",
        ]),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
