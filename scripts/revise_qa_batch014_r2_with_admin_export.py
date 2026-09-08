import csv
import hashlib
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
BATCH_ID = "qa_batch_014_r2"
QA_RUN_ID = "20260907_204400"

SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_014.xlsx"
QA_DATASET = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "qa_dataset.json"
CUSTOMIZER_AUDIT = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "customizer_audit.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_014_r2.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_014_r2.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    131: {
        "title": "Wolf Headdress Quilt Set",
        "meta_title": "Wolf Headdress Quilt Set",
        "meta_description": "Shop a wolf headdress quilt set with close-up yellow-eye portrait artwork, feather accents, matching shams and size panel.",
        "primary": "wolf headdress quilt set",
        "secondary": "wolf head quilt, feather wolf bedding, close up wolf quilt set",
        "cluster": "close-up wolf headdress quilt set",
        "detail": "close-up wolf head artwork with yellow eyes, headdress styling and feather accents on a tan background",
        "intent_role": "Wolf page separated by close-up headdress artwork, tan palette and feather accents.",
        "customizer": "A Custom Your Name field is present but not required; no number field is claimed.",
    },
    132: {
        "title": "Custom Teal Yellow Softball Comforter",
        "meta_title": "Custom Teal Yellow Softball Comforter",
        "meta_description": "Shop a teal yellow softball comforter with purple bedding, OLIVIA sample text, required name and number fields, and pillowcases.",
        "primary": "custom teal yellow softball comforter",
        "secondary": "softball comforter with name, purple yellow softball bedding, custom softball number comforter",
        "cluster": "purple teal yellow softball comforter",
        "detail": "purple comforter with bright yellow softball artwork, teal-yellow splatter, OLIVIA sample text and batter silhouettes",
        "intent_role": "Softball page separated by purple/yellow palette and confirmed required name and number fields.",
        "customizer": "Required Custom Your Name and Custom Your Number fields are present in the customizer audit.",
    },
    133: {
        "title": "Desert Cactus Flower Quilt Set",
        "meta_title": "Desert Cactus Flower Quilt Set",
        "meta_description": "Shop a desert cactus quilt set with potted cactus patchwork, flower blocks, green cream tones, matching shams and size chart.",
        "primary": "desert cactus flower quilt set",
        "secondary": "cactus patchwork quilt, potted cactus bedding, desert flower quilt set",
        "cluster": "potted desert cactus flower patchwork quilt set",
        "detail": "potted cactus and flower blocks in green, cream, tan and terracotta patchwork",
        "intent_role": "Cactus page separated by potted cactus blocks and desert flower patchwork.",
        "customizer": "The visible field is a general Customize Your Quilt text box; no name or number field is claimed.",
    },
    134: {
        "title": "Purple Floral Dragonfly Quilt Set",
        "meta_title": "Purple Floral Dragonfly Quilt Set",
        "meta_description": "Shop a purple floral dragonfly quilt set with stained-glass style dragonflies, oval frame, matching shams and size chart.",
        "primary": "purple floral dragonfly quilt set",
        "secondary": "dragonfly stained glass quilt, purple dragonfly bedding, floral dragonfly quilt",
        "cluster": "purple floral stained-glass dragonfly quilt set",
        "detail": "two dragonflies over purple flowers inside an oval stained-glass style frame",
        "intent_role": "Dragonfly page separated by purple flowers, paired dragonflies and oval stained-glass frame.",
        "customizer": "The visible field is a general Customize Your Quilt text box; no name or number field is claimed.",
    },
    135: {
        "title": "Sunburst Dragonfly Pond Quilt Set",
        "meta_title": "Sunburst Dragonfly Pond Quilt Set",
        "meta_description": "Shop a dragonfly pond quilt set with yellow sunburst, water lily artwork, green reeds, matching shams and size chart.",
        "primary": "sunburst dragonfly pond quilt set",
        "secondary": "dragonfly water lily quilt, yellow dragonfly bedding, pond dragonfly quilt",
        "cluster": "yellow sunburst dragonfly water lily quilt set",
        "detail": "dragonfly over yellow sunburst and water lily pond artwork with green reeds",
        "intent_role": "Dragonfly page separated by yellow sunburst, water lily and pond reeds.",
        "customizer": "The visible field is a general Customize Your Quilt text box; no name or number field is claimed.",
    },
    136: {
        "title": "Stained Glass Dragonfly Pond Quilt",
        "meta_title": "Stained Glass Dragonfly Pond Quilt",
        "meta_description": "Shop a stained-glass dragonfly pond quilt with blue round frame, yellow lily artwork, green reeds and matching shams.",
        "primary": "stained glass dragonfly pond quilt",
        "secondary": "blue dragonfly quilt set, lily pond dragonfly bedding, dragonfly patchwork quilt",
        "cluster": "blue framed dragonfly lily pond quilt set",
        "detail": "large dragonfly and yellow lily inside a blue circular stained-glass style pond frame",
        "intent_role": "Dragonfly page separated by blue circular frame, close dragonfly wing and yellow lily.",
        "customizer": "The visible field is a general Customize Your Quilt text box; no name or number field is claimed.",
    },
    137: {
        "title": "Purple Vintage Dragonfly Quilt Set",
        "meta_title": "Purple Vintage Dragonfly Quilt Set",
        "meta_description": "Shop a purple vintage dragonfly quilt set with bright floral panels, stained-glass look, matching shams and size chart.",
        "primary": "purple vintage dragonfly quilt set",
        "secondary": "purple dragonfly quilt, vintage dragonfly bedding, floral dragonfly quilt set",
        "cluster": "purple vintage floral dragonfly quilt set",
        "detail": "large green dragonfly with purple border, multicolor stained-glass panels and floral artwork",
        "intent_role": "Dragonfly page separated by purple border, large centered dragonfly and bright floral panels.",
        "customizer": "The visible field is a general Customize Your Quilt text box; no name or number field is claimed.",
    },
    138: {
        "title": "Colorful Elephant Patchwork Quilt",
        "meta_title": "Colorful Elephant Patchwork Quilt",
        "meta_description": "Shop a colorful elephant patchwork quilt with light blue background, multicolor fabric blocks, matching shams and size chart.",
        "primary": "colorful elephant patchwork quilt",
        "secondary": "elephant patchwork bedding, light blue elephant quilt, animal patchwork quilt set",
        "cluster": "light blue colorful elephant patchwork quilt set",
        "detail": "large elephant made from multicolor patchwork blocks on a light blue quilt background",
        "intent_role": "Elephant page separated by light blue base and multicolor elephant patchwork.",
        "customizer": "The visible field is a general Customize Your Quilt text box; no name or number field is claimed.",
    },
    139: {
        "title": "Navy Yellow Elephant Quilt Set",
        "meta_title": "Navy Yellow Elephant Quilt Set",
        "meta_description": "Shop a navy yellow elephant quilt set with floral patchwork elephant artwork, dark blue background, matching shams and size chart.",
        "primary": "navy yellow elephant quilt set",
        "secondary": "yellow elephant patchwork quilt, navy elephant bedding, floral elephant quilt set",
        "cluster": "navy yellow floral elephant patchwork quilt set",
        "detail": "yellow and cream patchwork elephant on a dark navy background with small floral accents",
        "intent_role": "Elephant page separated by navy base, yellow patchwork elephant and floral accents.",
        "customizer": "The visible field is a general Customize Your Quilt text box; no name or number field is claimed.",
    },
    140: {
        "title": "Rainbow Fantasy Dragon Quilt Set",
        "meta_title": "Rainbow Fantasy Dragon Quilt Set",
        "meta_description": "Shop a rainbow fantasy dragon quilt set with blue-orange dragon scales, close-up head artwork, matching shams and size chart.",
        "primary": "rainbow fantasy dragon quilt set",
        "secondary": "dragon scale quilt, fantasy dragon bedding, colorful dragon quilt set",
        "cluster": "rainbow blue orange fantasy dragon quilt set",
        "detail": "large fantasy dragon head with blue, orange, purple and gold scale artwork",
        "intent_role": "Dragon page separated by rainbow scale palette and close-up fantasy dragon head.",
        "customizer": "A Custom Your Name field is present but not required; no number field is claimed.",
    },
}


COMMON_QUILT_INCLUDED = ("Three-piece bedding set panel listing one premium quilt and two optional standard shams.", "Three piece quilt set information panel")
COMMON_SIZE = ("Premium quilt set size chart showing available bed sizes and measurements.", "Quilt set size chart")
COMMON_BEDSPREAD = ("Bedspread features panel showing quilt layers and anti-wrinkle, anti-static and machine-washable icons.", "Quilt construction feature panel")


def quilt7(label, alt_root):
    return {
        1: (f"Bedroom mockup of {label} with matching shams.", alt_root),
        2: (f"Room mockup of {label} with premium printed craft callout and fabric inset.", f"{alt_root} room mockup"),
        3: (f"Optional pillow shams panel showing {label}.", f"{alt_root} pillow shams"),
        4: (f"High-quality fabric panel showing {label} with quilt detail and pillow mockup.", f"{alt_root} fabric detail panel"),
        5: COMMON_SIZE,
        6: COMMON_BEDSPREAD,
        7: (f"Second bedroom mockup of {label} with matching shams.", f"{alt_root} bedroom mockup"),
    }


IMAGE_DETAILS = {
    131: {
        1: ("Bedroom mockup of wolf headdress quilt with yellow eyes, feather accents and matching shams.", "Wolf headdress quilt set"),
        2: ("Close-up view of wolf face, yellow eyes, headdress pattern and quilt stitching.", "Wolf headdress quilt close-up"),
        3: ("Pillow sham mockup showing wolf head with headdress and feather accents.", "Wolf headdress pillow sham"),
        4: ("Room mockup of wolf headdress quilt set on a modern bed.", "Wolf headdress bedding set"),
        5: COMMON_QUILT_INCLUDED,
    },
    132: {
        1: ("Bedroom mockup of purple and bright yellow softball comforter with OLIVIA sample text and matching pillowcases.", "Purple yellow softball comforter with OLIVIA"),
        2: ("Angled bedroom mockup of softball comforter with yellow ball graphic, purple base and OLIVIA sample text.", "Teal yellow softball comforter on bed"),
        3: ("Pillowcase close-up showing yellow batter silhouettes and softball stitch graphics on purple background.", "Softball pillowcases with batter silhouettes"),
        4: ("Folded comforter mockup showing yellow softball artwork and purple splatter design.", "Folded purple yellow softball comforter"),
        5: ("Close-up corner of softball comforter showing white underside and yellow ball stitching.", "Softball comforter fabric close-up"),
        6: ("Feature panel listing machine washable, soft touch, lightweight and breathable comforter qualities.", "Softball comforter feature panel"),
        7: ("Included items panel showing one comforter and two pillowcases.", "Softball comforter set included items"),
    },
    133: quilt7("potted cactus and flower patchwork quilt in green, cream and terracotta tones", "Desert cactus flower quilt set"),
    134: quilt7("purple floral dragonfly quilt with oval stained-glass style artwork", "Purple floral dragonfly quilt set"),
    135: quilt7("yellow sunburst dragonfly pond quilt with water lily and green reeds", "Sunburst dragonfly pond quilt set"),
    136: quilt7("blue framed dragonfly lily pond quilt with stained-glass style panels", "Stained glass dragonfly pond quilt"),
    137: {
        1: ("Bedroom mockup of purple vintage dragonfly quilt with floral stained-glass style artwork and matching shams.", "Purple vintage dragonfly quilt set"),
        2: ("Room mockup of purple dragonfly quilt with large green dragonfly and floral panels.", "Purple dragonfly quilt room mockup"),
        3: ("Optional pillow shams panel showing purple dragonfly artwork.", "Purple dragonfly pillow shams"),
        4: COMMON_BEDSPREAD,
        5: ("High-quality fabric panel showing purple dragonfly quilt detail and pillow mockup.", "Purple dragonfly quilt fabric panel"),
        6: COMMON_SIZE,
        7: ("Second bedroom mockup of purple vintage dragonfly quilt with matching shams.", "Purple vintage dragonfly bedroom mockup"),
    },
    138: quilt7("light blue elephant patchwork quilt with multicolor elephant artwork", "Colorful elephant patchwork quilt"),
    139: quilt7("navy yellow elephant quilt with floral patchwork elephant artwork", "Navy yellow elephant quilt set"),
    140: {
        1: ("Bedroom mockup of rainbow fantasy dragon quilt with blue-orange dragon head and matching shams.", "Rainbow fantasy dragon quilt set"),
        2: ("Overhead bedroom mockup of colorful dragon quilt with blue, orange and purple scale artwork.", "Colorful dragon quilt bedroom mockup"),
        3: ("Room mockup of rainbow dragon quilt with premium printed craft callout and fabric inset.", "Rainbow dragon quilt room mockup"),
        4: ("Optional pillow shams panel showing dragon head and scale artwork.", "Rainbow dragon pillow shams"),
        5: ("High-quality fabric panel showing colorful dragon scale quilt detail and pillow mockup.", "Dragon scale quilt fabric detail panel"),
        6: COMMON_BEDSPREAD,
        7: COMMON_SIZE,
    },
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


def load_qa():
    data = json.loads(QA_DATASET.read_text(encoding="utf-8"))
    products = data["QA_Products"]
    customizer = json.loads(CUSTOMIZER_AUDIT.read_text(encoding="utf-8"))
    return {
        "product_keys": [item["product_key"] for item in products],
        "handle_by_key": {item["product_key"]: item["handle"] for item in products},
        "pos_by_key": {item["product_key"]: int(item["inventory_position"]) for item in products},
        "customizer_by_pos": {int(item["inventory_position"]): item for item in customizer},
    }


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
            "Body (HTML)": first.get("Body (HTML)", ""),
            "Type": first.get("Type", ""),
            "SEO Title": first.get("SEO Title", ""),
            "SEO Description": first.get("SEO Description", ""),
            "Option1 Name": first.get("Option1 Name", ""),
            "Option2 Name": first.get("Option2 Name", ""),
            "Option3 Name": first.get("Option3 Name", ""),
            "Status": first.get("Status", ""),
            "images": images,
        }
    return admin


def option_text(admin_row):
    values = [admin_row["Option1 Name"], admin_row["Option2 Name"], admin_row["Option3 Name"]]
    values = [v for v in values if v]
    return ", ".join(values) if values else "Shopify export shows no named option group."


def description(pos, admin_row):
    item = PRODUCT_UPDATES[pos]
    return (
        f"<p>{item['title']} features {item['detail']}. "
        "This product page is written around the specific visible artwork rather than a broad template.</p>"
        "<h3>Design Details</h3>"
        f"<ul><li>Artwork: {item['detail']}.</li>"
        "<li>Gallery images show the main bed mockup plus detail, feature, care, size or included-item panels where present.</li></ul>"
        "<h3>Options and Fit</h3>"
        f"<ul><li>Available option groups: {option_text(admin_row)}.</li>"
        f"<li>{item['customizer']}</li>"
        "<li>Choose the product type and size shown on the product page before checkout.</li></ul>"
    )


def main():
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE, OUTPUT)
    qa = load_qa()
    admin = load_admin(set(qa["handle_by_key"].values()))
    admin_hash = sha256(ADMIN_EXPORT)
    now = datetime.now().astimezone().isoformat()

    wb = load_workbook(OUTPUT)
    product_key_set = set(qa["product_keys"])

    ws = wb["SEO_Products"]
    idx = headers(ws)
    revised_products = 0
    for row_num in range(2, ws.max_row + 1):
        pk = text(ws.cell(row_num, idx["product_key"]).value)
        if pk not in product_key_set:
            continue
        pos = qa["pos_by_key"][pk]
        handle = qa["handle_by_key"][pk]
        item = PRODUCT_UPDATES[pos]
        admin_row = admin[handle]
        ws.cell(row_num, idx["title_current"]).value = admin_row["Title"]
        ws.cell(row_num, idx["h1_current"]).value = admin_row["Title"]
        ws.cell(row_num, idx["rendered_title_current"]).value = admin_row["SEO Title"] or admin_row["Title"]
        ws.cell(row_num, idx["meta_description_current"]).value = admin_row["SEO Description"]
        ws.cell(row_num, idx["product_type"]).value = admin_row["Type"] or "Bedding"
        ws.cell(row_num, idx["title_proposed"]).value = item["title"]
        ws.cell(row_num, idx["meta_title_seo"]).value = item["meta_title"]
        ws.cell(row_num, idx["meta_title_length"]).value = len(item["meta_title"])
        ws.cell(row_num, idx["meta_description_seo"]).value = item["meta_description"]
        ws.cell(row_num, idx["meta_description_length"]).value = len(item["meta_description"])
        ws.cell(row_num, idx["description_proposed_html"]).value = description(pos, admin_row)
        ws.cell(row_num, idx["primary_keyword"]).value = item["primary"]
        ws.cell(row_num, idx["secondary_keywords"]).value = item["secondary"]
        ws.cell(row_num, idx["meta_keyword"]).value = item["primary"]
        ws.cell(row_num, idx["keyword_strategy"]).value = item["intent_role"]
        ws.cell(row_num, idx["buyer_search_summary"]).value = (
            f"US English buyer intent targets {item['cluster']}; no volume or ranking claim is made."
        )
        ws.cell(row_num, idx["keyword_evidence_level"]).value = "SERP_ONLY"
        ws.cell(row_num, idx["revision"]).value = "r2"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields"]:
            ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            f"R2 after Sang QA batch 014: used products_export_1.csv sha256:{admin_hash}; "
            "removed internal workflow copy, tied customization wording only to customizer evidence, "
            "split dragonfly/elephant/quilt intents by visible motif, and rewrote image observations/alts from inspected contact sheets. "
            "Still NEEDS_REVIEW."
        )
        revised_products += 1

    ws = wb["Image_Audit"]
    idx = headers(ws)
    revised_images = 0
    admin_alt_matches = 0
    handle_to_pk = {v: k for k, v in qa["handle_by_key"].items()}
    for row_num in range(2, ws.max_row + 1):
        handle = text(ws.cell(row_num, idx["Handle"]).value)
        pk = handle_to_pk.get(handle)
        if pk not in product_key_set:
            continue
        pos = qa["pos_by_key"][pk]
        image_number = int(ws.cell(row_num, idx["image_number"]).value)
        image_url = text(ws.cell(row_num, idx["image_url_export"]).value or ws.cell(row_num, idx["image_url"]).value)
        admin_image = admin.get(handle, {}).get("images", {}).get(normalize_url(image_url))
        if admin_image:
            ws.cell(row_num, idx["image_url_export"]).value = admin_image["src"]
            ws.cell(row_num, idx["alt_current"]).value = admin_image["alt"]
            admin_alt_matches += 1
        observation, alt = IMAGE_DETAILS[pos][image_number]
        ws.cell(row_num, idx["observed_visual_details"]).value = observation
        ws.cell(row_num, idx["alt_proposed"]).value = alt
        ws.cell(row_num, idx["alt_action"]).value = "SET"
        ws.cell(row_num, idx["viewed_status"]).value = "VIEWED"
        ws.cell(row_num, idx["revision"]).value = "r2"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields", "approved_revision"]:
            if col_name in idx:
                ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            "R2: observation and alt rewritten from inspected contact sheet; admin image URL and current alt matched from Shopify CSV where available."
        )
        revised_images += 1

    ws = wb["Keyword_Map"]
    idx = headers(ws)
    keyword_counts = defaultdict(int)
    keyword_rows = 0
    for row_num in range(2, ws.max_row + 1):
        pk = text(ws.cell(row_num, idx["product_key"]).value)
        if pk not in product_key_set:
            continue
        pos = qa["pos_by_key"][pk]
        item = PRODUCT_UPDATES[pos]
        role = text(ws.cell(row_num, idx["keyword_role"]).value)
        secondaries = [part.strip() for part in item["secondary"].split(",")]
        keyword = item["primary"] if role == "PRIMARY" else secondaries[min(keyword_counts[pk], len(secondaries) - 1)]
        if role != "PRIMARY":
            keyword_counts[pk] += 1
        ws.cell(row_num, idx["keyword"]).value = keyword
        ws.cell(row_num, idx["semantic_cluster"]).value = item["cluster"]
        ws.cell(row_num, idx["intent"]).value = "Commercial product intent"
        ws.cell(row_num, idx["target_page_type"]).value = "PRODUCT"
        ws.cell(row_num, idx["decision_reason"]).value = item["intent_role"]
        ws.cell(row_num, idx["demand_evidence"]).value = "SERP intent evidence only; no search-volume claim."
        ws.cell(row_num, idx["validation_source"]).value = (
            "Sang QA batch 014 + products_export_1.csv admin baseline + customizer_audit.json + contact-sheet inspection"
        )
        ws.cell(row_num, idx["checked_at"]).value = now
        ws.cell(row_num, idx["possible_overlap_with_other_products"]).value = item["intent_role"]
        ws.cell(row_num, idx["mapping_reason"]).value = (
            "R2 separates nearby pages by visible motif, wording, color palette and proof level for customization."
        )
        ws.cell(row_num, idx["mapping_status"]).value = "CANDIDATE_MAPPED_NEEDS_RECHECK"
        ws.cell(row_num, idx["mapping_version"]).value = "r2"
        keyword_rows += 1

    ws = wb["Buyer_Search_Research"]
    idx = headers(ws)
    for row_num in range(2, ws.max_row + 1):
        pk = text(ws.cell(row_num, idx["product_key"]).value)
        if pk not in product_key_set:
            continue
        pos = qa["pos_by_key"][pk]
        item = PRODUCT_UPDATES[pos]
        ws.cell(row_num, idx["purchase_context"]).value = f"US buyer looking for {item['cluster']} as a specific product page."
        ws.cell(row_num, idx["jtbd_statement"]).value = (
            f"When shopping for {item['cluster']}, the buyer wants the page to confirm exact artwork, product type, size choices, "
            "and only customization options that are actually supported."
        )
        ws.cell(row_num, idx["functional_motivation"]).value = (
            "Confirm product type, visible motif, size/care panels, supported custom fields and image accuracy before purchase."
        )
        ws.cell(row_num, idx["emotional_social_motivation"]).value = (
            "Choose a wolf, softball, cactus, dragonfly, elephant or dragon bedding item without confusing it with another similar design."
        )
        ws.cell(row_num, idx["purchase_concerns"]).value = (
            "Avoid unsupported name, number, photo, upload, cultural-origin or material claims; confirm exact size and product type."
        )
        ws.cell(row_num, idx["source_refs"]).value = (
            f"products_export_1.csv sha256:{admin_hash}; Sang QA {QA_RUN_ID}; customizer_audit.json; contact sheets 131-140"
        )
        ws.cell(row_num, idx["observed_at"]).value = now
        ws.cell(row_num, idx["research_status"]).value = "REVISED_NEEDS_RECHECK"
        ws.cell(row_num, idx["limitations"]).value = (
            "No Search Console, paid keyword volume or internal site-search export supplied; proof remains SERP-only."
        )
        ws.cell(row_num, idx["seo_application"]).value = f"Use {item['primary']} with intent role: {item['intent_role']}"

    ws = wb["Product_Evidence"]
    idx = headers(ws)
    pk_by_pos = {qa["pos_by_key"][pk]: pk for pk in qa["product_keys"]}
    for row_num in range(2, ws.max_row + 1):
        evidence_id = text(ws.cell(row_num, idx["evidence_id"]).value)
        match = re.search(r"_(\d{3})$", evidence_id)
        if not match:
            continue
        pos = int(match.group(1))
        if pos not in PRODUCT_UPDATES:
            continue
        pk = pk_by_pos[pos]
        handle = qa["handle_by_key"][pk]
        admin_row = admin[handle]
        item = PRODUCT_UPDATES[pos]
        ws.cell(row_num, idx["sources_accessed"]).value = (
            f"Original storefront/product evidence; Sang QA {QA_RUN_ID}; products_export_1.csv sha256:{admin_hash}; "
            "customizer_audit.json; inspected contact sheets"
        )
        ws.cell(row_num, idx["current_H1"]).value = admin_row["Title"]
        ws.cell(row_num, idx["current_meta_title"]).value = admin_row["SEO Title"] or "EMPTY"
        ws.cell(row_num, idx["verified_product_facts"]).value = (
            f"Admin export matched handle {handle}; type={admin_row['Type']}; options={option_text(admin_row)}; "
            f"status={admin_row['Status']}; image_count={len(admin_row['images'])}; design={item['detail']}; "
            f"customizer_note={item['customizer']}"
        )
        ws.cell(row_num, idx["factual_conflicts"]).value = "Sang QA issues addressed in r2; independent QA recheck still required."
        ws.cell(row_num, idx["confidence_and_reason"]).value = (
            "R2 uses inspected contact sheets, customizer audit and Shopify admin CSV baseline; no approval or import created."
        )

    ws = wb["README_QA"]
    idx = headers(ws)
    for metric, value, definition in [
        ("qa_batch_014_r2_revision", "r2", "Separate 10-product revision after Sang QA; source workbook was not modified."),
        ("qa_batch_014_r2_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_014_r2_scope", "inventory positions 131-140", "No products outside qa_batch_014 were revised."),
        ("qa_batch_014_r2_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_014_r2_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_014_r2_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
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
        "source_qa": str(QA_DATASET.relative_to(ROOT)),
        "admin_export": str(ADMIN_EXPORT.relative_to(ROOT)),
        "admin_export_sha256": admin_hash,
        "revision_workbook": str(OUTPUT.relative_to(ROOT)),
        "scope": "qa_batch_014 only; inventory positions 131-140",
        "revision": "r2",
        "revised_products": revised_products,
        "revised_images": revised_images,
        "admin_alt_matches": admin_alt_matches,
        "keyword_rows_revised": keyword_rows,
        "review_status": "NEEDS_REVIEW",
        "approved_created": False,
        "shopify_import_created": False,
        "next_step": "Hold for combined QA handoff after revisions through qa_batch_020, per user request.",
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text(
        "\n".join(
            [
                "# Revision qa_batch_014_r2",
                "",
                "Artifact chính thức theo quy trình từng lô 10 sản phẩm.",
                "",
                f"- Nguồn r1: `{SOURCE.relative_to(ROOT)}`",
                f"- QA của Sang: `{QA_DATASET.relative_to(ROOT)}`",
                f"- Shopify admin export: `{ADMIN_EXPORT.relative_to(ROOT)}`",
                f"- Admin export SHA-256: `{admin_hash}`",
                f"- Workbook r2: `{OUTPUT.relative_to(ROOT)}`",
                f"- Sản phẩm sửa: {revised_products}",
                f"- Ảnh sửa/refresh: {revised_images}",
                f"- Admin image alt match: {admin_alt_matches}/{revised_images}",
                "- Sửa trọng yếu: bỏ nội dung nội bộ, chỉ dùng claim name/number khi customizer audit xác nhận, tránh wording văn hóa không có bằng chứng, tách intent cho wolf/softball/cactus/dragonfly/elephant/dragon và viết lại mô tả/alt theo contact sheets.",
                "- Trạng thái: `NEEDS_REVIEW`; không tạo `APPROVED`; không tạo file import Shopify.",
                "- Bước tiếp theo: giữ lại để gửi Sang QA chung sau khi revision đến `qa_batch_020`, theo yêu cầu người dùng.",
                "",
            ]
        ),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
