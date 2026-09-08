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
BATCH_ID = "qa_batch_017_r3"
QA_RUN_ID = "20260907_212312"

SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / "qa_batch_017_r2" / "SEO_Product_Optimization_qa_batch_017_r2.xlsx"
QA_DATASET = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "qa_dataset.json"
CUSTOMIZER_AUDIT = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "customizer_audit.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_017_r3.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_017_r3.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    161: {
        "title": "Blue Raven Celtic Knot Quilt",
        "meta_title": "Blue Raven Celtic Knot Quilt",
        "meta_description": "Shop a blue Norse raven quilt with Celtic knot frame, moon patchwork panels, matching shams and size chart.",
        "primary": "blue raven Celtic knot quilt",
        "secondary": "Norse raven Celtic quilt, Viking raven quilt set, Celtic knot raven bedding",
        "cluster": "blue Norse raven Celtic knot quilt",
        "detail": "blue and black raven artwork inside a Celtic knot frame with moon and patchwork panels",
        "intent_role": "Norse raven page separated by blue palette, Celtic knot frame and raven-with-moon artwork.",
        "customizer": "The visible field is a general Customize Your Item text box; no name, number or upload field is claimed.",
    },
    162: {
        "title": "Valhalla Viking Shield Quilt",
        "meta_title": "Valhalla Viking Shield Quilt",
        "meta_description": "Shop a black Valhalla Viking shield quilt with crossed axes, runic artwork, matching shams and size chart.",
        "primary": "Valhalla Viking shield quilt",
        "secondary": "Viking shield quilt, crossed axes bedding, Norse Valhalla quilt",
        "cluster": "black Valhalla Viking shield crossed axes quilt",
        "detail": "Viking shield and crossed axes artwork with Valhalla text on a black runic background",
        "intent_role": "Viking page separated by shield, crossed axes, Valhalla wording and black runic background.",
        "customizer": "The visible field is a general Customize Your Item text box; no name, number or upload field is claimed.",
    },
    163: {
        "title": "Blue Gold Owl Night Quilt Set",
        "meta_title": "Blue Gold Owl Night Quilt Set",
        "meta_description": "Shop a blue gold owl night quilt set with moon frame artwork, starry patchwork border, shams and size chart.",
        "primary": "blue gold owl night quilt set",
        "secondary": "owl patchwork quilt, owl moon bedding, animal quilt set",
        "cluster": "blue gold owl moon patchwork quilt set",
        "detail": "owl perched on a branch inside a moon-style round frame with navy and gold patchwork details",
        "intent_role": "Owl page separated by blue-gold palette, moon frame, branch perch and starry patchwork border.",
        "customizer": "The visible field is a general Customize Your Quilt text box; no name, number or upload field is claimed.",
    },
    164: {
        "title": "Witch Moon Halloween Comforter Set",
        "meta_title": "Witch Moon Halloween Comforter Set",
        "meta_description": "Shop a witch moon Halloween comforter set with black white artwork, bats, pumpkins, sheets and size chart.",
        "primary": "witch moon Halloween comforter set",
        "secondary": "black white Halloween bedding, Halloween comforter set with sheets, twin Halloween comforter set",
        "cluster": "black white witch moon Halloween comforter set",
        "detail": "black and white Halloween design with flying witch, moon, bats, pumpkins and HALLOWEEN text",
        "intent_role": "Halloween comforter page separated by witch-and-moon artwork and black-white palette.",
        "customizer": "No verified personalization, name, number or upload field is claimed for this product.",
    },
    165: {
        "title": "Ghost Pumpkin Halloween Comforter Set",
        "meta_title": "Ghost Pumpkin Halloween Comforter Set",
        "meta_description": "Shop a ghost pumpkin Halloween comforter set with dark blue pattern, spiderwebs, bats, sheets and size chart.",
        "primary": "ghost pumpkin Halloween comforter set",
        "secondary": "dark blue Halloween bedding, pumpkin ghost comforter, twin Halloween comforter set",
        "cluster": "dark blue ghost pumpkin Halloween comforter set",
        "detail": "dark blue and gray Halloween pattern with white ghosts, orange pumpkins, spiderwebs, bats and stars",
        "intent_role": "Halloween comforter page separated by dark blue palette, ghost-pumpkin motif and spiderweb details.",
        "customizer": "No verified personalization, name, number or upload field is claimed for this product.",
    },
    166: {
        "title": "Black Ghost Halloween Comforter Set",
        "meta_title": "Black Ghost Halloween Comforter Set",
        "meta_description": "Shop a black ghost Halloween comforter set with white ghost pattern, matching pieces, fitted sheet panel and size chart.",
        "primary": "black ghost Halloween comforter set",
        "secondary": "white ghost bedding set, spooky ghost comforter, twin Halloween comforter set",
        "cluster": "black white ghost Halloween comforter set",
        "detail": "black Halloween comforter design covered with simple white ghost faces and matching pillow pattern",
        "intent_role": "Halloween comforter page separated by black base and repeated white ghost motif.",
        "customizer": "No verified personalization, name, number or upload field is claimed for this product.",
    },
    167: {
        "title": "Haunted Pumpkin Ghost Comforter Set",
        "meta_title": "Haunted Pumpkin Ghost Comforter Set",
        "meta_description": "Shop a haunted pumpkin ghost comforter set with moon, graveyard silhouettes, sheets and Halloween size chart.",
        "primary": "haunted pumpkin ghost comforter set",
        "secondary": "blue Halloween bedding set, pumpkin ghost bedding, twin Halloween comforter set",
        "cluster": "blue haunted pumpkin ghost Halloween comforter set",
        "detail": "blue Halloween scene with ghosts, jack-o-lantern pumpkins, yellow moon, bare trees and graveyard silhouettes",
        "intent_role": "Halloween comforter page separated by blue scene, moon, pumpkins, ghosts and graveyard artwork.",
        "customizer": "No verified personalization, name, number or upload field is claimed for this product.",
    },
    168: {
        "title": "Pink Cute Ghost Halloween Comforter Set",
        "meta_title": "Pink Cute Ghost Halloween Comforter Set",
        "meta_description": "Shop a pink cute ghost Halloween comforter set with witch hats, candy motifs, sheets and size chart.",
        "primary": "pink cute ghost Halloween comforter set",
        "secondary": "pink Halloween bedding, cute ghost bedding set, twin Halloween comforter set",
        "cluster": "pink cute ghost witch hat Halloween comforter set",
        "detail": "pink Halloween pattern with cute white ghosts, witch hats, candy, stars and small seasonal icons",
        "intent_role": "Halloween comforter page separated by pink palette and cute ghost-and-witch-hat pattern.",
        "customizer": "No verified personalization, name, number or upload field is claimed for this product.",
    },
    169: {
        "title": "Red Handprint Halloween Comforter Set",
        "meta_title": "Red Handprint Halloween Comforter Set",
        "meta_description": "Shop a red handprint Halloween comforter set with splatter-style artwork, sheets, fitted sheet panel and size chart.",
        "primary": "red handprint Halloween comforter set",
        "secondary": "red splatter bedding set, scary Halloween comforter, twin Halloween comforter set",
        "cluster": "red handprint splatter Halloween comforter set",
        "detail": "white comforter design with red splatter-style marks and red handprints plus matching red sheet pieces",
        "intent_role": "Halloween comforter page separated by red handprint and splatter-style artwork.",
        "customizer": "No verified personalization, name, number or upload field is claimed for this product.",
    },
    170: {
        "title": "Haunted House Pumpkin Comforter Set",
        "meta_title": "Haunted House Pumpkin Comforter Set",
        "meta_description": "Shop a haunted house pumpkin comforter set with black cat, spiderwebs, autumn leaves, sheets and size chart.",
        "primary": "haunted house pumpkin comforter set",
        "secondary": "white Halloween bedding set, pumpkin cat comforter, twin Halloween comforter set",
        "cluster": "white haunted house pumpkin cat Halloween comforter set",
        "detail": "white Halloween pattern with haunted houses, pumpkins, black cats, spiderwebs, bare trees and autumn leaves",
        "intent_role": "Halloween comforter page separated by white base, haunted house, pumpkin and black cat pattern.",
        "customizer": "No verified personalization, name, number or upload field is claimed for this product.",
    },
}


COMMON_QUILT_SIZE = ("Premium quilt set size chart showing available bed sizes and measurements.", "Quilt set size chart")
COMMON_QUILT_BEDSPREAD = ("Bedspread features panel showing quilt layers and comfort feature icons.", "Quilt construction feature panel")
COMMON_COMFORTER_MATTRESS = ("Fitted sheet panel showing standard and deep mattress fit up to 14 inches.", "Fitted sheet deep mattress panel")


IMAGE_DETAILS = {
    161: {
        1: ("Bedroom mockup of blue raven quilt with Celtic knot frame, moon panel and matching shams.", "Blue raven Celtic knot quilt"),
        2: ("Feature mockup of Norse raven quilt with cozy, breathable, lightweight and washable icons.", "Blue raven quilt feature panel"),
        3: ("All seasons panel showing blue raven Celtic knot quilt details and magnified texture circles.", "Norse raven Celtic knot detail panel"),
        4: ("Precision stitching panel showing raven, moon and Celtic knot close-ups.", "Blue raven quilt stitching panel"),
        5: ("Layer panel showing microfiber woven top, filling microfiber and back microfiber.", "Raven quilt microfiber layer panel"),
        6: COMMON_QUILT_SIZE,
        7: ("Fabric detail panel showing blue raven Celtic knot print and quilt surface close-ups.", "Blue raven Celtic quilt fabric panel"),
    },
    162: {
        1: ("Bedroom mockup of black Valhalla Viking shield quilt with crossed axes and matching shams.", "Valhalla Viking shield quilt"),
        2: ("Feature mockup of Viking shield quilt with cozy, breathable, lightweight and washable icons.", "Viking shield quilt feature panel"),
        3: ("All seasons panel showing Valhalla shield, crossed axes and runic background details.", "Valhalla shield quilt detail panel"),
        4: ("Precision stitching panel showing Viking shield, crossed axes and black fabric close-ups.", "Viking shield quilt stitching panel"),
        5: ("Layer panel showing microfiber woven top, filling microfiber and back microfiber.", "Viking shield quilt microfiber layer panel"),
        6: COMMON_QUILT_SIZE,
        7: ("Fabric detail panel showing black Valhalla shield quilt print and quilt surface close-ups.", "Valhalla Viking quilt fabric panel"),
    },
    163: {
        1: ("Bedroom mockup of blue gold owl night quilt set with matching shams.", "Blue gold owl night quilt set"),
        2: ("Feature mockup of owl moon quilt with cozy, breathable, lightweight and washable icons.", "Owl moon quilt feature panel"),
        3: ("All seasons panel showing owl, moon frame and starry blue patchwork details.", "Owl night quilt detail panel"),
        4: ("Precision stitching panel showing owl, branch and blue gold patchwork close-ups.", "Owl patchwork quilt stitching panel"),
        5: ("Layer panel showing microfiber woven top, filling microfiber and back microfiber.", "Owl quilt microfiber layer panel"),
        6: COMMON_QUILT_SIZE,
        7: ("Fabric detail panel showing owl night print and quilt surface close-ups.", "Blue gold owl quilt fabric panel"),
    },
    164: {
        1: ("Bedroom mockup of black white witch moon Halloween comforter set with bats and pumpkins.", "Witch moon Halloween comforter set"),
        2: ("Close-up of black comforter folded over white witch moon Halloween print.", "Witch Halloween comforter close-up"),
        3: ("Warm and cozy room mockup of witch moon Halloween bedding with pumpkins.", "Witch moon Halloween bedding room mockup"),
        4: ("Front bedroom mockup showing witch, moon, bats, pumpkins and HALLOWEEN text.", "Black white witch Halloween bedding"),
        5: COMMON_COMFORTER_MATTRESS,
        6: ("Package include panel with witch moon Halloween bedding and size table.", "Witch Halloween comforter package panel"),
    },
    165: {
        1: ("Bedroom mockup of dark blue ghost pumpkin Halloween comforter set with matching pillows.", "Ghost pumpkin Halloween comforter set"),
        2: ("Close-up of dark blue ghost pumpkin bedding pattern with spiderwebs and bats.", "Ghost pumpkin comforter close-up"),
        3: ("Warm and cozy room mockup of dark blue Halloween bedding with ghosts and pumpkins.", "Dark blue Halloween bedding room mockup"),
        4: COMMON_COMFORTER_MATTRESS,
        5: ("Package include panel with ghost pumpkin Halloween bedding and size table.", "Ghost pumpkin comforter package panel"),
    },
    166: {
        1: ("Bedroom mockup of black Halloween comforter set covered with white ghost faces.", "Black ghost Halloween comforter set"),
        2: ("Close-up of black comforter folded over repeated white ghost pattern.", "Black ghost comforter close-up"),
        3: ("Warm and cozy room mockup of black white ghost Halloween bedding.", "Black white ghost Halloween bedding"),
        4: ("Front bedroom mockup showing black ghost comforter and matching pillow pattern.", "White ghost bedding set front view"),
        5: COMMON_COMFORTER_MATTRESS,
        6: ("Package include panel with black ghost Halloween bedding and size table.", "Black ghost comforter package panel"),
    },
    167: {
        1: ("Bedroom mockup of blue haunted pumpkin ghost comforter set with moon and graveyard silhouettes.", "Haunted pumpkin ghost comforter set"),
        2: ("Close-up of blue Halloween bedding with ghost, pumpkin, moon and bare tree details.", "Haunted ghost comforter close-up"),
        3: ("Warm and cozy room mockup of blue Halloween bedding with pumpkins and ghosts.", "Blue Halloween bedding room mockup"),
        4: ("Front bedroom mockup showing moon, ghosts, jack-o-lanterns, trees and graveyard silhouettes.", "Pumpkin ghost Halloween bedding"),
        5: COMMON_COMFORTER_MATTRESS,
        6: ("Package include panel with haunted pumpkin ghost bedding and size table.", "Haunted pumpkin comforter package panel"),
    },
    168: {
        1: ("Bedroom mockup of pink cute ghost Halloween comforter set with matching pillows.", "Pink cute ghost Halloween comforter set"),
        2: ("Close-up of pink ghost bedding pattern with witch hats, candy and small stars.", "Pink ghost comforter close-up"),
        3: ("Warm and cozy room mockup of pink Halloween bedding with cute ghosts.", "Pink Halloween bedding room mockup"),
        4: ("Front bedroom mockup showing pink ghost pattern with witch hats and candy motifs.", "Cute ghost Halloween bedding"),
        5: COMMON_COMFORTER_MATTRESS,
        6: ("Package include panel with pink cute ghost Halloween bedding and size table.", "Pink ghost comforter package panel"),
    },
    169: {
        1: ("Bedroom mockup of white red handprint Halloween comforter set with red sheet pieces.", "Red handprint Halloween comforter set"),
        2: ("Close-up of red handprint and splatter-style Halloween bedding pattern.", "Red handprint comforter close-up"),
        3: ("Warm and cozy room mockup of red handprint Halloween bedding.", "Red handprint Halloween bedding room mockup"),
        4: ("Front bedroom mockup showing red handprints and splatter-style marks on white bedding.", "Red splatter Halloween bedding"),
        5: COMMON_COMFORTER_MATTRESS,
        6: ("Package include panel with red handprint Halloween bedding and size table.", "Red handprint comforter package panel"),
    },
    170: {
        1: ("Bedroom mockup of white haunted house pumpkin comforter set with black cats and rust sheets.", "Haunted house pumpkin comforter set"),
        2: ("Close-up of white Halloween bedding with haunted houses, pumpkins, cats and spiderwebs.", "Haunted house comforter close-up"),
        3: ("Warm and cozy room mockup of white Halloween bedding with haunted houses and pumpkins.", "White Halloween bedding room mockup"),
        4: ("Front bedroom mockup showing haunted houses, pumpkins, black cats, spiderwebs and leaves.", "Haunted house pumpkin bedding"),
        5: COMMON_COMFORTER_MATTRESS,
        6: ("Package include panel with haunted house pumpkin bedding and size table.", "Haunted house comforter package panel"),
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


def customizer_sentence(audit):
    excerpt = audit.get("config_excerpt", "")
    text_matches = re.findall(
        r'\{"id":\s*"[^"]+",\s*"label":\s*"([^"]+)",\s*"required":\s*(true|false),\s*"minLength":\s*(\d+),\s*"maxLength":\s*(\d+)',
        excerpt,
    )
    option_matches = re.findall(
        r'\{"id":\s*"[^"]+",\s*"label":\s*"([^"]+)",\s*"required":\s*(true|false),\s*"defaultOptionId"',
        excerpt,
    )
    parts = []
    for label, required, min_length, max_length in text_matches:
        requirement = "required" if required == "true" else "optional"
        parts.append(f"Personalization uses a {requirement} {label} field, {min_length}-{max_length} characters.")
    useful_options = [label for label, _ in option_matches if label and "confirmation" not in label.lower()]
    if useful_options:
        parts.append(f"Additional visible customization option: {', '.join(useful_options)}.")
    if not text_matches:
        parts.append("No shopper text-entry field is described for this product.")
    labels = " ".join(label for label, *_ in text_matches).lower()
    if "number" in labels:
        parts.append("Names and numbers visible in mockups are sample artwork unless the matching input is entered.")
    elif "name" in labels:
        parts.append("Visible names in mockups are sample artwork unless the matching input is entered.")
    elif text_matches:
        parts.append("Visible custom text in mockups is sample artwork unless the matching input is entered.")
    else:
        parts.append("Displayed wording or artwork is part of the product design shown, not a shopper-entered text claim.")
    return " ".join(parts)


def description(pos, admin_row, ctext):
    item = PRODUCT_UPDATES[pos]
    return (
        f"<p>{item['title']} features {item['detail']}. "
        "The page copy focuses on the specific visible design, product type and option groups available in the product data.</p>"
        "<h3>Design Details</h3>"
        f"<ul><li>Artwork: {item['detail']}.</li>"
        "<li>Gallery images show the main product mockup plus detail, feature, fit, size or included-item panels where present.</li></ul>"
        "<h3>Options and Fit</h3>"
        f"<ul><li>Available option groups: {option_text(admin_row)}.</li>"
        f"<li>{ctext}</li>"
        "<li>Select the product type and size shown on the product page before checkout.</li></ul>"
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
        ctext = customizer_sentence(qa["customizer_by_pos"].get(pos, {}))
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
        ws.cell(row_num, idx["description_proposed_html"]).value = description(pos, admin_row, ctext)
        ws.cell(row_num, idx["primary_keyword"]).value = item["primary"]
        ws.cell(row_num, idx["secondary_keywords"]).value = item["secondary"]
        ws.cell(row_num, idx["meta_keyword"]).value = item["primary"]
        ws.cell(row_num, idx["keyword_strategy"]).value = item["intent_role"]
        ws.cell(row_num, idx["buyer_search_summary"]).value = (
            f"US English buyer intent targets {item['cluster']}; no search-volume, trend or ranking claim is made."
        )
        ws.cell(row_num, idx["keyword_evidence_level"]).value = "SERP_ONLY"
        ws.cell(row_num, idx["revision"]).value = "r3"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields"]:
            ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            f"R3 deep recheck after Sang QA batch 017: used products_export_1.csv sha256:{admin_hash}; "
            "removed internal workflow copy, filled publishable title values, avoided unsupported personalization and upload claims, "
            "split nearby Halloween and Norse/Owl intents by visible motif, and rewrote image observations/alts from inspected contact sheets. "
            "Still NEEDS_REVIEW."
        )
        revised_products += 1

    ws = wb["Image_Audit"]
    idx = headers(ws)
    revised_images = 0
    admin_alt_matches = 0
    handle_to_pk = {v: k for k, v in qa["handle_by_key"].items()}
    seen_alts = set()
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
        if alt in seen_alts:
            alt = f"{alt} for {PRODUCT_UPDATES[pos]['cluster']}"
        seen_alts.add(alt)
        ws.cell(row_num, idx["observed_visual_details"]).value = observation
        ws.cell(row_num, idx["alt_proposed"]).value = alt[:125]
        ws.cell(row_num, idx["alt_action"]).value = "SET"
        ws.cell(row_num, idx["viewed_status"]).value = "VIEWED"
        ws.cell(row_num, idx["revision"]).value = "r3"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields", "approved_revision"]:
            if col_name in idx:
                ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            "R3: observation and alt rechecked from inspected contact sheet; admin image URL and current alt matched from Shopify CSV where available."
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
            "Sang QA batch 017 + products_export_1.csv admin baseline + customizer audit + contact-sheet inspection"
        )
        ws.cell(row_num, idx["checked_at"]).value = now
        ws.cell(row_num, idx["possible_overlap_with_other_products"]).value = item["intent_role"]
        ws.cell(row_num, idx["mapping_reason"]).value = (
            "R3 separates nearby pages by visible motif, wording, color palette and proof level for customization."
        )
        ws.cell(row_num, idx["mapping_status"]).value = "CANDIDATE_MAPPED_NEEDS_RECHECK"
        ws.cell(row_num, idx["mapping_version"]).value = "r3"
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
            "and only options that are actually supported."
        )
        ws.cell(row_num, idx["functional_motivation"]).value = (
            "Confirm product type, visible motif, size or fit panels, supported options and image accuracy before purchase."
        )
        ws.cell(row_num, idx["emotional_social_motivation"]).value = (
            "Choose a Norse, owl or Halloween bedding item without confusing it with another similar design."
        )
        ws.cell(row_num, idx["purchase_concerns"]).value = (
            "Avoid unsupported name, number, upload, cultural-origin or material claims; confirm exact size and product type."
        )
        ws.cell(row_num, idx["source_refs"]).value = (
            f"products_export_1.csv sha256:{admin_hash}; Sang QA {QA_RUN_ID}; customizer audit; contact sheets 161-170"
        )
        ws.cell(row_num, idx["observed_at"]).value = now
        ws.cell(row_num, idx["research_status"]).value = "REVISED_R3_NEEDS_RECHECK"
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
        ctext = customizer_sentence(qa["customizer_by_pos"].get(pos, {}))
        ws.cell(row_num, idx["sources_accessed"]).value = (
            f"Original storefront/product evidence; Sang QA {QA_RUN_ID}; products_export_1.csv sha256:{admin_hash}; "
            "customizer audit; inspected contact sheets"
        )
        ws.cell(row_num, idx["current_H1"]).value = admin_row["Title"]
        ws.cell(row_num, idx["current_meta_title"]).value = admin_row["SEO Title"] or "EMPTY"
        ws.cell(row_num, idx["verified_product_facts"]).value = (
            f"Admin export matched handle {handle}; type={admin_row['Type']}; options={option_text(admin_row)}; "
            f"status={admin_row['Status']}; image_count={len(admin_row['images'])}; design={item['detail']}; "
            f"customizer={ctext}"
        )
        ws.cell(row_num, idx["factual_conflicts"]).value = "Sang QA issues addressed in r3; independent QA recheck still required."
        ws.cell(row_num, idx["confidence_and_reason"]).value = (
            "R3 uses inspected contact sheets, customizer audit and Shopify admin CSV baseline; no approval or import created."
        )

    ws = wb["README_QA"]
    idx = headers(ws)
    for metric, value, definition in [
        ("qa_batch_017_r3_revision", "r3", "Separate 10-product revision after Sang QA; source workbook was not modified."),
        ("qa_batch_017_r3_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_017_r3_scope", "inventory positions 161-170", "No products outside qa_batch_017 were revised."),
        ("qa_batch_017_r3_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_017_r3_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_017_r3_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
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
        "scope": "qa_batch_017 only; inventory positions 161-170",
        "revision": "r3",
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
                "# Revision qa_batch_017_r3",
                "",
                "Artifact chính thức theo quy trình từng lô 10 sản phẩm.",
                "",
                f"- Nguồn r2: `{SOURCE.relative_to(ROOT)}`",
                f"- QA của Sang: `{QA_DATASET.relative_to(ROOT)}`",
                f"- Shopify admin export: `{ADMIN_EXPORT.relative_to(ROOT)}`",
                f"- Admin export SHA-256: `{admin_hash}`",
                f"- Workbook r3: `{OUTPUT.relative_to(ROOT)}`",
                f"- Sản phẩm sửa: {revised_products}",
                f"- Ảnh sửa/refresh: {revised_images}",
                f"- Admin image alt match: {admin_alt_matches}/{revised_images}",
                "- Sửa trọng yếu: rà lại theo rút kinh nghiệm các batch trước, bỏ nội dung nội bộ, giữ title/H1 publishable, tránh claim custom/upload không có bằng chứng, tách intent cho Norse/owl/Halloween theo motif nhìn thấy và viết lại mô tả/alt theo contact sheets.",
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
