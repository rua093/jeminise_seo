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
BATCH_ID = "qa_batch_018_r3"
QA_RUN_ID = "20260907_213100"

SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / "qa_batch_018_r2" / "SEO_Product_Optimization_qa_batch_018_r2.xlsx"
QA_DATASET = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "qa_dataset.json"
CUSTOMIZER_AUDIT = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "customizer_audit.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_018_r3.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_018_r3.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    171: {
        "title": "Pink Gothic Skull Comforter Set",
        "meta_title": "Pink Gothic Skull Comforter Set",
        "meta_description": "Shop a pink black gothic skull comforter set with patchwork artwork, skeletons, ravens, sheets and size chart.",
        "primary": "pink gothic skull comforter set",
        "secondary": "gothic skull bedding, skeleton comforter set, pink black Halloween bedding",
        "cluster": "pink black gothic skull patchwork comforter set",
        "detail": "pink, black and gray patchwork design with skulls, skeleton figures, ravens, bare trees, stripes and checkerboard panels",
        "intent_role": "Halloween comforter page separated by pink-black gothic patchwork, skulls, skeletons and raven panels.",
        "customizer": "No customization control is shown for this product, so the copy only describes the visible artwork.",
        "custom_supported": False,
    },
    172: {
        "title": "Trick or Treat Haunted House Comforter Set",
        "meta_title": "Trick or Treat Haunted House Comforter Set",
        "meta_description": "Shop a haunted house Halloween comforter set with Trick or Treat text, pumpkins, ghosts, webs and sheets.",
        "primary": "trick or treat haunted house comforter set",
        "secondary": "haunted house bedding, trick or treat bedding, black Halloween comforter set",
        "cluster": "black Trick or Treat haunted house Halloween comforter set",
        "detail": "black haunted house scene with Trick or Treat text, orange moon, pumpkins, ghosts, skeletons, webs and bare trees",
        "intent_role": "Halloween comforter page separated by Trick or Treat wording and haunted house scene.",
        "customizer": "A customize button is present, but no working text or image field was captured; no customization claim is used.",
        "custom_supported": False,
    },
    173: {
        "title": "Cream Pumpkin Ghost Comforter Set",
        "meta_title": "Cream Pumpkin Ghost Comforter Set",
        "meta_description": "Shop a cream pumpkin ghost Halloween comforter set with bats, spiderwebs, black keys, sheets and size chart.",
        "primary": "cream pumpkin ghost comforter set",
        "secondary": "pumpkin ghost bedding, cream Halloween comforter, Halloween comforter set with sheets",
        "cluster": "cream pumpkin ghost Halloween comforter set",
        "detail": "cream Halloween pattern with orange jack-o-lantern pumpkins, white ghosts, black bats, spiderwebs and ornate keys",
        "intent_role": "Halloween comforter page separated by cream base, pumpkin-ghost pattern and black key details.",
        "customizer": "No customization control is shown for this product, so the copy only describes the visible artwork.",
        "custom_supported": False,
    },
    174: {
        "title": "Custom Photo Music Player Quilt",
        "meta_title": "Custom Photo Music Player Quilt",
        "meta_description": "Shop a custom photo music player quilt with song name, artist name, waveform artwork, optional shams and size chart.",
        "primary": "custom photo music player quilt",
        "secondary": "personalized music player quilt, custom song quilt, couple photo quilt",
        "cluster": "custom photo song name music player quilt",
        "detail": "music player interface quilt showing a couple image area, Song Name and Artist Name text, waveform, play controls and heart icon",
        "intent_role": "Personalized quilt page supported by required image input and song-name text field in the customizer audit.",
        "customizer": "Personalization fields captured: required Song Name, optional Artist Name, optional custom text and required Image on Quilt.",
        "custom_supported": True,
    },
    175: {
        "title": "Autumn Tree of Life Birds Quilt Set",
        "meta_title": "Autumn Tree of Life Birds Quilt Set",
        "meta_description": "Shop an autumn Tree of Life quilt set with birds, flowers, golden leaves, optional shams and size chart.",
        "primary": "autumn Tree of Life birds quilt set",
        "secondary": "autumn tree quilt, birds flowers quilt, Tree of Life bedding",
        "cluster": "autumn Tree of Life birds and flowers quilt set",
        "detail": "radiant autumn Tree of Life design with blue birds, golden leaves, orange foliage and colorful flowers",
        "intent_role": "Tree of Life page separated by autumn color palette, birds, flowers and golden landscape artwork.",
        "customizer": "A general Customize Your Quilt text box is present; no name, number or image-submission field is claimed.",
        "custom_supported": False,
    },
    176: {
        "title": "Custom Name Baseball Flag Bedding",
        "meta_title": "Custom Name Baseball Flag Bedding",
        "meta_description": "Shop custom name baseball flag bedding with red blue flag artwork, baseballs, required name field and product confirmation.",
        "primary": "custom name baseball flag bedding",
        "secondary": "personalized baseball bedding, baseball flag comforter set, baseball name bedding",
        "cluster": "custom name baseball American flag bedding design 01",
        "detail": "red, white and blue American flag bedding with baseballs and a large sample vertical name on the design",
        "intent_role": "Baseball bedding page supported by a required name customizer field and separated by vertical flag-name layout.",
        "customizer": "Required field captured: Customize Your Name, plus a product-type confirmation control.",
        "custom_supported": True,
    },
    177: {
        "title": "Custom Baseball Glove Bedding",
        "meta_title": "Custom Baseball Glove Bedding",
        "meta_description": "Shop custom baseball glove bedding with black background, glove and ball artwork, required name and number fields.",
        "primary": "custom baseball glove bedding",
        "secondary": "personalized baseball bedding, baseball glove comforter, custom name sports bedding",
        "cluster": "custom name baseball glove black bedding design 02",
        "detail": "black baseball bedding with a glove, baseball, wooden bat and cursive sample name across the comforter",
        "intent_role": "Baseball bedding page supported by a required name field and separated by black glove-and-bat layout.",
        "customizer": "Required field captured: Customize Your Name, plus a product-type confirmation control.",
        "custom_supported": True,
    },
    178: {
        "title": "Baseball Flag Glove Bedding Set",
        "meta_title": "Baseball Flag Glove Bedding Set",
        "meta_description": "Shop a baseball flag bedding set with glove, ball, distressed red blue stripes, feature panels and zipper detail.",
        "primary": "baseball flag glove bedding set",
        "secondary": "baseball bedding set, American flag baseball bedding, baseball glove comforter",
        "cluster": "baseball glove American flag bedding set design 03",
        "detail": "distressed American flag bedding with red and blue stripes, stars, baseball glove, ball and large sample name text",
        "intent_role": "Baseball bedding page separated by American flag background and glove-ball artwork; no live custom field was captured.",
        "customizer": "No customization control is shown for this product; visible name text is treated as sample artwork.",
        "custom_supported": False,
    },
    179: {
        "title": "Custom Baseball Home Quote Bedding",
        "meta_title": "Custom Baseball Home Quote Bedding",
        "meta_description": "Shop custom baseball quote bedding with There is no place like Home text, player silhouettes and required name field.",
        "primary": "custom baseball home quote bedding",
        "secondary": "baseball quote bedding, personalized sports bedding, custom name baseball comforter",
        "cluster": "custom name baseball quote home bedding design 04",
        "detail": "cream baseball bedding with player silhouettes, baseball stitches and There is no place like Home quote text",
        "intent_role": "Baseball bedding page supported by a required name field and separated by home quote typography.",
        "customizer": "Required field captured: Customize Your Name, plus a product-type confirmation control.",
        "custom_supported": True,
    },
    180: {
        "title": "Catcher American Flag Baseball Bedding",
        "meta_title": "Catcher American Flag Baseball Bedding",
        "meta_description": "Shop catcher baseball bedding with American flag artwork, catcher silhouette, sample number text and feature panels.",
        "primary": "catcher American flag baseball bedding",
        "secondary": "catcher baseball bedding, American flag sports bedding, baseball comforter set",
        "cluster": "catcher American flag baseball bedding design 05",
        "detail": "American flag baseball bedding with a catcher silhouette, red stripes, blue stars and sample name and number text",
        "intent_role": "Baseball bedding page separated by catcher silhouette and American flag layout; no live custom field was captured.",
        "customizer": "No customization control is shown for this product; visible name and number text are treated as sample artwork.",
        "custom_supported": False,
    },
}


COMMON_COMFORTER_MATTRESS = ("Fitted sheet panel showing standard and deep mattress fit up to 14 inches.", "Fitted sheet deep mattress panel")
COMMON_BASEBALL_WEAVING = ("High-density weaving panel comparing BeddingOutlet fabric with other weave samples.", "High-density weaving feature panel")
COMMON_BASEBALL_WASH = ("Machine washable care panel with washing machine, laundry basket and care text.", "Machine washable bedding care panel")
COMMON_BASEBALL_ZIPPER = ("Bottom zippered closure panel showing white zipper close-up.", "Bottom zippered closure panel")
COMMON_QUILT_SIZE = ("Premium quilt set size chart showing available bed sizes and measurements.", "Quilt set size chart")
COMMON_QUILT_BEDSPREAD = ("Bedspread features panel showing microfiber layers and comfort feature icons.", "Quilt construction feature panel")


IMAGE_DETAILS = {
    171: {
        1: ("Bedroom mockup of pink black gothic skull comforter set with patchwork panels and matching pillows.", "Pink gothic skull comforter set"),
        2: ("Close-up of gray comforter folded over pink black skull, raven and patchwork print.", "Gothic skull comforter close-up"),
        3: ("Warm and cozy room mockup of gothic skull Halloween bedding with pumpkins.", "Pink black gothic bedding room mockup"),
        4: ("Front bedroom mockup showing skulls, skeletons, ravens, stripes and checkerboard panels.", "Gothic skull patchwork bedding"),
        5: COMMON_COMFORTER_MATTRESS,
        6: ("Package include panel with gothic skull comforter set and size table.", "Gothic skull comforter package panel"),
    },
    172: {
        1: ("Bedroom mockup of Trick or Treat haunted house comforter set with ghosts and pumpkins.", "Trick or Treat haunted house comforter set"),
        2: ("Close-up of black haunted house bedding with orange moon, webs, ghosts and pumpkins.", "Haunted house comforter close-up"),
        3: ("Warm and cozy room mockup of haunted house Halloween bedding.", "Haunted house Halloween bedding room mockup"),
        4: ("Front bedroom mockup showing haunted house, Trick or Treat text, skeletons and pumpkins.", "Trick or Treat Halloween bedding"),
        5: COMMON_COMFORTER_MATTRESS,
        6: ("Package include panel with haunted house comforter set and size table.", "Haunted house comforter package panel"),
    },
    173: {
        1: ("Bedroom mockup of cream pumpkin ghost Halloween comforter set with matching pillows.", "Cream pumpkin ghost comforter set"),
        2: ("Close-up of cream Halloween bedding with pumpkins, ghosts, spiderwebs, bats and black keys.", "Pumpkin ghost comforter close-up"),
        3: ("Warm and cozy room mockup of cream pumpkin ghost Halloween bedding.", "Cream Halloween bedding room mockup"),
        4: ("Front bedroom mockup showing pumpkins, ghosts, bats, webs and ornate black keys.", "Cream pumpkin ghost Halloween bedding"),
        5: COMMON_COMFORTER_MATTRESS,
        6: ("Package include panel with cream pumpkin ghost comforter set and size table.", "Pumpkin ghost comforter package panel"),
    },
    174: {
        1: ("Person holding custom photo music player quilt with song name, artist name, waveform and play controls.", "Custom photo music player quilt"),
        2: COMMON_QUILT_BEDSPREAD,
        3: ("Bedroom mockup of music player quilt with couple image area, song text and matching pillows.", "Music player quilt bedroom mockup"),
        4: COMMON_QUILT_SIZE,
        5: ("Room mockup of music player quilt with printed craft callout and fabric inset.", "Music player quilt fabric mockup"),
        6: ("Optional pillow shams panel for music player quilt with lightweight, soft, anti-pill and anti-static icons.", "Music player quilt pillow shams"),
        7: ("High-quality fabric panel showing music player quilt layers and bed cover use text.", "Music player quilt fabric panel"),
    },
    175: {
        1: ("Bedroom mockup of autumn Tree of Life quilt set with birds, flowers and matching shams.", "Autumn Tree of Life birds quilt set"),
        2: ("Room mockup of autumn Tree of Life quilt with printed craft callout and fabric inset.", "Autumn Tree of Life quilt room mockup"),
        3: ("Optional pillow shams panel showing autumn Tree of Life artwork with birds and comfort icons.", "Tree of Life birds pillow shams"),
        4: ("High-quality fabric panel showing autumn leaves, blue bird and quilt layer close-up.", "Autumn Tree of Life fabric panel"),
        5: COMMON_QUILT_SIZE,
        6: COMMON_QUILT_BEDSPREAD,
        7: ("Second bedroom mockup of autumn Tree of Life quilt with golden leaves and blue birds.", "Autumn Tree of Life bedroom view"),
    },
    176: {
        1: ("Bedroom mockup of custom name baseball flag bedding with baseballs and vertical sample name.", "Custom name baseball flag bedding"),
        2: COMMON_BASEBALL_WEAVING,
        3: COMMON_BASEBALL_WASH,
        4: COMMON_BASEBALL_ZIPPER,
    },
    177: {
        1: ("Bedroom mockup of custom baseball glove bedding with black background, bat, ball and cursive sample name.", "Custom baseball glove bedding"),
        2: COMMON_BASEBALL_WEAVING,
        3: COMMON_BASEBALL_WASH,
        4: COMMON_BASEBALL_ZIPPER,
    },
    178: {
        1: ("Bedroom mockup of baseball flag glove bedding with distressed flag, glove, ball and sample name text.", "Baseball flag glove bedding set"),
        2: COMMON_BASEBALL_WEAVING,
        3: COMMON_BASEBALL_WASH,
        4: COMMON_BASEBALL_ZIPPER,
    },
    179: {
        1: ("Bedroom mockup of custom baseball Home quote bedding with player silhouettes and sample names on pillows.", "Custom baseball Home quote bedding"),
        2: COMMON_BASEBALL_WEAVING,
        3: COMMON_BASEBALL_WASH,
        4: COMMON_BASEBALL_ZIPPER,
    },
    180: {
        1: ("Bedroom mockup of catcher American flag baseball bedding with catcher silhouette, sample name and number text.", "Catcher American flag baseball bedding"),
        2: COMMON_BASEBALL_WEAVING,
        3: COMMON_BASEBALL_WASH,
        4: COMMON_BASEBALL_ZIPPER,
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
    image_labels = re.findall(r'"label":\s*"([^"]*Image[^"]*)"', excerpt)
    parts = []
    for label, required, min_length, max_length in text_matches:
        requirement = "required" if required == "true" else "optional"
        parts.append(f"Personalization uses a {requirement} {label} field, {min_length}-{max_length} characters.")
    useful_options = [label for label, _ in option_matches if label and "confirmation" not in label.lower()]
    if useful_options:
        parts.append(f"Additional visible customization option: {', '.join(useful_options)}.")
    if image_labels:
        parts.append(f"Image customization shown: {', '.join(dict.fromkeys(image_labels))}.")
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
            f"R3 deep recheck after Sang QA batch 018: used products_export_1.csv sha256:{admin_hash}; "
            "removed internal workflow copy, filled publishable title values, corrected unsupported personalization/upload claims, "
            "split nearby Halloween/baseball/quilt intents by visible motif and customizer proof, and rewrote image observations/alts from inspected contact sheets. "
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
            "Sang QA batch 018 + products_export_1.csv admin baseline + customizer audit + contact-sheet inspection"
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
            "and whether the relevant customization or option fields are actually available."
        )
        ws.cell(row_num, idx["functional_motivation"]).value = (
            "Confirm product type, visible motif, size or fit panels, supported options and image accuracy before purchase."
        )
        ws.cell(row_num, idx["emotional_social_motivation"]).value = (
            "Choose a Halloween, romantic music, Tree of Life or baseball bedding item without confusing it with another similar design."
        )
        ws.cell(row_num, idx["purchase_concerns"]).value = (
            "Avoid unsupported name, number, image, cultural-origin or material claims; confirm exact size and product type."
        )
        ws.cell(row_num, idx["source_refs"]).value = (
            f"products_export_1.csv sha256:{admin_hash}; Sang QA {QA_RUN_ID}; customizer audit; contact sheets 171-180"
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
        ("qa_batch_018_r3_revision", "r3", "Separate 10-product revision after Sang QA; source workbook was not modified."),
        ("qa_batch_018_r3_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_018_r3_scope", "inventory positions 171-180", "No products outside qa_batch_018 were revised."),
        ("qa_batch_018_r3_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_018_r3_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_018_r3_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
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
        "scope": "qa_batch_018 only; inventory positions 171-180",
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
                "# Revision qa_batch_018_r3",
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
                "- Sửa trọng yếu: rà lại theo rút kinh nghiệm các batch trước, bỏ nội dung nội bộ, giữ title/H1 publishable, sửa claim cá nhân hóa theo customizer audit, tách intent cho Halloween/music/Tree of Life/baseball theo motif nhìn thấy và viết lại mô tả/alt theo contact sheets.",
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
