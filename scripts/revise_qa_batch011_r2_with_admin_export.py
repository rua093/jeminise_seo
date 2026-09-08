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
BATCH_ID = "qa_batch_011_r2"
QA_RUN_ID = "20260907_181100"

SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_011.xlsx"
QA_DATASET = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "qa_dataset.json"
CUSTOMIZER_AUDIT = ROOT / "seo_runs" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "customizer_audit.json"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_011_r2.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_011_r2.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    101: {
        "title": "Keep On Truckin Flag Comforter",
        "meta_title": "Keep On Truckin Flag Comforter",
        "meta_description": "Shop a red semi truck comforter with American flag ribbons, Keep On Truckin text, YOUR NAME sample and size guide.",
        "primary": "Keep On Truckin flag comforter",
        "secondary": "red semi truck comforter, American flag trucker bedding, semi truck bedding set",
        "cluster": "red Keep On Truckin semi truck flag comforter",
        "detail": "red semi truck artwork with American flag ribbons, Keep On Truckin text and YOUR NAME sample panel",
        "intent_role": "Truck page separated by Keep On Truckin wording, red cab and flag ribbon background.",
        "customizer": "A required Enter Name field supports up to 35 characters, with a required product-confirmation choice.",
    },
    102: {
        "title": "Black Soccer Cleats Comforter Set",
        "meta_title": "Black Soccer Cleats Comforter Set",
        "meta_description": "Shop a grayscale soccer comforter set with black cleats, soccer ball artwork, Alexis sample text and 07 shams.",
        "primary": "black soccer cleats comforter set",
        "secondary": "soccer cleats bedding, grayscale soccer comforter, soccer ball cleats bedding",
        "cluster": "grayscale soccer cleats and ball comforter",
        "detail": "grayscale soccer ball and black cleats artwork with Alexis sample text and 07 pillow shams",
        "intent_role": "Soccer page separated by black cleats over ball artwork and grayscale palette.",
        "customizer": "Alexis and 07 are treated as sample artwork; the visible field is a general Customize Your Item text box.",
    },
    103: {
        "title": "Blue Soccer Goal Net Comforter Set",
        "meta_title": "Blue Soccer Goal Net Comforter Set",
        "meta_description": "Shop a blue soccer goal net comforter set with large ball artwork, Matthew sample text, 15 shams and size guide.",
        "primary": "blue soccer goal net comforter set",
        "secondary": "soccer goal bedding set, blue soccer ball comforter, soccer net comforter set",
        "cluster": "blue soccer ball in goal net comforter",
        "detail": "blue soccer ball in goal net artwork with Matthew sample text and 15 number on the bedding",
        "intent_role": "Soccer page separated by goal frame and blue net background with Matthew 15 sample art.",
        "customizer": "Matthew and 15 are treated as sample artwork; the visible field is a general Customize Your Item text box.",
    },
    104: {
        "title": "Kenzo Soccer Number Comforter Set",
        "meta_title": "Kenzo Soccer Number Comforter Set",
        "meta_description": "Shop a blue-white soccer comforter set with large number 10 ball artwork, Kenzo side text and bedding size panel.",
        "primary": "Kenzo soccer number comforter set",
        "secondary": "soccer number 10 bedding, blue soccer goal comforter, soccer ball comforter set",
        "cluster": "Kenzo number 10 soccer ball comforter",
        "detail": "blue and white soccer ball artwork with large number 10 and Kenzo sample text along the side",
        "intent_role": "Soccer page separated by Kenzo side text and large number 10 ball motif.",
        "customizer": "Kenzo and 10 are treated as sample artwork; the visible field is a general Customize Your Item text box.",
    },
    105: {
        "title": "Jayden Soccer Net Comforter Set",
        "meta_title": "Jayden Soccer Net Comforter Set",
        "meta_description": "Shop a multicolor soccer net comforter set with speed-line artwork, Jayden sample text, 07 ball and feature panels.",
        "primary": "Jayden soccer net comforter set",
        "secondary": "multicolor soccer bedding, soccer net speed line comforter, soccer number 07 bedding",
        "cluster": "multicolor Jayden soccer net comforter",
        "detail": "black, teal and orange speed-line soccer net artwork with Jayden sample text and 07 soccer ball",
        "intent_role": "Soccer page separated by multicolor speed lines, Jayden side text and 07 ball.",
        "customizer": "Jayden and 07 are treated as sample artwork; the visible field is a general Customize Your Item text box.",
    },
    106: {
        "title": "Matthew Water Splash Soccer Comforter",
        "meta_title": "Matthew Water Splash Soccer Comforter",
        "meta_description": "Shop a blue water splash soccer comforter set with Matthew sample text, 15 number shams and microfiber panels.",
        "primary": "Matthew water splash soccer comforter",
        "secondary": "blue soccer splash bedding, water splash soccer comforter, soccer ball bedding set",
        "cluster": "Matthew blue water splash soccer comforter",
        "detail": "blue water splash soccer ball artwork with Matthew sample text and 15 number on the shams",
        "intent_role": "Soccer page separated by water splash artwork, Matthew sample text and number 15.",
        "customizer": "Matthew and 15 are treated as sample artwork; the visible field is a general Customize Your Item text box.",
    },
    107: {
        "title": "Lucas Lightning Soccer Comforter Set",
        "meta_title": "Lucas Lightning Soccer Comforter Set",
        "meta_description": "Shop a dark lightning soccer comforter set with Lucas sample text, number 10 shams and blue electric splash artwork.",
        "primary": "Lucas lightning soccer comforter set",
        "secondary": "lightning soccer bedding, electric soccer comforter, dark blue soccer ball comforter",
        "cluster": "Lucas lightning splash soccer comforter",
        "detail": "dark blue soccer ball artwork with white lightning splash, Lucas sample text and number 10 shams",
        "intent_role": "Soccer page separated by lightning splash artwork, dark blue palette and Lucas 10 sample art.",
        "customizer": "Lucas and 10 are treated as sample artwork; the visible field is a general Customize Your Item text box.",
    },
    108: {
        "title": "Isaac Soccer Ball 07 Comforter Set",
        "meta_title": "Isaac Soccer Ball 07 Comforter Set",
        "meta_description": "Shop a white and blue soccer comforter set with paint-splash background, Isaac sample text and 07 ball artwork.",
        "primary": "Isaac soccer ball 07 comforter set",
        "secondary": "soccer number 07 bedding, white blue soccer comforter, soccer ball paint splash bedding",
        "cluster": "Isaac white blue soccer ball 07 comforter",
        "detail": "white and blue paint-splash soccer ball artwork with Isaac sample text and number 07",
        "intent_role": "Soccer page separated by white base, blue paint splash and Isaac 07 sample art.",
        "customizer": "Isaac and 07 are treated as sample artwork; the visible field is a general Customize Your Item text box.",
    },
    109: {
        "title": "James Green Soccer Goal Comforter",
        "meta_title": "James Green Soccer Goal Comforter",
        "meta_description": "Shop a green soccer goal comforter set with net splash artwork, James sample text, 07 shams and care panel.",
        "primary": "James green soccer goal comforter",
        "secondary": "green soccer bedding set, soccer goal splash comforter, soccer ball net bedding",
        "cluster": "James green soccer goal splash comforter",
        "detail": "green soccer ball goal splash artwork with James sample text, 07 shams and net-line background",
        "intent_role": "Soccer page separated by green ball, goal splash and James 07 sample art.",
        "customizer": "James and 07 are treated as sample artwork; the visible field is a general Customize Your Item text box.",
    },
    110: {
        "title": "Kenzo Black Soccer Net Comforter",
        "meta_title": "Kenzo Black Soccer Net Comforter",
        "meta_description": "Shop a black and white soccer net comforter set with Kenzo side text, 07 ball artwork and microfiber panels.",
        "primary": "Kenzo black soccer net comforter",
        "secondary": "black soccer net bedding, soccer number 07 comforter, black white soccer comforter set",
        "cluster": "Kenzo black and white soccer net comforter",
        "detail": "black and white soccer ball in net artwork with Kenzo sample side text and number 07",
        "intent_role": "Soccer page separated by black-white net background, Kenzo side text and 07 ball.",
        "customizer": "Kenzo and 07 are treated as sample artwork; the visible field is a general Customize Your Item text box.",
    },
}


COMMON_DUVET_PANEL = ("Comparison panel explaining duvet cover set and comforter set bedding options.", "Duvet cover and comforter set comparison panel")
COMMON_CARE = ("Stress-free easy-care panel listing wrinkle-free, stain-proof, anti-pilling and wash durability.", "Easy care panel for comforter bedding")
COMMON_SIZE = ("Size dimension chart showing twin, full, queen and king bedding measurements.", "Comforter set size dimension chart")


IMAGE_DETAILS = {
    101: {
        1: ("Bedroom mockup of red semi truck comforter with American flag ribbons, Keep On Truckin text and YOUR NAME panel.", "Keep On Truckin red semi truck comforter"),
        2: ("Bedroom mockup of red truck flag bedding set with matching shams and folded white top sheet.", "Red semi truck flag bedding set with shams"),
        3: COMMON_DUVET_PANEL,
        4: ("Microfiber feature panel with red semi truck pillow mockup and 3D printed pattern callouts.", "Red semi truck comforter feature panel"),
        5: ("Low-angle red truck bedding mockup with soft, lightweight, durable and breathable icons.", "Red truck comforter with feature icons"),
        6: COMMON_CARE,
        7: COMMON_SIZE,
    },
    102: {
        1: ("Bedroom mockup of grayscale soccer comforter with black cleats, soccer ball, Alexis text and 07 shams.", "Black soccer cleats comforter with Alexis 07"),
        2: ("Low-angle grayscale soccer cleats bedding mockup with soft, lightweight, durable and breathable icons.", "Soccer cleats bedding with feature icons"),
        3: ("Microfiber feature panel with soccer cleats pillow mockup and 3D printed pattern callouts.", "Soccer cleats comforter feature panel"),
        4: COMMON_DUVET_PANEL,
        5: ("Pillow sham mockup showing black soccer cleats and number 07 on white bedding.", "Soccer cleats pillow shams with 07"),
        6: ("Bedroom mockup of grayscale soccer cleats comforter with Alexis text.", "Grayscale soccer cleats comforter on bed"),
        7: COMMON_SIZE,
        8: COMMON_CARE,
    },
    103: {
        1: ("Bedroom mockup of blue soccer goal net comforter with large ball, Matthew text and 15 number.", "Blue soccer goal net comforter with Matthew 15"),
        2: ("Feature panel showing blue soccer ball bedding, microfiber label and 3D printed pattern callouts.", "Blue soccer goal net comforter feature panel"),
        3: ("Pillow sham mockup showing blue soccer ball and number 15 on white bedding.", "Blue soccer ball pillow shams with 15"),
        4: COMMON_DUVET_PANEL,
        5: ("Bedroom mockup of blue goal net soccer bedding with Matthew text and matching shams.", "Matthew blue soccer goal bedding set"),
        6: ("Low-angle blue soccer goal bedding mockup with soft, lightweight, durable and breathable icons.", "Blue soccer goal bedding with feature icons"),
        7: ("Close-up corner showing blue net artwork under white microfiber filling.", "Blue soccer goal comforter microfiber close-up"),
        8: COMMON_CARE,
    },
    104: {
        1: ("Bedroom mockup of blue-white soccer comforter with large number 10 ball and Kenzo side text.", "Kenzo soccer number 10 comforter set"),
        2: ("Feature panel showing Kenzo soccer pillow mockup, microfiber label and 3D printed pattern callouts.", "Kenzo soccer comforter feature panel"),
        3: ("Low-angle Kenzo number 10 soccer bedding mockup with feature icons.", "Kenzo soccer bedding with feature icons"),
        4: ("Pillow sham mockup showing Kenzo side text and number 10 soccer ball artwork.", "Kenzo soccer pillow shams with number 10"),
        5: COMMON_DUVET_PANEL,
        6: ("Bedroom mockup of Kenzo number 10 soccer bedding with matching shams.", "Kenzo soccer comforter set on bed"),
        7: COMMON_SIZE,
    },
    105: {
        1: ("Bedroom mockup of multicolor soccer net comforter with Jayden side text and number 07 ball.", "Jayden soccer net comforter with 07 ball"),
        2: ("Feature panel showing Jayden soccer net bedding details and 3D printed pattern callouts.", "Jayden soccer net comforter feature panel"),
        3: ("Pillow sham mockup showing number 07 ball with orange and teal speed lines.", "Soccer net pillow shams with 07"),
        4: ("Close-up corner showing soccer net artwork under white microfiber filling.", "Jayden soccer net comforter microfiber close-up"),
        5: COMMON_DUVET_PANEL,
        6: ("Bedroom mockup of Jayden soccer net bedding with multicolor speed-line artwork.", "Jayden multicolor soccer bedding set"),
        7: ("Low-angle soccer net bedding mockup with soft, lightweight, durable and breathable icons.", "Jayden soccer net bedding with feature icons"),
        8: COMMON_CARE,
    },
    106: {
        1: ("Bedroom mockup of blue water splash soccer comforter with Matthew text and 15 shams.", "Matthew water splash soccer comforter"),
        2: ("Feature panel showing water splash soccer bedding and 3D printed pattern callouts.", "Water splash soccer comforter feature panel"),
        3: COMMON_DUVET_PANEL,
        4: ("Bedroom mockup of Matthew water splash soccer bedding with matching 15 shams.", "Matthew water splash soccer bedding set"),
        5: ("Close-up corner showing water splash soccer artwork under white microfiber filling.", "Water splash soccer comforter microfiber close-up"),
        6: ("Low-angle water splash soccer bedding mockup with soft, lightweight, durable and breathable icons.", "Water splash soccer bedding with feature icons"),
        7: COMMON_CARE,
        8: COMMON_SIZE,
    },
    107: {
        1: ("Bedroom mockup of dark blue lightning soccer comforter with Lucas text and number 10 shams.", "Lucas lightning soccer comforter set"),
        2: ("Feature panel showing lightning soccer bedding and 3D printed pattern callouts.", "Lightning soccer comforter feature panel"),
        3: ("Pillow sham mockup showing number 10 soccer ball with lightning splash.", "Lightning soccer pillow shams with number 10"),
        4: COMMON_DUVET_PANEL,
        5: ("Close-up corner showing lightning soccer artwork under white microfiber filling.", "Lightning soccer comforter microfiber close-up"),
        6: ("Bedroom mockup of Lucas lightning soccer bedding with matching shams.", "Lucas lightning soccer bedding on bed"),
        7: ("Low-angle lightning soccer bedding mockup with soft, lightweight, durable and breathable icons.", "Lightning soccer bedding with feature icons"),
        8: COMMON_CARE,
    },
    108: {
        1: ("Bedroom mockup of white and blue soccer comforter with Isaac text and number 07 ball.", "Isaac soccer ball 07 comforter set"),
        2: ("Feature panel showing Isaac soccer ball bedding and 3D printed pattern callouts.", "Isaac soccer comforter feature panel"),
        3: ("Bedroom mockup of Isaac soccer ball bedding with blue paint-splash background.", "Isaac soccer comforter set on bed"),
        4: ("Close-up corner showing 07 soccer ball artwork under white microfiber filling.", "Isaac soccer ball comforter microfiber close-up"),
        5: COMMON_DUVET_PANEL,
        6: ("Low-angle Isaac soccer bedding mockup with soft, lightweight, durable and breathable icons.", "Isaac soccer bedding with feature icons"),
        7: COMMON_CARE,
        8: COMMON_SIZE,
    },
    109: {
        1: ("Bedroom mockup of green soccer goal comforter with James text, 07 shams and net splash artwork.", "James green soccer goal comforter"),
        2: ("Feature panel showing green soccer goal bedding and 3D printed pattern callouts.", "Green soccer goal comforter feature panel"),
        3: ("Pillow sham mockup showing green soccer ball with number 10 on white bedding.", "Green soccer pillow shams with number 10"),
        4: COMMON_DUVET_PANEL,
        5: ("Bedroom mockup of James green soccer goal bedding with matching 07 shams.", "James green soccer goal bedding set"),
        6: ("Low-angle green soccer goal bedding mockup with soft, lightweight, durable and breathable icons.", "Green soccer goal bedding with feature icons"),
        7: COMMON_CARE,
    },
    110: {
        1: ("Bedroom mockup of black and white soccer net comforter with Kenzo side text and number 07 ball.", "Kenzo black soccer net comforter with 07"),
        2: ("Bedroom mockup of Kenzo black soccer net bedding set with matching 07 shams.", "Kenzo black soccer net bedding set"),
        3: ("Feature panel showing black soccer net bedding and 3D printed pattern callouts.", "Black soccer net comforter feature panel"),
        4: COMMON_DUVET_PANEL,
        5: ("Pillow sham mockup showing black soccer ball and number 07 on white bedding.", "Black soccer ball pillow shams with 07"),
        6: ("Low-angle black soccer net bedding mockup with soft, lightweight, durable and breathable icons.", "Black soccer net bedding with feature icons"),
        7: ("Close-up corner showing black soccer ball artwork under white microfiber filling.", "Black soccer net comforter microfiber close-up"),
        8: COMMON_CARE,
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
        "data": data,
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
    type_text = admin_row["Type"] or "bedding"
    return (
        f"<p>{item['title']} features {item['detail']}. "
        f"It gives shoppers a specific {type_text.lower()} design instead of a generic soccer or truck print.</p>"
        "<h3>Design Details</h3>"
        f"<ul><li>Artwork: {item['detail']}.</li>"
        "<li>Gallery images show main mockups plus close-up, feature, care, bedding-type or size panels where present.</li></ul>"
        "<h3>Options and Fit</h3>"
        f"<ul><li>Available option groups: {option_text(admin_row)}.</li>"
        f"<li>{item['customizer']}</li>"
        "<li>Check the selected product type, size and visible customization field before checkout.</li></ul>"
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
        ws.cell(row_num, idx["product_type"]).value = admin_row["Type"] or "Blanket"
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
            f"R2 after Sang QA batch 011: used products_export_1.csv sha256:{admin_hash}; "
            "removed internal workflow copy, removed unsafe personalization claims from soccer products, "
            "kept product 101 name-field claim tied to Enter Name evidence, split soccer keyword intents by visible motif, "
            "and rewrote image observations/alts from inspected contact sheets. Still NEEDS_REVIEW."
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
            "Sang QA batch 011 + products_export_1.csv admin baseline + customizer_audit.json + contact-sheet inspection"
        )
        ws.cell(row_num, idx["checked_at"]).value = now
        ws.cell(row_num, idx["possible_overlap_with_other_products"]).value = item["intent_role"]
        ws.cell(row_num, idx["mapping_reason"]).value = (
            "R2 separates nearby soccer pages by visible motif, sample wording, color palette and proof level for customization."
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
            f"When shopping for {item['cluster']}, the buyer wants the page to confirm the exact artwork, product type, size choices, "
            "and only the customization options that are actually supported."
        )
        ws.cell(row_num, idx["functional_motivation"]).value = (
            "Confirm product type, visible motif, size/care panels, supported custom fields and image accuracy before purchase."
        )
        ws.cell(row_num, idx["emotional_social_motivation"]).value = (
            "Choose a sports or truck bedding item without confusing it with another similar design in the batch."
        )
        ws.cell(row_num, idx["purchase_concerns"]).value = (
            "Avoid unsupported name, number, photo or material claims; confirm exact size and product type."
        )
        ws.cell(row_num, idx["source_refs"]).value = (
            f"products_export_1.csv sha256:{admin_hash}; Sang QA {QA_RUN_ID}; customizer_audit.json; contact sheets 101-110"
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
        ("qa_batch_011_r2_revision", "r2", "Separate 10-product revision after Sang QA; source workbook was not modified."),
        ("qa_batch_011_r2_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_011_r2_scope", "inventory positions 101-110", "No products outside qa_batch_011 were revised."),
        ("qa_batch_011_r2_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_011_r2_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_011_r2_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
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
        "scope": "qa_batch_011 only; inventory positions 101-110",
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
                "# Revision qa_batch_011_r2",
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
                "- Sửa trọng yếu: bỏ nội dung nội bộ, gỡ claim personalization rủi ro cho nhóm soccer, giữ claim tên ở sản phẩm 101 theo bằng chứng Enter Name, tách keyword intent và viết lại mô tả/alt theo contact sheets.",
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
