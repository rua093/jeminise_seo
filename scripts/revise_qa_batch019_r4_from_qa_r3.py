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
BATCH_ID = "qa_batch_019_r4"
QA_RUN_ID = "20260908_111000"

SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / "qa_batch_019_r3" / "SEO_Product_Optimization_qa_batch_019_r3.xlsx"
QA_REPORT = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "SEO_QA_qa_batch_019_r3.md"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_019_r4.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_019_r4.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    181: {
        "title": "Custom Baseball Glove Name Bedding",
        "meta_title": "Custom Baseball Glove Name Bedding",
        "meta_description": "Shop custom baseball glove bedding with brown glove artwork, baseballs, required name field and feature panels.",
        "primary": "custom baseball glove name bedding",
        "secondary": "baseball glove bedding, personalized baseball comforter, custom name baseball bedding",
        "cluster": "custom name brown baseball glove bedding design 06",
        "detail": "brown baseball glove bedding with close-up baseball, glove leather texture, netting and sample name and number text",
        "intent_role": "Baseball bedding page supported by a required name field and separated by brown glove close-up artwork; visible numbers are sample artwork only.",
    },
    182: {
        "title": "Custom Baseball Flag Name Bedding",
        "meta_title": "Custom Baseball Flag Name Bedding",
        "meta_description": "Shop custom baseball flag bedding with large baseball, red white blue flag artwork and required name field.",
        "primary": "custom baseball flag name bedding",
        "secondary": "American flag baseball bedding, personalized baseball bedding, baseball name comforter",
        "cluster": "custom name baseball American flag bedding design 07",
        "detail": "red, white and blue American flag bedding with a large baseball and sample cursive name near the bottom edge",
        "intent_role": "Baseball bedding page supported by a required name field and separated by oversized ball on flag layout.",
    },
    183: {
        "title": "Custom Baseball Flag Glove Bedding",
        "meta_title": "Custom Baseball Flag Glove Bedding",
        "meta_description": "Shop custom baseball flag glove bedding with American flag background, baseball artwork and required name field.",
        "primary": "custom baseball flag glove bedding",
        "secondary": "baseball flag bedding, American flag baseball bedding, personalized sports bedding",
        "cluster": "custom name baseball flag glove bedding design 08",
        "detail": "distressed American flag baseball bedding with glove corners, oversized baseball, sample name and large sample number",
        "intent_role": "Baseball bedding page supported by a required name field and separated by large sample number artwork on the baseball layout.",
    },
    184: {
        "title": "Custom Close-Up Baseball Glove Bedding",
        "meta_title": "Custom Close-Up Baseball Glove Bedding",
        "meta_description": "Shop custom baseball glove bedding with close-up glove and ball artwork, required name field and feature panels.",
        "primary": "custom close-up baseball glove bedding",
        "secondary": "baseball glove comforter, personalized baseball bedding, custom name sports bedding",
        "cluster": "custom name close-up baseball glove bedding design 09",
        "detail": "close-up brown baseball glove and ball bedding with sample name printed across the baseball",
        "intent_role": "Baseball bedding page supported by a required name field and separated by full-frame glove close-up.",
    },
    185: {
        "title": "Custom Gray Baseball Flag Bedding",
        "meta_title": "Custom Gray Baseball Flag Bedding",
        "meta_description": "Shop custom gray baseball flag bedding with glove and ball artwork, required name field and feature panels.",
        "primary": "custom gray baseball flag bedding",
        "secondary": "gray baseball bedding, American flag baseball comforter, personalized baseball bedding",
        "cluster": "custom name gray baseball American flag bedding design 10",
        "detail": "gray baseball glove and ball artwork over a distressed American flag background with a large sample name",
        "intent_role": "Baseball bedding page supported by a required name field and separated by gray glove over flag artwork.",
    },
    186: {
        "title": "Custom Baseball Pattern Name Bedding",
        "meta_title": "Custom Baseball Pattern Name Bedding",
        "meta_description": "Shop custom baseball pattern bedding with red white blue stripes, many baseballs, required name field and feature panels.",
        "primary": "custom baseball pattern name bedding",
        "secondary": "red white blue baseball bedding, baseball number bedding, personalized baseball comforter",
        "cluster": "custom name baseball pattern red blue bedding design 11",
        "detail": "red, white and blue striped bedding with repeating baseballs, star panels, sample name and sample number on shams",
        "intent_role": "Baseball bedding page supported by a required name field and separated by repeated baseball pattern; visible numbers are sample artwork only.",
    },
    187: {
        "title": "Custom Pitcher Baseball Bedding",
        "meta_title": "Custom Pitcher Baseball Bedding",
        "meta_description": "Shop custom pitcher baseball bedding with silhouette artwork, baseball stitch border and required name field.",
        "primary": "custom pitcher baseball bedding",
        "secondary": "baseball player bedding, pitcher silhouette comforter, personalized baseball bedding",
        "cluster": "custom name pitcher silhouette baseball bedding design 12",
        "detail": "cream baseball bedding with black pitcher silhouette, red baseball stitch border and sample name at the foot",
        "intent_role": "Baseball bedding page supported by a required name field and separated by pitcher silhouette artwork.",
    },
    188: {
        "title": "Custom Black Baseball Glove Bedding",
        "meta_title": "Custom Black Baseball Glove Bedding",
        "meta_description": "Shop custom black baseball glove bedding with ball artwork, required name field, sample number artwork and feature panels.",
        "primary": "custom black baseball glove bedding",
        "secondary": "black baseball bedding, baseball glove comforter, personalized sports bedding",
        "cluster": "custom name black baseball glove bedding design 13",
        "detail": "black baseball bedding with grayscale glove and ball artwork, sample cursive name and sample number",
        "intent_role": "Baseball bedding page supported by a required name field and separated by black monochrome glove design; visible numbers are sample artwork only.",
    },
    189: {
        "title": "Custom Blue Stripe Baseball Bedding",
        "meta_title": "Custom Blue Stripe Baseball Bedding",
        "meta_description": "Shop custom blue stripe baseball bedding with oversized baseball artwork, required name field and feature panels.",
        "primary": "custom blue stripe baseball bedding",
        "secondary": "blue baseball bedding, personalized baseball comforter, custom name baseball bedding",
        "cluster": "custom name blue striped baseball bedding design 14",
        "detail": "light blue striped bedding with oversized baseball graphic, baseball pillow pattern and sample cursive name",
        "intent_role": "Baseball bedding page supported by a required name field and separated by blue striped baseball layout.",
    },
    190: {
        "title": "Custom Galaxy Baseball Name Bedding",
        "meta_title": "Custom Galaxy Baseball Name Bedding",
        "meta_description": "Shop custom galaxy baseball bedding with blue space background, baseball artwork, required name field and sample number artwork.",
        "primary": "custom galaxy baseball name bedding",
        "secondary": "galaxy baseball bedding, baseball comforter, personalized baseball bedding",
        "cluster": "custom name galaxy baseball bedding design 15",
        "detail": "dark blue galaxy baseball bedding with space background, large baseball, sample cursive name and sample number",
        "intent_role": "Baseball bedding page supported by a required name field and separated by galaxy background artwork; visible numbers are sample artwork only.",
    },
}


COMMON_WEAVING = ("High-density weaving panel comparing BeddingOutlet fabric with other weave samples.", "High-density weaving feature panel")
COMMON_WASH = ("Machine washable care panel with washing machine, laundry basket and care text.", "Machine washable bedding care panel")
COMMON_ZIPPER = ("Bottom zippered closure panel showing white zipper close-up.", "Bottom zippered closure panel")


IMAGE_DETAILS = {
    181: {
        1: ("Bedroom mockup of custom brown baseball glove bedding with sample name and number.", "Custom baseball glove name bedding"),
        2: COMMON_WEAVING,
        3: COMMON_WASH,
        4: COMMON_ZIPPER,
    },
    182: {
        1: ("Bedroom mockup of custom baseball flag bedding with large baseball and sample cursive name.", "Custom baseball flag name bedding"),
        2: COMMON_WEAVING,
        3: COMMON_WASH,
        4: COMMON_ZIPPER,
    },
    183: {
        1: ("Bedroom mockup of custom baseball flag bedding with glove corners, sample name and large number.", "Custom baseball name number bedding"),
        2: COMMON_WEAVING,
        3: COMMON_WASH,
        4: COMMON_ZIPPER,
    },
    184: {
        1: ("Bedroom mockup of custom close-up baseball glove bedding with sample name on the ball.", "Custom close-up baseball glove bedding"),
        2: COMMON_WEAVING,
        3: COMMON_WASH,
        4: COMMON_ZIPPER,
    },
    185: {
        1: ("Bedroom mockup of custom gray baseball flag bedding with glove, ball and sample name.", "Custom gray baseball flag bedding"),
        2: COMMON_WEAVING,
        3: COMMON_WASH,
        4: COMMON_ZIPPER,
    },
    186: {
        1: ("Bedroom mockup of custom baseball pattern bedding with red blue stripes, many baseballs and sample name.", "Custom baseball pattern name bedding"),
        2: COMMON_WASH,
        3: COMMON_ZIPPER,
        4: COMMON_WEAVING,
    },
    187: {
        1: ("Bedroom mockup of custom pitcher baseball bedding with silhouette, stitch border and sample name.", "Custom pitcher baseball bedding"),
        2: COMMON_WEAVING,
        3: COMMON_WASH,
        4: COMMON_ZIPPER,
    },
    188: {
        1: ("Bedroom mockup of custom black baseball glove bedding with grayscale ball, sample name and number.", "Custom black baseball glove bedding"),
        2: COMMON_ZIPPER,
        3: COMMON_WEAVING,
        4: COMMON_WASH,
    },
    189: {
        1: ("Bedroom mockup of custom blue stripe baseball bedding with oversized baseball and sample name.", "Custom blue stripe baseball bedding"),
        2: COMMON_WASH,
        3: COMMON_ZIPPER,
        4: COMMON_WEAVING,
    },
    190: {
        1: ("Bedroom mockup of custom galaxy baseball bedding with large baseball, sample name and number.", "Custom galaxy baseball name bedding"),
        2: COMMON_WEAVING,
        3: COMMON_WASH,
        4: COMMON_ZIPPER,
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
    wb = load_workbook(SOURCE, data_only=False)
    ws = wb["SEO_Products"]
    idx = headers(ws)
    products = []
    for row_num in range(2, ws.max_row + 1):
        evidence_id = text(ws.cell(row_num, idx["evidence_id"]).value)
        match = re.search(r"_(\d{3})$", evidence_id)
        if not match:
            continue
        pos = int(match.group(1))
        if pos in PRODUCT_UPDATES:
            products.append({
                "product_key": text(ws.cell(row_num, idx["product_key"]).value),
                "handle": text(ws.cell(row_num, idx["Handle"]).value),
                "inventory_position": pos,
            })
    products.sort(key=lambda item: item["inventory_position"])
    return {
        "product_keys": [item["product_key"] for item in products],
        "handle_by_key": {item["product_key"]: item["handle"] for item in products},
        "pos_by_key": {item["product_key"]: int(item["inventory_position"]) for item in products},
        "customizer_by_pos": {},
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


def customizer_sentence(pos):
    if pos in {181, 183, 186, 188, 190}:
        return (
            'Customize Your Name is required and supports 1 to 30 characters; enter "NO" if no custom name is wanted. '
            "Numbers shown in mockups are fixed sample artwork, not a separate number field."
        )
    return 'Customize Your Name is required and supports 1 to 30 characters; enter "NO" if no custom name is wanted.'


def description(pos, admin_row, ctext):
    item = PRODUCT_UPDATES[pos]
    number_note = (
        "<li>Numbers shown in the product mockups are fixed sample artwork; the available custom input is the name field.</li>"
        if pos in {181, 183, 186, 188, 190}
        else ""
    )
    return (
        f"<p>The {item['title']} features {item['detail']} for a personalized baseball bedding design.</p>"
        "<h3>Personalization and Product Type</h3>"
        f"<ul><li>{ctext}</li>"
        f"{number_note}"
        "<li>Choose Comforter or Duvet Cover where the product-type selector is shown.</li>"
        "<li>The Duvet Cover option includes a bottom zippered closure shown in the gallery feature panel.</li></ul>"
        "<h3>Sizes, Add-ons and Care</h3>"
        "<ul><li>Available bed sizes are selected on the product page, including Twin, Full, Queen and King options where shown.</li>"
        "<li>Choose None, 1 Pillowcase or 2 Pillowcases where the pillowcase selector is shown.</li>"
        "<li>An optional matching flat sheet is available where the sheet-cover selector is shown.</li>"
        "<li>Feature panels show high-density woven fabric, machine-washable care and zipper detail.</li>"
        "<li>Machine wash cold on gentle cycle and tumble dry low.</li></ul>"
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
        ctext = customizer_sentence(pos)
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
        ws.cell(row_num, idx["revision"]).value = "r4"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields"]:
            ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            f"R4 after qa_batch_019_r3 QA: used products_export_1.csv sha256:{admin_hash}; "
            "removed process-style description copy, restored Comforter/Duvet Cover, zipper, pillowcase and flat-sheet details, "
            "kept only verified name customization and treated visible numbers as fixed sample artwork. Still NEEDS_REVIEW."
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
        ws.cell(row_num, idx["revision"]).value = "r4"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields", "approved_revision"]:
            if col_name in idx:
                ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            "R4: observation and alt carried forward from r3 checked image audit; admin image URL and current alt matched from Shopify CSV where available."
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
            "qa_batch_019_r3 QA report + products_export_1.csv admin baseline + r4 description/specification cleanup"
        )
        ws.cell(row_num, idx["checked_at"]).value = now
        ws.cell(row_num, idx["possible_overlap_with_other_products"]).value = item["intent_role"]
        ws.cell(row_num, idx["mapping_reason"]).value = (
            "R4 separates nearby baseball pages by motif, layout, color palette and verified custom-name field."
        )
        ws.cell(row_num, idx["mapping_status"]).value = "CANDIDATE_MAPPED_NEEDS_RECHECK"
        ws.cell(row_num, idx["mapping_version"]).value = "r4"
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
            f"When shopping for {item['cluster']}, the buyer wants the page to confirm exact baseball artwork, product type, size choices, "
            "and the required name field before purchase."
        )
        ws.cell(row_num, idx["functional_motivation"]).value = (
            "Confirm product type, visible baseball motif, required custom name field, care panels and image accuracy before purchase."
        )
        ws.cell(row_num, idx["emotional_social_motivation"]).value = (
            "Choose a personalized baseball bedding item that matches a preferred visual style, such as flag, glove, pitcher or galaxy artwork."
        )
        ws.cell(row_num, idx["purchase_concerns"]).value = (
            "Confirm spelling before checkout; avoid unsupported team, player, league, material, custom-number or delivery claims."
        )
        ws.cell(row_num, idx["source_refs"]).value = (
            f"products_export_1.csv sha256:{admin_hash}; QA report {QA_RUN_ID}; checked images 181-190"
        )
        ws.cell(row_num, idx["observed_at"]).value = now
        ws.cell(row_num, idx["research_status"]).value = "REVISED_R4_NEEDS_RECHECK"
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
        ctext = customizer_sentence(pos)
        ws.cell(row_num, idx["sources_accessed"]).value = (
            f"Original storefront/product evidence; QA report {QA_RUN_ID}; products_export_1.csv sha256:{admin_hash}; "
            "checked image audit"
        )
        ws.cell(row_num, idx["current_H1"]).value = admin_row["Title"]
        ws.cell(row_num, idx["current_meta_title"]).value = admin_row["SEO Title"] or "EMPTY"
        ws.cell(row_num, idx["verified_product_facts"]).value = (
            f"Admin export matched handle {handle}; type={admin_row['Type']}; options={option_text(admin_row)}; "
            f"status={admin_row['Status']}; image_count={len(admin_row['images'])}; design={item['detail']}; "
            f"customizer={ctext}"
        )
        ws.cell(row_num, idx["factual_conflicts"]).value = "qa_batch_019_r3 QA description issues addressed in r4; independent QA recheck still required."
        ws.cell(row_num, idx["confidence_and_reason"]).value = (
            "R4 uses checked image audit, QA report recommendations and Shopify admin CSV baseline; no approval or import file created."
        )

    ws = wb["README_QA"]
    idx = headers(ws)
    for metric, value, definition in [
        ("qa_batch_019_r4_revision", "r4", "Separate 10-product revision after qa_batch_019_r3 QA; source workbook was not modified."),
        ("qa_batch_019_r4_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_019_r4_scope", "inventory positions 181-190", "No products outside qa_batch_019 were revised."),
        ("qa_batch_019_r4_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_019_r4_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_019_r4_admin_alt_matches", str(admin_alt_matches), "Image rows with current admin alt matched by image URL from Shopify CSV."),
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
        "source_qa_report": str(QA_REPORT.relative_to(ROOT)),
        "admin_export": str(ADMIN_EXPORT.relative_to(ROOT)),
        "admin_export_sha256": admin_hash,
        "revision_workbook": str(OUTPUT.relative_to(ROOT)),
        "scope": "qa_batch_019 only; inventory positions 181-190",
        "revision": "r4",
        "revised_products": revised_products,
        "revised_images": revised_images,
        "admin_alt_matches": admin_alt_matches,
        "keyword_rows_revised": keyword_rows,
        "review_status": "NEEDS_REVIEW",
        "approved_created": False,
        "shopify_import_created": False,
        "r4_changes": [
            "Rewrote all 10 product descriptions to remove process/meta commentary flagged by qa_batch_019_r3 QA.",
            "Added Comforter vs Duvet Cover explanation, including bottom zipper detail for Duvet Cover.",
            "Added pillowcase and flat-sheet option details from the QA storefront audit.",
            "Clarified that only Customize Your Name is available and visible numbers are fixed sample artwork."
        ],
        "next_step": "Sang QA lại qa_batch_019_r4 before any approval or import.",
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text(
        "\n".join(
            [
                "# Revision qa_batch_019_r4",
                "",
                "Artifact chính thức theo quy trình từng lô 10 sản phẩm.",
                "",
                f"- Nguồn r3: `{SOURCE.relative_to(ROOT)}`",
                f"- QA của Sang: `{QA_REPORT.relative_to(ROOT)}`",
                f"- Shopify admin export: `{ADMIN_EXPORT.relative_to(ROOT)}`",
                f"- Admin export SHA-256: `{admin_hash}`",
                f"- Workbook r4: `{OUTPUT.relative_to(ROOT)}`",
                f"- Sản phẩm sửa: {revised_products}",
                f"- Ảnh sửa/refresh: {revised_images}",
                f"- Admin image alt match: {admin_alt_matches}/{revised_images}",
                "- Sửa trọng yếu: xử lý lỗi D2 còn lại bằng cách bỏ câu meta/process trong mô tả, thêm Comforter/Duvet Cover, zipper, pillowcase/flat sheet và làm rõ chỉ có custom name, không có custom number.",
                "- Trạng thái: `NEEDS_REVIEW`; không tạo `APPROVED`; không tạo file import Shopify.",
                "- Bước tiếp theo: gửi Sang QA lại `qa_batch_019_r4` trước khi duyệt hay import.",
                "",
            ]
        ),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
