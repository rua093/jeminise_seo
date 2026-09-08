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
BATCH_ID = "qa_batch_001_r5"
QA_RUN_ID = "20260908_001100"

SOURCE = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / "qa_batch_001_r4" / "SEO_Product_Optimization_qa_batch_001_r4.xlsx"
QA_REPORT = ROOT / "resutls" / SHOP / RUN_ID / "qa" / QA_RUN_ID / "SEO_QA_qa_batch_001_r4.md"
ADMIN_EXPORT = ROOT / "products_export_1.csv"
RESULT_DIR = ROOT / "resutls" / SHOP / RUN_ID / "revisions" / BATCH_ID
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID / "revisions" / BATCH_ID
OUTPUT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_001_r5.xlsx"
REPORT = RESULT_DIR / "SEO_Product_Optimization_qa_batch_001_r5.md"
SUMMARY = RUN_DIR / "revision_summary.json"


PRODUCT_UPDATES = {
    1: {
        "title": "Cardinal Sunflower Autumn Quilt Set",
        "meta_title": "Cardinal Sunflower Autumn Quilt Set",
        "meta_description": "Bring fall color to the bedroom with red cardinals, large sunflowers, autumn leaves, size choices and optional shams.",
        "primary": "cardinal sunflower autumn quilt set",
        "secondary": "autumn cardinal quilt, sunflower cardinal bedding, cardinal bird quilt set",
        "cluster": "cardinal autumn floral quilt",
        "intent_role": "Product page for autumn cardinal sunflower quilt; link from cardinal quilt collection with anchor 'cardinal sunflower autumn quilt'; avoid Christmas cardinal memorial or rose cardinal remembrance targets.",
        "detail": "red cardinals perched above large yellow sunflowers with autumn leaves",
        "options": "Choose quilt size; pillowcases are optional according to the current Shopify export.",
    },
    2: {
        "title": "Personalized Teal Softball Comforter Set",
        "meta_title": "Personalized Teal Softball Comforter Set",
        "meta_description": "Customize a teal softball comforter with name and number fields, coordinated pillowcases and sporty bedroom artwork.",
        "primary": "personalized softball comforter set",
        "secondary": "teal softball bedding set, custom softball comforter, softball bedding with name and number",
        "cluster": "personalized softball comforter",
        "intent_role": "Product page for personalized teal softball comforter; link from softball bedding collection with anchor 'personalized softball comforter'; avoid generic quilt targets.",
        "detail": "teal, black and white softball patchwork with player silhouettes, number artwork and Eat Sleep Softball text",
        "options": "Choose comforter size and pillowcases. Live Customizer shows required Custom Your Name and Custom Your Number text fields, with instructions to enter NO if not wanted.",
    },
    3: {
        "title": "Christmas Cardinal Memorial Quilt Set",
        "meta_title": "Christmas Cardinal Memorial Quilt Set",
        "meta_description": "Create a comforting holiday bed with red cardinals, a decorated Christmas tree, memorial text, size choices and optional shams.",
        "primary": "Christmas cardinal memorial quilt set",
        "secondary": "cardinal Christmas quilt bedding, I am always with you quilt, cardinal remembrance bedding",
        "cluster": "Christmas cardinal memorial tree quilt",
        "intent_role": "Product page for Christmas cardinal memorial quilt beside a decorated tree; link from Christmas cardinal collection; avoid rose cardinal remembrance and generic cardinal quilt targets.",
        "detail": "red cardinals beside a decorated Christmas tree with I Am Always with You memorial text",
        "options": "Choose quilt size; pillowcases are optional according to the current Shopify export.",
    },
    4: {
        "title": "Cardinal Roses Memorial Quilt Set",
        "meta_title": "Cardinal Roses Memorial Quilt Set",
        "meta_description": "Choose a memorial-style cardinal quilt with a cardinal pair, red roses, remembrance text, size choices and optional shams.",
        "primary": "cardinal roses memorial quilt set",
        "secondary": "cardinal remembrance quilt, cardinal rose bedding, I am always with you cardinal quilt",
        "cluster": "cardinal roses remembrance quilt",
        "intent_role": "Product page for cardinal roses remembrance quilt; link from memorial cardinal collection; avoid Christmas tree cardinal memorial targets.",
        "detail": "a male and female cardinal pair on a branch with red roses and I Am Always with You text",
        "options": "Choose quilt size; pillowcases are optional according to the current Shopify export.",
    },
    5: {
        "title": "Colorful Cat Patchwork Quilt Set",
        "meta_title": "Colorful Cat Patchwork Quilt Set",
        "meta_description": "Style a cat lover's room with a colorful cat face patchwork quilt, feature graphics, size choices and optional pillowcases.",
        "primary": "colorful cat patchwork quilt set",
        "secondary": "cat patchwork bedding set, colorful cat quilt, cat face quilt set",
        "cluster": "colorful cat face patchwork quilt",
        "intent_role": "Product page for colorful cat-face patchwork quilt; link from cat quilt collection; avoid geometric sitting-cat and chicken/farmhouse anchors.",
        "detail": "a large colorful cat face built from bright patchwork blocks",
        "options": "Choose quilt size; pillowcases are purchased separately according to the current Shopify export.",
    },
    6: {
        "title": "Geometric Cat Patchwork Quilt Set",
        "meta_title": "Geometric Cat Patchwork Quilt Set",
        "meta_description": "Add a calmer cat design with a geometric sitting cat quilt, warm patchwork colors, size choices and optional pillowcases.",
        "primary": "geometric cat patchwork quilt set",
        "secondary": "cat patchwork bedding set, sitting cat quilt, cat quilt bedding",
        "cluster": "geometric sitting cat quilt",
        "intent_role": "Product page for geometric sitting-cat patchwork quilt; link from cat quilt collection; avoid colorful cat-face anchor.",
        "detail": "a sitting cat in geometric patchwork colors on a beige leaf background",
        "options": "Choose quilt size; pillowcases are purchased separately according to the current Shopify export.",
    },
    7: {
        "title": "Celtic Fantasy Tree Quilt Set",
        "meta_title": "Celtic Fantasy Tree Quilt Set",
        "meta_description": "Add fantasy woodland style with a twisting Celtic tree quilt, exposed roots, knotwork border, size choices and pillowcase options.",
        "primary": "Celtic fantasy tree quilt set",
        "secondary": "fantasy tree quilt bedding, Celtic tree quilt, knotwork tree bedding",
        "cluster": "Celtic fantasy tree quilt",
        "intent_role": "Product page for twisting Celtic fantasy tree quilt; link from Celtic tree quilt collection; avoid green Tree of Life medallion anchor.",
        "detail": "a twisting tree trunk with exposed roots, teal night sky accents and a Celtic knot border",
        "options": "Choose quilt size and pillowcase quantity according to the current Shopify export.",
    },
    8: {
        "title": "Celtic Tree of Life Quilt Set",
        "meta_title": "Celtic Tree of Life Quilt Set",
        "meta_description": "Give the bedroom a Celtic focal point with a green Tree of Life medallion quilt, knotwork border, sizes and optional shams.",
        "primary": "Celtic Tree of Life quilt set",
        "secondary": "Tree of Life bedding set, Celtic knotwork quilt, green tree quilt set",
        "cluster": "green Celtic Tree of Life quilt",
        "intent_role": "Product page for green Celtic Tree of Life medallion quilt; link from Tree of Life collection; avoid twisting fantasy tree anchor.",
        "detail": "a green Tree of Life medallion with Celtic knotwork border",
        "options": "Choose quilt size; pillowcases are optional according to the current Shopify export.",
    },
    9: {
        "title": "Farmhouse Chicken Patchwork Quilt Set",
        "meta_title": "Farmhouse Chicken Patchwork Quilt Set",
        "meta_description": "Add country charm with a farmhouse chicken quilt featuring hens, nest-and-egg artwork, patchwork accents and optional shams.",
        "primary": "farmhouse chicken patchwork quilt set",
        "secondary": "chicken bedding set, country chicken quilt, hen patchwork quilt",
        "cluster": "farmhouse chicken patchwork quilt",
        "intent_role": "Product page for farmhouse chicken and hen patchwork quilt; link from farmhouse or chicken bedding collection; avoid cat quilt anchors.",
        "detail": "farmhouse chicken and hen artwork with nest-and-egg details and black patchwork accents",
        "options": "Choose quilt size; pillowcases are purchased separately according to the current Shopify export.",
    },
    10: {
        "title": "Personalized God Says I Am Christian Bedding Set",
        "meta_title": "Personalized God Says I Am Christian Bedding Set",
        "meta_description": "Customize a God Says I Am Christian bedding set with an Enter Name field, birth-month flowers, size choices and pillowcase options.",
        "primary": "personalized God Says I Am Christian bedding set",
        "secondary": "Christian comforter set, God Says I Am bedding, personalized Bible verse bedding",
        "cluster": "personalized Christian affirmation bedding",
        "intent_role": "Product page for personalized God Says I Am Christian bedding; link from Christian bedding collection; only claim the verified Enter Name field up to 13 characters.",
        "detail": "God Says I Am Christian affirmation artwork with a custom name example, flowers and butterflies",
        "options": "Choose product type and size, pillowcases and optional flat sheet. Live Customizer shows a required Enter Name field with a 13-character limit.",
    },
}


IMAGE_DETAILS = {
    1: [
        ("Red cardinals perched above large yellow sunflowers on an autumn quilt bed mockup with matching shams.", "Cardinal sunflower autumn quilt set on bed"),
        ("Close-up of a red cardinal, sunflower petals, orange leaves and visible quilt stitching.", "Close-up of cardinal sunflower quilt stitching"),
        ("Matching pillow sham with two cardinals, sunflower artwork and autumn leaf accents.", "Cardinal sunflower pillow sham with autumn leaves"),
        ("Bedroom scene showing the cardinal sunflower quilt styled across a bed with coordinated pillows.", "Cardinal sunflower autumn quilt bedroom mockup"),
        ("Size and component panel showing quilt size choices and two optional standard shams.", "Cardinal sunflower quilt size and sham options"),
    ],
    2: [
        ("Teal, black and white softball comforter bed mockup with player silhouettes, Riley name, number 88 and Eat Sleep Softball artwork.", "Personalized teal softball comforter set"),
        ("Angled bedroom mockup of the teal softball comforter with coordinated pillowcases, Riley name and number 88.", "Teal softball comforter bedroom mockup"),
        ("Coordinated pillowcases showing softball player silhouettes and number 88 on teal patchwork artwork.", "Teal softball pillowcases with number 88"),
        ("Folded teal softball comforter showing patchwork print, softball graphics and white backing.", "Folded teal softball comforter with white backing"),
        ("Close-up of folded comforter edge with teal softball print on top and white backing underneath.", "Teal softball comforter white backing close-up"),
        ("Feature panel describing microfiber, cozy filling, lightweight feel and breathable warmth for the softball comforter.", "Softball comforter microfiber feature panel"),
        ("Included-items panel showing one softball comforter and two coordinated pillowcases.", "Softball comforter and two pillowcases included"),
    ],
    3: [
        ("Christmas memorial quilt bed mockup with red cardinals beside a decorated tree, ornaments, hearts and I Am Always with You text.", "Christmas cardinal memorial quilt set on bed"),
        ("Close-up of red cardinal, Christmas tree ornaments, snowflake details and quilt stitching.", "Close-up of Christmas cardinal quilt stitching"),
        ("Matching pillow sham with red cardinal, decorated Christmas tree and remembrance message.", "Christmas cardinal memorial pillow sham"),
        ("Bedroom scene showing the Christmas cardinal memorial quilt with coordinated pillow shams.", "Christmas cardinal memorial quilt bedroom mockup"),
        ("Size and component panel showing Christmas cardinal quilt sizes and two optional standard shams.", "Christmas cardinal quilt size and sham options"),
    ],
    4: [
        ("Memorial quilt bed mockup with a cardinal pair on a branch, red roses, butterflies and I Am Always with You text.", "Cardinal roses memorial quilt set on bed"),
        ("Close-up of the cardinal pair, branch, rose artwork and visible quilt stitching.", "Close-up of cardinal roses quilt stitching"),
        ("Matching pillow sham with cardinal pair, red roses and remembrance artwork.", "Cardinal roses memorial pillow sham"),
        ("Bedroom scene showing the cardinal roses memorial quilt styled with coordinated pillow shams.", "Cardinal roses memorial quilt bedroom mockup"),
        ("Size and component panel showing cardinal roses quilt sizes and two optional standard shams.", "Cardinal roses quilt size and sham options"),
    ],
    5: [
        ("Bed mockup with a large colorful cat face made from bright patchwork blocks on a dark background.", "Colorful cat face patchwork quilt set"),
        ("Bedroom mockup with colorful cat patchwork quilt and premium printed craft callout.", "Colorful cat patchwork quilt bedroom mockup"),
        ("Optional pillow shams panel showing coordinated colorful cat face sham artwork.", "Colorful cat patchwork optional pillow shams"),
        ("Fabric feature panel showing soft breathable skin-friendly material and all-season use notes.", "Colorful cat quilt fabric feature panel"),
        ("Premium quilt set size chart for the colorful cat patchwork design.", "Colorful cat patchwork quilt size chart"),
        ("Bedspread feature panel with microfiber layers and care-resistance icons.", "Colorful cat quilt bedspread feature panel"),
        ("Second bedroom mockup showing the colorful cat face quilt from a different room angle.", "Colorful cat quilt second bedroom mockup"),
    ],
    6: [
        ("Bed mockup with a geometric sitting cat in warm orange and red patchwork on a beige leaf background.", "Geometric sitting cat patchwork quilt set"),
        ("Bedroom mockup with geometric cat patchwork quilt and premium printed craft callout.", "Geometric cat patchwork quilt bedroom mockup"),
        ("Optional pillow shams panel showing coordinated geometric cat sham artwork.", "Geometric cat patchwork optional pillow shams"),
        ("Fabric feature panel showing soft breathable skin-friendly material and all-season use notes.", "Geometric cat quilt fabric feature panel"),
        ("Premium quilt set size chart for the geometric cat patchwork design.", "Geometric cat patchwork quilt size chart"),
        ("Bedspread feature panel with microfiber layers and care-resistance icons.", "Geometric cat quilt bedspread feature panel"),
        ("Second bedroom mockup showing the geometric cat quilt from a different room angle.", "Geometric cat quilt second bedroom mockup"),
    ],
    7: [
        ("Bed mockup with a twisting fantasy tree trunk, exposed roots, teal landscape accents and Celtic knot border.", "Celtic fantasy tree quilt set on bed"),
        ("Bedroom mockup showing the twisting Celtic fantasy tree quilt with coordinated pillow shams.", "Celtic fantasy tree quilt bedroom mockup"),
        ("Optional pillow shams panel showing coordinated Celtic fantasy tree artwork.", "Celtic fantasy tree optional pillow shams"),
        ("Fabric feature panel showing soft breathable skin-friendly material and all-season use notes.", "Celtic fantasy tree quilt fabric feature panel"),
        ("Premium quilt set size chart for the Celtic fantasy tree design.", "Celtic fantasy tree quilt size chart"),
        ("Bedspread feature panel with microfiber layers and care-resistance icons.", "Celtic fantasy tree bedspread feature panel"),
        ("Second bedroom mockup showing the Celtic fantasy tree quilt from a different room angle.", "Celtic fantasy tree second bedroom mockup"),
    ],
    8: [
        ("Bed mockup with a green Tree of Life medallion and circular Celtic knotwork border.", "Celtic Tree of Life quilt set on bed"),
        ("Bedroom mockup showing the green Celtic Tree of Life quilt with coordinated pillow shams.", "Celtic Tree of Life quilt bedroom mockup"),
        ("Optional pillow shams panel showing coordinated green Tree of Life artwork.", "Celtic Tree of Life optional pillow shams"),
        ("Fabric feature panel showing soft breathable skin-friendly material and all-season use notes.", "Celtic Tree of Life quilt fabric feature panel"),
        ("Premium quilt set size chart for the green Celtic Tree of Life design.", "Celtic Tree of Life quilt size chart"),
        ("Bedspread feature panel with microfiber layers and care-resistance icons.", "Celtic Tree of Life bedspread feature panel"),
        ("Second bedroom mockup showing the green Celtic Tree of Life quilt from a different room angle.", "Celtic Tree of Life second bedroom mockup"),
    ],
    9: [
        ("Bed mockup with farmhouse chicken, hen, rooster, nest and egg artwork with black patchwork accents.", "Farmhouse chicken patchwork quilt set"),
        ("Bedroom mockup with farmhouse chicken patchwork quilt and premium printed craft callout.", "Farmhouse chicken quilt bedroom mockup"),
        ("Optional pillow shams panel showing coordinated chicken and hen patchwork artwork.", "Farmhouse chicken optional pillow shams"),
        ("Fabric feature panel showing soft breathable skin-friendly material and all-season use notes.", "Farmhouse chicken quilt fabric feature panel"),
        ("Premium quilt set size chart for the farmhouse chicken patchwork design.", "Farmhouse chicken quilt size chart"),
        ("Bedspread feature panel with microfiber layers and care-resistance icons.", "Farmhouse chicken bedspread feature panel"),
        ("Second bedroom mockup showing the farmhouse chicken quilt from a different room angle.", "Farmhouse chicken second bedroom mockup"),
    ],
    10: [
        ("Personalized Christian bedding mockup with Jessica name, God Says I Am affirmation text, Bible verse labels, flowers and butterflies.", "Personalized God Says I Am Christian bedding set"),
        ("Bedroom mockup of the Christian bedding with Jessica pillowcases, purple affirmation blocks and floral artwork.", "God Says I Am Christian bedding bedroom mockup"),
        ("Product type comparison panel showing duvet cover set and comforter set options for all-season use.", "Christian bedding duvet cover and comforter options"),
        ("Feature panel showing microfiber callout, 3D printed pattern close-ups, soft filling and lightweight breathable notes.", "Christian bedding microfiber feature panel"),
        ("Birth month flowers panel showing flower illustrations for January through December.", "Christian bedding birth month flowers panel"),
        ("Close bedroom view showing the personalized Christian bedding with soft, lightweight, durable and breathable icons.", "Personalized Christian bedding feature icons"),
        ("Stress-free easy care panel listing wrinkle-free, stain-proof, anti-pilling and wash-fade notes.", "Christian bedding easy care feature panel"),
        ("Size dimension panel showing Twin, Full, Queen and King bedding dimensions.", "Christian bedding size dimension chart"),
    ],
}

DESCRIPTION_HTML = {
    1: (
        "<p>The Cardinal Sunflower Autumn Quilt Set brings a warm fall look to the bedroom with red cardinals, oversized yellow sunflowers and orange leaf accents across the quilt and matching sham artwork.</p>"
        "<h3>Design Details</h3><ul><li>Red cardinal birds sit above large sunflower artwork on an autumn leaf background.</li><li>The gallery includes a bed mockup, close-up quilt stitching, matching pillow sham artwork, a bedroom view and a size/component panel.</li></ul>"
        "<h3>Sizes and Set Details</h3><ul><li>Shown size choices include Throw 60x70, Twin 70x80, Queen 80x90 and King 91x102.</li><li>The panel shows one premium quilt with two standard shams available as optional add-ons.</li><li>The quilt panel presents a soft, breathable fabric surface suitable for year-round bedroom use.</li></ul>"
    ),
    2: (
        "<p>The Personalized Teal Softball Comforter Set pairs teal, black and white patchwork with softball graphics, player silhouettes, number artwork and the phrase Eat Sleep Softball.</p>"
        "<h3>Design Details</h3><ul><li>The main artwork shows softball player silhouettes, ball graphics, teal patchwork blocks and coordinated pillowcases.</li><li>Gallery images show the bed view, angled room view, pillowcases, folded comforter, fabric close-up, feature panel and included-items panel.</li></ul>"
        "<h3>Options and Set Details</h3><ul><li>Select a comforter size and pillowcase option before checkout.</li><li>The personalization fields request a name and number; enter NO where the product page allows if a field should not appear.</li><li>The set panel shows one comforter with two pillowcases, while the feature panel notes machine washing, soft filling and a lightweight breathable feel.</li></ul>"
    ),
    3: (
        "<p>The Christmas Cardinal Memorial Quilt Set features red cardinals beside a decorated Christmas tree with hearts, ornaments, snowflakes and I Am Always With You remembrance text.</p>"
        "<h3>Design Details</h3><ul><li>The artwork centers on a holiday tree scene with cardinals and memorial wording, making this page distinct from rose or branch cardinal designs.</li><li>The gallery includes a bed mockup, close-up quilt stitching, pillow sham artwork, room view and set-size panel.</li></ul>"
        "<h3>Sizes and Set Details</h3><ul><li>Shown size choices include Throw 60x70, Twin 70x80, Queen 80x90 and King 91x102.</li><li>The panel shows one premium quilt with two standard shams available as optional add-ons.</li><li>The design is best positioned for Christmas cardinal, remembrance and memorial bedding searches.</li></ul>"
    ),
    4: (
        "<p>The Cardinal Roses Memorial Quilt Set uses a pair of cardinals on a branch with red roses, butterflies and I Am Always With You text for a softer remembrance bedding theme.</p>"
        "<h3>Design Details</h3><ul><li>The rose branch artwork separates this quilt from Christmas tree cardinal and general winter cardinal pages.</li><li>The gallery includes the full bed mockup, close-up quilt texture, pillow sham, bedroom scene and size/component panel.</li></ul>"
        "<h3>Sizes and Set Details</h3><ul><li>Shown size choices include Throw 60x70, Twin 70x80, Queen 80x90 and King 91x102.</li><li>The panel shows one premium quilt with two standard shams available as optional add-ons.</li><li>Use this product page for cardinal rose remembrance intent rather than broader Christmas cardinal queries.</li></ul>"
    ),
    5: (
        "<p>The Colorful Cat Patchwork Quilt Set turns a large cat face into bright patchwork blocks on a dark background, creating a bold animal bedding piece for cat-themed rooms.</p>"
        "<h3>Design Details</h3><ul><li>The artwork focuses on a front-facing colorful cat face with multicolor patchwork panels.</li><li>The gallery includes room mockups, optional pillow sham artwork, fabric panel, size chart, feature panel and a second bedroom view.</li></ul>"
        "<h3>Sizes and Features</h3><ul><li>The size chart shows Throw 60x70, Twin 68x86, Full 80x90, Queen 90x90 and King 102x91.</li><li>Pillow shams are shown as optional add-ons.</li><li>The feature panels mention a soft breathable fabric feel, microfiber layers, anti-wrinkle, anti-static, fade-resistant, shrink-resistant and machine-washable care icons.</li></ul>"
    ),
    6: (
        "<p>The Geometric Cat Patchwork Quilt Set has a calmer sitting-cat design with warm orange, red and beige patchwork shapes over a leafy background.</p>"
        "<h3>Design Details</h3><ul><li>The sitting cat silhouette and geometric patchwork color blocks make this design separate from the colorful cat-face quilt.</li><li>The gallery includes room mockups, coordinated sham artwork, fabric panel, size chart, feature panel and a second bedroom view.</li></ul>"
        "<h3>Sizes and Features</h3><ul><li>The size chart shows Throw 60x70, Twin 68x86, Full 80x90, Queen 90x90 and King 102x91.</li><li>Pillow shams are shown as optional add-ons.</li><li>The feature panels mention a soft breathable fabric feel, microfiber layers, anti-wrinkle, anti-static, fade-resistant, shrink-resistant and machine-washable care icons.</li></ul>"
    ),
    7: (
        "<p>The Celtic Fantasy Tree Quilt Set features a twisting tree trunk, exposed roots, teal night-sky accents and a Celtic knot border for a fantasy woodland bedding look.</p>"
        "<h3>Design Details</h3><ul><li>The twisting trunk and exposed-root artwork distinguish this quilt from the round green Tree of Life medallion design.</li><li>The gallery includes bed and room mockups, optional sham artwork, fabric panel, size chart, feature panel and a second bedroom view.</li></ul>"
        "<h3>Sizes and Features</h3><ul><li>The size chart shows Throw 60x70, Twin 68x86, Full 80x90, Queen 90x90 and King 102x91.</li><li>Choose the quilt size and pillowcase quantity shown on the product page.</li><li>The feature panels mention microfiber layers, soft breathable fabric, anti-static, fade-resistant, shrink-resistant and machine-washable care icons.</li></ul>"
    ),
    8: (
        "<p>The Celtic Tree of Life Quilt Set centers on a green Tree of Life medallion framed by Celtic knotwork, giving the bed a symbolic nature-inspired focal point.</p>"
        "<h3>Design Details</h3><ul><li>The circular medallion composition separates this page from the twisting fantasy tree quilt.</li><li>The gallery includes bed and room mockups, optional sham artwork, fabric panel, size chart, feature panel and a second bedroom view.</li></ul>"
        "<h3>Sizes and Features</h3><ul><li>The size chart shows Throw 60x70, Twin 68x86, Full 80x90, Queen 90x90 and King 102x91.</li><li>Pillow shams are shown as optional add-ons.</li><li>The feature panels mention microfiber layers, soft breathable fabric, anti-wrinkle, anti-static, fade-resistant, shrink-resistant and machine-washable care icons.</li></ul>"
    ),
    9: (
        "<p>The Farmhouse Chicken Patchwork Quilt Set combines chicken, hen, rooster, nest and egg artwork with black patchwork accents for country-style animal bedding.</p>"
        "<h3>Design Details</h3><ul><li>The artwork focuses on farmhouse poultry scenes rather than general animal or cat patchwork themes.</li><li>The gallery includes bed and room mockups, optional sham artwork, fabric panel, size chart, feature panel and a second bedroom view.</li></ul>"
        "<h3>Sizes and Features</h3><ul><li>The size chart shows Throw 60x70, Twin 68x86, Full 80x90, Queen 90x90 and King 102x91.</li><li>Pillow shams are shown as optional add-ons.</li><li>The feature panels mention microfiber layers, soft breathable fabric, anti-wrinkle, anti-static, fade-resistant, shrink-resistant and machine-washable care icons.</li></ul>"
    ),
    10: (
        "<p>The Personalized God Says I Am Christian Bedding Set combines faith affirmation text with birth-month flowers, butterflies and a custom name area for a devotional bedroom theme.</p>"
        "<h3>Design Details</h3><ul><li>The artwork includes the God Says I Am affirmation layout, floral accents, butterflies and a name field shown in the mockups.</li><li>The gallery includes room views, product type comparison, microfiber feature panel, birth-month flower panel, care panel and size chart.</li></ul>"
        "<h3>Options and Set Details</h3><ul><li>Choose product type and size, pillowcases and optional flat sheet from the product selectors.</li><li>The name field is required and is limited to 13 characters.</li><li>The size chart shows Twin, Full, Queen and King dimensions; feature panels mention microfiber, soft filling, lightweight breathable feel and easy-care icons.</li></ul>"
    ),
}

CAT_SERP_URLS = "https://www.walmart.com/ip/20788770071; https://www.walmart.com/ip/20076702698"
CARDINAL_TREE_SERP = "https://www.etsy.com/market/christmas_cardinal_quilt; https://www.amazon.com/s?k=christmas+cardinal+quilt"
CARDINAL_ROSES_SERP = "https://www.etsy.com/market/cardinal_memorial_quilt; https://www.amazon.com/s?k=cardinal+memorial+quilt"


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


def safe_len(value):
    return len(text(value))


def make_description(pos, admin):
    return DESCRIPTION_HTML[pos]


def load_admin_export(handles):
    by_handle = defaultdict(list)
    with ADMIN_EXPORT.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["Handle"] in handles:
                by_handle[row["Handle"]].append(row)

    admin = {}
    image_by_handle_url = {}
    for handle, rows in by_handle.items():
        first = rows[0]
        images = {}
        for row in rows:
            src = row.get("Image Src", "")
            if not src:
                continue
            images[normalize_url(src)] = {
                "src": src,
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
            "rows": len(rows),
        }
        image_by_handle_url[handle] = images
    return admin, image_by_handle_url


def load_qa():
    wb = load_workbook(SOURCE, data_only=False)
    ws = wb["SEO_Products"]
    idx = headers(ws)
    products = []
    for row_num in range(2, ws.max_row + 1):
        pk = text(ws.cell(row_num, idx["product_key"]).value)
        handle = text(ws.cell(row_num, idx["Handle"]).value)
        evidence_id = text(ws.cell(row_num, idx["evidence_id"]).value)
        match = re.search(r"_(\d{3})$", evidence_id)
        if not pk or not handle or not match:
            continue
        pos = int(match.group(1))
        if pos in PRODUCT_UPDATES:
            products.append({"product_key": pk, "handle": handle, "inventory_position": pos})

    products.sort(key=lambda item: int(item["inventory_position"]))
    product_keys = [item["product_key"] for item in products]
    handle_by_key = {item["product_key"]: item["handle"] for item in products}
    pos_by_key = {item["product_key"]: int(item["inventory_position"]) for item in products}

    observation_by_key = {}
    image_number_by_media = {}
    handle_to_key = {v: k for k, v in handle_by_key.items()}
    ws = wb["Image_Audit"]
    idx = headers(ws)
    for row_num in range(2, ws.max_row + 1):
        handle = text(ws.cell(row_num, idx["Handle"]).value)
        pk = handle_to_key.get(handle)
        if not pk:
            continue
        media_id = text(ws.cell(row_num, idx["media_id"]).value)
        key = (pk, media_id)
        observation_by_key[key] = text(ws.cell(row_num, idx["observed_visual_details"]).value)
        image_number_by_media[key] = ws.cell(row_num, idx["image_number"]).value

    issue_refs_by_key = defaultdict(list)
    for pos in PRODUCT_UPDATES:
        issue_refs_by_key[product_keys[pos - 1]].append(f"R5-BATCH001-POS{pos:03d}")
    return {
        "product_keys": product_keys,
        "handle_by_key": handle_by_key,
        "pos_by_key": pos_by_key,
        "observation_by_key": observation_by_key,
        "image_number_by_media": image_number_by_media,
        "issue_refs_by_key": issue_refs_by_key,
    }


def compact_observation_to_alt(observation, fallback):
    value = re.sub(r"[:;].*$", "", text(observation)).strip()
    value = re.sub(r"\s+", " ", value)
    if not value:
        value = fallback
    return value[:125]


def update_workbook():
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE, OUTPUT)

    qa = load_qa()
    handles = set(qa["handle_by_key"].values())
    admin, image_admin = load_admin_export(handles)
    now = datetime.now().astimezone().isoformat()
    admin_hash = sha256(ADMIN_EXPORT)

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
        update = PRODUCT_UPDATES[pos]
        admin_row = admin[handle]

        ws.cell(row_num, idx["title_current"]).value = admin_row["Title"]
        ws.cell(row_num, idx["h1_current"]).value = admin_row["Title"]
        ws.cell(row_num, idx["rendered_title_current"]).value = admin_row["SEO Title"] or admin_row["Title"]
        ws.cell(row_num, idx["meta_description_current"]).value = admin_row["SEO Description"]
        ws.cell(row_num, idx["product_type"]).value = admin_row["Type"] or ws.cell(row_num, idx["product_type"]).value

        ws.cell(row_num, idx["primary_keyword"]).value = update["primary"]
        ws.cell(row_num, idx["secondary_keywords"]).value = update["secondary"]
        ws.cell(row_num, idx["keyword_strategy"]).value = update["intent_role"]
        ws.cell(row_num, idx["buyer_search_summary"]).value = (
            f"Buyer intent targets {update['cluster']} for US English product search; no search-volume claim is made."
        )
        ws.cell(row_num, idx["title_proposed"]).value = update["title"]
        ws.cell(row_num, idx["meta_title_seo"]).value = update["meta_title"]
        ws.cell(row_num, idx["meta_title_length"]).value = safe_len(update["meta_title"])
        ws.cell(row_num, idx["meta_description_seo"]).value = update["meta_description"]
        ws.cell(row_num, idx["meta_description_length"]).value = safe_len(update["meta_description"])
        ws.cell(row_num, idx["description_proposed_html"]).value = make_description(pos, admin_row)
        ws.cell(row_num, idx["meta_keyword"]).value = update["primary"]
        ws.cell(row_num, idx["revision"]).value = "r5"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields"]:
            if col_name in idx:
                ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            f"R5 after qa_batch_001_r4 QA: used products_export_1.csv (sha256:{admin_hash}) as admin baseline, "
            "rewrote product-specific customer HTML, documented collection/internal-link role, preserved verified customizer wording, and refreshed keyword/evidence mapping. "
            "Still NEEDS_REVIEW; no approval or import file created."
        )
        revised_products += 1

    ws = wb["Image_Audit"]
    idx = headers(ws)
    revised_images = 0
    admin_alt_matched = 0
    for row_num in range(2, ws.max_row + 1):
        handle = text(ws.cell(row_num, idx["Handle"]).value)
        pk = None
        for product_key, product_handle in qa["handle_by_key"].items():
            if product_handle == handle:
                pk = product_key
                break
        if pk not in product_key_set:
            continue
        media_id = text(ws.cell(row_num, idx["media_id"]).value)
        key = (pk, media_id)
        pos = qa["pos_by_key"][pk]
        update = PRODUCT_UPDATES[pos]
        image_url = text(ws.cell(row_num, idx["image_url_export"]).value or ws.cell(row_num, idx["image_url"]).value)
        admin_image = image_admin.get(handle, {}).get(normalize_url(image_url))

        if admin_image:
            ws.cell(row_num, idx["image_url_export"]).value = admin_image["src"]
            ws.cell(row_num, idx["alt_current"]).value = admin_image["alt"]
            admin_alt_matched += 1

        image_number = ws.cell(row_num, idx["image_number"]).value or qa["image_number_by_media"].get(key)
        try:
            image_number = int(image_number)
        except (TypeError, ValueError):
            image_number = revised_images + 1
        detail_items = IMAGE_DETAILS.get(pos, [])
        if 1 <= image_number <= len(detail_items):
            observation, proposed_alt = detail_items[image_number - 1]
        else:
            observation = qa["observation_by_key"].get(key) or f"{update['detail']} shown in product gallery image {image_number}."
            proposed_alt = compact_observation_to_alt(observation, f"{update['title']} image {image_number}")
        ws.cell(row_num, idx["observed_visual_details"]).value = observation
        ws.cell(row_num, idx["alt_proposed"]).value = proposed_alt[:125]
        ws.cell(row_num, idx["revision"]).value = "r5"
        ws.cell(row_num, idx["review_status"]).value = "NEEDS_REVIEW"
        for col_name in ["approved_by", "approved_at", "approved_fields", "approved_revision"]:
            if col_name in idx:
                ws.cell(row_num, idx[col_name]).value = ""
        ws.cell(row_num, idx["issues"]).value = (
            "R5: image-level observation and alt retained from r4 contact-sheet review; current admin alt and image URL matched from products_export_1.csv where possible. "
            "Requires QA before approval."
        )
        revised_images += 1

    ws = wb["Keyword_Map"]
    idx = headers(ws)
    keyword_rows_by_pk = defaultdict(int)
    for row_num in range(2, ws.max_row + 1):
        pk = text(ws.cell(row_num, idx["product_key"]).value)
        if pk not in product_key_set:
            continue
        pos = qa["pos_by_key"][pk]
        update = PRODUCT_UPDATES[pos]
        role = text(ws.cell(row_num, idx["keyword_role"]).value)
        if role == "PRIMARY":
            keyword = update["primary"]
        else:
            secondaries = [item.strip() for item in update["secondary"].split(",")]
            secondary_index = keyword_rows_by_pk[pk]
            keyword = secondaries[min(max(secondary_index, 0), len(secondaries) - 1)] if secondaries else update["primary"]
            keyword_rows_by_pk[pk] += 1
        ws.cell(row_num, idx["keyword"]).value = keyword
        ws.cell(row_num, idx["semantic_cluster"]).value = update["cluster"]
        ws.cell(row_num, idx["intent"]).value = "Commercial product intent"
        ws.cell(row_num, idx["target_page_type"]).value = "Product"
        ws.cell(row_num, idx["decision_reason"]).value = update["intent_role"]
        ws.cell(row_num, idx["demand_evidence"]).value = "SERP intent evidence only; no search-volume claim."
        ws.cell(row_num, idx["validation_source"]).value = "qa_batch_001_r4 QA report + products_export_1.csv admin baseline + r5 product-level intent cleanup"
        ws.cell(row_num, idx["checked_at"]).value = now
        if pos == 5:
            ws.cell(row_num, idx["representative_SERP_URLs"]).value = CAT_SERP_URLS
        elif pos == 3:
            ws.cell(row_num, idx["representative_SERP_URLs"]).value = CARDINAL_TREE_SERP
        elif pos == 4:
            ws.cell(row_num, idx["representative_SERP_URLs"]).value = CARDINAL_ROSES_SERP
        ws.cell(row_num, idx["possible_overlap_with_other_products"]).value = update["intent_role"]
        ws.cell(row_num, idx["mapping_reason"]).value = "R5 documents sibling/collection role, product-page target and no-cannibalization rule by visible motif and option set."
        ws.cell(row_num, idx["mapping_status"]).value = "CANDIDATE_MAPPED_NEEDS_RECHECK"
        ws.cell(row_num, idx["mapping_version"]).value = "r5"

    ws = wb["Buyer_Search_Research"]
    idx = headers(ws)
    for row_num in range(2, ws.max_row + 1):
        pk = text(ws.cell(row_num, idx["product_key"]).value)
        if pk not in product_key_set:
            continue
        pos = qa["pos_by_key"][pk]
        update = PRODUCT_UPDATES[pos]
        ws.cell(row_num, idx["purchase_context"]).value = f"US buyer looking for {update['cluster']} as themed bedding or a gift."
        ws.cell(row_num, idx["jtbd_statement"]).value = (
            f"When shopping for {update['cluster']}, the buyer wants a product page that clearly shows the design, "
            "available size/component choices and any verified customization fields."
        )
        ws.cell(row_num, idx["functional_motivation"]).value = "Confirm design, size, components, fabric/care claims and customization requirements before checkout."
        ws.cell(row_num, idx["emotional_social_motivation"]).value = "Match a room theme, sports identity, remembrance message, faith gift or animal/fantasy decor preference."
        ws.cell(row_num, idx["purchase_concerns"]).value = "Fit, included or separately purchased pillowcases, verified personalization fields, care claims and image accuracy."
        ws.cell(row_num, idx["source_refs"]).value = f"products_export_1.csv sha256:{admin_hash}; QA report qa_batch_001_r4 {QA_RUN_ID}; contact sheets 1-10"
        ws.cell(row_num, idx["observed_at"]).value = now
        ws.cell(row_num, idx["research_status"]).value = "REVISED_R5_NEEDS_RECHECK"
        ws.cell(row_num, idx["limitations"]).value = "No search-volume source supplied; Shopify admin CSV baseline is now available for stored title/meta/body/image alt."
        ws.cell(row_num, idx["seo_application"]).value = f"Use {update['primary']} with intent role: {update['intent_role']}"

    ws = wb["Product_Evidence"]
    idx = headers(ws)
    for row_num in range(2, ws.max_row + 1):
        evidence_id = text(ws.cell(row_num, idx["evidence_id"]).value)
        match = re.search(r"_(\d{3})$", evidence_id)
        if not match:
            continue
        pos = int(match.group(1))
        if pos not in PRODUCT_UPDATES:
            continue
        update = PRODUCT_UPDATES[pos]
        pk = qa["product_keys"][pos - 1]
        handle = qa["handle_by_key"][pk]
        admin_row = admin[handle]
        ws.cell(row_num, idx["sources_accessed"]).value = (
            f"Storefront/product evidence, contact sheets and QA report qa_batch_001_r4; products_export_1.csv sha256:{admin_hash}"
        )
        ws.cell(row_num, idx["current_H1"]).value = admin_row["Title"]
        ws.cell(row_num, idx["current_meta_title"]).value = admin_row["SEO Title"] or "EMPTY"
        ws.cell(row_num, idx["verified_product_facts"]).value = (
            f"Admin export matched handle {handle}; type={admin_row['Type']}; options="
            f"{admin_row['Option1 Name']}, {admin_row['Option2 Name']}, {admin_row['Option3 Name']}; "
            f"status={admin_row['Status']}; image_count={len(admin_row['images'])}; design={update['detail']}."
        )
        if pos == 5:
            ws.cell(row_num, idx["SERP_evidence_references"]).value = CAT_SERP_URLS
        elif pos == 3:
            ws.cell(row_num, idx["SERP_evidence_references"]).value = CARDINAL_TREE_SERP
        elif pos == 4:
            ws.cell(row_num, idx["SERP_evidence_references"]).value = CARDINAL_ROSES_SERP
        ws.cell(row_num, idx["factual_conflicts"]).value = "R5 resolves qa_batch_001_r4 QA body and mapping issues; needs independent QA recheck."
        ws.cell(row_num, idx["confidence_and_reason"]).value = (
            "R5 uses contact-sheet image review, verified customization fields and Shopify admin CSV baseline; no approval or import file created."
        )

    ws = wb["README_QA"]
    idx = headers(ws)
    rows = [
        ("qa_batch_001_r5_revision", "r5", "Deep review of latest qa_batch_001_r4; source r4 workbook was not modified."),
        ("qa_batch_001_r5_admin_export", str(ADMIN_EXPORT.relative_to(ROOT)), f"Admin CSV baseline used; sha256:{admin_hash}."),
        ("qa_batch_001_r5_scope", "inventory positions 1-10", "No products outside qa_batch_001 were revised."),
        ("qa_batch_001_r5_revised_products", str(revised_products), "SEO_Products rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_001_r5_revised_images", str(revised_images), "Image_Audit rows revised and kept NEEDS_REVIEW."),
        ("qa_batch_001_r5_admin_alt_matches", str(admin_alt_matched), "Image rows with current admin alt matched by image URL from Shopify CSV."),
    ]
    for metric, value, definition in rows:
        ws.append([
            metric if col == idx["metric"] else value if col == idx["value"] else definition if col == idx["definition"] else ""
            for col in range(1, ws.max_column + 1)
        ])

    wb.save(OUTPUT)
    load_workbook(OUTPUT).save(OUTPUT)

    summary = {
        "created_at": now,
        "batch_id": BATCH_ID,
        "source_revision": str(SOURCE.relative_to(ROOT)),
        "source_qa_r4_report": str(QA_REPORT.relative_to(ROOT)),
        "admin_export": str(ADMIN_EXPORT.relative_to(ROOT)),
        "admin_export_sha256": admin_hash,
        "revision_workbook": str(OUTPUT.relative_to(ROOT)),
        "scope": "qa_batch_001 only; inventory positions 1-10",
        "revision": "r5",
        "revised_products": revised_products,
        "revised_images": revised_images,
        "admin_alt_matches": admin_alt_matched,
        "review_status": "NEEDS_REVIEW",
        "approved_created": False,
        "shopify_import_created": False,
        "r5_changes": [
            "Rewrote description_proposed_html for all 10 products using product-specific verified facts, dimensions, components, feature panels and customization rules.",
            "Documented collection/internal-link role and no-cannibalization rule for sibling cardinal, cat and Celtic products.",
            "Cleaned product-level commercial intent references for products 3 and 4 without adding search-volume claims.",
            "Retained the 65-image r4 audit and carried all revised image rows forward as r5.",
        ],
        "next_step": "Sang QA lại qa_batch_001_r5 before any approval or import.",
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text(
        "\n".join(
            [
                "# Revision qa_batch_001_r5",
                "",
                "Artifact rà sâu chính thức theo quy trình từng lô 10 sản phẩm.",
                "",
                f"- Nguồn r4: `{SOURCE.relative_to(ROOT)}`",
                f"- QA r4 report: `{QA_REPORT.relative_to(ROOT)}`",
                f"- Shopify admin export: `{ADMIN_EXPORT.relative_to(ROOT)}`",
                f"- Admin export SHA-256: `{admin_hash}`",
                f"- Workbook r5: `{OUTPUT.relative_to(ROOT)}`",
                f"- Sản phẩm sửa: {revised_products}",
                f"- Ảnh sửa/refresh: {revised_images}",
                f"- Admin image alt match: {admin_alt_matched}/{revised_images}",
                "- Trạng thái: `NEEDS_REVIEW`; không tạo approval; không tạo file import Shopify.",
                "- Điểm sửa sâu: body HTML từng sản phẩm, collection/internal-link role, no-cannibalization rule, intent cleanup cho products 3-4, giữ audit 65 ảnh từ r4.",
                "- Bước tiếp theo: Sang QA lại `qa_batch_001_r5`.",
                "",
            ]
        ),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    update_workbook()
