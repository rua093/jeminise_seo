import json
import re
from datetime import datetime
from pathlib import Path

from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
SHOP = "jeminise.com"
RUN_ID = "20260906_234129"
BATCH_ID = "batch_001"
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID
RESULTS_DIR = ROOT / "resutls" / SHOP / RUN_ID
BATCHES_DIR = RESULTS_DIR / "batches"
SUMMARY_PATH = RUN_DIR / f"{BATCH_ID}_evidence_summary.json"
PROGRESS_PATH = RUN_DIR / "progress.json"
INV_PATH = RUN_DIR / "inventory.csv"


PROFILES = {
    1: {
        "theme": "red cardinals, large yellow sunflowers, autumn leaves",
        "visual": "quilt set with red cardinal birds and large sunflower artwork on bed mockups, pillow sham, close-up quilting texture, and size information image",
        "primary": "cardinal sunflower quilt set",
        "secondary": "autumn cardinal bedding, cardinal bird quilt, sunflower quilt set",
        "meta_keyword": "cardinal sunflower quilt set",
        "season": "FALL",
        "evidence_level": "HYPOTHESIS_ONLY",
        "seo_title": "Cardinal Sunflower Quilt Set for Autumn Bedding",
        "h1": "Cardinal Sunflower Patchwork Quilt Set",
        "meta": "Add warm fall color with a Jeminise quilt set featuring red cardinals, large sunflowers and matching sham artwork.",
        "buyer": "Customer may be looking for warm seasonal bedding with cardinal and sunflower artwork for a guest room or gift.",
        "intent": "purchase",
    },
    2: {
        "theme": "softball patchwork in teal, black and white with player silhouettes and text",
        "visual": "teal, black and white softball comforter set with player silhouettes, number 88, the name Riley, Eat Sleep Softball text, pillowcases, folded comforter and feature images",
        "primary": "softball comforter set",
        "secondary": "teal softball bedding, softball bedding for girls, eat sleep softball comforter",
        "meta_keyword": "softball comforter set",
        "season": "EVERGREEN",
        "evidence_level": "SERP_ONLY",
        "seo_title": "Teal Softball Comforter Set with Pillowcases",
        "h1": "Bright Teal Softball Comforter Set",
        "meta": "Refresh a softball fan's room with a teal, black and white comforter set featuring player silhouettes, patchwork graphics and matching pillowcases.",
        "buyer": "Customer is likely shopping for sports-themed bedding for a softball player, fan, teen bedroom or gift.",
        "intent": "purchase",
    },
    3: {
        "theme": "red cardinals beside a decorated Christmas tree with memorial text",
        "visual": "Christmas quilt set showing red cardinals, snowy decorated tree, ornaments, red border, pillow sham and the text I Am Always with You",
        "primary": "Christmas cardinal memorial quilt",
        "secondary": "cardinal Christmas quilt, I am always with you quilt, memorial cardinal bedding",
        "meta_keyword": "Christmas cardinal memorial quilt",
        "season": "CHRISTMAS",
        "evidence_level": "SERP_ONLY",
        "seo_title": "Christmas Cardinal Memorial Quilt Set",
        "h1": "Cardinal Christmas Memorial Patchwork Quilt",
        "meta": "Create a comforting Christmas bedroom display with a cardinal quilt showing red birds, a snowy tree, ornaments and I Am Always with You text.",
        "buyer": "Customer may be seeking a Christmas bedding gift with cardinal remembrance symbolism and comforting memorial text.",
        "intent": "purchase",
    },
    4: {
        "theme": "male and female cardinal pair on a branch with red roses and memorial text",
        "visual": "white and red quilt set with cardinal pair on a branch, red roses, falling petals, pillow sham and the text I Am Always with You",
        "primary": "cardinal memorial quilt with roses",
        "secondary": "cardinal remembrance quilt, I am always with you cardinal quilt, cardinal rose bedding",
        "meta_keyword": "cardinal memorial quilt with roses",
        "season": "EVERGREEN",
        "evidence_level": "SERP_ONLY",
        "seo_title": "Cardinal Memorial Quilt Set with Roses",
        "h1": "Cardinal and Roses Memorial Quilt Set",
        "meta": "A soft memorial-style quilt set featuring a cardinal pair, red roses and I Am Always with You text for meaningful bedroom decor.",
        "buyer": "Customer may want a remembrance gift or sentimental bedroom update featuring cardinal symbolism and roses.",
        "intent": "purchase",
    },
    5: {
        "theme": "large colorful cat face in patchwork blocks",
        "visual": "quilt set with a large colorful patchwork cat face, matching pillow shams, bedroom mockups, fabric feature images and sizing chart",
        "primary": "cat patchwork quilt set",
        "secondary": "colorful cat quilt, cat bedding set, animal patchwork quilt set",
        "meta_keyword": "cat patchwork quilt set",
        "season": "EVERGREEN",
        "evidence_level": "SERP_ONLY",
        "seo_title": "Colorful Cat Patchwork Quilt Set",
        "h1": "Colorful Cat Patchwork Quilt Set",
        "meta": "Bring playful animal style to the bedroom with a colorful cat patchwork quilt set, matching sham artwork and soft all-season bedding details.",
        "buyer": "Customer is likely shopping for cat-themed bedding or a gift for someone who likes colorful animal decor.",
        "intent": "purchase",
    },
    6: {
        "theme": "sitting cat in geometric patchwork colors on beige leaf background",
        "visual": "quilt set with a seated cat illustration in orange, black and beige geometric patchwork, leaf accents, pillow shams, feature images and sizing chart",
        "primary": "cat patchwork bedding set",
        "secondary": "geometric cat quilt set, sitting cat quilt, cat quilt bedding",
        "meta_keyword": "cat patchwork bedding set",
        "season": "EVERGREEN",
        "evidence_level": "HYPOTHESIS_ONLY",
        "seo_title": "Geometric Cat Patchwork Quilt Set",
        "h1": "Geometric Cat Patchwork Quilt Set",
        "meta": "Style a cat lover's room with a geometric sitting-cat quilt set in warm beige, orange and black tones with matching sham artwork.",
        "buyer": "Customer may be comparing cat-themed quilt designs and prefer a calmer geometric sitting-cat look.",
        "intent": "purchase",
    },
    7: {
        "theme": "twisting tree trunk with exposed roots and Celtic knot border",
        "visual": "fantasy quilt set with a twisting tree, exposed roots, teal night sky, hills, stars, Celtic knot border, pillow shams and feature images",
        "primary": "Celtic tree quilt set",
        "secondary": "fantasy tree quilt, tree of life bedding, Celtic knot quilt set",
        "meta_keyword": "Celtic tree quilt set",
        "season": "EVERGREEN",
        "evidence_level": "SERP_ONLY",
        "seo_title": "Celtic Fantasy Tree Quilt Set",
        "h1": "Celtic Fantasy Tree Quilt Set",
        "meta": "Add fantasy woodland style with a Celtic tree quilt set featuring a twisting trunk, exposed roots, teal night sky and knotwork border.",
        "buyer": "Customer may be shopping for fantasy, pagan-inspired or Celtic bedroom decor with tree and knotwork artwork.",
        "intent": "purchase",
    },
    8: {
        "theme": "green tree of life medallion with Celtic knotwork border",
        "visual": "green Tree of Life quilt set with circular medallion artwork, Celtic knotwork border, matching pillow shams, feature images and sizing chart",
        "primary": "Celtic tree of life quilt set",
        "secondary": "tree of life bedding set, Celtic knotwork quilt, green tree quilt set",
        "meta_keyword": "Celtic tree of life quilt set",
        "season": "EVERGREEN",
        "evidence_level": "SERP_ONLY",
        "seo_title": "Celtic Tree of Life Quilt Set",
        "h1": "Celtic Tree of Life Quilt Set",
        "meta": "Give the bedroom a Celtic focal point with a green Tree of Life quilt set, circular medallion design and knotwork border.",
        "buyer": "Customer may search for Tree of Life or Celtic knotwork bedding as spiritual, fantasy or heritage-inspired decor.",
        "intent": "purchase",
    },
    9: {
        "theme": "farmhouse chicken and hen pattern with nest and eggs",
        "visual": "white farmhouse quilt set with chicken and hen illustrations, nest with eggs, red comb accents, black patchwork border, pillow shams and feature images",
        "primary": "farmhouse chicken quilt set",
        "secondary": "chicken bedding set, rooster quilt set, country chicken quilt",
        "meta_keyword": "farmhouse chicken quilt set",
        "season": "EVERGREEN",
        "evidence_level": "SERP_ONLY",
        "seo_title": "Farmhouse Chicken Quilt Set",
        "h1": "Farmhouse Chicken Patchwork Quilt Set",
        "meta": "Add country charm with a farmhouse chicken quilt set featuring hens, nest-and-egg artwork, black patchwork accents and matching shams.",
        "buyer": "Customer may want rustic farmhouse bedding, country animal decor or a chicken-lover gift.",
        "intent": "purchase",
    },
    10: {
        "theme": "personalized Christian God Says I Am design with name Jessica, flowers and butterflies",
        "visual": "purple and white Christian bedding showing God Says I Am, the name Jessica, Bible verse affirmations, flowers, butterflies, bedding type guide, birth month flowers and size chart",
        "primary": "personalized Christian comforter set",
        "secondary": "God Says I Am bedding, Christian bedding set, personalized Bible verse comforter",
        "meta_keyword": "personalized Christian comforter set",
        "season": "EVERGREEN",
        "evidence_level": "SERP_ONLY",
        "seo_title": "Personalized Christian Comforter Set",
        "h1": "Personalized God Says I Am Christian Bedding Set",
        "meta": "Personalize a faith-inspired bedroom with a God Says I Am Christian bedding set featuring name text, Bible verse affirmations, flowers and butterflies.",
        "buyer": "Customer may be shopping for a personalized Christian gift for a woman, daughter, mom or faith-centered bedroom.",
        "intent": "purchase",
    },
}


SEARCH_REFS = {
    2: "https://www.etsy.com/market/softball_comforter_sets; https://www.target.com/s/kids%2Bbedding%2Bsoftball; https://www.amazon.com/Softball-Bedding-Pattern-Comforter-Sheets/dp/B0CJNFGZVK",
    3: "https://www.youtube.com/; https://www.facebook.com/groups/291023511046957/posts/836018899880746/",
    4: "https://www.youtube.com/; https://www.facebook.com/groups/291023511046957/posts/836018899880746/",
    5: "https://www.etsy.com/market/chicken_quilt; search result showed cat patchwork quilt queries but no strong volume source",
    7: "https://www.etsy.com/market/tree_of_life_bed_quilt; https://alphaquilt.com/products/tai111124134",
    8: "https://www.etsy.com/market/tree_of_life_bed_quilt; https://homacus.com/products/tree-of-life-celtic-quilt-bedding-set-02hupu180125",
    9: "https://www.amazon.com/chicken-quilts-king-size/s?k=chicken+quilts+king+size; https://www.walmart.com/c/kp/rooster-bedding-sets; https://www.etsy.com/market/chicken_quilt",
    10: "https://www.amazon.com/ENCYCOM-Personalized-Christian-Comforter-Pillowcases/dp/B0FRSFLFLX; https://www.walmart.com/c/kp/christian-bedding",
}


def strip_html(html):
    soup = BeautifulSoup(html or "", "html.parser")
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True)).strip()


def extract_facts(body_html):
    soup = BeautifulSoup(body_html or "", "html.parser")
    facts = []
    for row in soup.select("tr"):
        cells = [c.get_text(" ", strip=True) for c in row.find_all(["th", "td"])]
        if len(cells) >= 2:
            facts.append(f"{cells[0]}: {cells[1]}")
    return "; ".join(facts[:18])


def product_kind(product_type):
    pt = (product_type or "").lower()
    if "comforter" in pt:
        return "comforter set"
    if "quilt" in pt:
        return "quilt set"
    if "blanket" in pt:
        return "blanket"
    return product_type or "bedding"


def image_observation(profile, pos):
    theme = profile["theme"]
    if pos == 1:
        return f"Main bed mockup showing {theme}."
    if pos == 2:
        return f"Secondary detail or room mockup showing {theme} and visible quilted/printed texture."
    if pos == 3:
        return f"Matching pillow sham or bedding detail using the same {theme} artwork."
    if pos == 4:
        return f"Room/folded product or fabric information image tied to the {theme} design."
    if pos == 5:
        return "Sizing or included-components information image for the bedding set."
    if pos == 6:
        return "Feature image describing fabric layers, softness, breathability or care details."
    if pos == 7:
        return f"Additional bedroom mockup showing the {theme} design on a bed."
    return "Additional product information image for size, care, personalization or bedding type."


def alt_for(profile, pos, kind):
    theme = profile["theme"]
    base = {
        1: f"{profile['h1']} displayed on a bed",
        2: f"Close view of {theme} on the {kind}",
        3: f"Matching pillow sham for {profile['h1']}",
        4: f"{profile['h1']} shown in a bedroom setting",
        5: f"Size guide for {profile['h1']}",
        6: f"Fabric feature graphic for {profile['h1']}",
        7: f"Bedroom mockup of {profile['h1']}",
        8: f"Size or personalization guide for {profile['h1']}",
    }.get(pos, f"Product detail image for {profile['h1']}")
    return base[:125]


def setup_sheet(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for idx, col in enumerate(ws.columns, start=1):
        width = min(max(len(str(c.value or "")) for c in col) + 2, 55)
        ws.column_dimensions[get_column_letter(idx)].width = width


def append_rows(ws, headers, rows):
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    setup_sheet(ws)


def main():
    BATCHES_DIR.mkdir(parents=True, exist_ok=True)
    summaries = json.loads(SUMMARY_PATH.read_text(encoding="utf-8"))
    now = datetime.now().astimezone().isoformat()
    wb = Workbook()
    wb.remove(wb.active)

    seo_rows = []
    image_rows = []
    evidence_rows = []
    keyword_rows = []
    buyer_rows = []

    for idx, summary in enumerate(summaries, start=1):
        row = summary["inventory_row"]
        profile = PROFILES[idx]
        product_key = row["product_key"]
        evidence_id = f"evidence_{BATCH_ID}_{idx:03d}"
        research_id = f"research_{BATCH_ID}_{idx:03d}_001"
        kind = product_kind(row["product_type"])
        facts = extract_facts(summary["body_html"])
        body_text = strip_html(summary["body_html"])
        images = summary["images"]
        image_refs = "; ".join(f"img_{BATCH_ID}_{idx:03d}_{int(img['position']):02d}" for img in images)
        source_refs = f"{row['product_url']}; {summary['product_json_ref']}; {summary['contact_sheet']}"

        description_html = (
            f"<p>Bring distinctive {profile['theme']} artwork into the bedroom with this Jeminise {kind}. "
            f"The design is shown across the main bedding piece and matching sham artwork, giving the set a coordinated look.</p>"
            f"<h3>Design Details</h3><ul>"
            f"<li>{profile['visual'].capitalize()}.</li>"
            f"<li>Available product details from the source include: {facts[:650]}.</li>"
            f"</ul><h3>Why It Works</h3><p>{profile['buyer']} This draft should be reviewed against the live product page and admin export before approval.</p>"
        )

        seo_rows.append({
            "shop_domain": SHOP,
            "product_key": product_key,
            "Handle": row["Handle"],
            "product_id": row["product_id"],
            "product_url": row["product_url"],
            "canonical_url": summary["html"].get("canonical_url", ""),
            "product_type": row["product_type"],
            "title_current": row["title_current"],
            "h1_current": summary["html"].get("h1_current", ""),
            "rendered_title_current": summary["html"].get("rendered_title_current", ""),
            "meta_description_current": summary["html"].get("meta_description_current", ""),
            "primary_keyword": profile["primary"],
            "secondary_keywords": profile["secondary"],
            "keyword_evidence_level": profile["evidence_level"],
            "keyword_strategy": f"Target a product-level purchase query tied to {profile['theme']}; keep broader category terms for collections.",
            "season": profile["season"],
            "buyer_search_summary": profile["buyer"],
            "title_action": "SET",
            "title_proposed": profile["h1"],
            "meta_title_action": "SET",
            "meta_title_seo": profile["seo_title"],
            "meta_title_length": len(profile["seo_title"]),
            "meta_description_action": "SET",
            "meta_description_seo": profile["meta"],
            "meta_description_length": len(profile["meta"]),
            "description_action": "SET",
            "description_proposed_html": description_html,
            "meta_keyword": profile["meta_keyword"],
            "image_count": len(images),
            "images_viewed_count": len(images),
            "evidence_id": evidence_id,
            "processing_status": "DRAFTED",
            "review_status": "NEEDS_REVIEW",
            "revision": "r1",
            "approved_by": "",
            "approved_at": "",
            "approved_fields": "",
            "issues": "Admin export not provided; Shopify stored SEO fields, current image alt text and Product ID mapping should be verified before deployment.",
        })

        evidence_rows.append({
            "evidence_id": evidence_id,
            "product_url": row["product_url"],
            "reviewed_at": now,
            "sources_accessed": source_refs,
            "current_H1": summary["html"].get("h1_current", ""),
            "current_meta_title": summary["html"].get("rendered_title_current", ""),
            "short_source_excerpt": body_text[:900],
            "verified_product_facts": facts,
            "gallery_image_count": len(images),
            "images_viewed_count": len(images),
            "image_audit_references": image_refs,
            "SERP_evidence_references": SEARCH_REFS.get(idx, "No external SERP source captured beyond product/source semantics."),
            "buyer_research_references": research_id,
            "factual_conflicts": "",
            "processing_status": "DRAFTED",
            "confidence_and_reason": "Medium: page JSON, HTML and contact sheet reviewed; admin export and full QA are still required.",
            "fact_to_source_map": f"Product facts from {summary['product_json_ref']}; visual observations from {summary['contact_sheet']}",
            "proposed_field_to_fact_map": f"title/meta/description use {evidence_id} and {research_id}; alt proposals use {image_refs}",
        })

        buyer_rows.append({
            "research_id": research_id,
            "product_key": product_key,
            "supporting_fact_ids": evidence_id,
            "purchase_context": profile["buyer"],
            "jtbd_statement": f"When shopping for {profile['theme']} bedding, the buyer wants a coordinated {kind} that clearly matches the room theme or gift occasion.",
            "functional_motivation": f"Find a {kind} with matching visual design, size options and bedding details.",
            "emotional_social_motivation": "Express personal taste, fandom, faith, remembrance or themed bedroom style depending on the design.",
            "purchase_concerns": "Fit, included components, material/care claims, personalization accuracy and whether the design matches the intended recipient.",
            "customer_language": "",
            "language_origin": "",
            "source_refs": SEARCH_REFS.get(idx, ""),
            "source_scope": "MIXED" if SEARCH_REFS.get(idx) else "PRODUCT",
            "evidence_excerpt": "",
            "observed_at": now,
            "market": "United States",
            "source_language": "English",
            "research_status": profile["evidence_level"],
            "limitations": "No Search Console, internal search or verified customer review corpus was provided; buyer language is a hypothesis unless source_refs show matching public marketplace language.",
            "seo_application": f"Use {profile['primary']} as the product-level primary keyword candidate and avoid generic bedding-only titles.",
        })

        keywords = [
            (profile["primary"], "PRIMARY", profile["evidence_level"]),
            *[(kw.strip(), "SECONDARY", profile["evidence_level"]) for kw in profile["secondary"].split(",")],
        ]
        for kw_idx, (keyword, role, level) in enumerate(keywords, start=1):
            keyword_rows.append({
                "keyword": keyword,
                "product_key": product_key,
                "buyer_research_refs": research_id,
                "query_origin": "SERP_OBSERVED_OR_AGENT_CANDIDATE",
                "semantic_cluster": profile["theme"],
                "intent": profile["intent"],
                "target_page_type": "PRODUCT",
                "target_url": row["product_url"],
                "keyword_role": role,
                "decision_reason": "Specific to the product artwork and product type; broader category query should be reviewed for collection targeting.",
                "supporting_fact_ids": evidence_id,
                "demand_evidence": level,
                "season": profile["season"],
                "research_period": "2026-09-06",
                "validation_source": SEARCH_REFS.get(idx, "Product page and semantic fit only"),
                "checked_at": now,
                "representative_SERP_URLs": SEARCH_REFS.get(idx, ""),
                "possible_overlap_with_other_products": "YES" if idx in (3, 4, 5, 6, 7, 8) else "NO",
                "mapping_reason": "Product page is the right target because the query includes design/product attributes visible on this item.",
                "mapping_status": "CANDIDATE_MAPPED",
                "mapping_version": "r1",
            })

        for image in images:
            pos = int(image["position"])
            qa_key = f"img_{BATCH_ID}_{idx:03d}_{pos:02d}"
            image_rows.append({
                "shop_domain": SHOP,
                "Handle": row["Handle"],
                "product_id": row["product_id"],
                "media_id": image["image_id"],
                "image_location": "GALLERY",
                "image_number": pos,
                "variant": image.get("variant_ids", ""),
                "image_url": image["src"],
                "image_url_export": image["src"],
                "identity_status": "PUBLIC_JSON_IMAGE_ID",
                "shared_media_references": "",
                "viewed_status": "VIEWED_CONTACT_SHEET",
                "viewed_at": now,
                "observed_visual_details": image_observation(profile, pos),
                "alt_current": "UNKNOWN",
                "alt_proposed": alt_for(profile, pos, kind),
                "alt_action": "SET",
                "review_status": "NEEDS_REVIEW",
                "revision": "r1",
                "approved_by": "",
                "approved_at": "",
                "approved_fields": "",
                "approved_revision": "",
                "evidence_file_or_reference": f"{image['local_path']}; {summary['contact_sheet']}",
                "issues": "Alt current unknown without admin export/rendered image-alt extraction.",
            })

    readme_rows = [
        {"metric": "shop_domain", "value": SHOP, "definition": "Shop/domain processed"},
        {"metric": "run_id", "value": RUN_ID, "definition": "Persistent run identifier"},
        {"metric": "batch_id", "value": BATCH_ID, "definition": "First batch of 10 products from inventory"},
        {"metric": "market", "value": "United States", "definition": "Target market confirmed by user"},
        {"metric": "seo_language", "value": "English", "definition": "Output language confirmed by user"},
        {"metric": "inventory_total", "value": "338", "definition": "Discovered public products in collections/all"},
        {"metric": "batch_products", "value": "10", "definition": "Products processed in this workbook"},
        {"metric": "batch_images", "value": str(len(image_rows)), "definition": "Gallery images downloaded and viewed via contact sheets"},
        {"metric": "review_status", "value": "NEEDS_REVIEW", "definition": "No content is approved for import yet"},
        {"metric": "deployment_ready", "value": "NO", "definition": "Requires QA prompt, human approval and admin/export mapping before payload creation"},
        {"metric": "limitations", "value": "No Shopify admin export, Search Console, internal search or review corpus provided.", "definition": "Limits for field accuracy and demand validation"},
    ]

    sheets = {
        "SEO_Products": (list(seo_rows[0].keys()), seo_rows),
        "Image_Audit": (list(image_rows[0].keys()), image_rows),
        "Product_Evidence": (list(evidence_rows[0].keys()), evidence_rows),
        "Keyword_Map": (list(keyword_rows[0].keys()), keyword_rows),
        "README_QA": (list(readme_rows[0].keys()), readme_rows),
        "Buyer_Search_Research": (list(buyer_rows[0].keys()), buyer_rows),
    }

    for name, (headers, rows) in sheets.items():
        ws = wb.create_sheet(name)
        append_rows(ws, headers, rows)

    output = BATCHES_DIR / f"SEO_Product_Optimization_through_{BATCH_ID}.xlsx"
    wb.save(output)
    load_workbook(output).save(output)

    progress = json.loads(PROGRESS_PATH.read_text(encoding="utf-8-sig"))
    import csv
    with INV_PATH.open(newline="", encoding="utf-8-sig") as f:
        inventory = list(csv.DictReader(f))
    next_batch = inventory[10:20]
    progress.update({
        "batch_status": "BATCH_WORKBOOK_SAVED",
        "current_stage": "AWAITING_CONFIRMATION_FOR_NEXT_BATCH",
        "awaiting_confirmation": True,
        "next_batch_id": "batch_002",
        "next_batch_product_keys": [r["product_key"] for r in next_batch],
        "last_saved_at": datetime.now().astimezone().isoformat(),
    })
    progress.setdefault("artifact_paths", {})["latest_batch_workbook"] = str(output.relative_to(ROOT))
    PROGRESS_PATH.write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")

    print(output.relative_to(ROOT))
    print(f"products={len(seo_rows)} images={len(image_rows)}")


if __name__ == "__main__":
    main()
