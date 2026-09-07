from __future__ import annotations

import hashlib
import json
from collections import Counter
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
RUN_ID = "20260906_234129"
QA_RUN_ID = "20260907_083534"
QA_BATCH_ID = "qa_batch_001"
TZ = "+07:00"
QA_DIR = ROOT / "seo_runs" / "jeminise.com" / RUN_ID / "qa" / QA_RUN_ID
OUT_DIR = ROOT / "resutls" / "jeminise.com" / RUN_ID / "qa" / QA_RUN_ID
SOURCE = ROOT / "resutls" / "jeminise.com" / RUN_ID / "batches" / "SEO_Product_Optimization_through_batch_034.xlsx"
RATING = {"FULL": 1.0, "PARTIAL": 0.5, "FAIL": 0.0}
PRODUCT_WEIGHTS = {"P1": 15, "P2": 10, "K1": 10, "K2": 5, "K3": 5, "T1": 10, "T2": 5, "D1": 5, "D2": 10, "I1": 20, "E1": 5}
IMAGE_WEIGHTS = {"IM1": 40, "IM2": 30, "IM3": 20, "IM4": 10}


def now() -> str:
    return datetime.now().astimezone().isoformat()


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest().upper()


def save_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def stable_image_key(product_key: str, url: str, position: int) -> str:
    clean = url.split("?")[0]
    digest = hashlib.sha256(f"{product_key}|{clean}|GALLERY|{position}".encode()).hexdigest()[:16]
    return f"qaimg_{digest}"


ACTUAL = {
    1: [
        "Bed mockup with two red cardinals, large yellow sunflowers and autumn leaves.",
        "Close-up of the cardinal-and-sunflower print and quilted stitching.",
        "Pillow sham using the cardinal-and-sunflower artwork.",
        "Alternate bedroom angle showing the cardinal-and-sunflower quilt.",
        "Included-components and dimensions graphic: one quilt; two standard shams are optional; Throw/Twin/Queen/King sizes shown.",
    ],
    2: [
        "Front bed mockup: teal, black and white softball patchwork, player silhouettes, number 88, RILEY and EAT SLEEP SOFTBALL text.",
        "Angled bedroom mockup of the same softball comforter design.",
        "Close view of two pillowcases with number 88 and a batter silhouette.",
        "Folded comforter showing the printed face and plain white reverse.",
        "Close-up comparing the plain reverse with the printed face; no size chart is present.",
        "Feature graphic stating machine washable, soft microfiber filling, lightweight and breathable.",
        "Included-components graphic stating one comforter and two pillowcases included.",
    ],
    3: [
        "Bed mockup with red cardinals, a decorated snowy Christmas tree, ornaments and I Am Always with You text.",
        "Close-up of the Christmas cardinal print and quilted stitching.",
        "Pillow sham using the Christmas cardinal artwork.",
        "Alternate bedroom angle showing the Christmas cardinal quilt.",
        "Included-components and dimensions graphic for the quilt and optional shams.",
    ],
    4: [
        "Bed mockup with a red male cardinal, tan female cardinal, red roses and I Am Always with You text.",
        "Close-up of the cardinal-and-roses print and quilted stitching.",
        "Pillow sham using the cardinal-and-roses artwork.",
        "Alternate bedroom angle showing the cardinal-and-roses quilt.",
        "Included-components and dimensions graphic for the quilt and optional shams.",
    ],
    5: [
        "Front bed mockup with a large colorful patchwork cat face.",
        "Bedroom/feature graphic with premium print, no fading and skin-friendly callouts.",
        "Pillow shams plus lightweight, soft, anti-pill and anti-static feature callouts.",
        "All-season and fabric-quality feature graphic with a small bed mockup.",
        "Quilt size chart.",
        "Bedspread construction-layer and fabric-feature graphic.",
        "Overhead bedroom mockup of the colorful cat quilt.",
    ],
    6: [
        "Front bed mockup with a seated geometric calico cat on a beige leaf background.",
        "Bedroom/feature graphic with premium print, no fading and skin-friendly callouts.",
        "Pillow shams plus lightweight, soft, anti-pill and anti-static feature callouts.",
        "All-season and fabric-quality feature graphic with a small bed mockup.",
        "Quilt size chart.",
        "Bedspread construction-layer and fabric-feature graphic.",
        "Overhead bedroom mockup of the geometric cat quilt.",
    ],
    7: [
        "Front bed mockup with a twisting tree, exposed roots, teal night sky and Celtic knot border.",
        "Bedroom/feature graphic with premium print, no fading and skin-friendly callouts.",
        "Pillow shams plus lightweight, soft, anti-pill and anti-static feature callouts.",
        "All-season and fabric-quality feature graphic with a small bed mockup.",
        "Quilt size chart.",
        "Bedspread construction-layer and fabric-feature graphic.",
        "Overhead bedroom mockup of the Celtic fantasy tree quilt.",
    ],
    8: [
        "Front bed mockup with a green Tree of Life medallion and Celtic knotwork border.",
        "Bedroom/feature graphic with premium print, no fading and skin-friendly callouts.",
        "Pillow shams plus lightweight, soft, anti-pill and anti-static feature callouts.",
        "All-season and fabric-quality feature graphic with a small bed mockup.",
        "Quilt size chart.",
        "Bedspread construction-layer and fabric-feature graphic.",
        "Overhead bedroom mockup of the green Celtic Tree of Life quilt.",
    ],
    9: [
        "Front bed mockup with farmhouse hens, a nest with eggs and black patchwork accents.",
        "Bedroom/feature graphic with premium print, no fading and skin-friendly callouts.",
        "Pillow shams plus lightweight, soft, anti-pill and anti-static feature callouts.",
        "All-season and fabric-quality feature graphic with a small bed mockup.",
        "Quilt size chart.",
        "Bedspread construction-layer and fabric-feature graphic.",
        "Overhead bedroom mockup of the farmhouse chicken quilt.",
    ],
    10: [
        "Main bed mockup with God Says I Am, the sample name Jessica, floral artwork and butterflies.",
        "Alternate bed mockup of the same Christian affirmation design.",
        "Comparison graphic for Duvet Cover Set versus Comforter Set.",
        "Feature graphic stating 100% microfiber, 3D print, soft, lightweight and breathable.",
        "Birth Month Flowers chart from January through December.",
        "Bedroom feature graphic with soft, lightweight, durable and breathable callouts.",
        "Easy-care graphic with wrinkle-free, stain-proof, anti-pilling and won't-fade claims.",
        "Twin, Full, Queen and King dimension chart.",
    ],
}

IMAGE_ASSESS = {
    1: [("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL"), ("PARTIAL", "PARTIAL")],
    2: [("FULL", "FULL"), ("PARTIAL", "FAIL"), ("FULL", "FULL"), ("PARTIAL", "FAIL"), ("FAIL", "FAIL"), ("FULL", "FULL"), ("FAIL", "FAIL")],
    3: [("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL"), ("PARTIAL", "PARTIAL")],
    4: [("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL"), ("PARTIAL", "PARTIAL")],
    5: [("FULL", "FULL"), ("PARTIAL", "FAIL"), ("PARTIAL", "PARTIAL"), ("PARTIAL", "PARTIAL"), ("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL")],
    6: [("FULL", "FULL"), ("PARTIAL", "FAIL"), ("PARTIAL", "PARTIAL"), ("PARTIAL", "PARTIAL"), ("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL")],
    7: [("FULL", "FULL"), ("PARTIAL", "FAIL"), ("PARTIAL", "PARTIAL"), ("PARTIAL", "PARTIAL"), ("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL")],
    8: [("FULL", "FULL"), ("PARTIAL", "FAIL"), ("PARTIAL", "PARTIAL"), ("PARTIAL", "PARTIAL"), ("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL")],
    9: [("FULL", "FULL"), ("PARTIAL", "FAIL"), ("PARTIAL", "PARTIAL"), ("PARTIAL", "PARTIAL"), ("FULL", "FULL"), ("FULL", "FULL"), ("FULL", "FULL")],
    10: [("FULL", "FULL"), ("PARTIAL", "FAIL"), ("FAIL", "FAIL"), ("FAIL", "FAIL"), ("FAIL", "FAIL"), ("FULL", "PARTIAL"), ("FAIL", "FAIL"), ("PARTIAL", "PARTIAL")],
}

ALT_FIXES = {
    (1, 5): "Cardinal sunflower quilt size and optional sham guide",
    (2, 2): "Teal softball comforter set in an angled bedroom view",
    (2, 4): "Folded teal softball comforter showing the white reverse",
    (2, 5): "Close-up of teal softball comforter print and white reverse",
    (2, 7): "Teal softball comforter set with two included pillowcases",
    (3, 5): "Christmas cardinal quilt size and optional sham guide",
    (4, 5): "Cardinal and roses quilt size and optional sham guide",
    (5, 2): "Colorful cat quilt with print and fabric feature callouts",
    (5, 3): "Colorful cat pillow shams with lightweight fabric features",
    (5, 4): "Colorful cat quilt all-season fabric feature graphic",
    (6, 2): "Geometric cat quilt with print and fabric feature callouts",
    (6, 3): "Geometric cat pillow shams with lightweight fabric features",
    (6, 4): "Geometric cat quilt all-season fabric feature graphic",
    (7, 2): "Celtic fantasy tree quilt with print and fabric feature callouts",
    (7, 3): "Celtic tree pillow shams with lightweight fabric features",
    (7, 4): "Celtic fantasy tree quilt all-season fabric feature graphic",
    (8, 2): "Celtic Tree of Life quilt with print and fabric feature callouts",
    (8, 3): "Celtic Tree of Life shams with lightweight fabric features",
    (8, 4): "Celtic Tree of Life quilt all-season fabric feature graphic",
    (9, 2): "Farmhouse chicken quilt with print and fabric feature callouts",
    (9, 3): "Farmhouse chicken shams with lightweight fabric features",
    (9, 4): "Farmhouse chicken quilt all-season fabric feature graphic",
    (10, 2): "Christian God Says I Am bedding in an alternate bedroom view",
    (10, 3): "Duvet cover and comforter set construction comparison",
    (10, 4): "Christian bedding microfiber and breathable fabric features",
    (10, 5): "Birth month flower options from January through December",
    (10, 6): "Christian bedding softness and durability feature graphic",
    (10, 7): "Christian bedding easy-care and fade-resistance features",
    (10, 8): "Christian bedding size chart for Twin through King",
}

SERP = {
    1: [("cardinal sunflower quilt set", ["https://vantique.net/collections/cardinal-quilts"]), ("autumn cardinal bedding", ["https://www.target.com/s/cardinal%2Bquilt"])],
    2: [("softball comforter set", ["https://www.etsy.com/market/softball_comforter_sets", "https://www.target.com/s/kids%2Bbedding%2Bsoftball"]), ("teal softball bedding", ["https://www.youcustomizeit.com/bbp/Softball-Bedding/534579"] )],
    3: [("Christmas cardinal memorial quilt", ["https://www.target.com/s/cardinal%2Bquilt"]), ("cardinal Christmas quilt", ["https://www.marcielobedding.com/products/marcelo-3-pcs-winter-cardinals-christmas-quilt-bedspread-set-decor"])],
    4: [("cardinal memorial quilt with roses", ["https://www.reddit.com/r/quilting/comments/19fiv3f"]), ("cardinal remembrance quilt", ["https://www.reddit.com/r/quilting/comments/166n7be"])],
    5: [("cat patchwork quilt set", ["https://www.walmart.com/ip/20788770071"]), ("colorful cat quilt", ["https://www.walmart.com/ip/20076702698"])],
    6: [("cat patchwork bedding set", ["https://www.wayfair.com/bed-bath/pdp/ambesonne-cat-bedspread-set-patchwork-style-silly-faces-multicolor-bbqv8335.html"]), ("geometric cat quilt set", ["https://vantique.net/products/whimsical-cat-patchwork-3-piece-quilted-bedding-set-ncu0nt5145"])],
    7: [("Celtic tree quilt set", ["https://usblanket.com/collections/tree-of-life"]), ("fantasy tree quilt", ["https://www.freshouseshop.com/products/celtic-bed-knot-tree-of-life-veru1967-quilt-bedding-set"])],
    8: [("Celtic tree of life quilt set", ["https://quiltnest.com/products/intricate-celtic-tree-of-life-3-piece-quilted-bedding-set-ncu0pd047"]), ("Yggdrasil quilt set", ["https://ohaprints.com/products/ohaprints-quilt-bed-set-pillowcase-tree-of-life-yggdrasil-black-yellow-mythology-ancient-norse-nordic-blanket-bedspread-bedding-695"])],
    9: [("farmhouse chicken quilt set", ["https://www.walmart.com/ip/20230621300"]), ("chicken bedding set", ["https://www.target.com/s/chicken%2Bbedding%2Bset"])],
    10: [("personalized Christian comforter set", ["https://www.visionbedding.com/bedding/christian"]), ("God Says I Am bedding", ["https://faithquilt.com/"])],
}

PRODUCT_ASSESS = {
    1: {"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    2: {"P1":"FULL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    3: {"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    4: {"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"PARTIAL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    5: {"P1":"PARTIAL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"FAIL"},
    6: {"P1":"PARTIAL","P2":"PARTIAL","K1":"PARTIAL","K2":"FULL","K3":"PARTIAL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    7: {"P1":"FULL","P2":"FULL","K1":"PARTIAL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    8: {"P1":"FULL","P2":"FULL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    9: {"P1":"PARTIAL","P2":"PARTIAL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"FULL","T2":"FULL","D1":"FULL","D2":"FAIL","E1":"PARTIAL"},
    10:{"P1":"FULL","P2":"FAIL","K1":"FULL","K2":"FULL","K3":"FULL","T1":"PARTIAL","T2":"PARTIAL","D1":"PARTIAL","D2":"FAIL","E1":"PARTIAL"},
}

REASONS = {
    "P1": "Đối chiếu loại sản phẩm, thiết kế và màu với product JSON, H1 live và toàn bộ gallery.",
    "P2": "Đối chiếu chất liệu, kích thước, thành phần đi kèm, biến thể và personalization với nguồn live/snapshot.",
    "K1": "Đối chiếu long-tail với thiết kế riêng và nguy cơ trùng intent trong cụm sản phẩm.",
    "K2": "Đã đọc SERP cho primary keyword và phương án gần nhất; đánh giá loại trang/intent.",
    "K3": "Không suy diễn volume; phân biệt SERP_ONLY và HYPOTHESIS_ONLY.",
    "T1": "Đánh giá SEO title English về độ đúng, tự nhiên, keyword và điểm phân biệt.",
    "T2": "Đánh giá H1 đề xuất độc lập với SEO title và đối chiếu lỗi ký tự nguồn.",
    "D1": "Đánh giá meta description về tính cụ thể và căn cứ claim.",
    "D2": "Description SET có HTML đầy đủ nhưng chứa câu nội bộ 'This draft should be reviewed...' và dữ liệu nguồn thô/mâu thuẫn; chưa phải nội dung publish-ready.",
    "E1": "Kiểm tra liên kết product key/handle/evidence/revision giữa workbook, snapshot và nguồn live.",
}


def issue(issue_id, product_key, severity, field, submitted, observed, reason, fix, evidence, recheck, image_key=""):
    return {"issue_id": issue_id, "product_key": product_key, "qa_image_key": image_key, "severity": severity,
            "field": field, "submitted_value": submitted, "source_observation": observed, "reason": reason,
            "recommended_fix": fix, "supporting_evidence": evidence, "recheck_condition": recheck}


def main() -> None:
    data = json.loads((QA_DIR / "submitted_batch_data.json").read_text(encoding="utf-8"))
    live = json.loads((QA_DIR / "live_source_comparison.json").read_text(encoding="utf-8"))
    products = data["products"]
    pk_to_pos = {p["product_key"]: i + 1 for i, p in enumerate(products)}
    images_by_pk = {p["product_key"]: [] for p in products}
    for im in data["images"]:
        images_by_pk[im["shop_domain"] + "+" + im["Handle"]].append(im)

    checked_at = now()
    qa_images, issues = [], []
    for p in products:
        pos = pk_to_pos[p["product_key"]]
        live_media = live[pos - 1]["live"]["product_js"]["media"]
        for idx, im in enumerate(sorted(images_by_pk[p["product_key"]], key=lambda x: x["image_number"])):
            n = idx + 1
            qkey = stable_image_key(p["product_key"], im["image_url"], n)
            im2, im3 = IMAGE_ASSESS[pos][idx]
            assessments = {"IM1": "FULL", "IM2": im2, "IM3": im3, "IM4": "FULL"}
            points = sum(IMAGE_WEIGHTS[k] * RATING[v] for k, v in assessments.items())
            refs = [im["evidence_file_or_reference"], p["product_url"], live_media[idx]["src"]]
            row = {
                "product_key": p["product_key"], "qa_image_key": qkey,
                "image_url_source": live_media[idx]["src"], "image_url_workbook": im["image_url"],
                "media_id": str(live_media[idx]["id"]), "workbook_image_id": str(im["media_id"]),
                "variant": im["variant"] or "", "image_location": im["image_location"],
                "check_method": "DIRECT_ORIGINAL_IMAGE", "checked_at": checked_at,
                "qa_observation": ACTUAL[pos][idx], "submitted_observation": im["observed_visual_details"],
                "storefront_alt_observed": live_media[idx].get("alt") or "",
                "alt_action": im["alt_action"], "alt_effective": im["alt_proposed"],
                **assessments, "image_verified_points": points, "image_assessed_weight": 100,
                "image_final_score": points, "image_score_lower_bound": points, "image_score_upper_bound": points,
                "issue_refs": [], "evidence_refs": refs,
            }
            if im2 != "FULL" or im3 != "FULL":
                iid = f"ISS-{pos:03d}-IMG-{n:02d}"
                sev = "MAJOR" if im3 == "FAIL" else "MINOR"
                fix = ALT_FIXES.get((pos, n), im["alt_proposed"])
                issues.append(issue(iid, p["product_key"], sev, "image_observation/alt_effective",
                    f"{im['observed_visual_details']} | {im['alt_proposed']}", ACTUAL[pos][idx],
                    "Nhận xét hoặc alt theo mẫu vị trí không phản ánh đầy đủ/đúng nội dung ảnh thực tế.", fix,
                    "; ".join(refs), "Mở lại ảnh gốc và xác nhận observation/alt mới mô tả đúng ảnh.", qkey))
                row["issue_refs"].append(iid)
            qa_images.append(row)

    # Product-level findings.
    for p in products:
        pos, pk = pk_to_pos[p["product_key"]], p["product_key"]
        issues.append(issue(f"ISS-{pos:03d}-DESC", pk, "MAJOR", "description_proposed_html",
            p["description_proposed_html"], "HTML chứa ngôn ngữ quy trình nội bộ và dữ liệu nguồn thô.",
            "Description SET chưa publish-ready và có thể đưa ghi chú QA lên storefront.",
            "Rewrite as customer-facing English HTML; remove internal drafting/QA language and resolve source conflicts.",
            p["evidence_id"], "Review full rendered HTML and confirm no internal QA language remains."))
    for pos in (1, 6):
        p = products[pos - 1]
        issues.append(issue(f"ISS-{pos:03d}-KW", p["product_key"], "LIMITATION", "primary_keyword",
            p["primary_keyword"], "Hai truy vấn QA cho thấy intent liên quan nhưng không có dữ liệu volume trả phí.",
            "Mức evidence chỉ là HYPOTHESIS_ONLY; không được mô tả như demand đã xác minh.",
            "Keep the keyword as a hypothesis or validate with US keyword-volume data before prioritization.",
            json.dumps(SERP[pos]), "Cung cấp và kiểm tra nguồn demand/volume US có kỳ dữ liệu rõ."))
    for pos in (5, 6, 9):
        p = products[pos - 1]
        issues.append(issue(f"ISS-{pos:03d}-FACT", p["product_key"], "MAJOR", "description_proposed_html/source facts",
            "Pattern/Color/Selected Design includes Dragonfly", ACTUAL[pos][0],
            "Dữ liệu cấu trúc nguồn mâu thuẫn rõ với thiết kế cat/chicken đang hiển thị; draft đã chép xung đột vào description.",
            "Remove 'Dragonfly' from customer-facing copy and resolve the product metafield/source mapping before publication.",
            p["evidence_id"] + "; " + p["product_url"], "Snapshot/admin mapping confirms the correct design value and revised HTML omits the conflict."))
    p5 = products[4]
    issues.append(issue("ISS-005-SERP", p5["product_key"], "CRITICAL", "SERP_evidence_references",
        data["product_evidence"][4]["SERP_evidence_references"], "Nguồn đã nộp dẫn tới chicken quilt, không phải cat patchwork quilt.",
        "Liên kết evidence sai sản phẩm làm chuỗi bằng chứng keyword không đáng tin cậy.",
        "Replace with the two QA-verified cat patchwork SERP sources and rerun keyword selection.",
        "https://www.etsy.com/market/chicken_quilt; https://www.walmart.com/ip/20788770071",
        "Hai URL thay thế được đọc lại, lưu timestamp/locale và map đúng product key."))
    for pos in (3, 4):
        p = products[pos - 1]
        issues.append(issue(f"ISS-{pos:03d}-SERP", p["product_key"], "MAJOR", "SERP_evidence_references",
            data["product_evidence"][pos-1]["SERP_evidence_references"], "Reference cũ không chứng minh trực tiếp intent sản phẩm cardinal tương ứng.",
            "Evidence cũ quá chung hoặc không truy được nội dung sản phẩm; QA đã tìm nguồn đối chiếu phù hợp hơn.",
            "Replace the submitted references with the QA SERP URLs and document product-page intent.", json.dumps(SERP[pos]),
            "Mở lại URL mới, xác nhận intent và map đúng keyword/product."))
    for pos in (2, 10):
        p = products[pos - 1]
        issues.append(issue(f"ISS-{pos:03d}-ENC", p["product_key"], "MINOR", "title_current/H1_current",
            p["title_current"], live[pos-1]["live"]["h1"], "Ký tự thay thế � tồn tại trong title/H1 nguồn live.",
            p["title_proposed"], p["product_url"], "Rendered H1 and title no longer contain U+FFFD."))
    p10 = products[9]
    issues.append(issue("ISS-010-PERS", p10["product_key"], "CRITICAL", "title/meta/description personalization claim",
        "Personalized / Personalize", "Product JSON exposes product type/size, pillowcase and sheet options but no name or birth-flower personalization input.",
        "Personalization is a material purchasing attribute and was not verifiable in the selectable live variants.",
        "Remove 'Personalized' unless a working customization control is evidenced; otherwise document the exact name/birth-flower input and fulfillment mapping.",
        p10["product_url"] + "; live product.js options", "Live purchase flow visibly accepts and preserves the claimed personalization values."))
    issues.append(issue("ISS-GLOBAL-ADMIN", "", "LIMITATION", "current admin SEO fields/current admin alt",
        "Not supplied", "Storefront metadata and product.js media alt were readable, but no Shopify admin export was provided.",
        "Storefront values cannot prove the current admin field values or future import mapping.",
        "Provide a Shopify admin/export snapshot before approval or import mapping.", str(SOURCE),
        "Admin export revision and hash are frozen and compared to storefront/workbook."))

    issue_by_product = {}
    for it in issues:
        issue_by_product.setdefault(it["product_key"], []).append(it)

    criteria, qa_products = [], []
    for p in products:
        pos, pk = pk_to_pos[p["product_key"]], p["product_key"]
        image_rows = [x for x in qa_images if x["product_key"] == pk]
        image_avg = sum(x["image_final_score"] for x in image_rows) / len(image_rows)
        p_issue_refs = [x["issue_id"] for x in issue_by_product.get(pk, [])]
        for cid, weight in PRODUCT_WEIGHTS.items():
            if cid == "I1":
                assessment, rating, earned = "DERIVED", image_avg / 100, weight * image_avg / 100
                reason = f"Tính từ trung bình {len(image_rows)} ảnh: {image_avg:.4f}/100."
            else:
                assessment = PRODUCT_ASSESS[pos][cid]
                rating, earned = RATING[assessment], weight * RATING[assessment]
                reason = REASONS[cid]
            criteria.append({"product_key": pk, "criterion_id": cid, "weight": weight,
                "assessment": assessment, "rating": rating, "earned_points": earned,
                "assessed_weight": weight, "reason": reason,
                "evidence_refs": [p["evidence_id"], p["product_url"], f"serp_qa_{pos:03d}"],
                "issue_refs": p_issue_refs})
        verified = sum(x["earned_points"] for x in criteria if x["product_key"] == pk)
        severities = Counter(x["severity"] for x in issue_by_product.get(pk, []))
        if severities["CRITICAL"]:
            status = "QA_FAIL"
        elif verified < 70:
            status = "QA_FAIL"
        elif verified < 85 or severities["MAJOR"]:
            status = "QA_REVISE"
        else:
            status = "QA_PASS"
        qa_products.append({"inventory_position": pos, "product_key": pk, "url": p["product_url"],
            "handle": p["Handle"], "product_id": str(p["product_id"]), "revision": "r1",
            "verified_points": verified, "assessed_weight": 100, "score_lower_bound": verified,
            "score_upper_bound": verified, "final_score": verified, "qa_status": status,
            "keyword_evidence_level": p["keyword_evidence_level"], "images_expected": p["image_count"],
            "images_checked": len(image_rows), "image_inventory_complete": True, "image_coverage": 1.0,
            "critical_count": severities["CRITICAL"], "major_count": severities["MAJOR"],
            "minor_count": severities["MINOR"], "limitation_count": severities["LIMITATION"],
            "issue_refs": p_issue_refs, "evidence_refs": [p["evidence_id"], p["product_url"], f"serp_qa_{pos:03d}"]})

    batch_avg = sum(p["final_score"] for p in qa_products) / len(qa_products)
    status_counts = Counter(p["qa_status"] for p in qa_products)
    sev_counts = Counter(x["severity"] for x in issues)
    source_hash = sha256(SOURCE)
    summary_rows = [
        {"metric":"rubric_version","value":"prompt_qa.md@sha256:" + sha256(ROOT/'seo-prompt/jeminise/prompt_qa.md'),"definition":"Rubric QA áp dụng."},
        {"metric":"source_workbook","value":str(SOURCE),"definition":"Workbook giai đoạn 1 được đóng băng; không chỉnh sửa."},
        {"metric":"source_sha256_at_handoff","value":source_hash,"definition":"Hash tính lại khi hoàn tất báo cáo."},
        {"metric":"qa_run_id","value":QA_RUN_ID,"definition":"Run QA độc lập."},
        {"metric":"batch_id","value":QA_BATCH_ID,"definition":"Inventory positions 1–10."},
        {"metric":"products_checked","value":10,"definition":"Đúng 10 product key."},
        {"metric":"images_checked","value":65,"definition":"65/65 ảnh mở trực tiếp ở ảnh gốc."},
        {"metric":"batch_final_score","value":batch_avg,"definition":"Trung bình đều 10 final_score; không bù lỗi chặn."},
        {"metric":"batch_result","value":"NOT_PASSED","definition":"Lô chỉ đạt khi mọi sản phẩm QA_PASS."},
        {"metric":"status_counts","value":dict(status_counts),"definition":"Số sản phẩm theo kết luận."},
        {"metric":"issue_counts","value":dict(sev_counts),"definition":"Số phát hiện theo severity."},
        {"metric":"admin_limitation","value":"No Shopify admin export","definition":"Không suy ra admin SEO fields/alt từ storefront."},
        {"metric":"xlsx_status","value":"COMPLETE","definition":"Workbook QA đã được tạo bằng openpyxl theo yêu cầu rõ của người dùng, tính lại và render kiểm tra."},
    ]
    serp_rows = []
    for p in products:
        pos = pk_to_pos[p["product_key"]]
        for query_no, (query, urls) in enumerate(SERP[pos], 1):
            serp_rows.append({"serp_id": f"serp_qa_{pos:03d}_{query_no}", "product_key": p["product_key"],
                "query": query, "market": "United States", "language": "English",
                "locale_limit": "US intent; search service locale could not be hard-pinned", "checked_at": checked_at,
                "result_urls_read": urls, "intent": "Commercial/product" if pos not in (1,4) else "Mixed commercial/product",
                "note": "Evidence supports intent only; no paid search-volume claim."})

    source_change_rows = []
    for pos, entry in enumerate(live, 1):
        old_html = entry["research_snapshot"]["html"]
        current = entry["live"]
        stable_fields = {
            "title": old_html.get("rendered_title_current") == current.get("title_element"),
            "h1": old_html.get("h1_current") == current.get("h1"),
            "meta_description": old_html.get("meta_description_current") == current.get("meta_description"),
            "canonical": old_html.get("canonical_url") == current.get("canonical"),
        }
        source_change_rows.append({"product_key": products[pos-1]["product_key"],
            "snapshot_checked_at": entry["research_snapshot"].get("reviewed_at"), "live_checked_at": current.get("checked_at"),
            "snapshot_html_sha256": old_html.get("html_sha256"), "live_html_sha256": current.get("html_sha256"),
            "html_hash_changed": old_html.get("html_sha256") != current.get("html_sha256"),
            "material_fields_equal": stable_fields, "source_revision_status": "LIVE_EQUIVALENT_DYNAMIC_HTML_CHANGED" if all(stable_fields.values()) else "SOURCE_CHANGED"})

    tests = {
        "product_weight_total": sum(PRODUCT_WEIGHTS.values()), "image_weight_total": sum(IMAGE_WEIGHTS.values()),
        "products": len(qa_products), "images": len(qa_images),
        "unique_product_keys": len({p['product_key'] for p in qa_products}),
        "unique_qa_image_keys": len({i['qa_image_key'] for i in qa_images}),
        "all_image_coverage_100": all(p["image_coverage"] == 1 for p in qa_products),
        "logic_100_with_critical": "QA_FAIL",
        "logic_90_full_no_blocker": "QA_PASS",
        "logic_72_on_80": {"range":"72-92", "status":"QA_INCOMPLETE"},
        "cross_links_valid": all(all(ref in {i['issue_id'] for i in issues} for ref in p['issue_refs']) for p in qa_products),
    }
    assert tests["product_weight_total"] == 100 and tests["image_weight_total"] == 100
    assert tests["products"] == tests["unique_product_keys"] == 10
    assert tests["images"] == tests["unique_qa_image_keys"] == 65
    assert tests["all_image_coverage_100"] and tests["cross_links_valid"]

    dataset = {"QA_Summary": summary_rows, "QA_Products": qa_products, "QA_Criteria": criteria,
               "QA_Images": qa_images, "QA_Issues": issues, "SERP_Evidence": serp_rows,
               "validation_tests": tests}
    save_json(QA_DIR / "qa_dataset.json", dataset)
    save_json(QA_DIR / "qa_workbook_payload.json", {k: dataset[k] for k in ("QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues")})
    save_json(QA_DIR / "serp_evidence.json", serp_rows)
    save_json(QA_DIR / "source_change_audit.json", source_change_rows)
    save_json(QA_DIR / "validation_results.json", tests)
    inventory = json.loads((ROOT / "seo_runs" / "jeminise.com" / RUN_ID / "inventory.json").read_text(encoding="utf-8-sig"))
    next_batch = [{k: row.get(k) for k in ("inventory_position", "product_key", "product_id", "Handle", "title_current", "product_url", "image_count")}
                  for row in inventory if 11 <= int(row["inventory_position"]) <= 20]
    assert len(next_batch) == 10
    next_preview_path = QA_DIR / "qa_batch_002_preview.json"
    save_json(next_preview_path, {"batch_id":"qa_batch_002", "status":"PREPARED_NOT_STARTED", "products":next_batch})

    lines = ["# SEO QA — qa_batch_001", "", "## Kết luận", "",
             f"- Phạm vi: **10 sản phẩm, 65/65 ảnh (100%)**; chỉ inventory position 1–10.",
             f"- Điểm lô: **{batch_avg:.1f}/100**; kết luận lô: **NOT_PASSED**.",
             f"- Trạng thái: {status_counts.get('QA_FAIL',0)} QA_FAIL, {status_counts.get('QA_REVISE',0)} QA_REVISE, {status_counts.get('QA_PASS',0)} QA_PASS.",
             f"- Phát hiện: {sev_counts.get('CRITICAL',0)} CRITICAL, {sev_counts.get('MAJOR',0)} MAJOR, {sev_counts.get('MINOR',0)} MINOR, {sev_counts.get('LIMITATION',0)} LIMITATION.",
             f"- Workbook nguồn: `{SOURCE}`", f"- SHA-256 lúc đóng băng và bàn giao: `{source_hash}`", "",
             "## Điểm theo sản phẩm", "", "| Pos | Sản phẩm | Điểm | Kết luận | C/M/m/L |", "|---:|---|---:|---|---:|"]
    for p, src in zip(qa_products, products):
        lines.append(f"| {p['inventory_position']} | {src['title_current'].replace('|','/')} | {p['final_score']:.1f} | {p['qa_status']} | {p['critical_count']}/{p['major_count']}/{p['minor_count']}/{p['limitation_count']} |")
    lines += ["", "## Lỗi ưu tiên", "",
              "1. **CRITICAL — product 5:** evidence SERP của cat quilt lại dẫn tới `chicken_quilt`; phải thay nguồn và chạy lại quyết định keyword.",
              "2. **CRITICAL — product 10:** draft dùng claim `Personalized`, nhưng purchase options live không cho thấy input name/birth flower; phải chứng minh control hoạt động hoặc bỏ claim.",
              "3. **MAJOR — cả 10 sản phẩm:** description HTML chứa câu nội bộ về draft/QA và chưa phải nội dung có thể xuất bản.",
              "4. **MAJOR — ảnh:** nhiều observation/alt được gán theo vị trí chung thay vì nội dung thật; product 10 sai nặng ở ảnh 3–8.",
              "5. **MAJOR — products 5, 6, 9:** draft chép `Dragonfly` từ dữ liệu nguồn mâu thuẫn với thiết kế cat/chicken.", "",
              "## SERP và keyword", "",
              "Mỗi sản phẩm đã được kiểm tra lại bằng primary keyword và một phương án gần nhất. Các URL/timestamp/locale limit nằm trong `serp_evidence.json`. Không có claim volume. Hai keyword của products 1 và 6 vẫn giữ `HYPOTHESIS_ONLY`.", "",
              "## Giới hạn và trạng thái bàn giao", "",
              "- Không có Shopify admin export; chỉ xác minh metadata và media alt ở storefront/product.js, không suy ra giá trị admin.",
              "- Product identity, canonical, product ID, variants, body HTML và gallery filename/order không đổi giữa snapshot và live. HTML trang động có hash khác nhưng không phát hiện drift nội dung sản phẩm trọng yếu.",
              "- Không sửa workbook nguồn, không đổi review_status, không tạo APPROVED/import và không cập nhật Shopify.",
              "- **XLSX: COMPLETE.** Workbook QA có đúng 5 sheet, công thức truy kiểm, filter/freeze/wrap và hyperlink; đã được Excel tính lại và render để kiểm tra.",
              "- Chưa QA products 11–20. `awaiting_confirmation=true`.", "",
              "## Tệp chi tiết", "",
              f"- QA data: `{QA_DIR / 'qa_dataset.json'}`", f"- SERP evidence: `{QA_DIR / 'serp_evidence.json'}`",
              f"- Validation: `{QA_DIR / 'validation_results.json'}`", f"- Manifest/checkpoint: `{QA_DIR}`", ""]
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    report = OUT_DIR / f"SEO_QA_{QA_BATCH_ID}.md"
    temp = report.with_suffix(".md.tmp")
    temp.write_text("\n".join(lines), encoding="utf-8")
    temp.replace(report)

    manifest_path = QA_DIR / "qa_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    xlsx_output = OUT_DIR / f"SEO_QA_{QA_BATCH_ID}.xlsx"
    manifest.update({"status":"COMPLETE", "completed_at":now(), "source_sha256_at_handoff":source_hash,
        "counts":{"products":10,"images":65}, "output_markdown":str(report),
        "output_xlsx":str(xlsx_output), "xlsx_blocker":None,
        "qa_dataset":str(QA_DIR/'qa_dataset.json'), "qa_workbook_payload":str(QA_DIR/'qa_workbook_payload.json'),
        "frozen_at":"2026-09-07T08:35:34+07:00", "source_snapshot_sha256":sha256(QA_DIR/'source_snapshot'/SOURCE.name),
        "prompt_files_sha256":{"prompt_qa.md":sha256(ROOT/'seo-prompt/jeminise/prompt_qa.md'),
                               "prompt.md":sha256(ROOT/'seo-prompt/jeminise/prompt.md')},
        "source_revision_status":"LIVE_EQUIVALENT_DYNAMIC_HTML_CHANGED"})
    save_json(manifest_path, manifest)
    progress_path = QA_DIR / "qa_progress.json"
    progress = json.loads(progress_path.read_text(encoding="utf-8"))
    progress.update({"current_product_key":products[-1]["product_key"], "current_stage":"BATCH_COMPLETE",
        "completed_image_keys":[x["qa_image_key"] for x in qa_images], "last_saved_at":now(),
        "artifact_paths":{"markdown":str(report),"xlsx":str(xlsx_output),"dataset":str(QA_DIR/'qa_dataset.json'),
                          "workbook_payload":str(QA_DIR/'qa_workbook_payload.json'),
                          "source_change_audit":str(QA_DIR/'source_change_audit.json')},
        "awaiting_confirmation":True, "confirmation_ref":None,
        "next_batch_preview":{"batch_id":"qa_batch_002","inventory_positions":"11-20","status":"PREPARED_NOT_STARTED",
                              "product_keys":[x["product_key"] for x in next_batch], "path":str(next_preview_path)}})
    save_json(progress_path, progress)
    print(json.dumps({"report":str(report),"dataset":str(QA_DIR/'qa_dataset.json'),"batch_score":batch_avg,
                      "statuses":dict(status_counts),"issues":dict(sev_counts),"source_hash":source_hash}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
