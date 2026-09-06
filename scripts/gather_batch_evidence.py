import argparse
import csv
import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
SHOP = "jeminise.com"
RUN_ID = "20260906_234129"
RUN_DIR = ROOT / "seo_runs" / SHOP / RUN_ID
PRODUCT_EVIDENCE_DIR = RUN_DIR / "evidence" / "products"
IMAGE_EVIDENCE_DIR = RUN_DIR / "evidence" / "images"
CONTACT_DIR = RUN_DIR / "evidence" / "contact_sheets"
INV_PATH = RUN_DIR / "inventory.csv"
RAW_PATH = RUN_DIR / "products_all_raw.json"
PROGRESS_PATH = RUN_DIR / "progress.json"


def clean_filename(value: str) -> str:
    value = re.sub(r"[^a-zA-Z0-9._-]+", "-", value).strip("-")
    return value[:140] or "item"


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def fetch_html(url: str) -> dict:
    response = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0 SEO evidence bot"})
    response.raise_for_status()
    html = response.text
    soup = BeautifulSoup(html, "html.parser")
    h1 = soup.find("h1")
    title = soup.find("title")
    meta_desc = soup.find("meta", attrs={"name": "description"})
    canonical = soup.find("link", attrs={"rel": "canonical"})
    og_title = soup.find("meta", attrs={"property": "og:title"})
    og_desc = soup.find("meta", attrs={"property": "og:description"})
    return {
        "url": url,
        "status_code": response.status_code,
        "html_sha256": sha256_bytes(response.content),
        "rendered_title_current": title.get_text(strip=True) if title else "",
        "h1_current": h1.get_text(" ", strip=True) if h1 else "",
        "meta_description_current": meta_desc.get("content", "").strip() if meta_desc else "",
        "canonical_url": canonical.get("href", "").strip() if canonical else "",
        "og_title": og_title.get("content", "").strip() if og_title else "",
        "og_description": og_desc.get("content", "").strip() if og_desc else "",
        "html_file": "",
    }, html


def load_products_by_id():
    products = json.loads(RAW_PATH.read_text(encoding="utf-8-sig"))
    return {str(p["id"]): p for p in products}


def make_contact_sheet(image_records, output_path: Path, title: str):
    thumbs = []
    for rec in image_records:
        try:
            img = Image.open(rec["local_path"]).convert("RGB")
            img.thumbnail((420, 420))
            canvas = Image.new("RGB", (460, 520), "white")
            x = (460 - img.width) // 2
            canvas.paste(img, (x, 20))
            draw = ImageDraw.Draw(canvas)
            label = f"#{rec['position']} {rec['width']}x{rec['height']}"
            draw.text((14, 452), label, fill="black")
            draw.text((14, 476), Path(urlparse(rec["src"]).path).name[:55], fill="black")
            thumbs.append(canvas)
        except Exception as exc:
            rec["download_status"] = f"ERROR_CONTACT_SHEET: {exc}"
    if not thumbs:
        return
    cols = 2
    rows = (len(thumbs) + cols - 1) // cols
    header_h = 70
    sheet = Image.new("RGB", (cols * 460, header_h + rows * 520), "white")
    draw = ImageDraw.Draw(sheet)
    draw.text((14, 18), title[:120], fill="black")
    for idx, thumb in enumerate(thumbs):
        x = (idx % cols) * 460
        y = header_h + (idx // cols) * 520
        sheet.paste(thumb, (x, y))
    sheet.save(output_path, quality=92)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch-number", type=int, default=1)
    args = parser.parse_args()
    batch_number = args.batch_number
    batch_id = f"batch_{batch_number:03d}"
    start = (batch_number - 1) * 10
    end = start + 10

    PRODUCT_EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
    IMAGE_EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
    CONTACT_DIR.mkdir(parents=True, exist_ok=True)

    with INV_PATH.open(newline="", encoding="utf-8-sig") as f:
        inventory = list(csv.DictReader(f))
    batch = inventory[start:end]
    products_by_id = load_products_by_id()
    now = datetime.now(timezone.utc).isoformat()
    summaries = []

    for row in batch:
        product_id = row["product_id"]
        product = products_by_id[product_id]
        handle = row["Handle"]
        slug = clean_filename(f"{row['inventory_position']}_{handle}")
        product_dir = PRODUCT_EVIDENCE_DIR / slug
        image_dir = IMAGE_EVIDENCE_DIR / slug
        product_dir.mkdir(parents=True, exist_ok=True)
        image_dir.mkdir(parents=True, exist_ok=True)

        try:
            html_data, html = fetch_html(row["product_url"])
        except Exception as exc:
            html_data = {
                "url": row["product_url"],
                "status_code": "ERROR",
                "html_sha256": "",
                "rendered_title_current": "",
                "h1_current": "",
                "meta_description_current": "",
                "canonical_url": "",
                "og_title": "",
                "og_description": "",
                "html_file": "",
                "fetch_error": str(exc),
            }
            html = ""
        html_path = product_dir / "page.html"
        html_path.write_text(html, encoding="utf-8")
        html_data["html_file"] = str(html_path.relative_to(ROOT))

        product_json_path = product_dir / "product.json"
        product_json_path.write_text(json.dumps(product, ensure_ascii=False, indent=2), encoding="utf-8")

        image_records = []
        for image in product.get("images", []):
            src = image.get("src", "")
            position = image.get("position", "")
            ext = Path(urlparse(src).path).suffix or ".jpg"
            image_path = image_dir / f"image_{int(position):02d}_{image.get('id')}{ext}"
            rec = {
                "image_id": str(image.get("id", "")),
                "position": position,
                "src": src,
                "width": image.get("width", ""),
                "height": image.get("height", ""),
                "variant_ids": ",".join(str(v) for v in image.get("variant_ids", [])),
                "local_path": str(image_path.relative_to(ROOT)),
                "download_status": "DOWNLOADED",
                "sha256": "",
            }
            try:
                r = requests.get(src, timeout=45, headers={"User-Agent": "Mozilla/5.0 SEO image evidence bot"})
                r.raise_for_status()
                image_path.write_bytes(r.content)
                rec["sha256"] = sha256_bytes(r.content)
            except Exception as exc:
                rec["download_status"] = f"ERROR: {exc}"
            image_records.append(rec)

        contact_path = CONTACT_DIR / f"{slug}.jpg"
        make_contact_sheet(image_records, contact_path, f"{row['inventory_position']}. {row['title_current']}")

        summary = {
            "reviewed_at": now,
            "inventory_row": row,
            "html": html_data,
            "product_json_ref": str(product_json_path.relative_to(ROOT)),
            "body_html": product.get("body_html", ""),
            "options": product.get("options", []),
            "variants": product.get("variants", []),
            "images": image_records,
            "contact_sheet": str(contact_path.relative_to(ROOT)),
        }
        summary_path = product_dir / "evidence_summary.json"
        summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        summaries.append(summary)

    batch_summary_path = RUN_DIR / f"{batch_id}_evidence_summary.json"
    batch_summary_path.write_text(json.dumps(summaries, ensure_ascii=False, indent=2), encoding="utf-8")

    progress = json.loads(PROGRESS_PATH.read_text(encoding="utf-8-sig"))
    progress.update({
        "market": "United States",
        "seo_language": "English",
        "batch_id": batch_id,
        "batch_status": "EVIDENCE_GATHERED",
        "current_stage": "EVIDENCE_GATHERED",
        "last_saved_at": datetime.now().astimezone().isoformat(),
    })
    progress.setdefault("artifact_paths", {})["batch_evidence_summary"] = str(batch_summary_path.relative_to(ROOT))
    PROGRESS_PATH.write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Gathered evidence for {len(summaries)} products")
    print(batch_summary_path.relative_to(ROOT))
    for s in summaries:
        row = s["inventory_row"]
        print(f"{row['inventory_position']}. {row['title_current']} | images={len(s['images'])} | contact={s['contact_sheet']}")


if __name__ == "__main__":
    main()
