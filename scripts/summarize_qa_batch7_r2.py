from pathlib import Path
import json


qa = Path(__file__).resolve().parents[1] / "seo_runs/jeminise.com/20260906_234129/qa/20260907_164400"
data = json.loads((qa / "submitted_batch_data.json").read_text(encoding="utf-8"))
for position, product in enumerate(data["products"], 61):
    print("\n", position, product.get("title_proposed"))
    print("PRIMARY:", product.get("primary_keyword"))
    print("META:", product.get("meta_description_seo"))
    print("BODY:", product.get("description_proposed_html"))
    product_images = sorted(
        (row for row in data["images"] if row["Handle"] == product["Handle"]),
        key=lambda row: row["image_number"],
    )
    for image in product_images:
        print("IMG", image["image_number"], "OBS:", image["observed_visual_details"], "ALT:", image["alt_proposed"])
