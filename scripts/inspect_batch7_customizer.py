import concurrent.futures
import html
import json
import re
from pathlib import Path

import requests


ROOT = Path(__file__).resolve().parents[1]
QA_DIR = ROOT / "seo_runs/jeminise.com/20260906_234129/qa/20260907_104027"
START_POSITION = 61


def fetch(url):
    return requests.get(url, timeout=30).text


def main():
    rows = json.loads((QA_DIR / "live_source_comparison.json").read_text(encoding="utf-8"))
    urls = [row["inventory"]["product_url"] for row in rows]
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as pool:
        pages = list(pool.map(fetch, urls))
    output = []
    for position, (url, page) in enumerate(zip(urls, pages), START_POSITION):
        match = re.search(r'data-amzcustom-root[^>]*data-config="([^"]+)"', page)
        config = json.loads(html.unescape(match.group(1))) if match else {}
        config_text = json.dumps(config, ensure_ascii=False)
        personalization_nodes = []
        def walk(node):
            if isinstance(node, dict):
                rendered = json.dumps(node, ensure_ascii=False)
                if re.search(r"name|number", rendered, re.I) and any(
                    key in node for key in ("label", "name", "type", "id", "maxLength")
                ):
                    personalization_nodes.append(node)
                for value in node.values():
                    walk(value)
            elif isinstance(node, list):
                for value in node:
                    walk(value)
        walk(config)
        output.append({
            "inventory_position": position,
            "url": url,
            "customizer_root_present": bool(match),
            "config_bytes": len(config_text.encode("utf-8")),
            "surface_count": len(config.get("surfaces", [])),
            "name_mentions": len(re.findall("name", config_text, re.I)),
            "number_mentions": len(re.findall("number", config_text, re.I)),
            "text_or_input_mentions": len(re.findall("text|input", config_text, re.I)),
            "upload_endpoint_present": '/apps/amazon-customizer/upload' in page,
            "customize_button_present": 'class="amzcustom-open"' in page,
            "personalization_nodes": personalization_nodes[:6],
        })
    rendered = json.dumps(output, ensure_ascii=False, indent=2)
    (QA_DIR / "customizer_audit.json").write_text(rendered, encoding="utf-8")
    print(rendered)


if __name__ == "__main__":
    main()
