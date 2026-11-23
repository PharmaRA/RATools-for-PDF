"""Read-only checks for the built manual: pages, links, anchors and search assets."""
from __future__ import annotations

import argparse
from html.parser import HTMLParser
import json
import os
from pathlib import Path
import re
from urllib.parse import unquote, urljoin, urlsplit

WEBSITE = Path(__file__).resolve().parents[1]
ROOT = WEBSITE.parent


class Page(HTMLParser):
    def __init__(self, text: str):
        super().__init__(convert_charrefs=True)
        self.ids: set[str] = set()
        self.links: list[str] = []
        self.feed(text)

    def handle_starttag(self, tag, attrs):
        values = dict(attrs)
        if values.get("id"):
            self.ids.add(values["id"])
        for key in ("href", "src"):
            if values.get(key):
                self.links.append(values[key])
        if values.get("srcset"):
            self.links.extend(item.strip().split()[0] for item in values["srcset"].split(",") if item.strip())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base-path", default=os.environ.get("DOCS_BASE", "/RATools-for-PDF"))
    args = parser.parse_args()
    base = "/" + args.base_path.strip("/") if args.base_path.strip("/") else ""
    dist = WEBSITE / "dist"
    errors: list[str] = []
    pages = {file.resolve(): Page(file.read_text(encoding="utf-8")) for file in dist.rglob("*.html")}
    if not pages:
        print("FAIL: no built HTML. Run npm run build first.")
        return 1
    manual = sorted((ROOT / "docs/manual").glob("*.md"))
    for source in manual:
        route = dist / "manual" / ("" if source.stem == "index" else source.stem) / "index.html"
        if route.resolve() not in pages:
            errors.append(f"Missing manual route: {source.name}")
    for expected in ("index.html", "404.html", "pagefind/pagefind.js"):
        if not (dist / expected).is_file():
            errors.append(f"Missing required output: {expected}")
    if not list((dist / "pagefind").glob("*.pf_meta")):
        errors.append("Missing Pagefind language index")
    local_links = 0
    for file, page in pages.items():
        relative = file.relative_to(dist.resolve()).as_posix()
        route = relative.removesuffix("index.html") if relative.endswith("index.html") else relative
        current_url = f"https://docs.invalid{base}/{route}"
        for href in page.links:
            parsed = urlsplit(href)
            if parsed.scheme or parsed.netloc:
                continue
            resolved = urlsplit(urljoin(current_url, href))
            if base and not (resolved.path == base or resolved.path.startswith(base + "/")):
                errors.append(f"Escaped base: {relative} -> {href}")
                continue
            subpath = unquote(resolved.path[len(base):]).lstrip("/")
            target = (dist / subpath).resolve()
            if not target.is_relative_to(dist.resolve()):
                errors.append(f"Escaped output: {relative} -> {href}")
                continue
            if target.is_dir():
                target /= "index.html"
            local_links += 1
            if not target.is_file():
                errors.append(f"Missing target: {relative} -> {href}")
            elif resolved.fragment and target.suffix == ".html":
                if unquote(resolved.fragment) not in pages[target].ids:
                    errors.append(f"Missing anchor: {relative} -> {href}")
    catalog = (ROOT / "ratools_pdf/config/rules_catalog.py").read_text(encoding="utf-8")
    titles = re.findall(r'"id": "[^"]+", "title": "([^"]+)"', catalog)
    index = (ROOT / "docs/manual/feature-index.md").read_text(encoding="utf-8")
    for title in titles:
        if title not in index:
            errors.append(f"Undocumented rule: {title}")
    print(json.dumps({"manual_pages": len(manual), "html_pages": len(pages), "local_links_and_assets": local_links,
                      "rule_coverage": f"{len(titles) - sum(title not in index for title in titles)}/{len(titles)}",
                      "base": base or "/", "errors": errors}, ensure_ascii=False, indent=2))
    return int(bool(errors))


if __name__ == "__main__":
    raise SystemExit(main())
