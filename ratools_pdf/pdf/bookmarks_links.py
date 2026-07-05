import csv
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from urllib.parse import unquote, urlparse

import fitz

from ratools_pdf.config.paths import get_resource_path
from ratools_pdf.pdf.font_embedding_providers import get_font_embedding_provider


def _processor_cls():
    from ratools_pdf.pdf.processor import PDFProcessor
    return PDFProcessor


def export_bookmarks(pdf_path, csv_path):
    """Export bookmarks to CSV with enough action data for round-trip import."""
    doc = fitz.open(pdf_path)
    toc = doc.get_toc(simple=False)
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['Level', 'Title', 'Page', 'Kind', 'Uri', 'File', 'TargetPage', 'ToX', 'ToY', 'Zoom', 'NewWindow'])
        for item in toc:
            lvl, title, page, dest = item
            if not isinstance(dest, dict):
                dest = {}
            point = dest.get('to')
            writer.writerow([
                lvl,
                title,
                page,
                dest.get('kind', fitz.LINK_NONE),
                dest.get('uri', ''),
                dest.get('file', ''),
                dest.get('page', ''),
                getattr(point, 'x', ''),
                getattr(point, 'y', ''),
                dest.get('zoom', ''),
                dest.get('newWindow', False),
            ])
    doc.close()


def import_bookmarks(pdf_path, csv_path, output_path):
    """Import bookmarks from CSV while remaining compatible with the old 3-column format."""
    doc = fitz.open(pdf_path)
    new_toc = []

    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                lvl = int(row.get('Level', 1))
                title = row.get('Title', '')
                page = int(row.get('Page', 1))
                page = max(1, min(page, doc.page_count))

                kind_text = str(row.get('Kind', '') or '').strip()
                if not kind_text:
                    new_toc.append([lvl, title, page])
                    continue

                kind = int(kind_text)
                if kind == fitz.LINK_URI:
                    dest = {
                        'kind': fitz.LINK_URI,
                        'uri': row.get('Uri', ''),
                    }
                elif kind == fitz.LINK_GOTO:
                    target_page = int(row.get('TargetPage', page - 1) or (page - 1))
                    dest = {
                        'kind': fitz.LINK_GOTO,
                        'page': max(0, min(target_page, doc.page_count - 1)),
                        'to': fitz.Point(float(row.get('ToX', 72.0) or 72.0), float(row.get('ToY', 36.0) or 36.0)),
                        'zoom': float(row.get('Zoom', 0.0) or 0.0),
                    }
                elif kind == fitz.LINK_GOTOR:
                    target_page = int(row.get('TargetPage', 0) or 0)
                    dest = {
                        'kind': fitz.LINK_GOTOR,
                        'file': row.get('File', ''),
                        'page': max(0, target_page),
                        'to': fitz.Point(float(row.get('ToX', 72.0) or 72.0), float(row.get('ToY', 36.0) or 36.0)),
                        'zoom': float(row.get('Zoom', 0.0) or 0.0),
                        'newWindow': str(row.get('NewWindow', 'false')).lower() == 'true',
                    }
                elif kind == fitz.LINK_LAUNCH:
                    dest = {
                        'kind': fitz.LINK_LAUNCH,
                        'file': row.get('File', ''),
                        'newWindow': str(row.get('NewWindow', 'false')).lower() == 'true',
                    }
                else:
                    dest = {'kind': fitz.LINK_NONE}

                new_toc.append([lvl, title, page, dest])
            except (ValueError, TypeError):
                continue

    if new_toc:
        for i in range(len(new_toc)):
            if i == 0:
                new_toc[i][0] = 1
            else:
                prev_lvl = new_toc[i - 1][0]
                if new_toc[i][0] > prev_lvl + 1:
                    new_toc[i][0] = prev_lvl + 1

    doc.set_toc(new_toc)
    doc.save(output_path, garbage=2, deflate=True, use_objstms=1)
    doc.close()


# 导出与导入超链接 (JSON)


def export_links(pdf_path, json_path):
    """Export links to JSON with enough destination data for round-trip import."""
    doc = fitz.open(pdf_path)
    all_links = []

    for page in doc:
        for link in page.get_links():
            rect = link['from']
            target_point = link.get('to')
            link_dict = {
                'page_index': page.number,
                'rect': [rect.x0, rect.y0, rect.x1, rect.y1],
                'kind': link.get('kind', fitz.LINK_NONE),
                'uri': link.get('uri', ''),
                'file': link.get('file', ''),
                'target_page': link.get('page', 0),
                'zoom': link.get('zoom', 0.0),
                'to': [getattr(target_point, 'x', 0.0), getattr(target_point, 'y', 0.0)] if target_point else None,
                'new_window': bool(link.get('newWindow', False)),
            }
            all_links.append(link_dict)

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(all_links, f, indent=4, ensure_ascii=False)
    doc.close()


def import_links(pdf_path, json_path, output_path):
    """Rebuild links from JSON while preserving internal destinations and new-window flags."""
    doc = fitz.open(pdf_path)
    link_file_kind = getattr(fitz, "LINK_FILE", None)

    with open(json_path, 'r', encoding='utf-8') as f:
        links_data = json.load(f)

    for page in doc:
        for link in page.get_links():
            page.delete_link(link)

    for ld in links_data:
        p_idx = ld.get('page_index', 0)
        if 0 <= p_idx < doc.page_count:
            page = doc[p_idx]
            rect = fitz.Rect(ld['rect'])
            kind = ld['kind']

            new_link = {"kind": kind, "from": rect}
            if kind == fitz.LINK_URI:
                new_link["uri"] = ld.get('uri', '')
            elif link_file_kind is not None and kind == link_file_kind:
                new_link["file"] = ld.get('file', '')
            elif kind in [fitz.LINK_GOTO, fitz.LINK_GOTOR]:
                new_link["page"] = ld.get('target_page', 0)
                new_link["zoom"] = ld.get('zoom', 0.0)
                to_value = ld.get('to')
                if isinstance(to_value, (list, tuple)) and len(to_value) >= 2:
                    new_link["to"] = fitz.Point(float(to_value[0]), float(to_value[1]))
                if kind == fitz.LINK_GOTOR:
                    new_link["file"] = ld.get('file', '')
                    new_link["newWindow"] = bool(ld.get('new_window', False))
            elif kind == fitz.LINK_LAUNCH:
                new_link["file"] = ld.get('file', '')
                new_link["newWindow"] = bool(ld.get('new_window', False))

            try:
                page.insert_link(new_link)
            except Exception:
                pass

    doc.save(output_path, garbage=2, deflate=True, use_objstms=1)
    doc.close()


