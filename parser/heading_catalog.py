import re

from openpyxl import load_workbook

from config.common_rules import HEADING_CATALOG_EXCEL_PREFIX, OUTPUT_DIR, clean_text


def split_heading_number_and_title(heading_text):
    """把“1.2 Dashboard”拆成编号和标题名称；如果没有编号，则编号为空。"""
    text = clean_text(heading_text)
    match = re.match(r"^(\d+(?:\.\d+)*)\s+(.+)$", text)
    if match:
        return match.group(1), match.group(2)
    return "", text


def get_heading_rows(items):
    """从某个文档的读取结果中提取标题行。"""
    rows = []
    for item in items:
        if item.get("is_toc") or not item.get("is_heading"):
            continue

        heading_text = item.get("heading", "")
        heading_number, heading_title = split_heading_number_and_title(heading_text)
        heading_level = item.get("heading_level")
        heading_level_text = f"Heading {heading_level}" if heading_level else "未识别"

        rows.append({
            "heading_number": heading_number,
            "heading_title": heading_title,
            "heading_level": heading_level_text,
            "location": item.get("location", ""),
            "full_heading": heading_text,
        })

    return rows


def find_latest_heading_catalog():
    files = list(OUTPUT_DIR.glob(f"{HEADING_CATALOG_EXCEL_PREFIX}*.xlsx"))
    if not files:
        files = list(OUTPUT_DIR.parent.glob(f"{HEADING_CATALOG_EXCEL_PREFIX}*.xlsx"))
    return max(files, key=lambda p: p.stat().st_mtime)


def load_heading_catalog_workbook(file_path):
    return load_workbook(file_path)
