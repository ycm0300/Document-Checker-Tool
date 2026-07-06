import re

from config.common_rules import clean_text


FIGURE_REFERENCE_PATTERN = re.compile(r"\bFigure\s+(\d+(?:[-.]\d+)*)\b", re.IGNORECASE)
SECTION_NUMBER_PATTERN = re.compile(r"^(\d+)(?:\.\d+)*\s+")


def _figure_number_key(figure_number):
    return re.sub(r"[-.]", "", figure_number)


def _get_section_chapter_number(item):
    section_key = item.get("section_key") or item.get("heading", "")
    for part in str(section_key).split(">"):
        match = SECTION_NUMBER_PATTERN.match(clean_text(part))
        if match:
            return match.group(1)
    return None


def _display_figure_number(figure_number, item):
    if "-" in figure_number:
        return figure_number

    if "." in figure_number:
        return figure_number.replace(".", "-")

    chapter_number = _get_section_chapter_number(item)
    if not chapter_number or not figure_number.startswith(chapter_number):
        return figure_number

    rest = figure_number[len(chapter_number):]
    if not rest or rest == "0":
        return figure_number

    return f"{chapter_number}-{rest}"


def _iter_text_figure_mentions(text):
    for match in FIGURE_REFERENCE_PATTERN.finditer(text):
        yield match.group(1)


def _iter_figure_mentions(items):
    for item in items:
        if item.get("is_toc"):
            continue

        text = clean_text(item.get("text", ""))
        xml_text = clean_text(item.get("xml_text", ""))
        if not text:
            continue

        figure_numbers = list(_iter_text_figure_mentions(text))
        if not figure_numbers and xml_text and xml_text != text:
            figure_numbers = list(_iter_text_figure_mentions(xml_text))

        for figure_number in figure_numbers:
            yield _figure_number_key(figure_number), figure_number, item


def check_figure_references(items):
    """检查 Figure 编号是否至少出现两次：题注本身 + 正文引用。"""
    figure_mentions = {}

    for figure_key, figure_number, item in _iter_figure_mentions(items):
        figure_mentions.setdefault(figure_key, []).append((figure_number, item))

    issues = []
    for _, mentions in figure_mentions.items():
        if len(mentions) > 1:
            continue

        figure_number, item = mentions[0]
        display_figure_number = _display_figure_number(figure_number, item)
        issues.append({
            "file_name": item["file_name"],
            "issue_type": "引用问题",
            "heading": item["heading"],
            "section_key": item.get("section_key", item["heading"]),
            "location": item["location"],
            "content": item["text"],
            "issue": f"Figure {display_figure_number} 仅出现一次，可能未被正文引用",
        })

    return issues
