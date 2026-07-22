import re

from config.common_rules import clean_text


FIGURE_REFERENCE_PATTERN = re.compile(r"\bFigure\s+(\d+(?:[-.]\d+)*)\b", re.IGNORECASE)
SECTION_NUMBER_PATTERN = re.compile(r"^(\d+)(?:\.\d+)*\s+")


def _figure_number_key(figure_number):
    return tuple(re.split(r"[-.]", figure_number))


def _is_figure_caption(figure_number, item):
    """优先根据 Word 样式识别题注，样式缺失时再使用文本位置判断。"""
    style_name = str(item.get("paragraph_style", "")).casefold()
    if "caption" in style_name or "题注" in style_name:
        return True

    text = clean_text(item.get("text", ""))
    if re.match(r"^Figure\s*(?:[-.]?\s*)", text, re.IGNORECASE):
        normalized_number = re.escape(figure_number).replace(r"\-", "[-.]").replace(r"\.", "[-.]")
        return bool(re.match(rf"^Figure\s+{normalized_number}\b", text, re.IGNORECASE)) or text.startswith("Figure -")
    return False


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
    """分别识别 Figure 题注和正文引用，并按确定性输出问题或提示。"""
    figure_mentions = {}

    for figure_key, figure_number, item in _iter_figure_mentions(items):
        figure_mentions.setdefault(figure_key, []).append((figure_number, item))

    issues = []
    for _, mentions in figure_mentions.items():
        captions = [mention for mention in mentions if _is_figure_caption(*mention)]
        references = [mention for mention in mentions if not _is_figure_caption(*mention)]
        if captions and references:
            continue

        figure_number, item = mentions[0]
        display_figure_number = _display_figure_number(figure_number, item)
        if captions:
            issue_type = "引用提示"
            issue = f"Figure {display_figure_number} 存在图题，但未检测到正文中的显式编号引用"
        else:
            issue_type = "引用问题"
            issue = f"正文引用了 Figure {display_figure_number}，但未检测到对应图题"

        issues.append({
            "file_name": item["file_name"],
            "issue_type": issue_type,
            "heading": item["heading"],
            "section_key": item.get("section_key", item["heading"]),
            "location": item["location"],
            "content": item["text"],
            "issue": issue,
        })

    return issues
