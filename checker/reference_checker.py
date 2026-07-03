import re

from config.common_rules import clean_text


FIGURE_REFERENCE_PATTERN = re.compile(r"\bFigure\s+(\d+(?:[-.]\d+)*)\b", re.IGNORECASE)


def _iter_figure_mentions(items):
    for item in items:
        if item.get("is_toc"):
            continue

        text = clean_text(item.get("text", ""))
        if not text:
            continue

        for match in FIGURE_REFERENCE_PATTERN.finditer(text):
            yield match.group(1), item


def check_figure_references(items):
    """检查 Figure 编号是否至少出现两次：题注本身 + 正文引用。"""
    figure_mentions = {}

    for figure_number, item in _iter_figure_mentions(items):
        figure_mentions.setdefault(figure_number, []).append(item)

    issues = []
    for figure_number, mentions in figure_mentions.items():
        if len(mentions) > 1:
            continue

        item = mentions[0]
        issues.append({
            "file_name": item["file_name"],
            "issue_type": "引用问题",
            "heading": item["heading"],
            "section_key": item.get("section_key", item["heading"]),
            "location": item["location"],
            "content": item["text"],
            "issue": f"Figure {figure_number} 仅出现一次，可能未被正文引用",
        })

    return issues
