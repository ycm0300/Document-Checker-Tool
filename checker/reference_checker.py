import re

from config.common_rules import clean_text


FIGURE_REFERENCE_PATTERN = re.compile(r"\bFigure\s+(\d+(?:[-.]\d+)*)\b", re.IGNORECASE)
UNNUMBERED_FIGURE_CAPTION_PATTERN = re.compile(r"^Figure\s*[-–—]\s*(.+)$", re.IGNORECASE)


def _normalize_figure_title(title):
    return clean_text(title).strip(" .。:：;；").lower()


def _extract_numbered_figure_title(text, match):
    title = text[match.end():]
    title = re.split(r"[.;。；\r\n]", title, maxsplit=1)[0]
    return _normalize_figure_title(title)


def _collect_unnumbered_figure_titles(items):
    titles = set()

    for item in items:
        if item.get("is_toc"):
            continue

        text = clean_text(item.get("text", ""))
        match = UNNUMBERED_FIGURE_CAPTION_PATTERN.match(text)
        if match:
            titles.add(_normalize_figure_title(match.group(1)))

    return titles


def _iter_figure_mentions(items):
    for item in items:
        if item.get("is_toc"):
            continue

        text = clean_text(item.get("text", ""))
        if not text:
            continue

        for match in FIGURE_REFERENCE_PATTERN.finditer(text):
            yield match.group(1), _extract_numbered_figure_title(text, match), item


def check_figure_references(items):
    """检查 Figure 编号是否至少在正文中出现两次：标题本身 + 正文引用。"""
    figure_mentions = {}
    unnumbered_figure_titles = _collect_unnumbered_figure_titles(items)

    for figure_number, figure_title, item in _iter_figure_mentions(items):
        figure_mentions.setdefault(figure_number, []).append((figure_title, item))

    issues = []
    for figure_number, mentions in figure_mentions.items():
        if len(mentions) > 1:
            continue

        figure_title, item = mentions[0]
        if figure_title and figure_title in unnumbered_figure_titles:
            continue

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
