from config.common_rules import clean_text


MIN_DUPLICATE_TEXT_LENGTH = 30


def _is_checkable_paragraph(item):
    if item.get("source_type") != "正文":
        return False
    if item.get("is_heading") or item.get("is_toc"):
        return False

    text = clean_text(item.get("text", ""))
    return len(text) >= MIN_DUPLICATE_TEXT_LENGTH


def check_duplicate_paragraphs(items):
    """检查同一最近标题章节内完全相同的正文段落。"""
    first_seen = {}
    issues = []

    for item in items:
        if not _is_checkable_paragraph(item):
            continue

        section_key = item.get("section_key") or item.get("heading", "")
        text = clean_text(item.get("text", ""))
        duplicate_key = (section_key, text)
        if duplicate_key not in first_seen:
            first_seen[duplicate_key] = item
            continue

        first_item = first_seen[duplicate_key]
        issues.append({
            "file_name": item["file_name"],
            "issue_type": "内容质量",
            "heading": item["heading"],
            "section_key": item.get("section_key", item["heading"]),
            "location": item["location"],
            "content": item["text"],
            "issue": f"发现重复段落，首次出现于 {first_item['location']}",
        })

    return issues
