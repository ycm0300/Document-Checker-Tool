from config.vendor_rules import FORBIDDEN_VENDOR_KEYWORDS


def check_vendor(text):
    """检查 KAYTUS 文档中的其他厂商残留。"""
    issues = []
    lower_text = text.lower()

    for keyword in FORBIDDEN_VENDOR_KEYWORDS:
        if keyword.lower() in lower_text:
            issues.append(f"发现厂商残留：{keyword}")

    return issues


def check_item(item):
    text = item["text"]
    issues = []

    for issue in check_vendor(text):
        issues.append({
            "file_name": item["file_name"],
            "issue_type": "厂商残留",
            "heading": item["heading"],
            "location": item["location"],
            "content": text,
            "issue": issue,
        })

    return issues
