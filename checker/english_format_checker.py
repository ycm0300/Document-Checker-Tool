import re

from config.chinese_rules import CHINESE_PUNCTUATION
from config.english_rules import (
    MISSING_SPACE_AFTER_ENGLISH_PUNCTUATION_PATTERN,
    MULTIPLE_SPACES_PATTERN,
    SPACE_BEFORE_ENGLISH_PUNCTUATION_PATTERN,
)


def check_english_format(text):
    """检查常见英文格式问题。返回问题说明列表。"""
    issues = []

    for punctuation in CHINESE_PUNCTUATION:
        if punctuation in text:
            issues.append(f"发现中文标点：{punctuation}")

    if re.search(MULTIPLE_SPACES_PATTERN, text):
        issues.append("发现连续两个或多个空格")

    if re.search(SPACE_BEFORE_ENGLISH_PUNCTUATION_PATTERN, text):
        issues.append("英文标点前可能多了空格")

    if re.search(MISSING_SPACE_AFTER_ENGLISH_PUNCTUATION_PATTERN, text):
        issues.append("英文标点后可能缺少空格")

    return issues


def check_item(item):
    text = item["text"]
    issues = []

    for issue in check_english_format(text):
        issues.append({
            "file_name": item["file_name"],
            "issue_type": "英文格式",
            "heading": item["heading"],
            "section_key": item.get("section_key", item["heading"]),
            "location": item["location"],
            "content": text,
            "issue": issue,
        })

    return issues
