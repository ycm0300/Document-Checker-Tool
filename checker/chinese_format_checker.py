import re

from config.chinese_rules import CHINESE_CHARACTER_PATTERN


def has_chinese(text):
    """检查真实内容是否包含中文汉字。"""
    return re.search(CHINESE_CHARACTER_PATTERN, text) is not None


def check_item(item):
    text = item["text"]
    if not has_chinese(text):
        return []

    return [{
        "file_name": item["file_name"],
        "issue_type": "中文残留",
        "heading": item["heading"],
        "location": item["location"],
        "content": text,
        "issue": "发现中文汉字",
    }]
