import re

from config.chinese_rules import CHINESE_PUNCTUATION
from config.english_rules import (
    MISSING_SPACE_AFTER_ENGLISH_PUNCTUATION_PATTERN,
    MULTIPLE_SPACES_PATTERN,
    SPACE_BEFORE_ENGLISH_PUNCTUATION_PATTERN,
)


COMMAND_LINE_PATTERN = re.compile(
    r"^\s*(?:(?:sudo|env)\s+)?(?:"
    r"imcli|mcli|tar|bash|sh|systemctl|kubectl|docker|helm|"
    r"apt(?:-get)?|dnf|yum|rpm|dpkg|curl|wget|ssh|scp|cp|mv|"
    r"chmod|chown|mkdir|mount|umount|exportfs|ufw|openssl|md5sum|"
    r"sha256sum|Get-FileHash|Set-Location|cd"
    r")(?:\s|$)",
    re.IGNORECASE,
)

PROTECTED_PATTERNS = tuple(
    re.compile(pattern, re.IGNORECASE)
    for pattern in (
        r"https?://[^\s<>\"]+",
        r"\b(?:\d{1,3}\.){3}\d{1,3}(?:/\d{1,2})?\b",
        r"\b(?:XX\.){3}XX\b",
        r"\b[A-Za-z]:\\[^\s,;!?]+",
        r"(?<!\w)(?:\.\.?[/\\]|/)[^\s,;!?]+",
        r"\b[vV]?\d+(?:\.\d+){1,}(?:[-_][A-Za-z0-9.-]+)?\b",
        r"\b(?:[A-Za-z0-9_-]+\.)+(?:com|org|net|io|cn|edu|gov|local)\b",
        r"\b[A-Za-z0-9_-]+(?:\.[A-Za-z0-9_-]+)*\."
        r"(?:tar\.gz|tar\.bz2|tar\.xz|sh|ps1|py|js|ts|json|ya?ml|xml|ini|"
        r"conf|cfg|properties|service|log|txt|md|docx?|xlsx?|zip|gz|rpm|deb)\b",
        r"(?<!\w)\.(?:sh|ps1|py|js|ts|json|ya?ml|xml|ini|conf|cfg|"
        r"properties|service|log|txt|md|docx?|xlsx?|zip|gz|rpm|deb)\b",
        r"\b[A-Fa-f0-9]{16,}\b",
        r"(?<!\w)--[A-Za-z0-9][A-Za-z0-9_-]*",
    )
)


def is_code_like_item(item):
    """识别应保持原样的独立命令或代码段。"""
    if item.get("is_code_style"):
        return True
    return bool(COMMAND_LINE_PATTERN.match(item.get("text", "")))


def mask_protected_spans(text):
    """遮蔽 URL、IP、路径等标点具有语法含义的片段。"""
    masked = text
    for pattern in PROTECTED_PATTERNS:
        masked = pattern.sub(lambda match: "x" * len(match.group()), masked)
    return masked


def check_english_format(text, *, skip_code_format=False):
    """检查常见英文格式问题。返回问题说明列表。"""
    issues = []

    if skip_code_format:
        return issues

    for punctuation in CHINESE_PUNCTUATION:
        if punctuation in text:
            issues.append(f"发现中文标点：{punctuation}")

    format_text = mask_protected_spans(text)

    if re.search(MULTIPLE_SPACES_PATTERN, format_text):
        issues.append("发现连续两个或多个空格")

    if re.search(SPACE_BEFORE_ENGLISH_PUNCTUATION_PATTERN, format_text):
        issues.append("英文标点前可能多了空格")

    if re.search(MISSING_SPACE_AFTER_ENGLISH_PUNCTUATION_PATTERN, format_text):
        issues.append("英文标点后可能缺少空格")

    return issues


def check_item(item):
    text = item["text"]
    issues = []

    for issue in check_english_format(text, skip_code_format=is_code_like_item(item)):
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
