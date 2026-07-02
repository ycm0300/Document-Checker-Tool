# from datetime import datetime
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

SUMMARY_EXCEL_PREFIX = "总检查结果"
HEADING_CATALOG_EXCEL_PREFIX = "标题目录"
HEADING_CHECK_REPORT_PREFIX = "标题检查结果"


def clean_text(text):
    """清理文本中的多余换行和空格。"""
    if text is None:
        return ""
    text = text.replace("\u3000", " ")
    text = " ".join(text.split())
    return text.strip()


def clean_cell(value):
    if value is None:
        return ""
    return str(value).strip()


def get_summary_excel_name():
    return f"{SUMMARY_EXCEL_PREFIX}.xlsx"
    # 如需保留每次运行的历史报告，先取消文件顶部 datetime 导入的注释。
    # 然后改用下面两行：
    # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # return f"{SUMMARY_EXCEL_PREFIX}_{timestamp}.xlsx"


def get_heading_catalog_excel_name():
    return f"{HEADING_CATALOG_EXCEL_PREFIX}.xlsx"
    # 如需保留每次运行的历史报告，先取消文件顶部 datetime 导入的注释。
    # 然后改用下面两行：
    # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # return f"{HEADING_CATALOG_EXCEL_PREFIX}_{timestamp}.xlsx"


def get_heading_check_report_name():
    return f"{HEADING_CHECK_REPORT_PREFIX}.xlsx"
    # 如需保留每次运行的历史报告，先取消文件顶部 datetime 导入的注释。
    # 然后改用下面两行：
    # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # return f"{HEADING_CHECK_REPORT_PREFIX}_{timestamp}.xlsx"


def safe_sheet_name(name):
    invalid_chars = ["\\", "/", "?", "*", "[", "]", ":"]
    for ch in invalid_chars:
        name = name.replace(ch, "_")

    removable_suffixes = [
        " for KSManage SuperPOD V1.0",
        " for KSManage SuperPOD",
    ]
    for suffix in removable_suffixes:
        if name.endswith(suffix):
            name = name[: -len(suffix)]
            break

    return name[:31] if len(name) > 31 else name
