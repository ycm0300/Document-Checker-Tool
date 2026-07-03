# from datetime import datetime
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
UNCLASSIFIED_GROUP_NAME = "Unclassified"

SUMMARY_EXCEL_PREFIX = "总检查结果"
HEADING_CATALOG_EXCEL_PREFIX = "标题目录"
HEADING_CHECK_REPORT_PREFIX = "标题检查结果"


def get_document_group(file_name):
    """根据文件名判断文档所属批次。"""
    lower_name = file_name.lower()

    if "superpod" in lower_name:
        return "SuperPOD"

    if "ksmanage" in lower_name and "v2.5" in lower_name:
        return "KSManageV2.5"

    return UNCLASSIFIED_GROUP_NAME


def get_document_group_from_path(file_path):
    """优先使用 input 下的一级子文件夹作为批次名。"""
    relative_path = file_path.resolve().relative_to(INPUT_DIR.resolve())

    if len(relative_path.parts) > 1:
        return relative_path.parts[0]

    return get_document_group(file_path.name)


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
