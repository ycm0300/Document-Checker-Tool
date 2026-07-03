from checker.chinese_format_checker import check_item as check_chinese_item
from checker.content_quality_checker import check_duplicate_paragraphs
from checker.english_format_checker import check_item as check_english_format_item
from checker.heading_checker import check_heading_catalog_sheet
from checker.layout_checker import check_heading1_starts_new_page
from checker.reference_checker import check_figure_references
from checker.vendor_checker import check_item as check_vendor_item
from config.common_rules import (
    INPUT_DIR,
    OUTPUT_DIR,
    get_document_group_from_path,
    get_heading_catalog_excel_name,
    get_summary_excel_name,
)
from parser.heading_catalog import load_heading_catalog_workbook
from parser.word_reader import read_word_file
from report.excel_writer import (
    export_heading_catalog_excel,
    export_heading_check_report,
    export_summary_excel,
)
from report.txt_writer import export_read_result_to_txt


def check_items(items):
    issues = []
    for item in items:
        issues.extend(check_chinese_item(item))
        issues.extend(check_vendor_item(item))
        issues.extend(check_english_format_item(item))
    return issues


def get_heading_catalog_document_names(workbook):
    summary_sheet = workbook["汇总"]
    detail_sheet_names = [
        sheet_name for sheet_name in workbook.sheetnames
        if sheet_name != "汇总"
    ]
    document_names = {}

    for index, sheet_name in enumerate(detail_sheet_names, start=2):
        document_name = summary_sheet.cell(index, 2).value
        document_names[sheet_name] = document_name or sheet_name

    return document_names


def check_heading_catalog_file(heading_catalog_file, output_dir):
    workbook = load_heading_catalog_workbook(heading_catalog_file)
    results = {}
    document_names = get_heading_catalog_document_names(workbook)

    for sheet_name in workbook.sheetnames:
        if sheet_name == "汇总":
            continue

        sheet = workbook[sheet_name]
        results[sheet_name] = {
            "document_name": document_names.get(sheet_name, sheet_name),
            "issues": check_heading_catalog_sheet(sheet),
        }

    return export_heading_check_report(results, output_dir)


def group_word_files(word_files):
    grouped_files = {}
    for word_file in word_files:
        group_name = get_document_group_from_path(word_file)
        grouped_files.setdefault(group_name, []).append(word_file)
    return grouped_files


def process_group(group_name, word_files):
    output_dir = OUTPUT_DIR / group_name
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n开始处理分组：{group_name}，共 {len(word_files)} 个文档")

    all_file_issues = {}
    all_file_items = {}

    for word_file in word_files:
        print(f"正在读取：{word_file.name}")

        items = read_word_file(word_file)
        all_file_items[word_file.name] = items

        read_txt_file = output_dir / f"{word_file.stem}_读取结果.txt"
        export_read_result_to_txt(items, read_txt_file)
        print(f"读取结果已输出：{read_txt_file}")

        issues = check_items(items)
        issues.extend(check_duplicate_paragraphs(items))
        issues.extend(check_figure_references(items))
        issues.extend(check_heading1_starts_new_page(word_file))
        all_file_issues[word_file.name] = issues

        print(f"检查完成：{word_file.name}，发现 {len(issues)} 个问题")

    summary_excel_file = output_dir / get_summary_excel_name()
    export_summary_excel(all_file_issues, summary_excel_file)
    print(f"总检查结果已输出：{summary_excel_file}")

    heading_catalog_file = output_dir / get_heading_catalog_excel_name()
    export_heading_catalog_excel(all_file_items, heading_catalog_file)
    print(f"标题目录已输出：{heading_catalog_file}")

    check_heading_catalog_file(heading_catalog_file, output_dir)


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    word_files = sorted([
        file for file in INPUT_DIR.rglob("*.docx")
        if not file.name.startswith("~$")
    ], key=lambda file: file.name.lower())

    if not word_files:
        print("input 文件夹及其子文件夹里没有找到 .docx 文件")
        return

    grouped_files = group_word_files(word_files)
    for group_name, group_files in grouped_files.items():
        process_group(group_name, group_files)


if __name__ == "__main__":
    main()
