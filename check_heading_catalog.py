from openpyxl import load_workbook, Workbook
from pathlib import Path
from datetime import datetime
import re


BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

CATALOG_PREFIX = "标题目录"
REPORT_PREFIX = "标题检查结果"


def clean(v):
    if v is None:
        return ""
    return str(v).strip()


def parse_num(text):
    """
    提取纯编号：2.1 / 6.4.1
    """
    text = clean(text)
    m = re.match(r"^(\d+(?:\.\d+)*)", text)
    if not m:
        return None
    return tuple(map(int, m.group(1).split(".")))


def parse_heading_level(text):
    """
    提取 Heading 级别：Heading 4 -> 4
    """
    text = clean(text)
    m = re.match(r"^Heading\s+(\d+)$", text, re.IGNORECASE)
    if not m:
        return None
    return int(m.group(1))


def is_relative_heading4_num(raw, num, heading_level):
    """
    Heading 4 中形如 1. / 2. / 3. 的相对编号。
    这类编号只在当前父级标题下检查连续性，不做全文重复检查。
    """
    text = clean(raw)
    return heading_level == 4 and len(num) == 1 and re.match(r"^\d+\.?$", text) is not None


def find_nearest_parent(ancestors, heading_level):
    """
    查找当前标题最近的上级标题。
    Heading 4 的相对编号可能直接挂在 Heading 2 下，不一定有 Heading 3。
    """
    for level in range(heading_level - 1, 0, -1):
        if level in ancestors:
            return ancestors[level]
    return ("未知父级",)


def find_latest():
    files = list(OUTPUT_DIR.glob(f"{CATALOG_PREFIX}*.xlsx"))
    if not files:
        files = list(BASE_DIR.glob(f"{CATALOG_PREFIX}*.xlsx"))
    return max(files, key=lambda p: p.stat().st_mtime)


def check_sheet(ws):
    headers = [clean(c.value) for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers)}

    if "标题编号" not in idx:
        raise ValueError(f"{ws.title} 缺少：标题编号")

    col_num = idx["标题编号"]
    col_level = idx.get("Heading级别")
    col_full_heading = idx.get("完整标题")

    seen = set()
    issues = []

    last_level = {}
    relative_heading4_last = {}
    ancestors = {}

    rows = []

    # 收集数据
    for r in range(2, ws.max_row + 1):
        val = ws.cell(r, col_num).value
        num = parse_num(val)
        if not num:
            continue

        heading_level = None
        if col_level:
            heading_level = parse_heading_level(ws.cell(r, col_level).value)

        full_heading = clean(ws.cell(r, col_full_heading).value) if col_full_heading else clean(val)

        rows.append((r, num, val, heading_level, full_heading))

    # =========================
    # 纯编号分析
    # =========================
    for i, (r, num, raw, heading_level, full_heading) in enumerate(rows):
        is_relative_heading4 = is_relative_heading4_num(raw, num, heading_level)

        if heading_level and not is_relative_heading4:
            ancestors[heading_level] = num
            for level in list(ancestors):
                if level > heading_level:
                    ancestors.pop(level, None)

        if is_relative_heading4:
            parent_key = find_nearest_parent(ancestors, heading_level)
            parent_text = ".".join(map(str, parent_key))
            cur = num[-1]

            if parent_key in relative_heading4_last:
                expect = relative_heading4_last[parent_key] + 1
                if cur != expect:
                    issues.append([
                        "不连续",
                        ".".join(map(str, num)),
                        raw,
                        full_heading,
                        f"第{r}行，父级 {parent_text} 下应为 {expect}."
                    ])
            elif cur != 1:
                issues.append([
                    "不连续",
                    ".".join(map(str, num)),
                    raw,
                    full_heading,
                    f"第{r}行，父级 {parent_text} 下应为 1."
                ])

            relative_heading4_last[parent_key] = cur
            continue

        # -------- 重复检查 --------
        if num in seen:
            issues.append([
                "重复",
                ".".join(map(str, num)),
                raw,
                full_heading,
                f"第{r}行重复出现"
            ])
            continue

        seen.add(num)

        # -------- 连续性检查 --------
        parent = num[:-1]
        cur = num[-1]

        if parent in last_level:
            expect = last_level[parent] + 1
            if cur != expect:
                issues.append([
                    "不连续",
                    ".".join(map(str, num)),
                    raw,
                    full_heading,
                    f"应为 {'.'.join(map(str, parent + (expect,)))}"
                ])

        last_level[parent] = cur

    return issues


def write_report(all_results):
    wb = Workbook()
    ws0 = wb.active
    ws0.title = "汇总"

    ws0.append(["文档", "问题数"])

    for sheet, issues in all_results.items():
        ws0.append([sheet, len(issues)])

    for sheet, issues in all_results.items():
        ws = wb.create_sheet(sheet[:31])
        ws.append(["类型", "编号", "原始内容", "完整标题", "说明"])

        if not issues:
            ws.append(["通过", "", "", "", "无问题"])
        else:
            for i in issues:
                ws.append(i)

    out = OUTPUT_DIR / f"{REPORT_PREFIX}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(out)
    print("输出：", out)


def main():
    file = find_latest()
    wb = load_workbook(file)

    results = {}

    for sheet in wb.sheetnames:
        if sheet == "汇总":
            continue

        ws = wb[sheet]
        results[sheet] = check_sheet(ws)

    write_report(results)


if __name__ == "__main__":
    main()
