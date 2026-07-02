import re

from config.common_rules import clean_cell, clean_text


HEADING_NUMBER_PATTERN = re.compile(r"^(\d+(?:\.\d+)*)\b")
RELATIVE_LEVEL4_PATTERN = re.compile(r"^(\d+)\.?\s+.+")
MAX_HEADING_DEPTH = 4


def parse_heading_number(text):
    """提取标题开头的数字编号，例如 2.3.1。"""
    match = HEADING_NUMBER_PATTERN.match(clean_text(text))
    if not match:
        return None
    try:
        return tuple(int(part) for part in match.group(1).split("."))
    except ValueError:
        return None


def parse_relative_level4_number(text):
    """识别三级标题下形如“1. 标题”的相对四级编号。"""
    match = RELATIVE_LEVEL4_PATTERN.match(clean_text(text))
    if not match:
        return None
    return int(match.group(1))


def is_toc_like_heading(text):
    text = clean_text(text)
    return re.match(r"^\d+(?:\.\d+)*\s+.+\s+\d+$", text) is not None


def make_heading_issue(item, issue):
    return {
        "file_name": item["file_name"],
        "issue_type": "标题层级",
        "heading": item["heading"],
        "section_key": item.get("section_key", item["heading"]),
        "location": item["location"],
        "content": item["heading"],
        "issue": issue,
    }


def format_heading_number(number):
    return ".".join(str(part) for part in number)


def number_from_style_level(level, style_counters):
    """按 Word Heading 样式层级重新生成标题编号。"""
    if level is None or level < 1 or level > MAX_HEADING_DEPTH:
        return None

    for upper_level in range(1, level):
        if style_counters[upper_level] == 0:
            style_counters[upper_level] = 1

    style_counters[level] += 1
    for lower_level in range(level + 1, MAX_HEADING_DEPTH + 1):
        style_counters[lower_level] = 0

    return tuple(style_counters[1:level + 1])


def check_heading_structure(items):
    """检查标题连续性、层级深度、重复编号和编号倒退。"""
    issues = []
    seen_numbers = {}
    max_number_by_parent = {}
    style_counters = [0] * (MAX_HEADING_DEPTH + 1)
    current_level3_number = None

    for item in items:
        if item.get("is_toc") or not item.get("is_heading"):
            continue

        heading_text = item["heading"]
        source_text = item.get("text", heading_text)
        if is_toc_like_heading(heading_text):
            continue

        heading_level = item.get("heading_level")
        heading_number = parse_heading_number(heading_text)
        explicit_number = parse_heading_number(source_text)
        relative_level4 = parse_relative_level4_number(heading_text)

        if heading_level is not None and heading_number is None:
            continue

        is_relative_level4 = False
        if heading_level == 4 and relative_level4 is not None and current_level3_number is not None:
            number = current_level3_number + (relative_level4,)
            is_relative_level4 = True
        elif heading_level is not None:
            number = number_from_style_level(heading_level, style_counters)
        elif relative_level4 is not None and current_level3_number is not None:
            number = current_level3_number + (relative_level4,)
        else:
            number = explicit_number

        if number is None:
            continue

        if len(number) == 3:
            current_level3_number = number
        elif len(number) < 3:
            current_level3_number = None

        number_text = format_heading_number(number)
        parent = number[:-1]
        current = number[-1]

        if len(number) > MAX_HEADING_DEPTH:
            issues.append(make_heading_issue(
                item,
                f"标题层级过深：{number_text}，当前最多建议 {MAX_HEADING_DEPTH} 级标题",
            ))

        if len(number) > 1 and parent not in seen_numbers and not is_relative_level4 and heading_level is None:
            parent_text = format_heading_number(parent)
            issues.append(make_heading_issue(
                item,
                f"标题层级跳跃：{number_text} 缺少上级标题 {parent_text}",
            ))

        if number in seen_numbers:
            issues.append(make_heading_issue(
                item,
                f"标题重复：{number_text} 已在 {seen_numbers[number]} 出现过",
            ))
            continue

        previous_max = max_number_by_parent.get(parent)
        if previous_max is None:
            if current != 1:
                expected = format_heading_number(parent + (1,))
                issues.append(make_heading_issue(
                    item,
                    f"标题连续性异常：当前为 {number_text}，该层级应从 {expected} 开始",
                ))
        elif current < previous_max:
            issues.append(make_heading_issue(
                item,
                f"标题倒退：同级标题已出现到 {format_heading_number(parent + (previous_max,))}，当前又出现 {number_text}",
            ))
        elif current > previous_max + 1:
            expected = format_heading_number(parent + (previous_max + 1,))
            issues.append(make_heading_issue(
                item,
                f"标题连续性异常：{format_heading_number(parent + (previous_max,))} 后出现 {number_text}，疑似缺少 {expected}",
            ))

        seen_numbers[number] = item["location"]
        max_number_by_parent[parent] = max(previous_max or 0, current)

    return issues


def parse_num(text):
    """提取纯编号：2.1 / 6.4.1。"""
    text = clean_cell(text)
    match = re.match(r"^(\d+(?:\.\d+)*)", text)
    if not match:
        return None
    return tuple(map(int, match.group(1).split(".")))


def parse_heading_level(text):
    """提取 Heading 级别：Heading 4 -> 4。"""
    text = clean_cell(text)
    match = re.match(r"^Heading\s+(\d+)$", text, re.IGNORECASE)
    if not match:
        return None
    return int(match.group(1))


def is_relative_heading4_num(raw, num, heading_level):
    text = clean_cell(raw)
    return heading_level == 4 and len(num) == 1 and re.match(r"^\d+\.?$", text) is not None


def find_nearest_parent(ancestors, heading_level):
    for level in range(heading_level - 1, 0, -1):
        if level in ancestors:
            return ancestors[level]
    return ("未知父级",)


def check_heading_catalog_sheet(ws):
    headers = [clean_cell(c.value) for c in ws[1]]
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

    for row_index in range(2, ws.max_row + 1):
        value = ws.cell(row_index, col_num).value
        num = parse_num(value)
        if not num:
            continue

        heading_level = None
        if col_level:
            heading_level = parse_heading_level(ws.cell(row_index, col_level).value)

        full_heading = clean_cell(ws.cell(row_index, col_full_heading).value) if col_full_heading else clean_cell(value)

        rows.append((row_index, num, value, heading_level, full_heading))

    for row_index, num, raw, heading_level, full_heading in rows:
        is_relative_heading4 = is_relative_heading4_num(raw, num, heading_level)

        if heading_level and not is_relative_heading4:
            ancestors[heading_level] = num
            for level in list(ancestors):
                if level > heading_level:
                    ancestors.pop(level, None)

        if is_relative_heading4:
            parent_key = find_nearest_parent(ancestors, heading_level)
            parent_text = ".".join(map(str, parent_key))
            current = num[-1]

            if parent_key in relative_heading4_last:
                expected = relative_heading4_last[parent_key] + 1
                if current != expected:
                    issues.append([
                        "不连续",
                        ".".join(map(str, num)),
                        raw,
                        full_heading,
                        f"第{row_index}行，父级 {parent_text} 下应为 {expected}."
                    ])
            elif current != 1:
                issues.append([
                    "不连续",
                    ".".join(map(str, num)),
                    raw,
                    full_heading,
                    f"第{row_index}行，父级 {parent_text} 下应为 1."
                ])

            relative_heading4_last[parent_key] = current
            continue

        if num in seen:
            issues.append([
                "重复",
                ".".join(map(str, num)),
                raw,
                full_heading,
                f"第{row_index}行重复出现"
            ])
            continue

        seen.add(num)

        parent = num[:-1]
        current = num[-1]

        if parent in last_level:
            expected = last_level[parent] + 1
            if current != expected:
                issues.append([
                    "不连续",
                    ".".join(map(str, num)),
                    raw,
                    full_heading,
                    f"应为 {'.'.join(map(str, parent + (expected,)))}"
                ])

        last_level[parent] = current

    return issues
