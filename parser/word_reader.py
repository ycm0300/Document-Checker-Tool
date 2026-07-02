import re

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from config.common_rules import clean_text


def is_heading(paragraph):
    """判断一个段落是否像章节标题。"""
    text = clean_text(paragraph.text)
    if not text:
        return False

    style_name = paragraph.style.name if paragraph.style else ""
    style_name_lower = style_name.lower()

    if style_name_lower.startswith("heading") or style_name.startswith("标题"):
        return True

    if re.match(r"^\d+(\.\d+)*\s+.+", text):
        return True

    return False


def get_heading_level(paragraph):
    """返回 Word 标题样式层级：Heading 1 -> 1。"""
    style_name = paragraph.style.name if paragraph.style else ""
    match = re.match(r"^(?:Heading|标题)\s*(\d+)$", style_name, re.IGNORECASE)
    if not match:
        return None
    return int(match.group(1))


def is_toc_paragraph(paragraph):
    """识别目录段落；目录不参与正文标题层级检查。"""
    raw_text = paragraph.text or ""
    text = clean_text(raw_text)
    if not text:
        return False

    style_name = paragraph.style.name if paragraph.style else ""
    style_name_lower = style_name.lower()
    if style_name_lower.startswith("toc") or style_name.startswith("目录"):
        return True

    if re.match(r"^\d+(?:\.\d+)*\s+.+\t\d+$", raw_text):
        return True
    if re.match(r"^\d+(?:\.\d+)*\s+.+\.{2,}\s*\d+$", text):
        return True
    if re.match(r"^\d+(?:\.\d+)*\s+.+\s+\d+$", text):
        return True

    return False


def xml_val(element):
    """读取 XML 元素里的 w:val。"""
    if element is None:
        return None
    return element.get(qn("w:val"))


def int_or_default(value, default=0):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def get_paragraph_num_pr(paragraph):
    """
    获取段落的编号属性 numPr。
    先读段落自身的编号；如果没有，再读样式里的编号。
    """
    p_pr = paragraph._p.pPr
    if p_pr is not None and p_pr.numPr is not None:
        return p_pr.numPr

    style = paragraph.style
    while style is not None:
        style_p_pr = style.element.pPr
        if style_p_pr is not None and style_p_pr.numPr is not None:
            return style_p_pr.numPr
        style = style.base_style

    return None


def get_paragraph_num_id_and_level(paragraph):
    """返回段落编号的 numId 和 ilvl。"""
    num_pr = get_paragraph_num_pr(paragraph)
    if num_pr is None:
        return None, None

    num_id = xml_val(num_pr.numId)
    ilvl = xml_val(num_pr.ilvl)

    if num_id is not None and ilvl is None:
        ilvl = "0"

    return num_id, ilvl


def build_numbering_info(document):
    """
    解析 word/numbering.xml。
    返回：
    - num_to_abstract: numId -> abstractNumId
    - abstract_levels: abstractNumId -> ilvl -> 编号格式信息
    - num_overrides: numId -> ilvl -> 起始编号覆盖
    """
    num_to_abstract = {}
    abstract_levels = {}
    num_overrides = {}

    try:
        numbering = document.part.numbering_part.element
    except Exception:
        return num_to_abstract, abstract_levels, num_overrides

    for abstract_num in numbering.findall(qn("w:abstractNum")):
        abstract_id = abstract_num.get(qn("w:abstractNumId"))
        if abstract_id is None:
            continue

        abstract_levels[abstract_id] = {}
        for lvl in abstract_num.findall(qn("w:lvl")):
            ilvl = lvl.get(qn("w:ilvl"))
            if ilvl is None:
                continue

            start = 1
            num_fmt = "decimal"
            lvl_text = None

            start_el = lvl.find(qn("w:start"))
            if start_el is not None and xml_val(start_el) is not None:
                start = int_or_default(xml_val(start_el), 1)

            num_fmt_el = lvl.find(qn("w:numFmt"))
            if num_fmt_el is not None and xml_val(num_fmt_el):
                num_fmt = xml_val(num_fmt_el)

            lvl_text_el = lvl.find(qn("w:lvlText"))
            if lvl_text_el is not None and xml_val(lvl_text_el):
                lvl_text = xml_val(lvl_text_el)

            abstract_levels[abstract_id][ilvl] = {
                "start": start,
                "num_fmt": num_fmt,
                "lvl_text": lvl_text,
            }

    for num in numbering.findall(qn("w:num")):
        num_id = num.get(qn("w:numId"))
        abstract_el = num.find(qn("w:abstractNumId"))
        abstract_id = xml_val(abstract_el)
        if num_id is not None and abstract_id is not None:
            num_to_abstract[num_id] = abstract_id

        for lvl_override in num.findall(qn("w:lvlOverride")):
            ilvl = lvl_override.get(qn("w:ilvl"))
            start_override_el = lvl_override.find(qn("w:startOverride"))
            if num_id is not None and ilvl is not None and start_override_el is not None:
                num_overrides.setdefault(num_id, {})[ilvl] = int_or_default(xml_val(start_override_el), 1)

    return num_to_abstract, abstract_levels, num_overrides


def number_to_letters(number, upper=True):
    """1 -> A/a, 2 -> B/b, 27 -> AA/aa。"""
    if number <= 0:
        return str(number)
    letters = []
    while number:
        number -= 1
        letters.append(chr(ord("A") + number % 26))
        number //= 26
    result = "".join(reversed(letters))
    return result if upper else result.lower()


def format_number(number, num_fmt):
    """按 Word 编号格式简单转换。常见标题一般都是 decimal。"""
    if num_fmt in ("decimal", "decimalZero"):
        return str(number).zfill(2) if num_fmt == "decimalZero" else str(number)
    if num_fmt == "upperLetter":
        return number_to_letters(number, upper=True)
    if num_fmt == "lowerLetter":
        return number_to_letters(number, upper=False)
    return str(number)


def make_number_string(num_id, ilvl, numbering_state):
    """根据 numId + ilvl 生成当前段落的自动编号字符串。"""
    if num_id is None or ilvl is None:
        return ""

    num_to_abstract = numbering_state["num_to_abstract"]
    abstract_levels = numbering_state["abstract_levels"]
    num_overrides = numbering_state["num_overrides"]
    counters = numbering_state["counters"]

    abstract_id = num_to_abstract.get(str(num_id))
    if abstract_id is None:
        return ""

    level = int_or_default(ilvl, 0)
    levels = abstract_levels.get(abstract_id, {})
    level_info = levels.get(str(level))
    if not level_info:
        return ""

    num_fmt = level_info.get("num_fmt", "decimal")
    lvl_text = level_info.get("lvl_text") or f"%{level + 1}."

    if num_fmt == "bullet":
        return ""

    counter_key = str(num_id)
    if counter_key not in counters:
        counters[counter_key] = {}

    for i in range(level + 1):
        if i not in counters[counter_key]:
            start = levels.get(str(i), {}).get("start", 1)
            if str(i) in num_overrides.get(str(num_id), {}):
                start = num_overrides[str(num_id)][str(i)]
            counters[counter_key][i] = start - 1 if i == level else start

    counters[counter_key][level] += 1

    for i in list(counters[counter_key].keys()):
        if i > level:
            counters[counter_key].pop(i, None)

    number_string = lvl_text
    for i in range(9):
        placeholder = f"%{i + 1}"
        if placeholder not in number_string:
            continue
        value = counters[counter_key].get(i)
        level_fmt = levels.get(str(i), {}).get("num_fmt", "decimal")
        if value is None:
            value = levels.get(str(i), {}).get("start", 1)
        replacement = format_number(value, level_fmt)
        number_string = number_string.replace(placeholder, replacement)

    return number_string.strip().rstrip(".").strip()


def get_heading_text(paragraph, numbering_state=None):
    """获取标题文本：优先拼接 Word 自动编号；失败则保留原标题文字。"""
    text = clean_text(paragraph.text)
    if not text:
        return "未识别到章节"

    if numbering_state is not None:
        num_id, ilvl = get_paragraph_num_id_and_level(paragraph)
        number = make_number_string(num_id, ilvl, numbering_state)
        if number and not re.match(r"^" + re.escape(number) + r"(\s+|\.)", text):
            return f"{number} {text}"

    return text if text else "未识别到章节"


def make_section_key(heading_stack):
    headings = [
        heading_stack[level]
        for level in sorted(heading_stack)
        if heading_stack.get(level)
    ]
    return " > ".join(headings) if headings else "未识别到章节"


def iter_block_items(parent):
    """按 Word 原始顺序遍历正文中的段落和表格。"""
    if isinstance(parent, DocxDocument):
        parent_elm = parent.element.body
    else:
        parent_elm = parent._tc

    for child in parent_elm.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)
        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


def read_word_file(file_path):
    """
    读取 Word 文档内容。
    返回列表，每条结构：{file_name, heading, location, source_type, text}
    """
    document = Document(file_path)
    file_name = file_path.name
    results = []
    current_heading = "未识别到章节"
    heading_stack = {}
    current_section_key = current_heading
    paragraph_index = 0
    table_index = 0

    num_to_abstract, abstract_levels, num_overrides = build_numbering_info(document)
    numbering_state = {
        "num_to_abstract": num_to_abstract,
        "abstract_levels": abstract_levels,
        "num_overrides": num_overrides,
        "counters": {},
    }

    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            text = clean_text(block.text)
            if not text:
                continue

            paragraph_index += 1

            is_toc = is_toc_paragraph(block)
            heading_level = get_heading_level(block)
            if is_heading(block):
                current_heading = get_heading_text(block, numbering_state)
                if not is_toc:
                    if heading_level is not None:
                        heading_stack[heading_level] = current_heading
                        for level in list(heading_stack):
                            if level > heading_level:
                                heading_stack.pop(level, None)
                    else:
                        heading_stack = {1: current_heading}
                    current_section_key = make_section_key(heading_stack)

                results.append({
                    "file_name": file_name,
                    "heading": current_heading,
                    "section_key": current_section_key,
                    "location": f"正文-段落{paragraph_index}",
                    "source_type": "正文",
                    "text": text,
                    "is_heading": not is_toc,
                    "is_toc": is_toc,
                    "heading_level": heading_level,
                })
            else:
                results.append({
                    "file_name": file_name,
                    "heading": current_heading,
                    "section_key": current_section_key,
                    "location": f"正文-段落{paragraph_index}",
                    "source_type": "正文",
                    "text": text,
                    "is_heading": False,
                    "is_toc": False,
                    "heading_level": None,
                })

        elif isinstance(block, Table):
            table_index += 1
            for row_index, row in enumerate(block.rows, start=1):
                row_texts = []
                for cell in row.cells:
                    cell_text = clean_text(cell.text)
                    if cell_text:
                        row_texts.append(cell_text)

                if row_texts:
                    results.append({
                        "file_name": file_name,
                        "heading": current_heading,
                        "section_key": current_section_key,
                        "location": f"正文-表格{table_index}-第{row_index}行",
                        "source_type": "正文表格",
                        "text": " | ".join(row_texts),
                    })

    for section_index, section in enumerate(document.sections, start=1):
        for p_index, para in enumerate(section.header.paragraphs, start=1):
            text = clean_text(para.text)
            if text:
                results.append({
                    "file_name": file_name,
                    "heading": "页眉",
                    "section_key": "页眉",
                    "location": f"第{section_index}节页眉-段落{p_index}",
                    "source_type": "页眉",
                    "text": text,
                })

        for table_index, table in enumerate(section.header.tables, start=1):
            for row_index, row in enumerate(table.rows, start=1):
                row_texts = []
                for cell in row.cells:
                    cell_text = clean_text(cell.text)
                    if cell_text:
                        row_texts.append(cell_text)
                if row_texts:
                    results.append({
                        "file_name": file_name,
                        "heading": "页眉",
                        "section_key": "页眉",
                        "location": f"第{section_index}节页眉-表格{table_index}-第{row_index}行",
                        "source_type": "页眉表格",
                        "text": " | ".join(row_texts),
                    })

        for p_index, para in enumerate(section.footer.paragraphs, start=1):
            text = clean_text(para.text)
            if text:
                results.append({
                    "file_name": file_name,
                    "heading": "页脚",
                    "section_key": "页脚",
                    "location": f"第{section_index}节页脚-段落{p_index}",
                    "source_type": "页脚",
                    "text": text,
                })

        for table_index, table in enumerate(section.footer.tables, start=1):
            for row_index, row in enumerate(table.rows, start=1):
                row_texts = []
                for cell in row.cells:
                    cell_text = clean_text(cell.text)
                    if cell_text:
                        row_texts.append(cell_text)
                if row_texts:
                    results.append({
                        "file_name": file_name,
                        "heading": "页脚",
                        "section_key": "页脚",
                        "location": f"第{section_index}节页脚-表格{table_index}-第{row_index}行",
                        "source_type": "页脚表格",
                        "text": " | ".join(row_texts),
                    })

    return results
