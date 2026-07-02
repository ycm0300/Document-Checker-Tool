def export_read_result_to_txt(items, output_file):
    """每个文档单独输出全文读取结果，方便调试。"""
    with open(output_file, "w", encoding="utf-8") as file:
        for index, item in enumerate(items, start=1):
            file.write(f"{index}. 文档：{item['file_name']}\n")
            file.write(f"   章节：{item['heading']}\n")
            file.write(f"   位置：{item['location']}\n")
            file.write(f"   内容：{item['text']}\n")
            file.write("\n")
