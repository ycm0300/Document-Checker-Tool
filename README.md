# document-compliance-checker

一个用于英文 Word 文档内容合规检查的小工具。

## 当前版本

V0.4.0

## 功能说明

当前项目主要用于批量读取 `.docx` 文档内容，检查文档中的常见合规问题，并生成检查结果报告。

主要功能：

- 读取 Word 正文、表格、页眉、页脚内容。
- 识别 Word 标题样式、自动编号标题和手动编号标题。
- 导出每个文档的读取结果文本。
- 检查英文文档中的中文残留。
- 检查厂商名称、网址、邮箱、电话、地址等残留信息。
- 检查中文标点、连续空格、英文标点空格等常见英文格式问题。
- 检查一级标题是否按要求新起页面。
- 检查 Figure 标题是否可能未被正文引用。
- 检查同一最小小节内是否存在重复正文段落。
- 导出总检查结果 Excel。
- 导出标题目录 Excel。
- 对标题目录进行编号连续性、重复编号和相对编号检查。

## V0.4.0 更新

本版本主要新增文档版式、引用关系和内容质量检查，并优化报告定位方式。

主要调整：

- 新增一级标题分页检查，识别 Heading 1 / 标题 1 是否设置为新页开始。
- 新增一级标题前多余空段落检查，识别分页符或分节符与一级标题之间存在空段落的情况。
- 新增 Figure 引用检查，提示只出现一次、可能未被正文引用的 Figure 编号。
- 新增同一最小小节内重复正文段落检查，减少跨章节标准句式带来的误报。
- 优化相对四级标题的问题定位显示，例如 `8.1.1 SuperPod > 1 Super Node Details`。
- 报告文件改为固定文件名，每次运行覆盖刷新，避免输出目录产生过多历史报告。

完整版本记录见 [CHANGELOG.md](CHANGELOG.md)。


## 目录说明

```text
document-compliance-checker/
├── checker/                # 检查逻辑
│   ├── chinese_format_checker.py
│   ├── content_quality_checker.py
│   ├── english_format_checker.py
│   ├── heading_checker.py
│   ├── layout_checker.py
│   ├── reference_checker.py
│   └── vendor_checker.py
├── config/                 # 公共配置和检查规则
│   ├── chinese_rules.py
│   ├── common_rules.py
│   ├── english_rules.py
│   └── vendor_rules.py
├── docs/                   # 项目文档目录
├── input/                  # 输入 Word 文档目录，可按文档分组放入子目录
├── output/                 # 输出检查结果目录，会按文档分组生成子目录
├── parser/                 # Word 读取和标题目录解析
│   ├── heading_catalog.py
│   └── word_reader.py
├── report/                 # 报告导出
│   ├── excel_writer.py
│   └── txt_writer.py
├── main.py                 # 主程序入口
├── README.md               # 项目说明
├── CHANGELOG.md            # 版本变更记录
└── .gitignore              # Git 忽略配置
```

## 使用说明

1. 将需要检查的 `.docx` 文档按分组放入 `input` 下的子目录，例如 `input/SuperPOD` 和 `input/KSManageV2.5`。
2. 运行主程序：

```bash
python main.py
```

3. 程序会扫描 `input` 及其子目录，并优先使用 `input` 下的一级子目录名作为分组名，检查结果输出到 `output` 下同名目录。

当前分组规则：

- `input/SuperPOD/*.docx`：输出到 `output/SuperPOD`
- `input/KSManageV2.5/*.docx`：输出到 `output/KSManageV2.5`
- `input/其他文件夹/*.docx`：输出到 `output/其他文件夹`
- 直接放在 `input` 根目录的文档，会按文件名规则兜底分到 `SuperPOD`、`KSManageV2.5` 或 `Unclassified`

## 输出文件

每个输出分组目录下会生成：

- `*_读取结果.txt`
- `总检查结果.xlsx`
- `标题目录.xlsx`
- `标题检查结果.xlsx`

## 规则维护

- 中文残留规则：`config/chinese_rules.py`
- 英文格式规则：`config/english_rules.py`
- 厂商残留关键词：`config/vendor_rules.py`
- 路径和输出文件名配置：`config/common_rules.py`

## 说明

`input` 和 `output` 目录中的实际文件不会提交到仓库，只保留目录结构。
