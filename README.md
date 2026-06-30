# document-compliance-checker

一个用于英文 Word 文档内容合规检查的小工具。

## 当前版本

V0.2.0

## 功能说明

当前项目主要用于批量读取 `.docx` 文档内容，并生成检查结果报告。

主要功能：

- 读取 Word 正文、表格、页眉、页脚内容。
- 识别 Word 标题样式和手动编号标题。
- 导出每个文档的读取结果文本。
- 检查英文文档中的中文残留。
- 检查厂商名称、网址、邮箱、电话等残留信息。
- 检查常见英文格式问题。
- 导出总检查结果 Excel。
- 导出标题目录 Excel。
- 对标题目录进行编号连续性和重复检查。

## V0.2.0 更新

- 新增 `check_heading_catalog.py`，用于检查标题目录编号。
- 标题检查结果增加 `完整标题` 列，方便定位问题。
- 修复 Heading 4 中 `1.`、`2.`、`3.` 这类相对编号被误判为重复的问题。
- Heading 4 相对编号现在按最近的上级标题范围检查连续性，不再按全文编号重复判断。
- 保留 `1.2.3.4` 这类完整多级编号的原有检查逻辑。

## 目录说明

```text
document-compliance-checker/
├── input/                  # 输入 Word 文档目录，实际文档不上传
├── output/                 # 输出检查结果目录，实际结果不上传
├── read_word.py            # 主程序：读取 Word 并生成检查结果、标题目录
├── check_heading_catalog.py # 标题目录编号检查脚本
├── README.md               # 项目说明
├── CHANGELOG.md            # 版本变更记录
└── .gitignore              # Git 忽略配置
```

## 使用说明

1. 将需要检查的 `.docx` 文档放入 `input` 目录。
2. 运行主程序，生成读取结果、总检查结果和标题目录：

```bash
python read_word.py
```

3. 如需检查标题目录编号，继续运行：

```bash
python check_heading_catalog.py
```

检查结果会输出到 `output` 目录。

## 输出文件

主程序会生成：

- `*_读取结果.txt`
- `总检查结果_YYYYMMDD_HHMMSS.xlsx`
- `标题目录_YYYYMMDD_HHMMSS.xlsx`

标题检查脚本会生成：

- `标题检查结果_YYYYMMDD_HHMMSS.xlsx`

## 说明

`input` 和 `output` 目录中的实际文件不会提交到仓库，只保留目录结构。
