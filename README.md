# document-compliance-checker

一个用于英文 Word 文档内容合规检查的小工具。

项目同时提供两种使用方式：

- 命令行模式：批量扫描 `input` 目录并将报告输出到 `output`。
- Web 页面模式：在浏览器中上传 Word 文档，实时查看检查进度、预览结果并下载报告。

## 当前版本

V1.0.0-rc.1

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
- 通过 Web 页面上传一个或多个 Word 文档并直接执行检查。
- 实时展示当前文件、检查阶段和整体检查进度。
- 在页面中预览各 Sheet，并下载 Excel 或 Markdown 结果。

## V1.0.0-rc.1 更新

本版本重点优化前端结构、Web 并发持久化和检查准确性。

主要调整：

- 将文档上传、检查进度和历史任务结果请求从 `App.vue` 抽离到独立 API 模块。
- 优化 Web 上传文件列表，逐项展示 Word 文件名并支持长列表滚动。
- 为同名文档归档和同文件组合任务替换增加互斥保护，避免并发请求产生目录竞争。
- 英文格式检查跳过命令和代码样式内容，并保护 URL、IP、路径、版本号等特殊文本，减少误报。
- Figure 引用检查区分图题与正文引用，并修复 Word 自动编号 `Figure 61` 与正文 `Figure 6-1` 无法匹配的问题。
- 新增英文格式和 Figure 引用自动化测试，覆盖主要误报与编号匹配场景。
- 新增前端现状分析与分阶段重构说明文档。

## V0.5.0 更新

本版本新增完整的 Web 检查流程和检查进度展示。

主要调整：

- 新增 FastAPI 后端，支持多文档上传、检查结果恢复和 Excel 下载。
- 新增 Vue Web 页面，可直接上传 `.docx` 文件并预览检查结果。
- 新增实时进度条，展示上传、文档读取、规则检查和报告生成阶段。
- 同一文件名组合再次上传时沿用任务编号，并在成功后安全替换旧结果。
- 每个 Web 文档单独归档原文、问题数据和更新时间。
- 新增 `start.ps1`，可一键启动 Python API 和前端开发服务器。
- 补充 Web 启动、使用、存储和接口说明文档。

## V0.4.1 更新

本版本主要优化分组运行、附录标题识别和 Figure 引用匹配。

主要调整：

- 支持按输入分组单独运行，例如 `python main.py SuperPOD`。
- 支持直接传入输入分组路径，例如 `python main.py .\input\SuperPOD\`。
- 修复附录类字母标题层级识别，例如 `A. Exit Status Codes` 会作为附录一级标题处理。
- 优化 Figure 引用检查，兼容 Word 自动编号字段中 `Figure - Title` / `Figure 11 Title` 与正文 `Figure 1-1 Title` 的匹配。

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

Web 上传协议、文件存储、接口返回结构及当前不足详见 [Web 文档检查系统分析](docs/WEB_SYSTEM_ANALYSIS.md)。


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
│   ├── WEB_SYSTEM_ANALYSIS.md # Web 上传、存储和接口分析
│   └── FRONTEND_REFACTOR_ANALYSIS.md # 前端现状和重构方案
├── Frontend/
│   └── frontend/           # Vue 3 + Vite Web 页面
│       ├── src/
│       │   ├── api/documentCheckers.js # Web API 请求封装
│       │   └── App.vue     # 上传、进度条和结果预览页面
│       └── vite.config.js  # 开发服务器及 API 代理配置
├── input/                  # 输入 Word 文档目录，可按文档分组放入子目录
├── output/                 # 输出检查结果目录，会按文档分组生成子目录
├── parser/                 # Word 读取和标题目录解析
│   ├── heading_catalog.py
│   └── word_reader.py
├── report/                 # 报告导出
│   ├── excel_writer.py
│   └── txt_writer.py
├── tests/                  # 英文格式和 Figure 引用自动化测试
├── main.py                 # 主程序入口
├── web_api.py              # FastAPI 上传、进度、结果和下载接口
├── web_service.py          # Web 文档检查及 Excel 数据转换服务
├── requirements.txt        # Python 依赖
├── start.ps1               # Web 前后端一键启动脚本
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

如果只需要检查某个分组，可以在命令后面加分组名：

```bash
python main.py SuperPOD
python main.py KSManageV2.5
```

也可以直接传入 `input` 下的分组路径：

```bash
python main.py .\input\SuperPOD\
python main.py .\input\KSManageV2.5\
```

也可以一次检查多个分组：

```bash
python main.py SuperPOD KSManageV2.5
```

当前分组规则：

- `input/SuperPOD/*.docx`：输出到 `output/SuperPOD`
- `input/KSManageV2.5/*.docx`：输出到 `output/KSManageV2.5`
- `input/其他文件夹/*.docx`：输出到 `output/其他文件夹`
- 直接放在 `input` 根目录的文档，会按文件名规则兜底分到 `SuperPOD`、`KSManageV2.5` 或 `Unclassified`

## Web 页面使用说明

### 1. 环境要求

- Windows PowerShell
- Python 3.10 或更高版本
- Node.js 20.19 或更高版本（推荐使用当前 Node.js LTS，并包含 npm）

可以用以下命令确认环境已经安装：

```powershell
python --version
node --version
npm --version
```

### 2. 首次安装依赖

在项目根目录打开 PowerShell，依次运行：

```powershell
python -m pip install -r requirements.txt
cd Frontend\frontend
npm install
cd ..\..
```

依赖只需要安装一次。

### 3. 一键启动（推荐）

在项目根目录运行：

```powershell
.\start.ps1
```

脚本会自动启动 Python API 和前端开发服务器。终端显示前端地址后，在浏览器打开：

```text
http://localhost:5173/
```

如果 PowerShell 提示禁止运行脚本，可仅为当前终端临时放开限制后再启动：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\start.ps1
```

需要停止服务时，在运行 `start.ps1` 的终端按 `Ctrl+C`；脚本退出时会同时关闭它启动的 Python API。

### 4. 手动启动

需要分别查看前后端日志时，可以打开两个 PowerShell 终端。

终端一：在项目根目录启动 Python API：

```powershell
python -m uvicorn web_api:app --reload
```

终端二：启动前端页面：

```powershell
cd Frontend\frontend
npm start
```

然后使用浏览器打开 `http://localhost:5173/`。

### 5. 页面操作

1. 点击“选择 Word 文档”，选择一个或多个 `.docx` 文件。
2. 点击“开始检查”。
3. 页面会通过进度条显示当前百分比、处理阶段和正在检查的文件名。
4. 检查完成后，可切换 Sheet 预览结果或下载检查报告。

页面支持：

- 实时展示上传、文档读取、规则检查和报告生成进度。
- 按 Sheet 切换和预览 Markdown 检查结果。
- 下载 Python 生成的多 Sheet Excel。
- 下载当前 Sheet 的 Markdown。
- 下载合并全部 Sheet 的 Markdown。

如果页面提示“无法连接 Python 检查服务”，请确认后端终端仍在运行，并访问 `http://127.0.0.1:8000/api/health`。正常情况下会返回 `{"status":"ok"}`。

### 6. Web 检查结果位置

网页检查结果保存在 `output/web_jobs/<任务编号>/`，不会覆盖命令行按分组生成的原有报告。
同一文件名或同一批文件名组合再次上传时，网页检查会沿用原任务编号，并用最新上传文件和检查结果覆盖旧任务；不同文件组合会创建新任务。
此外，每个网页文档会按文件名独立归档到 `output/web_documents/<文档标识>/`。重新检查同名文件时只替换该文档的 Word 副本和问题数据，未上传的其他文档不会被修改；本次下载的 Excel 仍只包含本次上传的文件。

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
