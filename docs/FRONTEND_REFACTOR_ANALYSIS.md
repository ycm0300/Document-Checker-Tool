# 前端现状分析与分阶段重构方案

本文基于当前项目代码进行静态分析，重点关注 Vue 前端，并说明前后端连接方式。本文只描述现状和重构建议，不包含重构后的完整代码。

主要参考文件：

- `Frontend/frontend/src/App.vue`
- `Frontend/frontend/src/main.js`
- `Frontend/frontend/src/style.css`
- `Frontend/frontend/src/components/HelloWorld.vue`
- `Frontend/frontend/vite.config.js`
- `web_api.py`
- `web_service.py`
- `report/excel_writer.py`

## 1. 完整业务流程

### 1.1 用户选择 Word 文档

页面使用隐藏的文件输入框接收文件：

```html
<input type="file" accept=".docx" multiple>
```

用户选择后，`handleWordFiles()` 将 `FileList` 转成数组并保存到 `selectedFiles`，同时重置当前批次的已检查状态并清除旧错误。

`accept=".docx"` 只是浏览器文件选择提示，真正的格式检查仍由后端完成。

### 1.2 前端发起检查请求

用户点击“开始检查”后，`startCheck()` 会：

1. 判断当前选择是否已经检查。
2. 判断是否至少选择了一个文件。
3. 设置加载状态和初始进度。
4. 生成一个 32 位十六进制 `progressId`。
5. 启动每 300 ms 一次的进度轮询。
6. 创建 `FormData`。
7. 将所有文件以相同字段名 `files` 加入表单。
8. 调用 `POST /api/check`。

请求格式为：

```http
POST /api/check
Content-Type: multipart/form-data; boundary=...
X-Progress-ID: <progressId>
```

浏览器自动生成 multipart boundary，前端不应手动填写 `Content-Type`。

### 1.3 Vite 连接 Python 后端

开发环境中，`vite.config.js` 将所有 `/api` 请求代理到：

```text
http://127.0.0.1:8000
```

因此浏览器请求 `/api/check`，实际由 FastAPI 接收。这样前端不需要写死完整后端地址，也避免本地开发时额外处理跨域地址。

### 1.4 FastAPI 接收并保存文件

`web_api.py` 使用以下参数接收上传内容：

```python
files: list[UploadFile] = File(...)
x_progress_id: str | None = Header(default=None)
```

后端依次执行：

1. 检查是否上传文件。
2. 检查所有文件名扩展名是否为 `.docx`。
3. 根据文件名组合查找历史任务。
4. 创建临时任务目录。
5. 按 1 MB 分块读取并保存文件。
6. 检查每个文件是否超过 50 MB。
7. 将文件保存到 `output/web_jobs/<任务编号>/uploads/`。

同一批上传出现重名文件时，后续文件会增加 `_2`、`_3` 等后缀。

### 1.5 执行 Python 检查

后端通过 `asyncio.to_thread()` 将同步检查放到工作线程中，使 FastAPI 在检查期间仍能响应进度查询。

`web_service.py` 对每个文档依次执行：

1. 读取 Word 内容。
2. 文本格式检查。
3. 重复内容检查。
4. Figure 引用检查。
5. 标题版式检查。

多个文档目前按顺序检查，不是并行检查。

### 1.6 获取检查进度

前端每 300 ms 请求：

```http
GET /api/check-progress/{progressId}
```

后端返回示例：

```json
{
  "percent": 60,
  "message": "正在检查重复内容：文档A.docx",
  "currentFile": "文档A.docx"
}
```

进度保存在 Python 进程内的 `CHECK_PROGRESS` 字典中，不会写入磁盘。服务重启后进度会丢失，多 worker 部署时也可能查询不到另一个进程中的进度。

当前百分比是按检查阶段权重估算，不是按真实段落数量或实际耗时精确计算。

### 1.7 生成 Excel 报告

检查结束后：

1. Python 使用 `openpyxl` 生成 `总检查结果.xlsx`。
2. 第一个 Sheet 为“汇总”。
3. 后续每个文档对应一个问题明细 Sheet。
4. 后端重新读取 Excel。
5. 将所有 Sheet 转换为二维字符串数组。
6. 将 Sheet 数组和 Excel 下载地址放入 JSON 返回前端。

Excel 完全由后端生成，前端不生成或修改 Excel。

### 1.8 前端展示结果

收到成功响应后，`applyJobResult()` 会：

1. 将 Sheet 二维数组转换成 Markdown 表格。
2. 保存任务编号、报告名称和 Excel 下载地址。
3. 默认选择第一个 Sheet。
4. 将任务编号保存到 `localStorage`。
5. 使用 `markdown-it` 将 Markdown 渲染成 HTML。
6. 通过 `v-html` 展示结果。

检查成功响应主要结构为：

```json
{
  "jobId": "0123456789abcdef0123456789abcdef",
  "fileName": "总检查结果.xlsx",
  "documentCount": 2,
  "sheets": [
    {
      "name": "汇总",
      "rows": []
    }
  ],
  "excelDownloadUrl": "/api/jobs/0123456789abcdef0123456789abcdef/excel",
  "replacedExistingTask": false
}
```

### 1.9 恢复上次结果

页面挂载时执行 `restoreLastResult()`：

1. 从 `localStorage` 读取上次 `jobId`。
2. 请求 `GET /api/jobs/{jobId}`。
3. 恢复 Sheet 数据。
4. 恢复上次选中的 Sheet。

### 1.10 下载结果

当前有三种下载方式：

- Excel：访问后端 `/api/jobs/{jobId}/excel`。
- 当前 Sheet Markdown：前端创建 Blob 下载。
- 全部 Sheet Markdown：前端合并 Markdown 后创建 Blob 下载。

## 2. `Frontend/frontend/src` 文件职责

| 文件 | 当前职责 | 使用情况 |
| --- | --- | --- |
| `App.vue` | 上传、API 请求、进度轮询、结果转换、缓存、下载、页面结构和大部分样式 | 核心文件，正在使用 |
| `main.js` | 创建 Vue 应用、加载全局样式、挂载 `App.vue` | 正在使用 |
| `style.css` | 全局字体、背景、盒模型和基础表单样式 | 正在使用 |
| `components/HelloWorld.vue` | Vite/Vue 初始模板示例和计数器 | 未被导入 |
| `assets/hero.png` | `HelloWorld.vue` 的示例图片 | 只被未使用组件引用 |
| `assets/vite.svg` | Vite 示例图标 | 只被未使用组件引用 |
| `assets/vue.svg` | Vue 示例图标 | 只被未使用组件引用 |

当前前端没有 Vue Router、Pinia、API 服务目录、composables、业务组件或测试目录。

`HelloWorld.vue` 和三个图片资源属于脚手架遗留内容。本次只分析，不删除这些文件。

## 3. `App.vue` 分析

### 3.1 状态

#### DOM 状态

- `fileInput`：隐藏文件输入框的 DOM 引用。

#### 上传状态

- `selectedFiles`：用户选择的文件。
- `currentSelectionChecked`：当前选择是否已经检查。

#### 任务与结果状态

- `currentJobId`：当前后端任务编号。
- `resultFileName`：Excel 报告名称。
- `excelDownloadUrl`：Excel 下载地址。
- `sheets`：已转换为 Markdown 的 Sheet 列表。
- `selectedSheetName`：当前选择的 Sheet。

#### 请求与进度状态

- `errorMessage`：页面错误文本。
- `isLoading`：检查或恢复任务时的公共加载状态。
- `checkProgress`：当前检查百分比。
- `progressMessage`：当前检查阶段说明。

#### 缓存键

- `LAST_JOB_KEY`
- `LAST_SHEET_KEY`

#### 派生状态

- `selectedSheet`：根据名称获得当前 Sheet。
- `renderedMarkdown`：将当前 Sheet Markdown 渲染成 HTML。

### 3.2 函数

| 函数 | 职责 |
| --- | --- |
| `escapeCell()` | 转义 Markdown 表格单元格 |
| `rowsToMarkdown()` | 将二维 Sheet 数据转换为 Markdown 表格 |
| `applyJobResult()` | 应用后端结果并保存任务编号 |
| `handleWordFiles()` | 接收用户选择的文件 |
| `startCheck()` | 校验、启动轮询、上传、处理错误并应用结果 |
| `restoreLastResult()` | 恢复上次任务 |
| `clearCurrentResult()` | 清理页面状态和本地缓存 |
| `downloadMarkdown()` | 创建 Markdown Blob 并触发下载 |
| `safeFileName()` | 替换下载文件名中的非法字符 |
| `downloadCurrentSheet()` | 下载当前 Sheet Markdown |
| `downloadAllSheets()` | 合并并下载全部 Sheet Markdown |
| `downloadOriginalExcel()` | 通过后端地址下载 Excel |

此外还有：

- `watch(selectedSheetName)`：保存最后选择的 Sheet。
- `onMounted(restoreLastResult)`：页面加载后恢复历史结果。

### 3.3 `App.vue` 当前承担的职责

`App.vue` 同时负责：

- 文件选择
- 输入校验
- API 请求
- multipart 请求构造
- API 错误解析
- 进度 ID 生成
- 定时轮询
- 任务恢复
- `localStorage` 管理
- 后端数据到页面数据的转换
- Markdown 表格生成
- Markdown HTML 渲染
- Excel 下载
- Markdown 下载
- 页面布局
- 交互状态
- 响应式样式

### 3.4 是否职责过多

是。

问题不仅是文件约 368 行，而是 HTTP 层、业务流程、缓存、数据转换、下载、页面结构和样式全部集中在一个组件里。修改任何一部分都需要进入 `App.vue`，并容易影响其他功能。

## 4. 当前问题

### 4.1 重复代码

#### 重复的请求处理

`startCheck()` 和 `restoreLastResult()` 都包含以下模式：

```text
fetch → response.json() → 检查 response.ok → 抛出错误
```

进度请求也直接调用 `fetch`，但使用另一套处理方式。JSON 解析、错误处理和接口约定没有统一入口。

#### 重复的状态清理

检查失败和用户主动清除都会重置 Sheet、报告名称、下载地址等状态，但两处范围不一致。

检查失败时没有清除 `currentJobId` 和旧的 `localStorage`，可能出现页面没有结果，但仍保存旧任务编号的情况。

#### 重复的下载动作

Markdown 和 Excel 下载都需要创建 `<a>`、设置地址、设置文件名并调用 `click()`，但逻辑分散在不同函数中。

#### 分散的本地缓存操作

任务编号和 Sheet 名称的读取、写入、删除分散在 `applyJobResult()`、`restoreLastResult()`、`clearCurrentResult()` 和 `watch()` 中。

### 4.2 过长函数

`startCheck()` 同时负责：

- 前置校验
- 加载状态
- 进度初始化
- 进度 ID
- 定时器生命周期
- FormData
- 上传请求
- 响应解析
- 结果应用
- 完成状态
- 错误处理
- 定时器清理

函数长度不是极端大，但业务职责明显过多。

`restoreLastResult()` 也混合了缓存读取、接口请求、HTTP 状态分类、缓存清理和页面错误文案。

### 4.3 请求、业务和展示混在一起

`startCheck()` 中同时存在：

```text
HTTP 请求细节
  + 检查任务业务流程
  + Vue 页面状态
  + 页面错误文案
  + 结果展示转换
```

这使 API 无法独立测试，也让后端错误结构变化直接影响页面组件。

### 4.4 状态不容易维护

#### `isLoading` 含义过宽

`isLoading` 同时表示：

- 正在检查文档
- 正在恢复历史结果

但模板把它直接显示成“Python 正在读取并检查 Word 文档”。因此恢复历史结果时，也可能错误显示为正在检查，并显示 0% 进度。

以后可以区分 `isChecking` 和 `isRestoring`，或统一为明确的状态枚举：

```text
idle / restoring / uploading / checking / success / error
```

#### 多个任务字段必须手工同步

以下状态存在强关联：

- `currentJobId`
- `resultFileName`
- `excelDownloadUrl`
- `sheets`
- `selectedSheetName`

任何一个字段漏清理都会产生不一致。

#### `currentSelectionChecked` 表达能力有限

它只表示当前选择是否检查过，与文件列表、任务编号和任务状态存在间接关系。如果以后支持重新检查、取消或检查选项，这个布尔值会更难维护。

#### `sheets` 只保留 Markdown

API 返回二维 rows，但前端立即转换并只保留 Markdown。以后增加表格筛选、分页或其他展示方式时，需要改变整个状态结构。

### 4.5 定时器和轮询问题

#### 异步 `setInterval` 可能产生重叠请求

如果一次请求超过 300 ms，下一次请求仍会启动，可能出现多个同时进行的进度请求。

后果包括：

- 增加后端压力。
- 旧响应较晚返回，覆盖更新进度。
- 网络慢时积累请求。

更稳妥的方式是一次请求完成后，再通过 `setTimeout` 安排下一次请求。

#### 轮询早于后端进度初始化

前端先启动轮询，再发起主检查请求。首次轮询可能返回 404，目前被静默忽略。

#### 轮询错误全部忽略

网络断开、服务停止、服务器异常和返回格式错误都不会显示，也不会停止轮询。

#### 没有组件卸载清理

定时器是 `startCheck()` 的局部变量，只在主请求 `finally` 中清理。没有使用 `onBeforeUnmount()`。

#### 没有请求取消和超时

如果主检查请求长期不返回，页面会一直加载和轮询，用户不能取消，也没有超时提示。

#### 轮询频率较高

300 ms 约等于每秒 3.3 次请求。本机单用户影响较小，但不适合直接扩展到多用户部署。

### 4.6 错误处理不完整

- 前端假定所有响应都能通过 `response.json()` 解析。
- 如果服务器返回 HTML 或纯文本，JSON 解析错误会掩盖真正原因。
- 网络错误通过 `error.message.includes('fetch')` 判断，不同浏览器并不稳定。
- 抛出的对象不一定有 `message`，再次调用 `includes()` 可能产生新错误。
- 进度请求的所有错误都被忽略。
- Excel 下载失败不会在当前页面显示错误。
- 检查失败后没有完整清理任务编号和本地缓存。
- 后端 HTTP 500 直接返回原始异常文本，公开部署时可能暴露内部信息。

## 5. 分阶段重构方案

基本原则：

- 每次只重构一个关注点。
- 不改变现有 API 路径、字段和响应结构。
- 不同时修改页面视觉设计。
- 每阶段单独提交并完成回归验证。
- 前三个阶段不拆页面组件。

## 6. 阶段一：抽离 API 请求

### 为什么需要修改

`App.vue` 当前直接知道 API 路径、HTTP 方法、请求头、FormData 字段、JSON 解析和错误结构。

API 层边界最清晰，抽离时不需要改变模板和后端接口，风险最低。完成后，后续下载、轮询和组件拆分都会更简单。

### 修改文件

- `Frontend/frontend/src/App.vue`

### 新增文件

建议新增：

```text
Frontend/frontend/src/api/documentChecker.js
```

建议承载：

- `checkDocuments(files, progressId)`
- `getCheckProgress(progressId)`
- `getJobResult(jobId)`
- 统一 JSON 和纯文本响应解析
- 统一抛出带 `status`、`detail` 的 API 错误

第一阶段不要同时抽离轮询、下载、Markdown 转换和 Vue 状态。

### 风险

- FormData 字段名 `files` 被误改。
- `X-Progress-ID` 请求头遗漏。
- 错误对象结构变化导致恢复任务的 404 判断失效。
- API 模块返回的数据层级与原来不一致。

### 如何验证

1. 单文件上传成功。
2. 多文件上传成功。
3. 进度正常显示。
4. Sheet 数量和内容不变。
5. 后端未启动时仍显示连接错误。
6. 非 DOCX 文件仍显示格式错误。
7. 上次任务恢复成功。
8. 不存在的历史任务仍能清理缓存并提示。
9. `npm run build` 通过。

## 7. 阶段二：抽离下载逻辑

### 为什么需要修改

下载是独立浏览器能力，与任务检查状态没有直接关系。抽离后可以统一 Blob、URL、文件名和编码处理，并减少 `App.vue` 的重复代码。

### 修改文件

- `Frontend/frontend/src/App.vue`

### 新增文件

建议新增：

```text
Frontend/frontend/src/utils/download.js
```

建议承载：

- `safeFileName(name)`
- `downloadText(content, fileName, mimeType)`
- `downloadUrl(url, fileName)`

Markdown 内容如何组合可以暂时留在 `App.vue`，本阶段只抽离通用下载动作。

### 风险

- UTF-8 BOM 遗漏，中文 Markdown 可能乱码。
- `URL.revokeObjectURL()` 调用时机变化。
- Excel 下载文件名变化。
- 非法字符替换规则不一致。

### 如何验证

1. 下载当前 Sheet Markdown。
2. 下载全部 Sheet Markdown。
3. 确认中文不乱码。
4. 确认 Markdown 标题和表格内容不变。
5. 下载 Excel 并确认文件名和内容不变。
6. 验证包含非法字符的 Sheet 下载名称。
7. `npm run build` 通过。

## 8. 阶段三：抽离任务进度和轮询逻辑

### 为什么需要修改

轮询存在请求重叠、卸载清理、错误忽略、取消和超时等问题，适合抽成 Vue composable，使计时器和生命周期有明确所有者。

### 修改文件

- `Frontend/frontend/src/App.vue`
- `Frontend/frontend/src/api/documentChecker.js`

### 新增文件

建议新增：

```text
Frontend/frontend/src/composables/useCheckProgress.js
```

建议承载：

- `progress`
- `message`
- `currentFile`
- `isPolling`
- `startPolling(progressId)`
- `stopPolling()`
- `resetProgress()`
- 组件卸载时自动清理

建议用请求完成后递归 `setTimeout` 的方式代替异步 `setInterval`，避免请求重叠。

后端进度接口保持不变：

```text
GET /api/check-progress/{progressId}
```

### 风险

- 轮询停止过早，看不到最后进度。
- 检查成功或失败后没有正确停止轮询。
- composable 状态和主任务状态不同步。
- 页面卸载后旧请求回调仍然更新状态。
- 调整轮询方式后刷新节奏略有变化。

### 如何验证

1. 单文件全过程进度正常。
2. 多文件时文件名正确切换。
3. 百分比不倒退。
4. 成功后不再产生进度请求。
5. 失败后不再产生进度请求。
6. 模拟慢速接口，确认进度请求不重叠。
7. 组件卸载时确认定时器停止。
8. 后端暂时不可用时不会无限快速重试。
9. 主检查接口和响应结构不变。

## 9. 阶段四：拆分 Vue 组件

### 为什么最后拆组件

如果先拆组件，API、下载和轮询逻辑会随着组件一起移动，同时改变数据流、DOM 和样式，风险较高。

前三阶段完成后，组件边界会更清楚，此时再拆分页面更安全。

### 修改文件

- `Frontend/frontend/src/App.vue`

根据样式拆分方式，可能还需要修改：

- `Frontend/frontend/src/style.css`

### 新增文件

建议新增：

```text
Frontend/frontend/src/components/UploadPanel.vue
Frontend/frontend/src/components/CheckProgress.vue
Frontend/frontend/src/components/ResultPreview.vue
```

#### `UploadPanel.vue`

- 文件选择区域
- 已选文件数量和名称
- 开始检查按钮
- 下载按钮入口
- 通过 props 接收状态
- 通过 emits 通知选择、开始和下载

#### `CheckProgress.vue`

- 显示百分比
- 显示阶段文本
- 保留 ARIA 进度属性
- 只负责展示，不请求接口

#### `ResultPreview.vue`

- Sheet 选择
- Markdown HTML 展示
- 下载当前 Sheet
- 清除结果事件

最终 `App.vue` 只保留页面级任务协调、结果状态和组件间数据传递。

### 风险

- props 和 emits 名称不一致。
- Sheet 双向选择行为变化。
- scoped CSS 拆分后样式失效。
- DOM 结构变化造成布局回归。
- 文件输入框的 ref 和点击触发关系失效。
- 按钮禁用条件遗漏。

### 如何验证

1. 桌面布局一致。
2. 小于 900 px 时响应式布局一致。
3. 上传区域仍能打开文件选择器。
4. 所有按钮禁用条件一致。
5. 错误信息位置和样式一致。
6. 进度条 ARIA 属性保留。
7. Sheet 切换和 Markdown 展示一致。
8. 三种下载功能一致。
9. 清除结果行为一致。
10. `npm run build` 通过。

## 10. 建议提交顺序

每个阶段单独提交：

```text
refactor: extract document checker api
refactor: extract file download utilities
refactor: extract check progress polling
refactor: split document checker components
```

这些提交中不要同时进行以下修改：

- 修改后端接口
- 更改响应字段
- 改页面视觉设计
- 删除脚手架文件
- 更改后端存储策略
- 引入全局状态管理

以上事项应作为独立任务处理。

## 11. 当前还缺少的信息

当前静态分析所需的核心文件已经具备，没有对关键流程进行猜测。

如果下一步需要制定自动化测试和生产部署方案，还需要确认：

- 是否允许引入 Vitest 和 Vue Test Utils。
- 需要支持哪些浏览器。
- 是否需要配置 CI。
- 是否需要生产部署，而不只是 Vite 本地开发服务器。
- 后端未来是否会启用多个 Uvicorn worker。
- 是否需要重新检查、取消任务或删除服务器文件。

## 12. 简单总结

### 当前最主要的问题

`App.vue` 同时负责接口请求、检查流程、进度轮询、缓存、Markdown 转换、下载、页面结构和样式，职责过多，逻辑耦合较强。

### 第一项应该重构什么

先将所有后端 API 请求抽离到独立的：

```text
src/api/documentChecker.js
```

### 为什么从 API 请求开始

API 请求的边界最清楚，抽离时不需要改变页面结构和后端接口，风险最低。API 层稳定后，下载逻辑、进度轮询和组件拆分都会更容易，也更方便测试。
