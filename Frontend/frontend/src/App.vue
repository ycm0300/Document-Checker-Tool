<script setup>
import { computed, ref } from 'vue'
import MarkdownIt from 'markdown-it'
import * as XLSX from 'xlsx'

const markdown = new MarkdownIt({ html: false, linkify: true, typographer: true })
const fileInput = ref(null)
const resultFileName = ref('')
const resultFile = ref(null)
const sheets = ref([])
const selectedSheetName = ref('')
const errorMessage = ref('')
const isLoading = ref(false)

const selectedSheet = computed(() =>
  sheets.value.find((sheet) => sheet.name === selectedSheetName.value),
)
const renderedMarkdown = computed(() =>
  selectedSheet.value ? markdown.render(selectedSheet.value.markdown) : '',
)

function escapeCell(value) {
  if (value === null || value === undefined) return ''
  return String(value)
    .replace(/\\/g, '\\\\')
    .replace(/\|/g, '\\|')
    .replace(/\r?\n/g, '<br>')
    .trim()
}

function rowsToMarkdown(name, rows) {
  const normalizedRows = rows.filter((row) => row.some((cell) => escapeCell(cell)))
  if (!normalizedRows.length) return `## ${name}\n\n此 Sheet 暂无数据。\n`

  const columnCount = Math.max(...normalizedRows.map((row) => row.length))
  const paddedRows = normalizedRows.map((row) =>
    Array.from({ length: columnCount }, (_, index) => escapeCell(row[index])),
  )
  const [header, ...body] = paddedRows
  const lines = [
    `## ${name}`,
    '',
    `| ${header.join(' | ')} |`,
    `| ${header.map(() => '---').join(' | ')} |`,
    ...body.map((row) => `| ${row.join(' | ')} |`),
  ]
  return `${lines.join('\n')}\n`
}

async function handleResultFile(event) {
  const file = event.target.files?.[0]
  if (!file) return

  errorMessage.value = ''
  isLoading.value = true
  try {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' })
    sheets.value = workbook.SheetNames.map((name) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
        header: 1,
        defval: '',
        raw: false,
      })
      return { name, markdown: rowsToMarkdown(name, rows) }
    })
    resultFileName.value = file.name
    resultFile.value = file
    selectedSheetName.value = sheets.value[0]?.name ?? ''
  } catch (error) {
    sheets.value = []
    selectedSheetName.value = ''
    resultFileName.value = ''
    resultFile.value = null
    errorMessage.value = `无法读取 Excel 文件：${error.message}`
  } finally {
    isLoading.value = false
  }
}

function downloadMarkdown(content, fileName) {
  const blob = new Blob(['\uFEFF', content], { type: 'text/markdown;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  link.click()
  URL.revokeObjectURL(url)
}

function safeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_')
}

function downloadCurrentSheet() {
  if (!selectedSheet.value) return
  downloadMarkdown(
    `# 文档合规检查结果\n\n${selectedSheet.value.markdown}`,
    `${safeFileName(selectedSheet.value.name)}.md`,
  )
}

function downloadAllSheets() {
  if (!sheets.value.length) return
  const title = resultFileName.value.replace(/\.xlsx?$/i, '') || '文档合规检查结果'
  const content = `# ${title}\n\n${sheets.value.map((sheet) => sheet.markdown).join('\n')}`
  downloadMarkdown(content, `${safeFileName(title)}.md`)
}

function downloadOriginalExcel() {
  if (!resultFile.value) return
  const url = URL.createObjectURL(resultFile.value)
  const link = document.createElement('a')
  link.href = url
  link.download = resultFile.value.name
  link.click()
  URL.revokeObjectURL(url)
}
</script>

<template>
  <main class="app-shell">
    <header class="page-header">
      <div>
        <p class="eyebrow">DOCUMENT COMPLIANCE</p>
        <h1>文档合规检查工具</h1>
        <p class="subtitle">导入 Python 生成的多 Sheet Excel，按 Sheet 查看和导出 Markdown。</p>
      </div>
      <span class="status-badge">本地处理 · 文件不会上传</span>
    </header>

    <section class="workspace">
      <aside class="panel control-panel">
        <div class="panel-heading">
          <span class="step">1</span>
          <div><h2>导入检查结果</h2><p>支持 .xlsx 和 .xls 文件</p></div>
        </div>

        <input
          ref="fileInput"
          class="sr-only"
          type="file"
          accept=".xlsx,.xls"
          @change="handleResultFile"
        />
        <button class="upload-box" type="button" @click="fileInput.click()">
          <span class="upload-icon">↑</span>
          <strong>{{ resultFileName || '选择检查结果 Excel' }}</strong>
          <small>{{ resultFileName ? '点击可重新选择文件' : 'Python 输出的总检查结果.xlsx' }}</small>
        </button>

        <p v-if="errorMessage" class="error-message">{{ errorMessage }}</p>

        <div class="info-card">
          <span>Sheet 数量</span><strong>{{ sheets.length }}</strong>
        </div>

        <div class="download-actions">
          <button class="excel-button" type="button" :disabled="!resultFile" @click="downloadOriginalExcel">
            下载原始 Excel
          </button>
          <button class="primary-button" type="button" :disabled="!sheets.length" @click="downloadAllSheets">
            下载完整 Markdown
          </button>
        </div>
      </aside>

      <section class="panel preview-panel">
        <div class="preview-toolbar">
          <div>
            <p class="section-label">MARKDOWN PREVIEW</p>
            <h2>检查结果预览</h2>
          </div>
          <div class="toolbar-actions">
            <label>
              <span>Sheet</span>
              <select v-model="selectedSheetName" :disabled="!sheets.length">
                <option v-if="!sheets.length" value="">请先导入 Excel</option>
                <option v-for="sheet in sheets" :key="sheet.name" :value="sheet.name">
                  {{ sheet.name }}
                </option>
              </select>
            </label>
            <button class="secondary-button" type="button" :disabled="!selectedSheet" @click="downloadCurrentSheet">
              下载当前 MD
            </button>
          </div>
        </div>

        <div v-if="isLoading" class="empty-state">正在读取 Excel...</div>
        <article v-else-if="selectedSheet" class="markdown-body" v-html="renderedMarkdown"></article>
        <div v-else class="empty-state">
          <div class="empty-icon">MD</div>
          <strong>暂无预览内容</strong>
          <p>从左侧选择 Python 生成的检查结果 Excel</p>
        </div>
      </section>
    </section>
  </main>
</template>

<style scoped>
.app-shell { min-height: 100vh; padding: 36px; background: #f3f6fb; color: #172033; }
.page-header { max-width: 1440px; margin: 0 auto 24px; display: flex; justify-content: space-between; align-items: end; gap: 24px; }
.eyebrow,.section-label { margin: 0 0 7px; color: #2563eb; font-size: 11px; font-weight: 800; letter-spacing: .12em; }
h1 { margin: 0; font-size: 30px; letter-spacing: -.03em; }
.subtitle { margin: 8px 0 0; color: #667085; font-size: 14px; }
.status-badge { padding: 8px 12px; border: 1px solid #cfe0ff; border-radius: 999px; background: #eef5ff; color: #2459a9; font-size: 12px; }
.workspace { max-width: 1440px; margin: 0 auto; display: grid; grid-template-columns: 300px minmax(0,1fr); gap: 20px; align-items: start; }
.panel { box-sizing: border-box; border: 1px solid #e2e8f0; border-radius: 14px; background: #fff; box-shadow: 0 8px 24px rgba(35,55,80,.05); }
.control-panel { padding: 20px; }
.panel-heading { display: flex; align-items: center; gap: 11px; margin-bottom: 18px; }
.panel-heading h2,.preview-toolbar h2 { margin: 0; font-size: 17px; }
.panel-heading p { margin: 3px 0 0; color: #98a2b3; font-size: 12px; }
.step { display: grid; width: 30px; height: 30px; place-items: center; border-radius: 9px; background: #2563eb; color: white; font-weight: 700; }
.upload-box { width: 100%; min-height: 150px; padding: 18px; display: flex; flex-direction: column; justify-content: center; align-items: center; gap: 8px; border: 1px dashed #adc6ee; border-radius: 12px; background: #f8fbff; color: #344054; cursor: pointer; }
.upload-box:hover { border-color: #2563eb; background: #f0f6ff; }
.upload-icon { display: grid; width: 36px; height: 36px; place-items: center; border-radius: 50%; background: #e6efff; color: #2563eb; font-size: 22px; }
.upload-box small { color: #98a2b3; }
.info-card { margin: 16px 0; padding: 12px 14px; display: flex; justify-content: space-between; border-radius: 9px; background: #f7f8fa; color: #667085; font-size: 13px; }
.info-card strong { color: #172033; }
.primary-button,.secondary-button { height: 38px; padding: 0 14px; border-radius: 8px; font-weight: 650; cursor: pointer; }
.primary-button { width: 100%; border: 0; background: #2563eb; color: white; }
.secondary-button { border: 1px solid #d5dce7; background: #fff; color: #344054; white-space: nowrap; }
.download-actions { display: grid; gap: 9px; }
.excel-button { width: 100%; height: 38px; padding: 0 14px; border: 1px solid #b8d8c5; border-radius: 8px; background: #f0faf4; color: #187044; font-weight: 650; cursor: pointer; }
.excel-button:hover:not(:disabled) { background: #e4f6eb; border-color: #80bd99; }
button:disabled { opacity: .45; cursor: not-allowed; }
.error-message { color: #b42318; font-size: 12px; line-height: 1.5; }
.preview-panel { min-width: 0; overflow: hidden; }
.preview-toolbar { padding: 18px 20px; display: flex; justify-content: space-between; align-items: center; gap: 20px; border-bottom: 1px solid #e8edf4; }
.toolbar-actions { display: flex; align-items: end; gap: 10px; }
.toolbar-actions label { display: grid; gap: 5px; color: #667085; font-size: 11px; }
select { min-width: 250px; height: 38px; padding: 0 34px 0 11px; border: 1px solid #d5dce7; border-radius: 8px; background: #fff; color: #344054; }
.empty-state { min-height: 510px; display: flex; flex-direction: column; align-items: center; justify-content: center; color: #98a2b3; }
.empty-state strong { color: #667085; }
.empty-state p { margin: 7px 0 0; font-size: 13px; }
.empty-icon { margin-bottom: 12px; display: grid; width: 54px; height: 54px; place-items: center; border: 1px solid #d8e1ef; border-radius: 14px; background: #f7f9fc; color: #4773b8; font-weight: 800; }
.markdown-body { min-height: 510px; max-height: calc(100vh - 210px); padding: 24px; overflow: auto; color: #344054; font-size: 14px; line-height: 1.65; }
.markdown-body :deep(h2) { margin: 0 0 18px; color: #172033; font-size: 22px; }
.markdown-body :deep(table) { width: max-content; min-width: 100%; border-collapse: collapse; font-size: 13px; }
.markdown-body :deep(th),.markdown-body :deep(td) { max-width: 440px; padding: 10px 12px; border: 1px solid #dfe5ee; text-align: left; vertical-align: top; white-space: normal; overflow-wrap: anywhere; }
.markdown-body :deep(th) { position: sticky; top: 0; background: #f0f5fc; color: #344054; }
.markdown-body :deep(tr:nth-child(even) td) { background: #fafbfd; }
.sr-only { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0,0,0,0); }
@media (max-width: 900px) { .app-shell { padding: 22px; } .workspace { grid-template-columns: 1fr; } .page-header { align-items: start; } .status-badge { display: none; } .preview-toolbar { align-items: start; flex-direction: column; } .toolbar-actions { width: 100%; } .toolbar-actions label { flex: 1; } select { width: 100%; min-width: 0; } }
</style>
