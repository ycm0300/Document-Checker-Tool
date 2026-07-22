<script setup>
import { computed, onMounted, ref, watch } from 'vue'
import MarkdownIt from 'markdown-it'
import { checkDocuments } from './api/documentCheckers'
import { getCheckResult } from './api/documentCheckers'
import { getJobResult } from './api/documentCheckers'

const markdown = new MarkdownIt({ html: false, linkify: true, typographer: true })
const fileInput = ref(null)
const selectedFiles = ref([])
const currentJobId = ref('')
const resultFileName = ref('')
const excelDownloadUrl = ref('')
const sheets = ref([])
const selectedSheetName = ref('')
const errorMessage = ref('')
const isLoading = ref(false)
const checkProgress = ref(0)
const progressMessage = ref('')
const currentSelectionChecked = ref(false)
const LAST_JOB_KEY = 'document-checker:last-job-id'
const LAST_SHEET_KEY = 'document-checker:last-sheet-name'

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

function applyJobResult(data, preferredSheetName = '') {
  sheets.value = data.sheets.map((sheet) => ({
    name: sheet.name,
    markdown: rowsToMarkdown(sheet.name, sheet.rows),
  }))
  currentJobId.value = data.jobId
  resultFileName.value = data.fileName
  excelDownloadUrl.value = data.excelDownloadUrl
  selectedSheetName.value = sheets.value.some((sheet) => sheet.name === preferredSheetName)
    ? preferredSheetName
    : (sheets.value[0]?.name ?? '')
  localStorage.setItem(LAST_JOB_KEY, data.jobId)
}

function handleWordFiles(event) {
  selectedFiles.value = Array.from(event.target.files ?? [])
  currentSelectionChecked.value = false
  errorMessage.value = ''
}

async function startCheck() {
  if (currentSelectionChecked.value) return
  if (!selectedFiles.value.length) {
    errorMessage.value = '请先选择至少一个 Word 文档。'
    return
  }
  errorMessage.value = ''
  isLoading.value = true
  checkProgress.value = 1
  progressMessage.value = '准备上传文档'
  const progressId = crypto.randomUUID().replaceAll('-', '')
  const progressTimer = window.setInterval(async () => {
    try {
      const progress = await getCheckResult(progressId)
      checkProgress.value = progress.percent
      progressMessage.value = progress.message
    } catch {
      // 主检查请求负责展示连接错误，进度轮询失败无需重复提示。
    }
  }, 300)
  try {
    const data = await checkDocuments(
      selectedFiles.value,
      progressId,
    )

    applyJobResult(data)
    currentSelectionChecked.value = true
    checkProgress.value = 100
    progressMessage.value = '检查完成'
  } catch (error) {
    sheets.value = []
    selectedSheetName.value = ''
    resultFileName.value = ''
    excelDownloadUrl.value = ''
    errorMessage.value = error.message.includes('fetch')
      ? '无法连接 Python 检查服务，请确认后端已启动。'
      : error.message
  } finally {
    window.clearInterval(progressTimer)
    isLoading.value = false
  }
}

async function restoreLastResult() {
  const jobId = localStorage.getItem(LAST_JOB_KEY)
  if (!jobId) return

  isLoading.value = true
  errorMessage.value = ''
  try {
    const data = await getJobResult(jobId)
    applyJobResult(
      data,
      localStorage.getItem(LAST_SHEET_KEY) || '',
)
  } catch (error) {
    if (error.status === 404) {
      localStorage.removeItem(LAST_JOB_KEY)
      localStorage.removeItem(LAST_SHEET_KEY)
      errorMessage.value = '上次检查结果已不存在，请重新上传文档检查。'
    } else {
      errorMessage.value = '暂时无法恢复上次结果，请确认 Python 检查服务已启动。'
    }
  } finally {
    isLoading.value = false
  }
}

function clearCurrentResult() {
  currentJobId.value = ''
  resultFileName.value = ''
  excelDownloadUrl.value = ''
  sheets.value = []
  selectedSheetName.value = ''
  selectedFiles.value = []
  currentSelectionChecked.value = false
  checkProgress.value = 0
  progressMessage.value = ''
  if (fileInput.value) fileInput.value.value = ''
  errorMessage.value = ''
  localStorage.removeItem(LAST_JOB_KEY)
  localStorage.removeItem(LAST_SHEET_KEY)
}

watch(selectedSheetName, (name) => {
  if (name) localStorage.setItem(LAST_SHEET_KEY, name)
})

onMounted(restoreLastResult)

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
  if (!excelDownloadUrl.value) return
  const link = document.createElement('a')
  link.href = excelDownloadUrl.value
  link.download = resultFileName.value || '总检查结果.xlsx'
  link.click()
}
</script>

<template>
  <main class="app-shell">
    <header class="page-header">
      <div>
        <p class="eyebrow">DOCUMENT COMPLIANCE</p>
        <h1>文档合规检查工具</h1>
        <p class="subtitle">上传 Word 文档，调用 Python 检查并按 Sheet 预览结果。</p>
      </div>
      <span class="status-badge">本机 Python 服务</span>
    </header>

    <section class="workspace">
      <aside class="panel control-panel">
        <div class="panel-heading">
          <span class="step">1</span>
          <div><h2>上传待检文档</h2><p>支持一个或多个 .docx 文件</p></div>
        </div>

        <input
          ref="fileInput"
          class="sr-only"
          type="file"
          accept=".docx"
          multiple
          @change="handleWordFiles"
        />
        <button class="upload-box" type="button" @click="fileInput.click()">
          <span class="upload-icon">↑</span>
          <strong>{{ selectedFiles.length ? `已选择 ${selectedFiles.length} 个文档` : '选择 Word 文档' }}</strong>
          <div v-if="selectedFiles.length" class="selected-file-list">
            <div
              v-for="file in selectedFiles"
              :key="`${file.name}-${file.size}-${file.lastModified}`"
              class="selected-file-row"
              :title="file.name"
            >
              <span class="word-file-icon" aria-hidden="true">W</span>
              <span class="selected-file-name">{{ file.name }}</span>
            </div>
          </div>
          <small v-else>可同时选择多个 .docx 文件</small>
        </button>

        <p v-if="errorMessage" class="error-message">{{ errorMessage }}</p>

        <div class="info-card">
          <span>待检查文档</span><strong>{{ selectedFiles.length }}</strong>
        </div>

        <div v-if="isLoading" class="progress-card" aria-live="polite">
          <div class="progress-label">
            <span>{{ progressMessage || '正在检查文档' }}</span>
            <strong>{{ checkProgress }}%</strong>
          </div>
          <div
            class="progress-track"
            role="progressbar"
            aria-label="文件检查进度"
            :aria-valuenow="checkProgress"
            aria-valuemin="0"
            aria-valuemax="100"
          >
            <span class="progress-fill" :style="{ width: `${checkProgress}%` }"></span>
          </div>
        </div>

        <div class="download-actions">
          <button class="primary-button" type="button" :disabled="!selectedFiles.length || isLoading || currentSelectionChecked" @click="startCheck">
            {{ isLoading ? '正在检查…' : currentSelectionChecked ? '本批文档已检查' : '开始检查' }}
          </button>
          <button class="excel-button" type="button" :disabled="!excelDownloadUrl" @click="downloadOriginalExcel">
            下载检查结果 Excel
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
                <option v-if="!sheets.length" value="">暂无检查结果</option>
                <option v-for="sheet in sheets" :key="sheet.name" :value="sheet.name">
                  {{ sheet.name }}
                </option>
              </select>
            </label>
            <button class="secondary-button" type="button" :disabled="!selectedSheet" @click="downloadCurrentSheet">
              下载当前 MD
            </button>
            <button class="clear-button" type="button" :disabled="!currentJobId" @click="clearCurrentResult">
              清除结果
            </button>
          </div>
        </div>

        <div v-if="isLoading" class="empty-state">Python 正在读取并检查 Word 文档，请稍候...</div>
        <article v-else-if="selectedSheet" class="markdown-body" v-html="renderedMarkdown"></article>
        <div v-else class="empty-state">
          <div class="empty-icon">MD</div>
          <strong>暂无预览内容</strong>
          <p>从左侧选择 Word 文档并开始检查</p>
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
.workspace { max-width: 1440px; margin: 0 auto; display: grid; grid-template-columns: 300px minmax(0,1fr); gap: 20px; align-items: stretch; }
.panel { box-sizing: border-box; border: 1px solid #e2e8f0; border-radius: 14px; background: #fff; box-shadow: 0 8px 24px rgba(35,55,80,.05); }
.control-panel { padding: 20px; }
.panel-heading { display: flex; align-items: center; gap: 11px; margin-bottom: 18px; }
.panel-heading h2,.preview-toolbar h2 { margin: 0; font-size: 17px; }
.panel-heading p { margin: 3px 0 0; color: #98a2b3; font-size: 12px; }
.step { display: grid; width: 30px; height: 30px; place-items: center; border-radius: 9px; background: #2563eb; color: white; font-weight: 700; }
.upload-box { width: 100%; min-height: 150px; padding: 18px; display: flex; flex-direction: column; justify-content: center; align-items: stretch; gap: 8px; border: 1px dashed #adc6ee; border-radius: 12px; background: #f8fbff; color: #344054; cursor: pointer; }
.upload-box:hover { border-color: #2563eb; background: #f0f6ff; }
.upload-icon { align-self: center; display: grid; width: 36px; height: 36px; place-items: center; border-radius: 50%; background: #e6efff; color: #2563eb; font-size: 22px; }
.upload-box > strong,.upload-box > small { text-align: center; }
.upload-box small { color: #98a2b3; }
.selected-file-list { max-height: 228px; margin-top: 4px; display: grid; gap: 6px; overflow-y: auto; scrollbar-width: thin; }
.selected-file-row { min-width: 0; height: 32px; padding: 0 9px; display: flex; align-items: center; gap: 8px; border: 1px solid #dbe6f5; border-radius: 7px; background: #fff; text-align: left; }
.word-file-icon { flex: none; display: grid; width: 20px; height: 22px; place-items: center; border-radius: 3px; background: #2563eb; color: #fff; font-size: 11px; font-weight: 800; }
.selected-file-name { min-width: 0; flex: 1; overflow: hidden; color: #52637a; font-size: 12px; text-overflow: ellipsis; white-space: nowrap; }
.info-card { margin: 16px 0; padding: 12px 14px; display: flex; justify-content: space-between; border-radius: 9px; background: #f7f8fa; color: #667085; font-size: 13px; }
.info-card strong { color: #172033; }
.progress-card { margin: -4px 0 16px; }
.progress-label { margin-bottom: 7px; display: flex; justify-content: space-between; gap: 10px; color: #667085; font-size: 12px; }
.progress-label span { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.progress-label strong { flex: none; color: #2563eb; }
.progress-track { height: 8px; overflow: hidden; border-radius: 999px; background: #e8eef8; }
.progress-fill { display: block; height: 100%; border-radius: inherit; background: linear-gradient(90deg, #2563eb, #60a5fa); transition: width .25s ease; }
.primary-button,.secondary-button { height: 38px; padding: 0 14px; border-radius: 8px; font-weight: 650; cursor: pointer; }
.primary-button { width: 100%; border: 0; background: #2563eb; color: white; }
.secondary-button { border: 1px solid #d5dce7; background: #fff; color: #344054; white-space: nowrap; }
.clear-button { height: 38px; padding: 0 13px; border: 1px solid #efc5c2; border-radius: 8px; background: #fff8f7; color: #b42318; font-weight: 650; cursor: pointer; white-space: nowrap; }
.download-actions { display: grid; gap: 9px; }
.excel-button { width: 100%; height: 38px; padding: 0 14px; border: 1px solid #b8d8c5; border-radius: 8px; background: #f0faf4; color: #187044; font-weight: 650; cursor: pointer; }
.excel-button:hover:not(:disabled) { background: #e4f6eb; border-color: #80bd99; }
button:disabled { opacity: .45; cursor: not-allowed; }
.error-message { color: #b42318; font-size: 12px; line-height: 1.5; }
.preview-panel { min-width: 0; display: flex; flex-direction: column; overflow: hidden; }
.preview-toolbar { padding: 18px 20px; display: flex; justify-content: space-between; align-items: center; gap: 20px; border-bottom: 1px solid #e8edf4; }
.toolbar-actions { display: flex; align-items: end; gap: 10px; }
.toolbar-actions label { display: grid; gap: 5px; color: #667085; font-size: 11px; }
select { min-width: 250px; height: 38px; padding: 0 34px 0 11px; border: 1px solid #d5dce7; border-radius: 8px; background: #fff; color: #344054; }
.empty-state { min-height: 510px; flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; color: #98a2b3; }
.empty-state strong { color: #667085; }
.empty-state p { margin: 7px 0 0; font-size: 13px; }
.empty-icon { margin-bottom: 12px; display: grid; width: 54px; height: 54px; place-items: center; border: 1px solid #d8e1ef; border-radius: 14px; background: #f7f9fc; color: #4773b8; font-weight: 800; }
.markdown-body { min-height: 510px; max-height: calc(100vh - 210px); flex: 1; padding: 24px; overflow: auto; color: #344054; font-size: 14px; line-height: 1.65; }
.markdown-body :deep(h2) { margin: 0 0 18px; color: #172033; font-size: 22px; }
.markdown-body :deep(table) { width: max-content; min-width: 100%; border-collapse: collapse; font-size: 13px; }
.markdown-body :deep(th),.markdown-body :deep(td) { max-width: 440px; padding: 10px 12px; border: 1px solid #dfe5ee; text-align: left; vertical-align: top; white-space: normal; overflow-wrap: anywhere; }
.markdown-body :deep(th) { background: #f0f5fc; color: #344054; }
.markdown-body :deep(tr:nth-child(even) td) { background: #fafbfd; }
.sr-only { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0,0,0,0); }
@media (max-width: 900px) { .app-shell { padding: 22px; } .workspace { grid-template-columns: 1fr; align-items: start; } .panel { width: 100%; } .page-header { align-items: start; } .status-badge { display: none; } .preview-toolbar { align-items: start; flex-direction: column; } .toolbar-actions { width: 100%; } .toolbar-actions label { flex: 1; } select { width: 100%; min-width: 0; } }
</style>
