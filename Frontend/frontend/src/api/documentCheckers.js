export async function checkDocuments(files, progressId) {
  // 上传并检查文档 POST /api/check
  const formData = new FormData()
  files.forEach((file) => formData.append('files', file))
  const response = await fetch('/api/check', {
    method: 'POST',
    headers: { 'X-Progress-ID': progressId },
    body: formData,
  })
  const data = await response.json()
  if (!response.ok) throw new Error(data.detail || '检查请求失败')
  return data
}

export async function getCheckResult(progressId) {
  // 查询进度条接口
  const response = await fetch(`/api/check-progress/${progressId}`)
  if (!response.ok) return
  const progress = await response.json()
  return progress
}

export async function getJobResult(jobId) {
  // 恢复历史任务结果
  const response = await fetch(`/api/jobs/${jobId}`)
  const data = await response.json()
  if (!response.ok) throw Object.assign(new Error(data.detail || '无法恢复上次检查结果'), { status: response.status })
  return data
}
