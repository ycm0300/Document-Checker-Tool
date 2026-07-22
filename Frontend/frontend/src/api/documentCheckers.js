export async function checkDocuments(files, progressId) {
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