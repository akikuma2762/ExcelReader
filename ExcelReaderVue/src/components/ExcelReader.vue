<template>
  <div class="container">
    <h1>Excel 讀取器</h1>

    <div class="upload-section">
      <div class="upload-area" @drop="handleDrop" @dragover.prevent @dragenter.prevent>
        <input
          ref="fileInput"
          type="file"
          accept=".xlsx,.xls"
          @change="handleFileSelect"
          style="display: none"
        />
        <button @click="fileInput?.click()" class="upload-btn">
          選擇檔案
        </button>
        <p>或拖拽 Excel 檔案到此處</p>
        <p class="file-info">支援格式：.xlsx, .xls</p>
      </div>

      <div class="button-group">
        <button @click="loadSampleData" class="sample-btn">
          載入範例資料
        </button>
        <button @click="downloadSampleFile" class="download-btn">
          下載範例檔案
        </button>
      </div>
    </div>

    <div v-if="loading" class="loading">
      上傳中...
    </div>

    <div v-if="message" class="message" :class="messageType">
      {{ message }}
    </div>

    <div v-if="excelData" class="data-section">
      <h2>{{ excelData.fileName }}</h2>
      <p>
        工作表：{{ excelData.worksheetName }} |
        總行數：{{ excelData.totalRows }} |
        總欄數：{{ excelData.totalColumns }}
      </p>

      <div v-if="excelData.availableWorksheets.length > 1" class="worksheet-info">
        <p>可用工作表：{{ excelData.availableWorksheets.join(', ') }}</p>
      </div>

      <div class="table-container">
        <table class="data-table">
          <thead>
            <tr>
              <template v-for="(header, index) in excelData.headers[0]" :key="index">
                <th
                  v-if="shouldRenderCell(header)"
                  :style="getHeaderStyle(header)"
                  :colspan="header.colSpan || 1"
                  :rowspan="header.rowSpan || 1"
                >
                  <span v-if="header.isRichText" v-html="renderRichText(header)"></span>
                  <span v-else v-html="formatTextWithLineBreaks(getDisplayValue(header))"></span>
                  <div class="format-info" v-if="showFormatInfo">
                    <small>格式: {{ header.formatCode || '一般' }}</small>
                    <small v-if="header.isRichText" style="color: orange;">Rich Text</small>
                  </div>
                </th>
              </template>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(row, rowIndex) in excelData.rows" :key="rowIndex">
              <template v-for="(cell, cellIndex) in row" :key="cellIndex">
                <td
                  v-if="shouldRenderCell(cell)"
                  :class="getCellClass(cell)"
                  :style="getCellStyle(cell)"
                  :title="getCellTooltip(cell)"
                  :colspan="cell.colSpan || 1"
                  :rowspan="cell.rowSpan || 1"
                >
                  <span v-if="cell.isRichText" v-html="renderRichText(cell)"></span>
                  <span v-else v-html="formatTextWithLineBreaks(getDisplayValue(cell))"></span>
                </td>
              </template>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="format-controls">
        <label>
          <input type="checkbox" v-model="showFormatInfo" />
          顯示格式信息
        </label>
        <label>
          <input type="checkbox" v-model="showOriginalValue" />
          顯示原始值
        </label>
      </div>

      <div class="json-section">
        <h3>JSON 資料：</h3>
        <button @click="toggleJsonView" class="toggle-btn">
          {{ showJson ? '隱藏' : '顯示' }} JSON
        </button>
        <pre v-if="showJson" class="json-display">{{ JSON.stringify(excelData, null, 2) }}</pre>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue'
import axios from 'axios'

interface RichTextPart {
  text: string
  fontBold?: boolean
  fontItalic?: boolean
  fontUnderline?: boolean
  fontSize?: number
  fontName?: string
  fontColor?: string
}

interface ExcelCellInfo {
  value: string | number | boolean | Date | null
  displayText: string
  formatCode: string
  dataType: string
  fontBold?: boolean
  fontSize?: number
  fontName?: string
  backgroundColor?: string
  fontColor?: string
  textAlign?: string
  columnWidth?: number
  richText?: RichTextPart[]
  isRichText?: boolean
  rowSpan?: number
  colSpan?: number
  isMerged?: boolean
  isMainMergedCell?: boolean
}

interface ExcelData {
  headers: ExcelCellInfo[][]
  rows: ExcelCellInfo[][]
  totalRows: number
  totalColumns: number
  fileName: string
  worksheetName: string
  availableWorksheets: string[]
}

interface UploadResponse {
  success: boolean
  message: string
  data?: ExcelData
}

const loading = ref<boolean>(false)
const message = ref<string>('')
const messageType = ref<'success' | 'error' | ''>('')
const excelData = ref<ExcelData | null>(null)
const showJson = ref<boolean>(false)
const fileInput = ref<HTMLInputElement | null>(null)
const showFormatInfo = ref<boolean>(false)
const showOriginalValue = ref<boolean>(false)

const API_BASE_URL = 'http://localhost:5280/api' // API伺服器URL

const clearMessage = () => {
  setTimeout(() => {
    message.value = ''
    messageType.value = ''
  }, 5000)
}

const handleFileSelect = (event: Event) => {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (file) {
    uploadFile(file)
  }
}

const handleDrop = (event: DragEvent) => {
  event.preventDefault()
  const file = event.dataTransfer?.files[0]
  if (file) {
    uploadFile(file)
  }
}

const uploadFile = async (file: File) => {
  if (!file) return

  loading.value = true
  message.value = ''
  excelData.value = null

  const formData = new FormData()
  formData.append('file', file)

  try {
    const response = await axios.post<UploadResponse>(`${API_BASE_URL}/excel/upload`, formData, {
      headers: {
        'Content-Type': 'multipart/form-data'
      }
    })

    if (response.data.success) {
      excelData.value = response.data.data || null
      message.value = response.data.message
      messageType.value = 'success'
    } else {
      message.value = response.data.message
      messageType.value = 'error'
    }
  } catch (error: unknown) {
    const axiosError = error as { response?: { data?: { message?: string } }; message?: string }
    message.value = `上傳失敗：${axiosError.response?.data?.message || axiosError.message || '未知錯誤'}`
    messageType.value = 'error'
  } finally {
    loading.value = false
    clearMessage()
  }
}

const loadSampleData = async () => {
  loading.value = true
  message.value = ''
  excelData.value = null

  try {
    const response = await axios.get<ExcelData>(`${API_BASE_URL}/excel/sample`)
    excelData.value = response.data
    message.value = '已載入範例資料'
    messageType.value = 'success'
  } catch (error: unknown) {
    const axiosError = error as { message?: string }
    message.value = `載入範例資料失敗：${axiosError.message || '未知錯誤'}`
    messageType.value = 'error'
  } finally {
    loading.value = false
    clearMessage()
  }
}

const toggleJsonView = () => {
  showJson.value = !showJson.value
}

const downloadSampleFile = async () => {
  try {
    const response = await axios.get(`${API_BASE_URL}/excel/download-sample`, {
      responseType: 'blob'
    })

    const url = window.URL.createObjectURL(new Blob([response.data]))
    const link = document.createElement('a')
    link.href = url
    link.setAttribute('download', '範例員工資料.xlsx')
    document.body.appendChild(link)
    link.click()
    link.remove()
    window.URL.revokeObjectURL(url)

    message.value = '範例檔案已下載'
    messageType.value = 'success'
    clearMessage()
  } catch (error: unknown) {
    const axiosError = error as { message?: string }
    message.value = `下載失敗：${axiosError.message || '未知錯誤'}`
    messageType.value = 'error'
    clearMessage()
  }
}

const getDisplayValue = (cell: ExcelCellInfo): string => {
  if (showOriginalValue.value) {
    return cell.value?.toString() || ''
  }
  return cell.displayText || ''
}

// 新增：渲染Rich Text的HTML
const renderRichText = (cell: ExcelCellInfo): string => {
  if (!cell.isRichText || !cell.richText) {
    // 處理一般文字的換行
    return formatTextWithLineBreaks(cell.displayText || '')
  }

  return cell.richText.map((part: RichTextPart) => {
    // HTML轉義文字內容以防止XSS，並處理換行
    let html = formatTextWithLineBreaks(escapeHtml(part.text))
    const styles: string[] = []

    if (part.fontBold) styles.push('font-weight: bold')
    if (part.fontItalic) styles.push('font-style: italic')
    if (part.fontUnderline) styles.push('text-decoration: underline')
    if (part.fontSize && part.fontSize > 0) styles.push(`font-size: ${part.fontSize}pt`)
    if (part.fontName && part.fontName.trim()) styles.push(`font-family: ${part.fontName}`)
    if (part.fontColor) styles.push(`color: ${part.fontColor}`)

    if (styles.length > 0) {
      html = `<span style="${styles.join('; ')}">${html}</span>`
    }

    return html
  }).join('')
}

// 處理文字換行的函數
const formatTextWithLineBreaks = (text: string): string => {
  return text.replace(/\r\n/g, '<br>').replace(/\n/g, '<br>').replace(/\r/g, '<br>')
}

// HTML轉義函數以防止XSS攻擊
const escapeHtml = (text: string): string => {
  const div = document.createElement('div')
  div.textContent = text
  return div.innerHTML
}

// 將 Excel 欄寬轉換為像素寬度
const convertExcelWidthToPixels = (excelWidth: number): number => {
  // Excel 欄寬是以字符為單位，1 字符 ≈ 7 像素（基於 Arial 10pt）
  // 但實際轉換會考慮padding和borders，所以使用較精確的公式
  return Math.round(excelWidth * 7.5)
}

const getHeaderStyle = (header: ExcelCellInfo) => {
  const style: Record<string, string> = {}

  if (header.fontBold) {
    style.fontWeight = 'bold'
  }

  if (header.fontSize) {
    style.fontSize = `${header.fontSize}px`
  }

  if (header.fontName) {
    style.fontFamily = `"${header.fontName}"`
  }

  if (header.backgroundColor) {
    style.backgroundColor = `#${header.backgroundColor}`
  }

  if (header.fontColor) {
    style.color = `#${header.fontColor}`
  }

  if (header.textAlign) {
    style.textAlign = header.textAlign
  }

  if (header.columnWidth) {
    style.width = `${convertExcelWidthToPixels(header.columnWidth)}px`
  }

  return style
}

const getCellClass = (cell: ExcelCellInfo): string => {
  const classes = ['cell']

  switch (cell.dataType) {
    case 'DateTime':
      classes.push('cell-date')
      break
    case 'Number':
    case 'Integer':
      classes.push('cell-number')
      break
    case 'Boolean':
      classes.push('cell-boolean')
      break
    case 'Empty':
      classes.push('cell-empty')
      break
    default:
      classes.push('cell-text')
  }

  return classes.join(' ')
}

const getCellStyle = (cell: ExcelCellInfo) => {
  const style: Record<string, string> = {}

  if (cell.fontBold) {
    style.fontWeight = 'bold'
  }

  if (cell.fontSize) {
    style.fontSize = `${cell.fontSize}px`
  }

  if (cell.fontName) {
    style.fontFamily = `"${cell.fontName}"`
  }

  if (cell.backgroundColor) {
    style.backgroundColor = `#${cell.backgroundColor}`
  }

  if (cell.fontColor) {
    style.color = `#${cell.fontColor}`
  }

  if (cell.textAlign) {
    style.textAlign = cell.textAlign
  }

  if (cell.columnWidth) {
    style.width = `${convertExcelWidthToPixels(cell.columnWidth)}px`
  }

  return style
}

const getCellTooltip = (cell: ExcelCellInfo): string => {
  const parts = []

  parts.push(`類型: ${cell.dataType}`)

  if (cell.formatCode) {
    parts.push(`格式: ${cell.formatCode}`)
  }

  if (cell.value !== null && cell.value !== undefined) {
    parts.push(`原始值: ${cell.value}`)
  }

  if (cell.displayText) {
    parts.push(`顯示文字: ${cell.displayText}`)
  }

  if (cell.isRichText && cell.richText) {
    parts.push(`Rich Text 片段數: ${cell.richText.length}`)
  }

  if (cell.isMerged && cell.rowSpan && cell.colSpan) {
    parts.push(`合併儲存格: ${cell.rowSpan}行 x ${cell.colSpan}欄`)
  }

  return parts.join('\n')
}

const shouldRenderCell = (cell: ExcelCellInfo): boolean => {
  // 如果不是合併儲存格，正常顯示
  if (!cell.isMerged) {
    return true
  }
  
  // 如果是合併儲存格，只顯示主儲存格
  return cell.isMainMergedCell === true
}
</script>

<style scoped>
.container {
  max-width: 1200px;
  margin: 0 auto;
  padding: 20px;
  font-family: Arial, sans-serif;
}

h1 {
  text-align: center;
  color: #333;
  margin-bottom: 30px;
}

.upload-section {
  text-align: center;
  margin-bottom: 30px;
}

.upload-area {
  border: 2px dashed #ccc;
  border-radius: 8px;
  padding: 40px;
  margin-bottom: 20px;
  transition: border-color 0.3s;
}

.upload-area:hover {
  border-color: #007bff;
}

.upload-btn {
  background-color: #007bff;
  color: white;
  border: none;
  padding: 12px 24px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 16px;
  margin-bottom: 10px;
}

.upload-btn:hover {
  background-color: #0056b3;
}

.button-group {
  display: flex;
  gap: 10px;
  justify-content: center;
  flex-wrap: wrap;
}

.sample-btn {
  background-color: #28a745;
  color: white;
  border: none;
  padding: 10px 20px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
}

.sample-btn:hover {
  background-color: #218838;
}

.download-btn {
  background-color: #17a2b8;
  color: white;
  border: none;
  padding: 10px 20px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
}

.download-btn:hover {
  background-color: #138496;
}

.file-info {
  color: #666;
  font-size: 14px;
  margin: 0;
}

.loading {
  text-align: center;
  color: #007bff;
  font-weight: bold;
  margin: 20px 0;
}

.message {
  padding: 12px;
  border-radius: 4px;
  margin: 20px 0;
  text-align: center;
}

.message.success {
  background-color: #d4edda;
  color: #155724;
  border: 1px solid #c3e6cb;
}

.message.error {
  background-color: #f8d7da;
  color: #721c24;
  border: 1px solid #f5c6cb;
}

.data-section {
  margin-top: 30px;
}

.data-section h2 {
  color: #333;
  margin-bottom: 10px;
}

.table-container {
  overflow-x: auto;
  margin: 20px 0;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.data-table {
  width: 100%;
  border-collapse: collapse;
  min-width: 600px;
}

.data-table th,
.data-table td {
  border: 1px solid #ddd;
  padding: 12px;
  text-align: left;
}

.data-table th {
  background-color: #f8f9fa;
  font-weight: bold;
  position: sticky;
  top: 0;
}

.data-table tr:nth-child(even) {
  background-color: #f8f9fa;
}

.data-table tr:hover {
  background-color: #e9ecef;
}

.json-section {
  margin-top: 30px;
}

.toggle-btn {
  background-color: #6c757d;
  color: white;
  border: none;
  padding: 8px 16px;
  border-radius: 4px;
  cursor: pointer;
  margin-bottom: 15px;
}

.toggle-btn:hover {
  background-color: #545b62;
}

.json-display {
  background-color: #f8f9fa;
  border: 1px solid #ddd;
  border-radius: 4px;
  padding: 15px;
  max-height: 400px;
  overflow-y: auto;
  font-family: 'Courier New', monospace;
  font-size: 12px;
  line-height: 1.4;
}

.worksheet-info {
  margin: 10px 0;
  padding: 8px;
  background-color: #e9ecef;
  border-radius: 4px;
  font-size: 14px;
}

.format-info {
  margin-top: 4px;
  opacity: 0.7;
}

.format-controls {
  margin: 20px 0;
  display: flex;
  gap: 20px;
  flex-wrap: wrap;
}

.format-controls label {
  display: flex;
  align-items: center;
  gap: 5px;
  font-size: 14px;
  cursor: pointer;
}

.format-controls input[type="checkbox"] {
  cursor: pointer;
}

/* 儲存格類型樣式 */
.cell-date {
  color: #007bff;
}

.cell-number {
  color: #28a745;
  text-align: right;
}

.cell-boolean {
  color: #dc3545;
  text-align: center;
}

.cell-empty {
  background-color: #f8f9fa;
  font-style: italic;
}

.cell-text {
  color: #333;
}

@media (max-width: 768px) {
  .container {
    padding: 10px;
  }

  .upload-area {
    padding: 20px;
  }

  .data-table {
    font-size: 14px;
  }

  .data-table th,
  .data-table td {
    padding: 8px;
  }
}
</style>
