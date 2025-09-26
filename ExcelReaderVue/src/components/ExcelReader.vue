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

      <!-- 標頭類型選擇 -->
      <div class="header-type-controls">
        <label class="header-type-label">
          標頭類型：
          <select v-model="headerType" @change="onHeaderTypeChange" class="header-type-select">
            <option value="column">Excel 欄位標頭 (A, B, C, D...)</option>
            <option value="content">工作表內容標頭 (第一行內容)</option>
          </select>
        </label>
      </div>

      <div class="table-container">
        <table class="data-table">
          <thead>
            <tr>
              <template v-for="(header, index) in getCurrentHeaders()" :key="index">
                <!-- Excel 欄位標頭（簡單字串） -->
                <th v-if="headerType === 'column'" class="column-header">
                  {{ header }}
                </th>
                <!-- 工作表內容標頭（ExcelCellInfo 物件） -->
                <th
                  v-else-if="headerType === 'content' && shouldRenderCell(header as ExcelCellInfo)"
                  :style="getHeaderStyle(header as ExcelCellInfo)"
                  :colspan="(header as ExcelCellInfo).dimensions?.colSpan || 1"
                  :rowspan="(header as ExcelCellInfo).dimensions?.rowSpan || 1"
                >
                  <span v-if="(header as ExcelCellInfo).metadata?.isRichText" v-html="renderRichText(header as ExcelCellInfo)"></span>
                  <span v-else v-html="formatTextWithLineBreaks(getDisplayValue(header as ExcelCellInfo))"></span>
                  <div class="format-info" v-if="showFormatInfo">
                    <small>格式: {{ (header as ExcelCellInfo).numberFormat || '一般' }}</small>
                    <small v-if="(header as ExcelCellInfo).metadata?.isRichText" style="color: orange;">Rich Text</small>
                  </div>
                  <div class="position-info" v-if="showPositionInfo">
                    <small>位置: {{ (header as ExcelCellInfo).position?.address || '未知' }}</small>
                    <small v-if="(header as ExcelCellInfo).formula">公式: {{ (header as ExcelCellInfo).formula }}</small>
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
                  :colspan="cell.dimensions?.colSpan || 1"
                  :rowspan="cell.dimensions?.rowSpan || 1"
                >
                  <div class="cell-content">
                    <span v-if="cell.metadata?.isRichText" v-html="renderRichText(cell)"></span>
                    <span v-else v-html="formatTextWithLineBreaks(getDisplayValue(cell))"></span>
                    <div class="position-info" v-if="showPositionInfo && (cell.position?.address || cell.formula)">
                      <small v-if="cell.position?.address">{{ cell.position.address }}</small>
                      <small v-if="cell.formula" style="color: green;">{{ cell.formula }}</small>
                    </div>
                  </div>
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
        <label>
          <input type="checkbox" v-model="showAdvancedFormatting" />
          顯示進階格式 (邊框、對齊等)
        </label>
        <label>
          <input type="checkbox" v-model="showPositionInfo" />
          顯示位置資訊
        </label>
      </div>

      <div class="json-section">
        <h3>JSON 資料：</h3>
        <div class="json-controls">
          <button @click="toggleJsonView" class="toggle-btn">
            {{ showJson ? '隱藏' : '顯示' }} JSON
          </button>
          <button @click="downloadJson" class="download-json-btn" :disabled="!excelData">
            下載 JSON
          </button>
        </div>
        <pre v-if="showJson" class="json-display">{{ JSON.stringify(excelData, null, 2) }}</pre>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue'
import axios from 'axios'
import type {
  ExcelCellInfo,
  ExcelData,
  UploadResponse,
  RichTextPart
} from '@/types'



const loading = ref<boolean>(false)
const message = ref<string>('')
const messageType = ref<'success' | 'error' | ''>('')
const excelData = ref<ExcelData | null>(null)
const showJson = ref<boolean>(false)
const fileInput = ref<HTMLInputElement | null>(null)
const showFormatInfo = ref<boolean>(false)
const showOriginalValue = ref<boolean>(false)
const showAdvancedFormatting = ref<boolean>(false)
const showPositionInfo = ref<boolean>(false)
const headerType = ref<'column' | 'content'>('column') // 默認顯示 Excel 欄位標頭

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

const downloadJson = () => {
  if (!excelData.value) {
    message.value = '沒有可下載的資料'
    messageType.value = 'error'
    clearMessage()
    return
  }

  try {
    // 創建JSON字符串
    const jsonString = JSON.stringify(excelData.value, null, 2)
    const blob = new Blob([jsonString], { type: 'application/json' })
    const url = window.URL.createObjectURL(blob)
    
    // 創建下載連結
    const link = document.createElement('a')
    link.href = url
    
    // 生成檔案名稱，使用Excel檔案名稱作為基礎
    const fileName = excelData.value.fileName ? 
      `${excelData.value.fileName.replace(/\.[^/.]+$/, '')}.json` : 
      'excel-data.json'
    
    link.setAttribute('download', fileName)
    document.body.appendChild(link)
    link.click()
    link.remove()
    window.URL.revokeObjectURL(url)

    message.value = 'JSON檔案已下載'
    messageType.value = 'success'
    clearMessage()
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : '未知錯誤'
    message.value = `下載失敗：${errorMessage}`
    messageType.value = 'error'
    clearMessage()
  }
}

const onHeaderTypeChange = () => {
  // 當標頭類型改變時，可以在這裡添加額外的邏輯
  // 例如：重新渲染表格或顯示通知
}

const getCurrentHeaders = () => {
  if (!excelData.value || !excelData.value.headers) return []
  
  if (headerType.value === 'column') {
    // 返回 Excel 欄位標頭 (A, B, C...)
    return excelData.value.headers[0] || []
  } else {
    // 返回工作表內容標頭（第一行內容）
    return excelData.value.headers[1] || []
  }
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
  return cell.text || ''
}

// 新增：渲染Rich Text的HTML
const renderRichText = (cell: ExcelCellInfo): string => {
  if (!cell.metadata?.isRichText || !cell.richText) {
    // 處理一般文字的換行
    return formatTextWithLineBreaks(cell.text || '')
  }

  return cell.richText.map((part: RichTextPart) => {
    // HTML轉義文字內容以防止XSS，並處理換行
    let html = formatTextWithLineBreaks(escapeHtml(part.text))
    const styles: string[] = []

    if (part.bold) styles.push('font-weight: bold')
    if (part.italic) styles.push('font-style: italic')
    if (part.underLine) styles.push('text-decoration: underline')
    if (part.size && part.size > 0) styles.push(`font-size: ${part.size}pt`)
    if (part.fontName && part.fontName.trim()) styles.push(`font-family: ${part.fontName}`)
    if (part.color) styles.push(`color: ${part.color}`)

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

// 將Excel邊框樣式轉換為CSS邊框樣式
const convertBorderStyle = (excelStyle?: string): string => {
  if (!excelStyle || excelStyle === 'None') return 'none'

  const styleMap: Record<string, string> = {
    'Thin': '1px solid',
    'Thick': '3px solid',
    'Medium': '2px solid',
    'Dotted': '1px dotted',
    'Dashed': '1px dashed',
    'Double': '3px double',
    'Hair': '1px solid'
  }

  return styleMap[excelStyle] || '1px solid'
}

// 獲取儲存格的邊框樣式
const getCellBorderStyle = (cell: ExcelCellInfo): Record<string, string> => {
  const borderStyles: Record<string, string> = {}

  if (cell.border?.top?.style && cell.border.top.style !== 'None') {
    const color = cell.border.top.color ? `#${cell.border.top.color}` : '#000000'
    borderStyles.borderTop = `${convertBorderStyle(cell.border.top.style)} ${color}`
  }

  if (cell.border?.bottom?.style && cell.border.bottom.style !== 'None') {
    const color = cell.border.bottom.color ? `#${cell.border.bottom.color}` : '#000000'
    borderStyles.borderBottom = `${convertBorderStyle(cell.border.bottom.style)} ${color}`
  }

  if (cell.border?.left?.style && cell.border.left.style !== 'None') {
    const color = cell.border.left.color ? `#${cell.border.left.color}` : '#000000'
    borderStyles.borderLeft = `${convertBorderStyle(cell.border.left.style)} ${color}`
  }

  if (cell.border?.right?.style && cell.border.right.style !== 'None') {
    const color = cell.border.right.color ? `#${cell.border.right.color}` : '#000000'
    borderStyles.borderRight = `${convertBorderStyle(cell.border.right.style)} ${color}`
  }

  return borderStyles
}

const getHeaderStyle = (header: ExcelCellInfo) => {
  const style: Record<string, string> = {}

  // 字體樣式
  if (header.font?.bold) {
    style.fontWeight = 'bold'
  }

  if (header.font?.italic) {
    style.fontStyle = 'italic'
  }

  if (header.font?.size) {
    style.fontSize = `${header.font.size}px`
  }

  if (header.font?.name) {
    style.fontFamily = `"${header.font.name}"`
  }

  if (header.font?.strike) {
    style.textDecoration = 'line-through'
  }

  // 顏色樣式
  if (header.fill?.backgroundColor) {
    style.backgroundColor = `#${header.fill.backgroundColor}`
  }

  if (header.font?.color) {
    style.color = `#${header.font.color}`
  }

  // 對齊樣式
  if (header.alignment?.horizontal) {
    style.textAlign = header.alignment.horizontal.toLowerCase()
  }

  if (header.alignment?.vertical) {
    style.verticalAlign = header.alignment.vertical.toLowerCase()
  }

  if (header.alignment?.wrapText) {
    style.whiteSpace = 'pre-wrap'
  }

  // 尺寸
  if (header.dimensions?.columnWidth) {
    style.width = `${convertExcelWidthToPixels(header.dimensions.columnWidth)}px`
  }

  if (header.dimensions?.rowHeight) {
    style.height = `${header.dimensions.rowHeight}px`
  }

  // 邊框樣式 (僅在顯示進階格式時應用)
  if (showAdvancedFormatting.value) {
    Object.assign(style, getCellBorderStyle(header))
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

  // 字體樣式
  if (cell.font?.bold) {
    style.fontWeight = 'bold'
  }

  if (cell.font?.italic) {
    style.fontStyle = 'italic'
  }

  if (cell.font?.size) {
    style.fontSize = `${cell.font.size}px`
  }

  if (cell.font?.name) {
    style.fontFamily = `"${cell.font.name}"`
  }

  if (cell.font?.strike) {
    style.textDecoration = 'line-through'
  }

  // 顏色樣式
  if (cell.fill?.backgroundColor) {
    style.backgroundColor = `#${cell.fill.backgroundColor}`
  }

  if (cell.font?.color) {
    style.color = `#${cell.font.color}`
  }

  // 對齊樣式
  if (cell.alignment?.horizontal) {
    style.textAlign = cell.alignment.horizontal.toLowerCase()
  }

  if (cell.alignment?.vertical) {
    style.verticalAlign = cell.alignment.vertical.toLowerCase()
  }

  if (cell.alignment?.wrapText) {
    style.whiteSpace = 'pre-wrap'
  }

  // 尺寸
  if (cell.dimensions?.columnWidth) {
    style.width = `${convertExcelWidthToPixels(cell.dimensions.columnWidth)}px`
  }

  if (cell.dimensions?.rowHeight) {
    style.height = `${cell.dimensions.rowHeight}px`
  }

  // 邊框樣式 (僅在顯示進階格式時應用)
  if (showAdvancedFormatting.value) {
    Object.assign(style, getCellBorderStyle(cell))
  }

  return style
}

const getCellTooltip = (cell: ExcelCellInfo): string => {
  const parts = []

  // 基本資訊
  parts.push(`位置: ${cell.position?.address || '未知'}`)
  parts.push(`類型: ${cell.dataType}`)
  parts.push(`值類型: ${cell.valueType || '未知'}`)

  // 格式資訊
  if (cell.numberFormat) {
    parts.push(`數字格式: ${cell.numberFormat}`)
  }

  if (cell.numberFormatId) {
    parts.push(`格式ID: ${cell.numberFormatId}`)
  }

  // 值資訊
  if (cell.value !== null && cell.value !== undefined) {
    parts.push(`原始值: ${cell.value}`)
  }

  if (cell.text) {
    parts.push(`顯示文字: ${cell.text}`)
  }

  if (cell.formula) {
    parts.push(`公式: ${cell.formula}`)
  }

  // 字體資訊
  if (cell.font?.name || cell.font?.size) {
    const fontInfo = []
    if (cell.font.name) fontInfo.push(`字體: ${cell.font.name}`)
    if (cell.font.size) fontInfo.push(`大小: ${cell.font.size}pt`)
    if (cell.font.bold) fontInfo.push('粗體')
    if (cell.font.italic) fontInfo.push('斜體')
    if (fontInfo.length > 0) parts.push(fontInfo.join(', '))
  }

  // 對齊資訊
  if (cell.alignment?.horizontal || cell.alignment?.vertical) {
    const alignInfo = []
    if (cell.alignment.horizontal) alignInfo.push(`水平: ${cell.alignment.horizontal}`)
    if (cell.alignment.vertical) alignInfo.push(`垂直: ${cell.alignment.vertical}`)
    if (cell.alignment.wrapText) alignInfo.push('自動換行')
    if (alignInfo.length > 0) parts.push(`對齊: ${alignInfo.join(', ')}`)
  }

  // Rich Text 資訊
  if (cell.metadata?.isRichText && cell.richText) {
    parts.push(`Rich Text 片段數: ${cell.richText.length}`)
  }

  // 合併儲存格資訊
  if (cell.dimensions?.isMerged && cell.dimensions?.rowSpan && cell.dimensions?.colSpan) {
    parts.push(`合併儲存格: ${cell.dimensions.rowSpan}行 x ${cell.dimensions.colSpan}欄`)
  }

  // 尺寸資訊
  if (cell.dimensions?.columnWidth || cell.dimensions?.rowHeight) {
    const sizeInfo = []
    if (cell.dimensions.columnWidth) sizeInfo.push(`欄寬: ${cell.dimensions.columnWidth.toFixed(2)}`)
    if (cell.dimensions.rowHeight) sizeInfo.push(`行高: ${cell.dimensions.rowHeight.toFixed(2)}`)
    if (sizeInfo.length > 0) parts.push(`尺寸: ${sizeInfo.join(', ')}`)
  }

  // 註解資訊
  if (cell.comment) {
    parts.push(`註解: ${cell.comment.text || '無內容'}`)
    if (cell.comment.author) parts.push(`註解作者: ${cell.comment.author}`)
  }

  // 超連結資訊
  if (cell.hyperlink) {
    parts.push(`超連結: ${cell.hyperlink.originalString || cell.hyperlink.absoluteUri || '無連結'}`)
  }

  // 樣式資訊
  if (cell.metadata?.styleId || cell.metadata?.styleName) {
    const styleInfo = []
    if (cell.metadata.styleId) styleInfo.push(`ID: ${cell.metadata.styleId}`)
    if (cell.metadata.styleName) styleInfo.push(`名稱: ${cell.metadata.styleName}`)
    if (styleInfo.length > 0) parts.push(`樣式: ${styleInfo.join(', ')}`)
  }

  return parts.join('\n')
}

const shouldRenderCell = (cell: ExcelCellInfo): boolean => {
  // 如果不是合併儲存格，正常顯示
  if (!cell.dimensions?.isMerged) {
    return true
  }

  // 如果是合併儲存格，只顯示主儲存格
  return cell.dimensions?.isMainMergedCell === true
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
  padding: 2px;
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

.json-controls {
  display: flex;
  gap: 10px;
  margin-bottom: 15px;
  flex-wrap: wrap;
}

.download-json-btn {
  background-color: #28a745;
  color: white;
  border: none;
  padding: 8px 16px;
  border-radius: 4px;
  cursor: pointer;
}

.download-json-btn:hover:not(:disabled) {
  background-color: #218838;
}

.download-json-btn:disabled {
  background-color: #6c757d;
  cursor: not-allowed;
  opacity: 0.6;
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

.position-info {
  margin-top: 2px;
  opacity: 0.6;
  font-size: 10px;
}

.position-info small {
  display: block;
  color: #666;
}

.cell-content {
  position: relative;
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

/* 標頭類型控制 */
.header-type-controls {
  margin: 15px 0;
  padding: 10px;
  background-color: #f8f9fa;
  border-radius: 5px;
  border-left: 4px solid #007bff;
}

.header-type-label {
  display: flex;
  align-items: center;
  gap: 10px;
  font-weight: 500;
  color: #333;
}

.header-type-select {
  padding: 5px 10px;
  border: 1px solid #ddd;
  border-radius: 4px;
  background-color: white;
  font-size: 14px;
  cursor: pointer;
}

.header-type-select:focus {
  outline: none;
  border-color: #007bff;
  box-shadow: 0 0 0 2px rgba(0, 123, 255, 0.25);
}

/* Excel 欄位標頭樣式 */
.column-header {
  background-color: #007bff !important;
  color: white !important;
  text-align: center !important;
  font-weight: bold !important;
  font-size: 14px !important;
  min-width: 40px;
}
</style>
