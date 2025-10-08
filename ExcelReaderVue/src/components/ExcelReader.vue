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
                <!-- Excel 欄位標頭（包含寬度的物件） -->
                <th v-if="headerType === 'column'" class="column-header" :style="getColumnHeaderStyle(header)">
                  {{ getColumnHeaderName(header) }}
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
                    <!-- 圖片顯示 -->
                    <div v-if="cell.images && cell.images.length > 0" class="cell-images">
                      <div v-for="(image, imageIndex) in cell.images" :key="imageIndex" class="image-container">
                        <!-- 檢查是否為佔位圖片 -->
                        <div v-if="isPlaceholderImage(image)" class="placeholder-image">
                          <div class="placeholder-content">
                            <div class="placeholder-icon">🖼️</div>
                            <div class="placeholder-text">
                              <strong>DISPIMG 圖片</strong><br>
                              <small>{{ image.fileName }}</small><br>
                              <small style="color: #dc3545;">圖片資料無法存取</small><br>
                              <small style="color: #6c757d;">EPPlus 7.1.0 限制</small>
                            </div>
                          </div>
                        </div>
                        <!-- 正常圖片 -->
                        <!-- EMF 格式 (已轉換為 PNG) -->
                        <div v-else-if="image.imageType.toLowerCase() === 'emf'" class="emf-converted-container">
                          <img
                            :src="`data:image/png;base64,${image.base64Data}`"
                            :alt="image.name"
                            :title="`${image.name} - EMF 格式已轉換為 PNG: ${image.width}x${image.height}px, ${formatFileSize(image.fileSize)}`"
                            class="cell-image emf-converted"
                            :style="{
                              width: image.width > 0 ? image.width + 'px' : 'auto',
                              height: image.height > 0 ? image.height + 'px' : 'auto'
                            }"
                            @click="openImageModal(image)"
                          />
                          <div class="emf-badge">EMF→PNG</div>
                        </div>
                        <!-- 一般圖片 -->
                        <img
                          v-else
                          :src="`data:image/${image.imageType.toLowerCase()};base64,${image.base64Data}`"
                          :alt="image.name"
                          :title="`${image.name} - Excel顯示: ${image.width}x${image.height}px, 原始: ${image.originalWidth}x${image.originalHeight}px, ${formatFileSize(image.fileSize)}`"
                          class="cell-image"
                          :style="{
                            width: image.width > 0 ? image.width + 'px' : 'auto',
                            height: image.height > 0 ? image.height + 'px' : 'auto'
                          }"
                          @click="openImageModal(image)"
                          @error="handleImageError"
                        />
                        <div v-if="showImageInfo" class="image-info">
                          <small>{{ image.name }} ({{ image.imageType }})</small>
                          <small>{{ image.width }}x{{ image.height }}</small>
                          <small v-if="isPlaceholderImage(image)" style="color: #dc3545;">佔位圖片</small>
                        </div>
                      </div>
                    </div>
                    <!-- 文字內容 -->
                    <div class="text-content" v-if="!getDisplayValue(cell).includes('#VALUE!')">
                      <span v-if="cell.metadata?.isRichText" v-html="renderRichText(cell)"></span>
                      <span v-else v-html="formatTextWithLineBreaks(getDisplayValue(cell))"></span>
                    </div>
                    <!-- 🆕 浮動物件資訊 -->
                    <div class="floating-objects-info" v-if="showFloatingObjectInfo && cell.floatingObjects && cell.floatingObjects.length > 0">
                      <div v-for="(obj, idx) in cell.floatingObjects" :key="idx" class="floating-object-item">
                        <small class="floating-object-badge">{{ obj.objectType }}</small>
                        <small class="floating-object-name">{{ obj.name }}</small>
                        <div v-if="obj.text" class="floating-object-text">
                          <small>📝 {{ obj.text }}</small>
                        </div>
                        <small class="floating-object-position" v-if="obj.fromCell && obj.toCell">
                          {{ obj.fromCell.address }} → {{ obj.toCell.address }}
                        </small>
                      </div>
                    </div>
                    <!-- 位置資訊 -->
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
        <label>
          <input type="checkbox" v-model="showImageInfo" />
          顯示圖片資訊
        </label>
        <label>
          <input type="checkbox" v-model="showFloatingObjectInfo" />
          顯示浮動物件資訊
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

  <!-- 圖片模態框 -->
  <div v-if="showImageModal && selectedImage" class="image-modal" @click="closeImageModal">
    <div class="modal-content" @click.stop>
      <div class="modal-header">
        <h3>{{ selectedImage.name }}</h3>
        <button @click="closeImageModal" class="close-btn">×</button>
      </div>
      <div class="modal-body">
        <!-- EMF 格式 (已轉換) -->
        <div v-if="selectedImage.imageType.toLowerCase() === 'emf'">
          <img
            :src="`data:image/png;base64,${selectedImage.base64Data}`"
            :alt="selectedImage.name"
            class="modal-image emf-converted-modal"
          />
          <div class="emf-modal-info">
            <div class="emf-info-badge">✅ EMF 格式已自動轉換為 PNG</div>
            <p>原始格式：Enhanced Metafile (.emf) - Windows 向量圖形格式</p>
            <p>為了在瀏覽器中正常顯示，系統已自動將此圖片轉換為 PNG 格式</p>
          </div>
        </div>
        <!-- 一般圖片 -->
        <img
          v-else
          :src="`data:image/${selectedImage.imageType.toLowerCase()};base64,${selectedImage.base64Data}`"
          :alt="selectedImage.name"
          class="modal-image"
        />
        <div class="image-details">
          <p><strong>類型:</strong> {{ selectedImage.imageType }}</p>
          <p><strong>尺寸:</strong> {{ selectedImage.width }} x {{ selectedImage.height }}</p>
          <p><strong>檔案大小:</strong> {{ formatFileSize(selectedImage.fileSize) }}</p>
          <p v-if="selectedImage.description"><strong>描述:</strong> {{ selectedImage.description }}</p>
          <p v-if="selectedImage.anchorCell"><strong>錨點儲存格:</strong> {{ selectedImage.anchorCell.address }}</p>
          <p v-if="selectedImage.hyperlinkAddress"><strong>超連結:</strong> <a :href="selectedImage.hyperlinkAddress" target="_blank">{{ selectedImage.hyperlinkAddress }}</a></p>
        </div>
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
  RichTextPart,
  ImageInfo
} from '@/types'

// 欄位標頭類型定義
interface ColumnHeader {
  name: string;
  width: number;
  index: number;
}



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
const showImageInfo = ref<boolean>(false)
const showFloatingObjectInfo = ref<boolean>(false) // 🆕 顯示浮動物件資訊
const headerType = ref<'column' | 'content'>('column') // 默認顯示 Excel 欄位標頭
const selectedImage = ref<ImageInfo | null>(null)
const showImageModal = ref<boolean>(false)

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

// 格式化文件大小
const formatFileSize = (bytes: number): string => {
  if (bytes === 0) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
}

// 打開圖片模態框
const openImageModal = (image: ImageInfo) => {
  selectedImage.value = image
  showImageModal.value = true
}

// 關閉圖片模態框
const closeImageModal = () => {
  selectedImage.value = null
  showImageModal.value = false
}

// 檢查是否為佔位圖片
const isPlaceholderImage = (image: ImageInfo): boolean => {
  // 檢查檔案名稱是否包含 dispimg
  if (image.fileName && image.fileName.toLowerCase().includes('dispimg')) {
    return true
  }

  // 檢查 Base64 資料是否為預設的佔位圖片
  const placeholderBase64 = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAGXRFWHRDb21tZW50AEltYWdlIG5vdCBmb3VuZMk4KcsAAAA+SURBVFiF7dAxAQAACAOg9VPgAAIAAEAAABAAAAQAAAEAAABAAAAQAAAEAAABAAAAQAAAEAAABAAAAQAAAECKDYwIAAAAAElFTkSuQmCC'
  if (image.base64Data === placeholderBase64) {
    return true
  }

  // 檢查檔案大小是否為 0 或 hyperlink 包含 DISPIMG
  if (image.fileSize === 0 || (image.hyperlinkAddress && image.hyperlinkAddress.includes('DISPIMG'))) {
    return true
  }

  return false
}

// 處理圖片載入錯誤
const handleImageError = (event: Event) => {
  const img = event.target as HTMLImageElement
  console.warn('圖片載入失敗:', img.src)
  img.style.display = 'none'
}

// 獲取儲存格的邊框樣式
const getCellBorderStyle = (cell: ExcelCellInfo): Record<string, string> => {
  const borderStyles: Record<string, string> = {}

  if (cell.border?.top?.style && cell.border.top.style !== 'None') {
    const color = cell.border.top.color ? `#${cell.border.top.color}` : '#000000'
    borderStyles.borderTop = `${convertBorderStyle(cell.border.top.style)} ${color} !important`
  }

  if (cell.border?.bottom?.style && cell.border.bottom.style !== 'None') {
    const color = cell.border.bottom.color ? `#${cell.border.bottom.color}` : '#000000'
    borderStyles.borderBottom = `${convertBorderStyle(cell.border.bottom.style)} ${color} !important`
  }

  if (cell.border?.left?.style && cell.border.left.style !== 'None') {
    const color = cell.border.left.color ? `#${cell.border.left.color}` : '#000000'
    borderStyles.borderLeft = `${convertBorderStyle(cell.border.left.style)} ${color} !important`
  }

  if (cell.border?.right?.style && cell.border.right.style !== 'None') {
    const color = cell.border.right.color ? `#${cell.border.right.color}` : '#000000'
    borderStyles.borderRight = `${convertBorderStyle(cell.border.right.style)} ${color} !important`
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

  // 邊框樣式 - 總是套用 Excel 的邊框設定
  const borderStyles = getCellBorderStyle(header)
  if (Object.keys(borderStyles).length > 0) {
    Object.assign(style, borderStyles)
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

  // 邊框樣式 - 總是套用 Excel 的邊框設定
  const borderStyles = getCellBorderStyle(cell)
  if (Object.keys(borderStyles).length > 0) {
    Object.assign(style, borderStyles)
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

// 獲取欄位標頭名稱（處理新的物件格式）
const getColumnHeaderName = (header: unknown): string => {
  // 如果是新的物件格式（包含 name, width, index）
  if (typeof header === 'object' && header !== null && 'name' in header) {
    return (header as ColumnHeader).name
  }

  // 如果是舊的字串格式
  if (typeof header === 'string') {
    return header
  }

  return ''
}

// 獲取欄位標頭樣式（包含寬度）
const getColumnHeaderStyle = (header: unknown): Record<string, string> => {
  const style: Record<string, string> = {}

  // 如果是新的物件格式且有寬度資訊
  if (typeof header === 'object' && header !== null && 'width' in header) {
    const columnHeader = header as ColumnHeader
    style.width = `${convertExcelWidthToPixels(columnHeader.width)}px`
  }

  return style
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
  /*excel thead已有固定寬度*/
  width: 0;
  border-collapse: collapse;
  min-width: 600px;
  table-layout: fixed ;
  margin: auto;
}

.data-table th,
.data-table td {
  /* 只設定默認邊框，如果沒有動態邊框的話 */
  border: 1px solid #ddd;
  padding: 2px;
  text-align: left;
  white-space: nowrap;
  /* 強制使用設定的高度，避免合併儲存格影響其他行的高度 */
  box-sizing: border-box;
  overflow: hidden;
}

/* 針對合併儲存格的特殊處理 */
.data-table td[rowspan] {
  /* 合併儲存格使用 top 對齊，避免影響其他儲存格 */
  vertical-align: top !important;
}

/* 確保沒有合併的儲存格能維持設定的高度 */
.data-table td:not([rowspan]) {
  /* 對於非合併儲存格，使用行內設定的高度 */
  height: auto;
  min-height: inherit;
}

/* 當有動態邊框時，讓動態邊框優先 */
/* .data-table td[style*="border"] 讓行內樣式生效 */

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
  display: inline-block;
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

/* 圖片顯示樣式 */
.cell-images {
  margin-bottom: 4px;
}

.image-container {
  display: inline-block;
  margin: 2px;
  text-align: center;
  width:100%;
}

.cell-image {
  cursor: pointer;
  border: 1px solid #ddd;
  border-radius: 4px;
  transition: transform 0.2s, box-shadow 0.2s;
}

.cell-image:hover {
  transform: scale(1.05);
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
}

.image-info {
  font-size: 10px;
  color: #666;
  margin-top: 2px;
}

.image-info small {
  display: block;
  line-height: 1.2;
}

.text-content {
  margin-top: 4px;
}

/* 🆕 浮動物件資訊樣式 */
.floating-objects-info {
  margin-top: 8px;
  padding: 6px;
  background-color: #f8f9fa;
  border-left: 3px solid #007bff;
  border-radius: 4px;
}

.floating-object-item {
  padding: 4px 0;
  border-bottom: 1px dashed #dee2e6;
}

.floating-object-item:last-child {
  border-bottom: none;
}

.floating-object-badge {
  display: inline-block;
  padding: 2px 6px;
  background-color: #007bff;
  color: white;
  border-radius: 3px;
  font-size: 10px;
  font-weight: bold;
  margin-right: 4px;
}

.floating-object-name {
  color: #495057;
  font-size: 11px;
  font-weight: 500;
}

.floating-object-text {
  margin: 4px 0;
  padding: 4px 8px;
  background-color: #fff;
  border-radius: 3px;
  border: 1px solid #dee2e6;
}

.floating-object-text small {
  color: #212529;
  font-size: 11px;
  line-height: 1.4;
  white-space: pre-wrap;
}

.floating-object-position {
  color: #6c757d;
  font-size: 10px;
  font-style: italic;
}

/* 佔位圖片樣式 */
.placeholder-image {
  display: inline-flex;
  align-items: center;
  padding: 8px 12px;
  border: 2px dashed #dc3545;
  border-radius: 8px;
  background-color: #f8f9fa;
  margin: 2px;
  max-width: 200px;
  cursor: pointer;
  transition: background-color 0.3s ease;
}

.placeholder-image:hover {
  background-color: #e9ecef;
}

.placeholder-content {
  display: flex;
  align-items: center;
  gap: 8px;
}

.placeholder-icon {
  font-size: 24px;
  color: #dc3545;
}

.placeholder-text {
  font-size: 12px;
  line-height: 1.3;
}

.placeholder-text strong {
  color: #495057;
  font-size: 13px;
}

/* 圖片模態框樣式 */
.image-modal {
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: rgba(0, 0, 0, 0.8);
  display: flex;
  justify-content: center;
  align-items: center;
  z-index: 1000;
}

.modal-content {
  background: white;
  border-radius: 8px;
  max-width: 90%;
  max-height: 90%;
  overflow: auto;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
}

.modal-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 16px 20px;
  border-bottom: 1px solid #eee;
}

.modal-header h3 {
  margin: 0;
  color: #333;
}

.close-btn {
  background: none;
  border: none;
  font-size: 24px;
  cursor: pointer;
  color: #666;
  padding: 0;
  width: 30px;
  height: 30px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.close-btn:hover {
  color: #000;
}

.modal-body {
  padding: 20px;
  text-align: center;
}

.modal-image {
  max-width: 100%;
  max-height: 60vh;
  border: 1px solid #ddd;
  border-radius: 4px;
}

.image-details {
  margin-top: 16px;
  text-align: left;
  background: #f8f9fa;
  padding: 12px;
  border-radius: 4px;
}

.image-details p {
  margin: 4px 0;
  font-size: 14px;
}

.image-details strong {
  color: #333;
}

.image-details a {
  color: #007bff;
  text-decoration: none;
}

.image-details a:hover {
  text-decoration: underline;
}

/* EMF 格式樣式 - 已轉換為 PNG */
.emf-converted-container {
  position: relative;
  display: inline-block;
}

.emf-converted {
  border: 2px solid #28a745;
  border-radius: 4px;
  box-shadow: 0 2px 4px rgba(40, 167, 69, 0.1);
}

.emf-badge {
  position: absolute;
  top: -8px;
  right: -8px;
  background: #28a745;
  color: white;
  font-size: 10px;
  font-weight: bold;
  padding: 2px 6px;
  border-radius: 10px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.2);
  z-index: 10;
}

/* EMF 格式樣式 - 舊版 (保留以防需要) */
.emf-placeholder {
  display: flex;
  align-items: center;
  padding: 8px;
  border: 2px dashed #ffc107;
  background: #fff3cd;
  border-radius: 4px;
  cursor: pointer;
  transition: background-color 0.2s;
  min-height: 60px;
  max-width: 200px;
}

.emf-placeholder:hover {
  background: #fff3a0;
}

.emf-icon {
  font-size: 24px;
  margin-right: 8px;
}

.emf-text {
  text-align: left;
}

.emf-text div:first-child {
  font-weight: bold;
  color: #856404;
}

.emf-note {
  font-size: 11px;
  color: #856404;
  opacity: 0.8;
}

/* EMF 模態框樣式 - 新版轉換後 */
.emf-converted-modal {
  border: 3px solid #28a745;
  border-radius: 8px;
  box-shadow: 0 4px 8px rgba(40, 167, 69, 0.2);
}

.emf-modal-info {
  margin-top: 16px;
  padding: 16px;
  background: #d4edda;
  border: 1px solid #c3e6cb;
  border-radius: 8px;
  text-align: left;
}

.emf-info-badge {
  display: inline-block;
  background: #28a745;
  color: white;
  font-weight: bold;
  padding: 6px 12px;
  border-radius: 20px;
  font-size: 14px;
  margin-bottom: 12px;
}

.emf-modal-info p {
  color: #155724;
  margin-bottom: 8px;
  line-height: 1.5;
}

/* EMF 模態框樣式 - 舊版 (保留) */
.emf-modal-placeholder {
  text-align: center;
  padding: 40px 20px;
  background: #fff3cd;
  border: 2px dashed #ffc107;
  border-radius: 8px;
  max-width: 500px;
  margin: 0 auto;
}

.emf-modal-icon {
  font-size: 64px;
  margin-bottom: 16px;
}

.emf-modal-content h4 {
  color: #856404;
  margin-bottom: 12px;
}

.emf-modal-content p {
  color: #856404;
  margin-bottom: 8px;
}

.emf-warning {
  background: #f8d7da;
  color: #721c24 !important;
  padding: 8px;
  border-radius: 4px;
  border: 1px solid #f5c6cb;
  margin: 12px 0 !important;
}

.emf-suggestions {
  text-align: left;
  background: white;
  padding: 16px;
  border-radius: 4px;
  margin-top: 16px;
  border: 1px solid #ffc107;
}

.emf-suggestions ul {
  margin: 8px 0;
  padding-left: 20px;
}

.emf-suggestions li {
  margin: 4px 0;
  color: #495057;
}
</style>
