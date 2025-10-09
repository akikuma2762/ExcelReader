# ExcelReaderVue - API 整合文檔

**版本:** 2.0.0  
**最後更新:** 2025年10月9日

---

## 📋 目錄

- [API 概述](#api-概述)
- [配置說明](#配置說明)
- [API 端點](#api-端點)
- [資料模型](#資料模型)
- [錯誤處理](#錯誤處理)
- [使用範例](#使用範例)
- [最佳實踐](#最佳實踐)

---

## API 概述

ExcelReaderVue 通過 HTTP API 與 ExcelReaderAPI 後端服務通訊。本文檔詳細說明前端如何整合和使用這些 API。

### API 架構

```
ExcelReaderVue (前端)
        │
        │ HTTP/HTTPS
        ▼
ExcelReaderAPI (後端)
        │
        ▼
    EPPlus 8.1.0
        │
        ▼
    Excel 檔案
```

### 通訊協定

- **協定**: HTTP/HTTPS
- **格式**: JSON (響應), multipart/form-data (檔案上傳)
- **編碼**: UTF-8
- **HTTP Library**: Axios 1.12.2

---

## 配置說明

### API 基礎 URL

在 `ExcelReader.vue` 中配置:

```typescript
// 開發環境
const API_BASE_URL = 'http://localhost:5000'

// 生產環境
const API_BASE_URL = import.meta.env.VITE_API_URL || 'https://api.yourdomain.com'
```

### 環境變數

建立 `.env` 檔案:

```bash
# .env.development
VITE_API_URL=http://localhost:5000

# .env.production
VITE_API_URL=https://api.yourdomain.com
```

使用環境變數:

```typescript
const API_BASE_URL = import.meta.env.VITE_API_URL
```

### Axios 配置

建立 Axios 實例:

```typescript
import axios from 'axios'

const apiClient = axios.create({
  baseURL: API_BASE_URL,
  timeout: 300000, // 5 分鐘
  headers: {
    'Accept': 'application/json'
  }
})

// 請求攔截器
apiClient.interceptors.request.use(
  config => {
    console.log('Request:', config.method?.toUpperCase(), config.url)
    return config
  },
  error => {
    return Promise.reject(error)
  }
)

// 響應攔截器
apiClient.interceptors.response.use(
  response => {
    console.log('Response:', response.status, response.config.url)
    return response
  },
  error => {
    console.error('API Error:', error.message)
    return Promise.reject(error)
  }
)
```

---

## API 端點

### 1. 上傳 Excel 檔案

**端點:** `POST /api/excel/upload`

**用途:** 上傳並解析 Excel 檔案

#### 請求

```typescript
const uploadFile = async (file: File) => {
  const formData = new FormData()
  formData.append('file', file)
  
  const response = await axios.post(
    `${API_BASE_URL}/api/excel/upload`,
    formData,
    {
      headers: {
        'Content-Type': 'multipart/form-data'
      },
      onUploadProgress: (progressEvent) => {
        if (progressEvent.total) {
          const percent = Math.round(
            (progressEvent.loaded * 100) / progressEvent.total
          )
          console.log(`上傳進度: ${percent}%`)
        }
      }
    }
  )
  
  return response.data
}
```

#### 響應 (成功)

```json
{
  "success": true,
  "message": "檔案上傳成功",
  "data": {
    "fileName": "test.xlsx",
    "fileSize": 123456,
    "worksheets": [
      {
        "name": "Sheet1",
        "index": 0,
        "rowCount": 100,
        "columnCount": 10,
        "cells": [...],
        "mergedCells": ["A1:B2"]
      }
    ],
    "totalWorksheets": 1,
    "processingTime": "1.234s"
  }
}
```

#### 響應 (錯誤)

```json
{
  "success": false,
  "message": "檔案格式不正確或檔案損壞",
  "error": {
    "code": "INVALID_FILE_FORMAT",
    "details": "The uploaded file is not a valid Excel file"
  }
}
```

---

### 2. 載入範例資料

**端點:** `GET /api/excel/sample`

**用途:** 獲取範例 Excel 資料

#### 請求

```typescript
const loadSampleData = async () => {
  const response = await axios.get(
    `${API_BASE_URL}/api/excel/sample`
  )
  
  return response.data
}
```

#### 響應

```json
{
  "fileName": "sample.xlsx",
  "fileSize": 0,
  "worksheets": [
    {
      "name": "Sample Sheet",
      "index": 0,
      "rowCount": 3,
      "columnCount": 3,
      "cells": [...]
    }
  ],
  "totalWorksheets": 1
}
```

---

## 資料模型

### TypeScript 介面定義

完整的型別定義位於 `src/types/excel.ts`:

```typescript
export interface ExcelData {
  fileName: string
  fileSize: number
  worksheets: Worksheet[]
  totalWorksheets: number
  processingTime?: string
}

export interface Worksheet {
  name: string
  index: number
  rowCount: number
  columnCount: number
  cells: ExcelCellInfo[]
  mergedCells: string[]
  metadata?: WorksheetMetadata
}

export interface ExcelCellInfo {
  position: CellPosition
  value: any
  text: string
  formula?: string
  dataType: string
  font: FontInfo
  alignment: AlignmentInfo
  border: BorderInfo
  fill: FillInfo
  dimensions: DimensionInfo
  images?: ImageInfo[]
  floatingObjects?: FloatingObjectInfo[]
  richText?: RichTextPart[]
  comment?: CommentInfo
  hyperlink?: HyperlinkInfo
  metadata?: CellMetadata
}

export interface CellPosition {
  row: number
  column: number
  address: string
}

export interface FontInfo {
  name: string
  size: number
  bold: boolean
  italic: boolean
  underLine: string
  strike: boolean
  color?: string
  colorTheme?: string
  colorTint?: number
}

export interface AlignmentInfo {
  horizontal: string
  vertical: string
  wrapText: boolean
  indent: number
  readingOrder: string
  textRotation: number
  shrinkToFit: boolean
}

export interface BorderInfo {
  top: BorderStyle
  bottom: BorderStyle
  left: BorderStyle
  right: BorderStyle
  diagonal: BorderStyle
  diagonalUp: boolean
  diagonalDown: boolean
}

export interface BorderStyle {
  style: string
  color?: string
}

export interface FillInfo {
  patternType: string
  backgroundColor?: string
  patternColor?: string
  backgroundColorTheme?: string
  backgroundColorTint?: number
}

export interface DimensionInfo {
  columnWidth: number
  rowHeight: number
  isMerged: boolean
  isMainMergedCell?: boolean
  rowSpan?: number
  colSpan?: number
  mergedRangeAddress?: string
}

export interface ImageInfo {
  name: string
  description?: string
  imageType: string
  width: number
  height: number
  originalWidth?: number
  originalHeight?: number
  left: number
  top: number
  base64Data: string
  fileName?: string
  fileSize: number
  anchorCell: CellPosition
  hyperlinkAddress?: string
  isInCellPicture?: boolean
  altText?: string
}

export interface FloatingObjectInfo {
  name: string
  description?: string
  objectType: string
  width: number
  height: number
  left: number
  top: number
  text?: string
  anchorCell: CellPosition
  fromCell: CellPosition
  toCell: CellPosition
  isFloating: boolean
  style?: string
  hyperlinkAddress?: string
}

export interface RichTextPart {
  text: string
  bold: boolean
  italic: boolean
  underLine: boolean
  strike: boolean
  size: number
  fontName: string
  color?: string
  verticalAlign: string
}
```

---

## 錯誤處理

### 錯誤類型

#### 1. 網路錯誤

```typescript
try {
  const data = await uploadFile(file)
} catch (error) {
  if (axios.isAxiosError(error)) {
    if (!error.response) {
      // 網路連線失敗
      console.error('網路連線錯誤:', error.message)
      alert('無法連線到伺服器,請檢查網路連線')
    }
  }
}
```

#### 2. HTTP 錯誤

```typescript
try {
  const data = await uploadFile(file)
} catch (error) {
  if (axios.isAxiosError(error)) {
    const status = error.response?.status
    
    switch (status) {
      case 400:
        alert('檔案格式不正確或檔案損壞')
        break
      case 413:
        alert('檔案大小超過限制 (最大 100MB)')
        break
      case 500:
        alert('伺服器處理錯誤,請稍後再試')
        break
      default:
        alert(`HTTP 錯誤: ${status}`)
    }
  }
}
```

#### 3. 業務邏輯錯誤

```typescript
try {
  const response = await uploadFile(file)
  
  if (!response.success) {
    // API 返回錯誤
    console.error('API Error:', response.error)
    alert(response.message || '處理失敗')
  } else {
    // 處理成功
    excelData.value = response.data
  }
} catch (error) {
  console.error('Upload failed:', error)
}
```

### 統一錯誤處理

```typescript
interface ApiError {
  message: string
  code?: string
  details?: string
}

const handleApiError = (error: unknown): ApiError => {
  if (axios.isAxiosError(error)) {
    // Axios 錯誤
    if (error.response) {
      // 伺服器返回錯誤響應
      return {
        message: error.response.data?.message || '請求失敗',
        code: error.response.data?.error?.code,
        details: error.response.data?.error?.details
      }
    } else if (error.request) {
      // 請求已發送但沒有收到響應
      return {
        message: '無法連線到伺服器',
        code: 'NETWORK_ERROR'
      }
    } else {
      // 請求配置錯誤
      return {
        message: error.message || '請求配置錯誤',
        code: 'CONFIG_ERROR'
      }
    }
  } else if (error instanceof Error) {
    // 一般錯誤
    return {
      message: error.message,
      code: 'UNKNOWN_ERROR'
    }
  } else {
    // 未知錯誤
    return {
      message: '未知錯誤',
      code: 'UNKNOWN_ERROR'
    }
  }
}

// 使用
try {
  const data = await uploadFile(file)
} catch (error) {
  const apiError = handleApiError(error)
  console.error('Error:', apiError)
  alert(apiError.message)
}
```

---

## 使用範例

### 完整的檔案上傳流程

```typescript
import { ref } from 'vue'
import axios from 'axios'
import type { ExcelData } from '@/types/excel'

const API_BASE_URL = 'http://localhost:5000'

// 狀態
const excelData = ref<ExcelData | null>(null)
const loading = ref(false)
const message = ref('')
const messageType = ref<'success' | 'error' | 'info'>('info')

// 檔案上傳
const handleFileUpload = async (file: File) => {
  // 重置狀態
  loading.value = true
  message.value = ''
  excelData.value = null
  
  try {
    // 1. 驗證檔案
    if (!validateFile(file)) {
      return
    }
    
    // 2. 建立 FormData
    const formData = new FormData()
    formData.append('file', file)
    
    // 3. 上傳檔案
    const response = await axios.post(
      `${API_BASE_URL}/api/excel/upload`,
      formData,
      {
        headers: {
          'Content-Type': 'multipart/form-data'
        },
        timeout: 300000, // 5 分鐘
        onUploadProgress: (progressEvent) => {
          if (progressEvent.total) {
            const percent = Math.round(
              (progressEvent.loaded * 100) / progressEvent.total
            )
            message.value = `上傳中... ${percent}%`
          }
        }
      }
    )
    
    // 4. 處理響應
    if (response.data.success) {
      excelData.value = response.data.data
      message.value = '檔案上傳成功!'
      messageType.value = 'success'
      
      console.log('Excel 資料:', excelData.value)
      console.log('處理時間:', excelData.value.processingTime)
    } else {
      throw new Error(response.data.message || '上傳失敗')
    }
    
  } catch (error) {
    // 5. 錯誤處理
    const apiError = handleApiError(error)
    message.value = apiError.message
    messageType.value = 'error'
    console.error('上傳錯誤:', apiError)
    
  } finally {
    loading.value = false
  }
}

// 檔案驗證
const validateFile = (file: File): boolean => {
  // 檢查檔案類型
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    message.value = '不支援的檔案格式,請上傳 .xlsx 或 .xls 檔案'
    messageType.value = 'error'
    loading.value = false
    return false
  }
  
  // 檢查檔案大小 (100MB)
  const maxSize = 100 * 1024 * 1024
  if (file.size > maxSize) {
    message.value = '檔案大小超過限制 (最大 100MB)'
    messageType.value = 'error'
    loading.value = false
    return false
  }
  
  return true
}
```

### 載入範例資料

```typescript
const loadSampleData = async () => {
  loading.value = true
  message.value = ''
  
  try {
    const response = await axios.get(
      `${API_BASE_URL}/api/excel/sample`
    )
    
    excelData.value = response.data
    message.value = '範例資料載入成功'
    messageType.value = 'success'
    
  } catch (error) {
    const apiError = handleApiError(error)
    message.value = apiError.message
    messageType.value = 'error'
    
  } finally {
    loading.value = false
  }
}
```

---

## 最佳實踐

### 1. 使用 TypeScript

```typescript
// ✅ 正確: 明確的型別定義
const excelData = ref<ExcelData | null>(null)

// ❌ 錯誤: 使用 any
const excelData = ref<any>(null)
```

### 2. 錯誤處理

```typescript
// ✅ 正確: 完整的錯誤處理
try {
  const data = await uploadFile(file)
  // 處理成功
} catch (error) {
  // 處理錯誤
  const apiError = handleApiError(error)
  console.error('Error:', apiError)
  alert(apiError.message)
}

// ❌ 錯誤: 忽略錯誤
const data = await uploadFile(file) // 可能拋出未捕獲的錯誤
```

### 3. 載入狀態

```typescript
// ✅ 正確: 顯示載入狀態
loading.value = true
try {
  const data = await uploadFile(file)
} finally {
  loading.value = false
}

// ❌ 錯誤: 沒有載入狀態
const data = await uploadFile(file) // 用戶不知道正在載入
```

### 4. 超時處理

```typescript
// ✅ 正確: 設定合理的超時時間
const response = await axios.post(url, data, {
  timeout: 300000 // 5 分鐘
})

// ❌ 錯誤: 沒有超時設定
const response = await axios.post(url, data) // 可能無限等待
```

### 5. 進度顯示

```typescript
// ✅ 正確: 顯示上傳進度
const response = await axios.post(url, formData, {
  onUploadProgress: (progressEvent) => {
    if (progressEvent.total) {
      const percent = Math.round(
        (progressEvent.loaded * 100) / progressEvent.total
      )
      console.log(`上傳進度: ${percent}%`)
    }
  }
})

// ❌ 錯誤: 沒有進度顯示
const response = await axios.post(url, formData) // 用戶不知道進度
```

### 6. 資料驗證

```typescript
// ✅ 正確: 驗證 API 響應
const response = await axios.post(url, data)
if (response.data && response.data.success) {
  // 處理資料
} else {
  throw new Error('無效的響應資料')
}

// ❌ 錯誤: 不驗證響應
const response = await axios.post(url, data)
excelData.value = response.data.data // 可能為 undefined
```

---

## 測試

### 單元測試範例

```typescript
import { describe, it, expect, vi } from 'vitest'
import axios from 'axios'

describe('API Integration', () => {
  it('should upload file successfully', async () => {
    // Mock axios
    const mockResponse = {
      data: {
        success: true,
        data: {
          fileName: 'test.xlsx',
          worksheets: []
        }
      }
    }
    vi.spyOn(axios, 'post').mockResolvedValue(mockResponse)
    
    // 測試
    const file = new File([''], 'test.xlsx')
    const result = await uploadFile(file)
    
    // 驗證
    expect(result.success).toBe(true)
    expect(result.data.fileName).toBe('test.xlsx')
  })
  
  it('should handle upload error', async () => {
    // Mock 錯誤
    vi.spyOn(axios, 'post').mockRejectedValue(
      new Error('Network error')
    )
    
    // 測試
    const file = new File([''], 'test.xlsx')
    
    // 驗證
    await expect(uploadFile(file)).rejects.toThrow('Network error')
  })
})
```

---

**文檔維護者:** ExcelReader Team  
**最後更新:** 2025年10月9日
