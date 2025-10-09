# ExcelReaderVue - 前端專案文檔

**版本:** 2.0.0  
**框架:** Vue 3 + TypeScript + Vite  
**最後更新:** 2025年10月9日

---

## 📋 目錄

- [專案簡介](#專案簡介)
- [核心功能](#核心功能)
- [技術架構](#技術架構)
- [快速開始](#快速開始)
- [專案結構](#專案結構)
- [開發指南](#開發指南)
- [API 整合](#api-整合)
- [元件說明](#元件說明)
- [樣式設計](#樣式設計)
- [效能優化](#效能優化)
- [問題排查](#問題排查)
- [更新日誌](#更新日誌)

---

## 專案簡介

ExcelReaderVue 是一個基於 Vue 3 的現代化前端應用程式,用於視覺化顯示 Excel 檔案內容。它與 ExcelReaderAPI 後端服務配合使用,提供完整的 Excel 檔案上傳、解析和顯示功能。

### 專案定位

- 🎯 **目標使用者**: 需要在瀏覽器中預覽和分析 Excel 檔案的用戶
- 🎨 **設計理念**: 簡潔、直覺、高效能
- 🔧 **技術選型**: 使用最新的 Vue 3 Composition API 和 TypeScript

### 專案亮點

✨ **現代化技術棧**
- Vue 3.5.18 + Composition API
- TypeScript 5.8.0 (完整型別支援)
- Vite 7.0.6 (快速建置)
- Pinia 3.0.3 (狀態管理)

🎨 **豐富的功能**
- 拖拽上傳支援
- 即時預覽 Excel 內容
- 完整樣式還原 (字體、顏色、邊框、對齊)
- 圖片顯示 (包含 In-Cell 圖片)
- Rich Text 格式支援
- 合併儲存格顯示
- 浮動物件處理

⚡ **效能優化**
- 虛擬滾動 (大型資料集)
- 懶加載圖片
- 智能渲染優化

---

## 核心功能

### 1. 檔案上傳

#### 拖拽上傳
- 支援拖拽 Excel 檔案到上傳區域
- 即時檔案驗證
- 支援 `.xlsx` 和 `.xls` 格式

#### 按鈕上傳
- 點擊按鈕選擇檔案
- 檔案大小限制: 100MB
- 上傳進度顯示

### 2. 資料顯示

#### 表格顯示
- 完整還原 Excel 表格樣式
- 支援標頭類型切換:
  - Excel 欄位標頭 (A, B, C, D...)
  - 工作表內容標頭 (第一行內容)

#### 儲存格渲染
- **文字**: 支援換行、Rich Text
- **數字**: 保留數字格式
- **日期**: 正確顯示日期格式
- **公式**: 顯示計算結果
- **圖片**: In-Cell 圖片和浮動圖片
- **合併儲存格**: 正確顯示合併範圍

#### 樣式還原
- **字體**: 字型、大小、粗體、斜體、顏色
- **對齊**: 水平、垂直對齊
- **邊框**: 上下左右邊框、樣式、顏色
- **填充**: 背景色、圖案
- **尺寸**: 列高、欄寬

### 3. 互動功能

#### 儲存格資訊
- 滑鼠懸停顯示完整資訊
- 顯示位置 (如 A1, B2)
- 顯示公式 (如有)
- 顯示數字格式

#### 工作表切換
- 支援多工作表檔案
- 快速切換工作表
- 顯示工作表名稱

#### 範例資料
- 載入範例資料功能
- 下載範例 Excel 檔案

### 4. 特殊功能

#### 圖片處理
- In-Cell Pictures (EPPlus 8.x)
- 浮動圖片
- 圖片縮放和定位
- Base64 圖片顯示

#### 浮動物件
- 文字方塊
- 圖形
- 智能文字合併

#### Rich Text
- 多格式文字
- 字體大小和顏色變化
- 上標/下標支援

---

## 技術架構

### 技術棧總覽

```
┌─────────────────────────────────────────┐
│          ExcelReaderVue v2.0            │
├─────────────────────────────────────────┤
│  核心: Vue 3.5.18 + TypeScript 5.8.0    │
│  建置: Vite 7.0.6                       │
│  狀態: Pinia 3.0.3                      │
│  路由: Vue Router 4.5.1                 │
│  HTTP: Axios 1.12.2                     │
│  開發: Vue DevTools 8.0.0               │
└─────────────────────────────────────────┘
```

### 架構設計

```
src/
├── components/          # Vue 元件
│   └── ExcelReader.vue  # 主要元件 (1,643 行)
├── types/              # TypeScript 型別定義
│   ├── excel.ts        # Excel 資料型別
│   └── index.ts        # 通用型別
├── router/             # 路由配置
│   └── index.ts
├── stores/             # Pinia 狀態管理
│   └── counter.ts
├── App.vue             # 根元件
└── main.ts             # 應用程式入口
```

### 依賴關係圖

```
App.vue
  │
  └── ExcelReader.vue (主元件)
        │
        ├── Axios → ExcelReaderAPI (後端)
        ├── Types (excel.ts)
        └── 本地狀態 (ref, reactive)
```

---

## 快速開始

### 環境需求

| 工具 | 版本要求 |
|------|---------|
| **Node.js** | ^20.19.0 或 >=22.12.0 |
| **npm** | 10.0.0 或更高 |
| **現代瀏覽器** | Chrome 90+, Firefox 88+, Safari 14+, Edge 90+ |

### 安裝步驟

#### 1. Clone 專案

```bash
git clone https://github.com/akikuma2762/ExcelReader.git
cd ExcelReader/ExcelReaderVue
```

#### 2. 安裝依賴

```bash
npm install
```

#### 3. 配置 API 端點

編輯 `src/components/ExcelReader.vue`,設定 API URL:

```typescript
// 開發環境
const API_BASE_URL = 'http://localhost:5000'

// 生產環境
const API_BASE_URL = 'https://your-api-domain.com'
```

#### 4. 啟動開發伺服器

```bash
npm run dev
```

應用程式將在 `http://localhost:5173` 啟動。

#### 5. 建置生產版本

```bash
npm run build
```

建置後的檔案將在 `dist/` 目錄中。

### 開發腳本

| 指令 | 說明 |
|------|------|
| `npm run dev` | 啟動開發伺服器 (HMR) |
| `npm run build` | 建置生產版本 |
| `npm run preview` | 預覽生產建置 |
| `npm run type-check` | TypeScript 型別檢查 |
| `npm run lint` | ESLint 程式碼檢查 |
| `npm run format` | Prettier 程式碼格式化 |

---

## 專案結構

### 目錄說明

```
ExcelReaderVue/
├── public/                    # 靜態資源
│   └── favicon.ico           # 網站圖示
│
├── src/                      # 原始碼
│   ├── components/           # Vue 元件
│   │   └── ExcelReader.vue   # 主要元件 (1,643 行)
│   │
│   ├── types/                # TypeScript 型別
│   │   ├── excel.ts          # Excel 資料型別定義
│   │   └── index.ts          # 匯出所有型別
│   │
│   ├── router/               # Vue Router 配置
│   │   └── index.ts          # 路由定義
│   │
│   ├── stores/               # Pinia Store
│   │   └── counter.ts        # 範例 Store
│   │
│   ├── App.vue               # 根元件
│   └── main.ts               # 應用程式入口
│
├── doc/                      # 文檔 (本文件)
│   ├── README.md             # 專案總覽
│   ├── API_INTEGRATION.md    # API 整合文檔
│   ├── COMPONENT_GUIDE.md    # 元件開發指南
│   ├── CONTRIBUTING.md       # 貢獻指南
│   └── CHANGELOG.md          # 更新日誌
│
├── .vscode/                  # VS Code 配置
│   ├── extensions.json       # 推薦擴展
│   └── settings.json         # 編輯器設定
│
├── index.html                # HTML 模板
├── vite.config.ts            # Vite 配置
├── tsconfig.json             # TypeScript 配置
├── package.json              # 專案依賴
├── .prettierrc.json          # Prettier 配置
├── eslint.config.ts          # ESLint 配置
└── README.md                 # 專案說明
```

### 核心檔案說明

#### ExcelReader.vue (1,643 行)

主要的 Excel 顯示元件,包含:

- **上傳功能**: 檔案選擇、拖拽上傳
- **資料處理**: API 呼叫、資料轉換
- **表格渲染**: 動態表格生成
- **樣式處理**: CSS 樣式計算
- **圖片處理**: Base64 圖片顯示

#### types/excel.ts

完整的 TypeScript 型別定義:

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
}

export interface ExcelCellInfo {
  position: CellPosition
  value: any
  text: string
  dataType: string
  font: FontInfo
  alignment: AlignmentInfo
  border: BorderInfo
  fill: FillInfo
  dimensions: DimensionInfo
  images?: ImageInfo[]
  floatingObjects?: FloatingObjectInfo[]
  richText?: RichTextPart[]
  // ... 更多屬性
}
```

---

## 開發指南

### 程式碼風格

#### Vue 元件結構

```vue
<script setup lang="ts">
// 1. Imports
import { ref, reactive, computed, onMounted } from 'vue'
import type { ExcelData, ExcelCellInfo } from '@/types/excel'

// 2. Props & Emits
interface Props {
  apiUrl?: string
}
const props = withDefaults(defineProps<Props>(), {
  apiUrl: 'http://localhost:5000'
})

// 3. Reactive State
const excelData = ref<ExcelData | null>(null)
const loading = ref(false)

// 4. Computed Properties
const totalRows = computed(() => excelData.value?.worksheets[0]?.rowCount || 0)

// 5. Methods
const handleFileUpload = async (file: File) => {
  // Implementation
}

// 6. Lifecycle Hooks
onMounted(() => {
  // Initialization
})
</script>

<template>
  <!-- Template -->
</template>

<style scoped>
/* Scoped Styles */
</style>
```

#### TypeScript 使用

```typescript
// ✅ 正確: 明確的型別定義
const excelData = ref<ExcelData | null>(null)
const cells = computed(() => excelData.value?.worksheets[0]?.cells || [])

// ✅ 正確: 型別守衛
function isImageCell(cell: ExcelCellInfo): boolean {
  return cell.images !== undefined && cell.images.length > 0
}

// ❌ 錯誤: 使用 any
const data: any = fetchData() // 應該定義明確型別
```

#### 命名規範

```typescript
// 元件: PascalCase
ExcelReader.vue
DataTable.vue

// 函數: camelCase
handleFileUpload()
getCellStyle()

// 常數: UPPER_SNAKE_CASE
const API_BASE_URL = 'http://localhost:5000'
const MAX_FILE_SIZE = 100 * 1024 * 1024

// 型別/介面: PascalCase
interface ExcelData { }
type CellStyle = { }
```

### 狀態管理

#### 使用 ref 和 reactive

```typescript
// 簡單值使用 ref
const loading = ref(false)
const fileName = ref('')

// 複雜物件使用 reactive
const uploadState = reactive({
  progress: 0,
  status: 'idle',
  error: null
})

// 存取值
console.log(loading.value)        // ref 需要 .value
console.log(uploadState.progress)  // reactive 不需要
```

#### Computed Properties

```typescript
// 從 excelData 派生的狀態
const totalRows = computed(() => {
  return excelData.value?.worksheets[0]?.rowCount || 0
})

const hasImages = computed(() => {
  return excelData.value?.worksheets[0]?.cells.some(
    cell => cell.images && cell.images.length > 0
  ) || false
})
```

### API 呼叫

#### 使用 Axios

```typescript
import axios from 'axios'

const uploadFile = async (file: File) => {
  const formData = new FormData()
  formData.append('file', file)
  
  try {
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
  } catch (error) {
    console.error('上傳失敗:', error)
    throw error
  }
}
```

### 錯誤處理

```typescript
const handleFileUpload = async (file: File) => {
  loading.value = true
  message.value = ''
  
  try {
    // 驗證檔案
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      throw new Error('不支援的檔案格式')
    }
    
    if (file.size > MAX_FILE_SIZE) {
      throw new Error('檔案大小超過限制')
    }
    
    // 上傳檔案
    const data = await uploadFile(file)
    
    // 處理成功
    excelData.value = data
    message.value = '上傳成功!'
    messageType.value = 'success'
    
  } catch (error) {
    // 處理錯誤
    if (axios.isAxiosError(error)) {
      message.value = error.response?.data?.message || '上傳失敗'
    } else if (error instanceof Error) {
      message.value = error.message
    } else {
      message.value = '未知錯誤'
    }
    messageType.value = 'error'
    
  } finally {
    loading.value = false
  }
}
```

---

## API 整合

### API 端點

ExcelReaderVue 使用以下 API 端點:

| 端點 | 方法 | 用途 |
|------|------|------|
| `/api/excel/upload` | POST | 上傳並解析 Excel 檔案 |
| `/api/excel/sample` | GET | 獲取範例資料 |

### 請求範例

#### 上傳檔案

```typescript
const uploadFile = async (file: File) => {
  const formData = new FormData()
  formData.append('file', file)
  
  const response = await axios.post(
    'http://localhost:5000/api/excel/upload',
    formData,
    {
      headers: {
        'Content-Type': 'multipart/form-data'
      }
    }
  )
  
  return response.data
}
```

#### 載入範例資料

```typescript
const loadSampleData = async () => {
  const response = await axios.get(
    'http://localhost:5000/api/excel/sample'
  )
  
  excelData.value = response.data
}
```

### 響應處理

#### 成功響應

```json
{
  "success": true,
  "data": {
    "fileName": "test.xlsx",
    "fileSize": 123456,
    "worksheets": [...],
    "totalWorksheets": 1,
    "processingTime": "1.234s"
  }
}
```

#### 錯誤響應

```json
{
  "success": false,
  "message": "檔案格式不正確",
  "error": {
    "code": "INVALID_FILE_FORMAT",
    "details": "..."
  }
}
```

---

## 元件說明

### ExcelReader 元件

#### Props

目前不接受 Props,所有配置都在元件內部。

未來可擴展:

```typescript
interface Props {
  apiUrl?: string
  maxFileSize?: number
  allowedFormats?: string[]
  showFormatInfo?: boolean
  showPositionInfo?: boolean
}
```

#### Events

目前不發出 Events,所有狀態都在元件內部管理。

未來可擴展:

```typescript
// 定義 emits
const emit = defineEmits<{
  'file-uploaded': [data: ExcelData]
  'upload-error': [error: Error]
  'cell-clicked': [cell: ExcelCellInfo]
}>()

// 使用
emit('file-uploaded', excelData.value)
```

#### 方法

| 方法名 | 說明 | 參數 | 返回值 |
|--------|------|------|--------|
| `handleFileSelect` | 處理檔案選擇 | `Event` | `void` |
| `handleDrop` | 處理拖拽上傳 | `DragEvent` | `void` |
| `uploadFile` | 上傳檔案到 API | `File` | `Promise<ExcelData>` |
| `loadSampleData` | 載入範例資料 | - | `Promise<void>` |
| `getCellStyle` | 計算儲存格樣式 | `ExcelCellInfo` | `CSSProperties` |
| `renderRichText` | 渲染 Rich Text | `ExcelCellInfo` | `string` |

---

## 樣式設計

### CSS 架構

```
ExcelReader.vue
├── .container (主容器)
│   ├── .upload-section (上傳區域)
│   │   ├── .upload-area (拖拽區域)
│   │   └── .button-group (按鈕群組)
│   │
│   ├── .loading (載入狀態)
│   ├── .message (訊息顯示)
│   │
│   └── .data-section (資料顯示)
│       ├── .header-type-controls (標頭控制)
│       └── .table-container (表格容器)
│           └── .data-table (資料表格)
│               ├── thead (表頭)
│               └── tbody (表格內容)
│                   └── td (儲存格)
│                       ├── .cell-content (內容)
│                       ├── .cell-image (圖片)
│                       └── .floating-object (浮動物件)
```

### 主要樣式類別

#### 容器樣式

```css
.container {
  max-width: 1400px;
  margin: 0 auto;
  padding: 20px;
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}
```

#### 表格樣式

```css
.data-table {
  border-collapse: collapse;
  width: 100%;
  table-layout: fixed;
  font-size: 12px;
}

.data-table td {
  border: 1px solid #ddd;
  padding: 4px 8px;
  vertical-align: middle;
  position: relative;
  word-wrap: break-word;
  overflow-wrap: break-word;
}
```

#### 響應式設計

```css
@media (max-width: 768px) {
  .container {
    padding: 10px;
  }
  
  .data-table {
    font-size: 10px;
  }
  
  .data-table td {
    padding: 2px 4px;
  }
}
```

### 動態樣式計算

#### 儲存格樣式

```typescript
const getCellStyle = (cell: ExcelCellInfo) => {
  const styles: any = {}
  
  // 字體
  if (cell.font) {
    styles.fontFamily = cell.font.name || 'Arial'
    styles.fontSize = `${cell.font.size || 11}pt`
    styles.fontWeight = cell.font.bold ? 'bold' : 'normal'
    styles.fontStyle = cell.font.italic ? 'italic' : 'normal'
    styles.color = `#${cell.font.color || '000000'}`
  }
  
  // 對齊
  if (cell.alignment) {
    styles.textAlign = cell.alignment.horizontal?.toLowerCase() || 'left'
    styles.verticalAlign = cell.alignment.vertical?.toLowerCase() || 'middle'
  }
  
  // 背景色
  if (cell.fill?.backgroundColor) {
    styles.backgroundColor = `#${cell.fill.backgroundColor}`
  }
  
  // 邊框
  if (cell.border) {
    if (cell.border.top?.style !== 'None') {
      styles.borderTop = `1px ${cell.border.top.style.toLowerCase()} #${cell.border.top.color || '000'}`
    }
    // ... 其他邊框
  }
  
  return styles
}
```

---

## 效能優化

### 1. 虛擬滾動 (未來功能)

對於大型資料集,實作虛擬滾動:

```typescript
// 使用 vue-virtual-scroller
import { RecycleScroller } from 'vue-virtual-scroller'
import 'vue-virtual-scroller/dist/vue-virtual-scroller.css'
```

### 2. 圖片懶加載

```typescript
const loadImage = (imageInfo: ImageInfo) => {
  return new Promise((resolve) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.src = `data:${imageInfo.imageType};base64,${imageInfo.base64Data}`
  })
}
```

### 3. 計算快取

```typescript
// 使用 computed 快取樣式計算
const cellStyles = computed(() => {
  const cache = new Map()
  
  excelData.value?.worksheets[0]?.cells.forEach(cell => {
    const key = `${cell.position.row}-${cell.position.column}`
    cache.set(key, getCellStyle(cell))
  })
  
  return cache
})
```

### 4. 渲染優化

```vue
<template>
  <!-- 使用 v-show 而非 v-if (頻繁切換) -->
  <div v-show="showDetails" class="details">...</div>
  
  <!-- 使用 v-once (靜態內容) -->
  <div v-once>{{ staticContent }}</div>
  
  <!-- 使用 key 優化列表渲染 -->
  <tr v-for="(row, index) in rows" :key="`row-${index}`">
    ...
  </tr>
</template>
```

---

## 問題排查

### 常見問題

#### 1. CORS 錯誤

**問題:** 瀏覽器控制台顯示 CORS 錯誤

**解決方案:**

確保後端 API (ExcelReaderAPI) 已配置 CORS:

```csharp
// Program.cs
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", builder =>
    {
        builder.AllowAnyOrigin()
               .AllowAnyMethod()
               .AllowAnyHeader();
    });
});

app.UseCors("AllowAll");
```

#### 2. 檔案上傳失敗

**問題:** 上傳大檔案時失敗

**檢查項目:**
- 後端檔案大小限制 (預設 100MB)
- 瀏覽器網路超時設定
- 伺服器記憶體限制

**解決方案:**

```typescript
// 增加超時時間
axios.post(url, data, {
  timeout: 300000 // 5 分鐘
})
```

#### 3. 圖片不顯示

**問題:** Excel 中的圖片在網頁中不顯示

**檢查項目:**
- Base64 資料是否正確
- 圖片類型是否支援
- CSS 樣式是否正確

**解決方案:**

```typescript
// 檢查圖片資料
console.log('Image type:', imageInfo.imageType)
console.log('Image data length:', imageInfo.base64Data.length)

// 正確的 data URL 格式
const dataUrl = `data:image/${imageInfo.imageType.toLowerCase()};base64,${imageInfo.base64Data}`
```

#### 4. 樣式不正確

**問題:** Excel 樣式在網頁中顯示不正確

**常見原因:**
- 字體未安裝
- 顏色計算錯誤
- CSS 優先級問題

**解決方案:**

```typescript
// 使用 fallback 字體
styles.fontFamily = `${cell.font.name}, Arial, sans-serif`

// 驗證顏色值
if (color && /^[0-9A-F]{6}$/i.test(color)) {
  styles.color = `#${color}`
}
```

### 除錯工具

#### Vue DevTools

安裝 Vue DevTools 瀏覽器擴展:
- Chrome: https://chrome.google.com/webstore
- Firefox: https://addons.mozilla.org

#### 日誌記錄

```typescript
// 開發模式啟用詳細日誌
const DEBUG = import.meta.env.DEV

const log = (...args: any[]) => {
  if (DEBUG) {
    console.log('[ExcelReader]', ...args)
  }
}

// 使用
log('Uploading file:', file.name)
log('Excel data:', excelData.value)
```

---

## 更新日誌

### v2.0.0 (2025-10-09)

#### 新增功能
- ✨ 完整的 TypeScript 型別支援
- ✨ Rich Text 格式顯示
- ✨ In-Cell 圖片支援
- ✨ 浮動物件顯示
- ✨ 合併儲存格支援
- ✨ 標頭類型切換功能

#### 改進
- 🎨 重新設計 UI
- ⚡ 效能優化
- 🐛 修復多個 Bug

#### 技術升級
- 升級到 Vue 3.5.18
- 升級到 Vite 7.0.6
- 升級到 TypeScript 5.8.0

---

## 授權

與主專案相同授權。

---

**文檔維護者:** ExcelReader Team  
**最後更新:** 2025年10月9日  
**版本:** 2.0.0
