# ExcelReaderVue - 元件開發指南

**版本:** 2.0.0  
**最後更新:** 2025年10月9日

---

## 元件架構

### ExcelReader.vue

主要元件,負責 Excel 檔案的上傳、解析和顯示。

**檔案位置:** `src/components/ExcelReader.vue`  
**程式碼行數:** 1,643 行

### 元件結構

```vue
<script setup lang="ts">
// 1. Imports
// 2. State Management
// 3. Computed Properties
// 4. Methods
// 5. Lifecycle Hooks
</script>

<template>
  <!-- UI Template -->
</template>

<style scoped>
  /* Component Styles */
</style>
```

---

## 核心功能

### 1. 檔案上傳

#### 拖拽上傳

```typescript
const handleDrop = (e: DragEvent) => {
  e.preventDefault()
  const files = e.dataTransfer?.files
  if (files && files.length > 0) {
    handleFileSelect({ target: { files } } as any)
  }
}
```

#### 按鈕上傳

```typescript
const handleFileSelect = async (event: Event) => {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  if (file) {
    await uploadFile(file)
  }
}
```

### 2. 資料處理

#### API 呼叫

```typescript
const uploadFile = async (file: File) => {
  const formData = new FormData()
  formData.append('file', file)
  
  const response = await axios.post(
    `${API_BASE_URL}/api/excel/upload`,
    formData
  )
  
  return response.data
}
```

#### 資料轉換

```typescript
const processExcelData = (data: ExcelData) => {
  // 轉換為表格格式
  const rows = convertToRows(data.worksheets[0].cells)
  return rows
}
```

### 3. 樣式計算

#### 儲存格樣式

```typescript
const getCellStyle = (cell: ExcelCellInfo) => {
  const styles: any = {}
  
  // 字體
  if (cell.font) {
    styles.fontFamily = cell.font.name
    styles.fontSize = `${cell.font.size}pt`
    styles.fontWeight = cell.font.bold ? 'bold' : 'normal'
    styles.color = `#${cell.font.color}`
  }
  
  // 對齊
  if (cell.alignment) {
    styles.textAlign = cell.alignment.horizontal?.toLowerCase()
    styles.verticalAlign = cell.alignment.vertical?.toLowerCase()
  }
  
  // 背景色
  if (cell.fill?.backgroundColor) {
    styles.backgroundColor = `#${cell.fill.backgroundColor}`
  }
  
  return styles
}
```

### 4. Rich Text 渲染

```typescript
const renderRichText = (cell: ExcelCellInfo): string => {
  if (!cell.richText) return cell.text
  
  return cell.richText.map(part => {
    let html = part.text
    if (part.bold) html = `<strong>${html}</strong>`
    if (part.italic) html = `<em>${html}</em>`
    if (part.underLine) html = `<u>${html}</u>`
    if (part.color) {
      html = `<span style="color: #${part.color}">${html}</span>`
    }
    return html
  }).join('')
}
```

---

## 擴展指南

### 新增功能

#### 1. 新增工作表切換功能

```typescript
// State
const currentWorksheetIndex = ref(0)

// Method
const switchWorksheet = (index: number) => {
  if (excelData.value && index < excelData.value.worksheets.length) {
    currentWorksheetIndex.value = index
  }
}

// Template
<template>
  <div class="worksheet-tabs">
    <button
      v-for="(ws, index) in excelData.worksheets"
      :key="index"
      @click="switchWorksheet(index)"
      :class="{ active: currentWorksheetIndex === index }"
    >
      {{ ws.name }}
    </button>
  </div>
</template>
```

#### 2. 新增匯出功能

```typescript
const exportToJSON = () => {
  const json = JSON.stringify(excelData.value, null, 2)
  const blob = new Blob([json], { type: 'application/json' })
  const url = URL.createObjectURL(blob)
  
  const a = document.createElement('a')
  a.href = url
  a.download = `${excelData.value?.fileName || 'export'}.json`
  a.click()
  
  URL.revokeObjectURL(url)
}
```

---

## 效能優化

### 1. 虛擬滾動

對於大型資料集,使用虛擬滾動減少 DOM 節點數量。

### 2. 懶加載圖片

延遲載入不在視窗內的圖片。

### 3. 計算快取

使用 `computed` 快取計算結果。

---

**文檔維護者:** ExcelReader Team
