# Excel 格式增強功能實作報告

## 新增功能

### 1. 📝 TextAlign 文字對齊屬性支援

#### 後端改進
- **ExcelCellInfo 模型**: 新增 `TextAlign` 屬性
- **ExcelController**: 
  - 新增 `GetTextAlign` 方法轉換 EPPlus 對齊格式為 CSS 值
  - 支援的對齊方式：
    - `Left` → "left"
    - `Center` → "center" 
    - `Right` → "right"
    - `Justify` → "justify"
    - `Fill` → "left"
    - `CenterContinuous` → "center"
    - `Distributed` → "justify"

#### 前端改進
- **TypeScript 介面**: ExcelCellInfo 新增 `textAlign?: string` 屬性
- **樣式渲染**: 
  - `getHeaderStyle` 和 `getCellStyle` 函數都支援 textAlign
  - 自動套用 CSS `text-align` 屬性

### 2. 🔤 換行文字處理

#### 核心功能
新增 `formatTextWithLineBreaks` 函數，支援多種換行格式：
```typescript
const formatTextWithLineBreaks = (text: string): string => {
  return text.replace(/\r\n/g, '<br>').replace(/\n/g, '<br>').replace(/\r/g, '<br>')
}
```

#### 應用範圍
- **Rich Text 內容**: 在 `renderRichText` 函數中處理每個文字片段
- **一般文字**: 透過 `v-html` 指令渲染，支援：
  - 表頭內容換行
  - 儲存格內容換行
  - 保持 HTML 安全性（經過轉義處理）

## 技術實作細節

### 後端 API 回應格式
```json
{
  "value": "範例文字\n第二行",
  "displayText": "範例文字\n第二行", 
  "textAlign": "center",
  "fontBold": true,
  "fontSize": 12,
  "fontName": "Arial"
}
```

### 前端渲染結果
```html
<td style="text-align: center; font-weight: bold; font-size: 12px; font-family: 'Arial'">
  <span>範例文字<br>第二行</span>
</td>
```

## 相容性與安全性

### ✅ HTML 安全性
- 所有文字內容經過 `escapeHtml` 函數轉義
- 防止 XSS 攻擊
- 只允許安全的 HTML 標籤（如 `<br>`）

### ✅ 向後相容
- 舊有功能完全保持
- 新屬性為可選項，不影響現有資料
- Rich Text 功能增強但不改變原有邏輯

### ✅ 瀏覽器支援
- CSS `text-align` 屬性廣泛支援
- `<br>` 標籤標準 HTML 元素
- Vue 3 `v-html` 指令標準功能

## 使用範例

### Excel 換行文字
```
儲存格內容：
第一行文字
第二行文字
第三行文字
```

### 網頁顯示結果
表格中會正確顯示為三行文字，保持原始 Excel 的換行格式。

### 對齊效果
- 左對齊文字將顯示為 `text-align: left`
- 居中文字將顯示為 `text-align: center`  
- 右對齊文字將顯示為 `text-align: right`

## 測試建議

1. **對齊測試**: 上傳包含不同對齊方式的 Excel 檔案
2. **換行測試**: 測試包含多行文字的儲存格
3. **Rich Text 換行**: 測試 Rich Text 格式中的換行處理
4. **混合格式**: 測試同時包含對齊、換行、字體樣式的複雜格式

這些改進讓 Excel 讀取器更接近原始 Excel 的顯示效果，大幅提升使用者體驗！