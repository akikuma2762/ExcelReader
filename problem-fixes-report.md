# 問題修復報告

## 已修復的問題

### 問題1: 後端EPPlus沒有記錄rowSpan與colSpan

**解決方案:**
1. **後端 (C#)**:
   - 在`ExcelCellInfo`類別中新增了三個屬性:
     - `RowSpan`: 行合併數量
     - `ColSpan`: 欄合併數量  
     - `IsMerged`: 是否為合併儲存格
   
   - 在`CreateCellInfo`方法中添加合併儲存格檢測:
   ```csharp
   // 檢查是否為合併儲存格
   if (cell.Merge)
   {
       cellInfo.IsMerged = true;
       cellInfo.RowSpan = cell.Rows;
       cellInfo.ColSpan = cell.Columns;
   }
   ```

2. **前端 (Vue + TypeScript)**:
   - 更新TypeScript介面添加合併儲存格屬性
   - 在表格模板中添加`colspan`和`rowspan`屬性支援
   - 在Tooltip中顯示合併儲存格信息

### 問題2: 前端字體名稱HTML轉義問題

**問題描述:**
```html
<th style="font-family: &quot;Arial Black&quot;;">
```
字體名稱中的引號被HTML轉義為`&quot;`，導致CSS無效。

**解決方案:**
1. **修正getHeaderStyle方法**:
   ```javascript
   if (header.fontName) {
     style.fontFamily = `"${header.fontName}"`  // 正確添加引號
   }
   ```

2. **Rich Text渲染已正確處理**:
   ```javascript
   if (part.fontName && part.fontName.trim()) 
     styles.push(`font-family: "${part.fontName}"`)
   ```

## 新功能特色

### 合併儲存格支援
- ✅ 自動檢測Excel中的合併儲存格
- ✅ 正確渲染colspan和rowspan
- ✅ Tooltip顯示合併信息 (`合併儲存格: 2行 x 3欄`)
- ✅ 保持合併儲存格的原始格式和內容

### 字體渲染改進
- ✅ 修正字體名稱CSS渲染問題
- ✅ 支援含空格的字體名稱 (如"Arial Black", "Times New Roman")
- ✅ Rich Text片段中每個字體都正確渲染
- ✅ 標題行和資料行統一使用正確的字體CSS

## 測試建議

### 合併儲存格測試
1. 上傳包含合併儲存格的Excel檔案
2. 確認表格中合併的儲存格正確顯示
3. 檢查Tooltip是否顯示合併信息
4. 驗證合併儲存格的格式是否保留

### 字體渲染測試  
1. 上傳包含特殊字體名稱的Excel (如Arial Black, Times New Roman)
2. 檢查瀏覽器開發者工具中CSS是否正確
3. 確認Rich Text中的字體變化正確顯示
4. 驗證標題行字體渲染正常

## 技術細節

### EPPlus合併儲存格API
- `cell.Merge`: 檢查儲存格是否為合併儲存格的一部分
- `cell.Rows`: 合併儲存格的行數 
- `cell.Columns`: 合併儲存格的欄數

### CSS字體名稱規範
- 包含空格的字體名稱必須用引號包圍
- 正確: `font-family: "Arial Black"`
- 錯誤: `font-family: Arial Black` 或 `font-family: &quot;Arial Black&quot;`

這些修復大幅提升了Excel格式保真度，特別是對於複雜的表格結構和字體格式。