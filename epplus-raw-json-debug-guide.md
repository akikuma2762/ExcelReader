# EPPlus 原始 JSON 調試工具

## 🔍 調試端點說明

我已經為您添加了一個專門的調試端點，用於查看 EPPlus 讀取 Excel 時的所有原始屬性。

### 📡 API 端點
```
POST /api/excel/debug-raw-data
Content-Type: multipart/form-data
```

### 🚀 使用方法

#### 1. 使用 Postman 或類似工具
```http
POST http://localhost:5280/api/excel/debug-raw-data
Content-Type: multipart/form-data

[上傳 Excel 檔案]
```

#### 2. 使用 cURL 命令
```bash
curl -X POST "http://localhost:5280/api/excel/debug-raw-data" \
     -F "file=@your-excel-file.xlsx"
```

### 📊 回應結構

調試端點會返回包含以下信息的完整 JSON：

#### 🗂️ 工作表信息
```json
{
  "fileName": "your-file.xlsx",
  "worksheetInfo": {
    "name": "Sheet1",
    "totalRows": 10,
    "totalColumns": 5,
    "defaultColWidth": 8.43,
    "defaultRowHeight": 15
  }
}
```

#### 📋 所有工作表清單
```json
{
  "allWorksheets": [
    {
      "name": "Sheet1",
      "index": 0,
      "state": "Visible"
    }
  ]
}
```

#### 🔬 儲存格詳細屬性 (前 5x5 儲存格)
```json
{
  "sampleCells": [
    [
      {
        "position": {
          "row": 1,
          "column": 1,
          "address": "A1"
        },
        "value": "標題文字",
        "text": "標題文字",
        "formula": "",
        "formulaR1C1": "",
        "valueType": "String",
        
        "numberFormat": {
          "format": "General",
          "numberFormatId": 0
        },
        
        "font": {
          "name": "Calibri",
          "size": 11,
          "bold": true,
          "italic": false,
          "underline": "None",
          "strike": false,
          "color": "FF000000",
          "colorTheme": null,
          "colorTint": 0,
          "charset": 0,
          "scheme": "Minor",
          "family": 2
        },
        
        "alignment": {
          "horizontal": "General",
          "vertical": "Bottom",
          "wrapText": false,
          "indent": 0,
          "readingOrder": "ContextDependent",
          "textRotation": 0,
          "shrinkToFit": false
        },
        
        "border": {
          "top": { "style": "None", "color": null },
          "bottom": { "style": "None", "color": null },
          "left": { "style": "None", "color": null },
          "right": { "style": "None", "color": null },
          "diagonal": { "style": "None", "color": null },
          "diagonalUp": false,
          "diagonalDown": false
        },
        
        "fill": {
          "patternType": "None",
          "backgroundColor": null,
          "patternColor": null,
          "backgroundColorTheme": null,
          "backgroundColorTint": 0
        },
        
        "dimensions": {
          "columnWidth": 12.5,
          "rowHeight": 15,
          "isMerged": false,
          "mergedRangeAddress": null
        },
        
        "richText": null,
        
        "comment": null,
        
        "hyperlink": null,
        
        "metadata": {
          "hasFormula": false,
          "isRichText": false,
          "styleId": 0,
          "styleName": "Normal",
          "rows": 1,
          "columns": 1,
          "start": { "row": 1, "column": 1, "address": "A1" },
          "end": { "row": 1, "column": 1, "address": "A1" }
        }
      }
    ]
  ]
}
```

## 🎯 可用屬性完整清單

從這個調試端點，您可以看到 EPPlus 提供的所有屬性：

### 📍 位置和範圍
- `Position.Row`, `Position.Column`, `Position.Address`
- `Metadata.Start`, `Metadata.End`
- `Metadata.Rows`, `Metadata.Columns`

### 💾 值和內容
- `Value` - 原始值
- `Text` - 顯示文字
- `Formula` - 公式 (如 "=A1+B1")
- `FormulaR1C1` - R1C1 格式的公式
- `ValueType` - 資料類型

### 🎨 格式化
- `NumberFormat.Format` - 數字格式 (如 "0.00%", "yyyy-mm-dd")
- `NumberFormat.NumberFormatId` - 格式 ID

### 🖋️ 字體樣式
- `Font.Name`, `Font.Size`
- `Font.Bold`, `Font.Italic`, `Font.Underline`, `Font.Strike`
- `Font.Color`, `Font.ColorTheme`, `Font.ColorTint`
- `Font.Charset`, `Font.Scheme`, `Font.Family`

### 📏 對齊方式
- `Alignment.Horizontal`, `Alignment.Vertical`
- `Alignment.WrapText`, `Alignment.Indent`
- `Alignment.TextRotation`, `Alignment.ShrinkToFit`
- `Alignment.ReadingOrder`

### 🔲 邊框
- `Border.Top/Bottom/Left/Right/Diagonal.Style`
- `Border.Top/Bottom/Left/Right/Diagonal.Color`
- `Border.DiagonalUp`, `Border.DiagonalDown`

### 🎨 填充/背景
- `Fill.PatternType`
- `Fill.BackgroundColor`, `Fill.PatternColor`
- `Fill.BackgroundColorTheme`, `Fill.BackgroundColorTint`

### 📐 尺寸和合併
- `Dimensions.ColumnWidth`, `Dimensions.RowHeight`
- `Dimensions.IsMerged`, `Dimensions.MergedRangeAddress`

### 🎨 Rich Text
- `RichText[]` - Rich Text 片段陣列
- 每個片段包含：`Text`, `Bold`, `Italic`, `UnderLine`, `Strike`, `Size`, `FontName`, `Color`, `VerticalAlign`

### 💬 註解和超連結
- `Comment.Text`, `Comment.Author`, `Comment.Visible`
- `Hyperlink.AbsoluteUri`, `Hyperlink.OriginalString`

### 🔧 元數據
- `Metadata.HasFormula`, `Metadata.IsRichText`
- `Metadata.StyleId`, `Metadata.StyleName`

## 🚀 使用建議

1. **先上傳簡單的 Excel 檔案**，查看基本屬性
2. **逐步測試複雜格式**，如合併儲存格、Rich Text、公式等
3. **對比不同儲存格**，了解屬性差異
4. **根據需求選擇屬性**，添加到您的主要 API 中

## 📝 下一步

基於調試結果，您可以：
1. 選擇需要的屬性添加到 `ExcelCellInfo` 模型
2. 在 `CreateCellInfo` 方法中實現新屬性的讀取
3. 在前端添加對應的顯示和樣式處理

這個調試端點讓您能完整了解 EPPlus 的所有可用屬性！