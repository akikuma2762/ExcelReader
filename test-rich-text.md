# Excel Rich Text 格式測試

## 問題分析

從附件文件比較可以看出：

### HTML文件 (QF-VQ-82101 (20241201 發行).html (12).json) 包含的信息：
```json
{
  "value": "deta International德大機械股份有限公司",
  "originalHTML": "d<span style=\"font-size:24pt;font-family:\"Arial Black\",Arial;font-weight:bold;color:#ff0000;\">e</span><span style=\"font-size:24pt;font-family:\"Arial Black\",Arial;font-weight:bold;\">ta</span><span style=\"font-size:24pt;font-family:\"Times New Roman\";font-weight:bold;\"> </span><span style=\"font-size:14pt;font-family:\"Times New Roman\";font-weight:bold;font-style:italic;\">International</span><span style=\"font-size:24pt;font-family:Arial;font-weight:bold;\">德大機械股份有限公司</span>"
}
```

這顯示了非常詳細的格式信息：
- "d": 普通文字
- "e": 紅色 (#ff0000) Arial Black 字體, 24pt, 粗體
- "ta": Arial Black 字體, 24pt, 粗體  
- " ": Times New Roman 字體, 24pt, 粗體
- "International": Times New Roman 字體, 14pt, 粗體, 斜體
- "德大機械股份有限公司": Arial 字體, 24pt, 粗體

### 我們的 EPPlus 輸出 (test.json):
```json
{
  "value": "deta International德大機械股份有限公司",
  "displayText": "deta International德大機械股份有限公司", 
  "formatCode": "General",
  "fontBold": true,
  "fontSize": 24,
  "fontName": "Arial Black",
  "backgroundColor": "",
  "fontColor": null
}
```

## 改進方案

我已經更新了代碼來支援 EPPlus 的 Rich Text 功能：

1. **後端 (C#)**:
   - 新增 `RichTextPart` 類別來儲存每個文字片段的格式
   - 更新 `ExcelCellInfo` 加入 `RichText` 和 `IsRichText` 屬性  
   - 修改 `CreateCellInfo` 方法來檢測並解析 Rich Text

2. **前端 (Vue)**:
   - 更新 TypeScript 接口來匹配後端結構
   - 新增 `renderRichText` 方法來將 Rich Text 轉換為 HTML
   - 更新表格渲染來支援 Rich Text 顯示

## 測試計劃

接下來需要：
1. 重新上傳包含 Rich Text 的 Excel 檔案
2. 確認 EPPlus 能正確檢測 `cell.IsRichText`
3. 驗證 Rich Text 格式是否被正確解析和顯示