# 🐛 Rich Text 第一個項目格式修正

## 問題描述

根據 `test.json` 的分析，發現 Rich Text 的第一個項目存在格式異常：

### 問題範例
```json
"richText": [
  {
    "text": "d",
    "bold": false,      // ❌ 應該是 true
    "italic": false,
    "underLine": false,
    "strike": false,
    "size": 0,          // ❌ 應該是 24
    "fontName": "",     // ❌ 應該是 "Arial Black"
    "color": null,
    "verticalAlign": "None"
  },
  {
    "text": "e",
    "bold": true,       // ✅ 正確
    "italic": false,
    "underLine": false,
    "strike": false,
    "size": 24,         // ✅ 正確
    "fontName": "Arial Black", // ✅ 正確
    "color": "#FF0000", // ✅ 正確
    "verticalAlign": "None"
  },
  // ... 其他項目
]
```

## 根本原因

這是 EPPlus 函式庫的一個已知問題：
- **第一個 Rich Text 片段**經常缺少格式資訊
- EPPlus 假設第一個片段繼承儲存格的預設樣式
- 但在實際輸出時沒有正確應用這些樣式

## 🔧 解決方案

### 智慧格式繼承
修改後端 Rich Text 處理邏輯，在第一個項目格式缺失時自動從儲存格樣式繼承：

```csharp
// 修正第一個 Rich Text 部分的格式問題
if (i == 0)
{
    if (size == 0 || string.IsNullOrEmpty(fontName) || (!bold && !italic))
    {
        // 從儲存格樣式繼承缺失的格式
        size = size == 0 ? cell.Style.Font.Size : size;
        fontName = string.IsNullOrEmpty(fontName) ? cell.Style.Font.Name : fontName;
        
        // 只有當 Rich Text 部分沒有設定格式時才繼承
        if (!richTextPart.Bold && cell.Style.Font.Bold)
            bold = true;
        if (!richTextPart.Italic && cell.Style.Font.Italic)
            italic = true;
    }
}
```

### 檢查條件
修正會在以下情況觸發：
1. **字體大小為 0**：`size == 0`
2. **字體名稱為空**：`string.IsNullOrEmpty(fontName)`
3. **沒有粗體或斜體**：`(!bold && !italic)`

### 繼承邏輯
- **字體大小**：從 `cell.Style.Font.Size` 繼承
- **字體名稱**：從 `cell.Style.Font.Name` 繼承  
- **粗體樣式**：如果儲存格樣式為粗體且 Rich Text 未設定，則應用粗體
- **斜體樣式**：如果儲存格樣式為斜體且 Rich Text 未設定，則應用斜體

## 📊 修正效果

### 修正前
```json
{
  "text": "d",
  "bold": false,
  "size": 0,
  "fontName": "",
  "color": null
}
```

### 修正後
```json
{
  "text": "d",
  "bold": true,         // ✅ 從儲存格樣式繼承
  "size": 24,           // ✅ 從儲存格樣式繼承
  "fontName": "Arial Black", // ✅ 從儲存格樣式繼承
  "color": null         // 保持原狀（可能沒有特定顏色）
}
```

## 🎯 影響範圍

### 直接影響
- **Rich Text 顯示**：第一個片段現在會正確顯示格式
- **視覺一致性**：整個 Rich Text 內容的格式更加統一
- **資料完整性**：避免遺失重要的格式資訊

### 相容性
- **向後相容**：不影響現有功能
- **非破壞性**：只修正缺失的格式，不改變正確的格式
- **選擇性應用**：只對第一個 Rich Text 項目進行檢查和修正

## 🧪 測試驗證

### 測試案例
1. **標準 Rich Text**：包含多個不同格式的文字片段
2. **第一項目缺失格式**：驗證是否正確繼承儲存格樣式
3. **混合格式**：確保不影響其他正確的格式項目

### 預期結果
- 第一個 Rich Text 項目顯示正確的字體、大小、粗體等格式
- 其他項目保持原有格式不變
- 整體 Rich Text 渲染效果符合 Excel 原始顯示

## 💡 技術細節

### EPPlus 行為
EPPlus 在處理 Rich Text 時：
- 假設第一個片段使用儲存格的預設樣式
- 但在 API 輸出中不會自動填入這些預設值
- 導致第一個片段看起來沒有格式

### 修正策略
- **檢測缺失**：識別格式資訊不完整的情況
- **智慧繼承**：從父級儲存格樣式補充缺失資訊
- **保持原狀**：不修改已經正確的格式資訊

這個修正確保了 Rich Text 資料的完整性和一致性，解決了您發現的重要問題！ 🎉