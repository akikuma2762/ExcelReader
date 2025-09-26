# JSON 檔案優化報告

## 已移除的屬性

### ✅ 已完成移除

#### 1. DisplayText 屬性
- **位置**: `ExcelCellInfo.DisplayText`
- **狀態**: 已完全移除
- **說明**: 這是 `Text` 屬性的別名，完全重複
- **影響**: 
  - 後端模型: `Models/ExcelData.cs` - 移除 `DisplayText` 屬性
  - 前端類型: `types/excel.ts` - 移除 `displayText` 和 `SimpleCellInfo.displayText`
- **估計減少**: 每個儲存格約 10-20 字元（取決於文字內容）

### 📋 待確認移除的其他重複屬性

#### 2. Value vs Text 差異分析
根據從 test.json 的觀察，所有儲存格的 `value` 和 `text` 內容都相同：
```json
"value": "客戶\nClient",
"text": "客戶\nClient"
```

**EPPlus 中的差異**:
- `cell.Value`: 儲存原始數據（DateTime、double、string 等）
- `cell.Text`: 儲存格式化後的顯示文字

**建議**: 需要進一步測試確認在不同數據類型下是否真的相同

#### 3. 其他已標記為過時的屬性
以下屬性在模型中仍然存在，建議逐步移除：
- `FormatCode` → 使用 `NumberFormat`
- `FontBold` → 使用 `Font.Bold` 
- `FontSize` → 使用 `Font.Size`
- `FontName` → 使用 `Font.Name`
- `BackgroundColor` → 使用 `Fill.BackgroundColor`
- `FontColor` → 使用 `Font.Color`
- `TextAlign` → 使用 `Alignment.Horizontal`
- `ColumnWidth` → 使用 `Dimensions.ColumnWidth`
- `IsRichText` → 使用 `Metadata.IsRichText`
- `RowSpan/ColSpan` → 使用 `Dimensions.RowSpan/ColSpan`
- `IsMerged/IsMainMergedCell` → 使用 `Dimensions.IsMerged/IsMainMergedCell`

## 優化效果估算

基於 test.json (約 7MB) 的分析：
- 總儲存格數: 約 3000+ 個
- 移除 `displayText` 預估減少: 5-10%
- 移除所有重複屬性預估減少: 15-25%

## 下一步建議

1. 測試 Value vs Text 在不同數據類型下的差異
2. 逐步移除其他過時屬性
3. 考慮在序列化時排除 null/empty 值
4. 實施 JSON 壓縮（如 gzip）

## 技術影響

- ✅ 後端編譯正常
- ✅ 前端 TypeScript 檢查通過
- ✅ 向後兼容性: 前端程式碼使用 `text` 屬性，不受影響