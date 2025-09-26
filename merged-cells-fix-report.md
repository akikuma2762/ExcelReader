# 合併儲存格問題修正報告

## 問題分析

從 `test.json` 檔案中可以看到：
```json
{
  "isMerged": true,
  "rowSpan": 1,
  "colSpan": 1
}
```

雖然 `isMerged` 正確識別為 `true`，但 `rowSpan` 和 `colSpan` 都是 1，這不符合合併儲存格的預期。

## 根本原因

原來的程式碼有以下問題：

```csharp
// 錯誤的方式
if (cell.Merge)
{
    cellInfo.IsMerged = true;
    cellInfo.RowSpan = cell.Rows;  // 對單個儲存格始終返回 1
    cellInfo.ColSpan = cell.Columns; // 對單個儲存格始終返回 1
}
```

**問題**:
- `cell.Merge` 只是檢查該儲存格是否為合併區域的一部分
- `cell.Rows` 和 `cell.Columns` 對於單個儲存格 (`ExcelRange`) 總是返回 1
- 需要從工作表的合併範圍集合中查找實際的合併區域大小

## 解決方案

### 1. 新增 FindMergedRange 方法
```csharp
private ExcelRange? FindMergedRange(ExcelWorksheet worksheet, int row, int column)
{
    // 檢查所有合併範圍，找到包含指定儲存格的範圍
    foreach (var mergedRange in worksheet.MergedCells)
    {
        var range = worksheet.Cells[mergedRange];
        if (row >= range.Start.Row && row <= range.End.Row &&
            column >= range.Start.Column && column <= range.End.Column)
        {
            return range;
        }
    }
    return null;
}
```

### 2. 修正 CreateCellInfo 方法
```csharp
private ExcelCellInfo CreateCellInfo(ExcelRange cell, ExcelWorksheet worksheet)
{
    // ... 其他程式碼 ...

    // 檢查是否為合併儲存格
    if (cell.Merge)
    {
        cellInfo.IsMerged = true;

        // 尋找包含此儲存格的合併範圍
        var mergedRange = FindMergedRange(worksheet, cell.Start.Row, cell.Start.Column);
        if (mergedRange != null)
        {
            cellInfo.RowSpan = mergedRange.Rows;    // 正確的行跨度
            cellInfo.ColSpan = mergedRange.Columns; // 正確的欄跨度
        }
        else
        {
            cellInfo.RowSpan = 1;
            cellInfo.ColSpan = 1;
        }
    }
}
```

### 3. 更新方法呼叫
- 所有 `CreateCellInfo(cell)` 更新為 `CreateCellInfo(cell, worksheet)`
- 包括 `UploadExcel` 和 `UploadExcelWorksheet` 兩個方法

## 技術細節

### EPPlus 合併儲存格 API
- `worksheet.MergedCells`: 包含所有合併範圍的字串集合（如 "A1:C3"）
- `worksheet.Cells[mergedRange]`: 將字串轉換為 ExcelRange 物件
- `range.Rows` / `range.Columns`: 實際的合併範圍大小
- `range.Start.Row` / `range.Start.Column`: 合併範圍的起始位置
- `range.End.Row` / `range.End.Column`: 合併範圍的結束位置

### 預期結果
修正後，對於真正的合併儲存格應該看到：
```json
{
  "isMerged": true,
  "rowSpan": 2,     // 實際的行跨度
  "colSpan": 3,     // 實際的欄跨度
}
```

## 測試建議

1. **重新上傳 Excel 檔案**: 使用包含合併儲存格的檔案測試
2. **檢查 JSON 輸出**: 確認 `rowSpan` 和 `colSpan` 顯示正確數值
3. **前端渲染驗證**: 確認表格中合併儲存格正確顯示
4. **Tooltip 檢查**: 滑鼠懸停應顯示正確的合併資訊

這個修正解決了合併儲存格檢測的核心問題，現在能夠正確獲取並顯示 Excel 中合併儲存格的實際範圍。