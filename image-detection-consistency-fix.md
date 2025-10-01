# 圖片檢測一致性修正報告

## 問題描述

在處理 Excel 檔案時，發現某些儲存格的 `DataType` 被標記為 `"Image"`，但 `Images` 欄位卻沒有實際的圖片資料（為 `null` 或空陣列）。

### 測試案例
- **檔案**：`QF-VQ-82203 鍊式刀庫品檢表 (2).xlsx`
- **問題儲存格**：B5-M5（合併儲存格）、J9
- **現象**：
  - `DataType` = `"Image"` ✅
  - `Images` = `null` 或 `[]` ❌

## 根本原因分析

系統中有兩個不同的圖片檢測方法，使用了**不一致的位置匹配邏輯**：

### 1. DetectCellContentType 方法（內容類型檢測）
```csharp
// 舊邏輯：使用寬鬆匹配（± 1 容錯範圍）
if (fromRow >= cellStartRow - 1 && fromRow <= cellEndRow + 1 &&
    fromCol >= cellStartCol - 1 && fromCol <= cellEndCol + 1)
{
    hasImages = true;  // 檢測到有圖片
    break;
}
```
- **用途**：快速判斷儲存格是否包含圖片
- **結果**：設置 `DataType = "Image"`
- **匹配策略**：寬鬆（± 1 行列容錯）

### 2. GetCellImages 方法（圖片資料獲取）
```csharp
// 新邏輯：使用精確匹配
bool shouldInclude = (fromRow >= cellStartRow && fromRow <= cellEndRow &&
                     fromCol >= cellStartCol && fromCol <= cellEndCol);
```
- **用途**：實際獲取圖片資料並填充到 `Images` 欄位
- **匹配策略**：精確（圖片起始點必須在儲存格範圍內）

### 不一致的後果

當圖片的位置在邊界附近時：
1. `DetectCellContentType` 使用寬鬆匹配，認為儲存格有圖片 → 設置 `DataType = "Image"`
2. `GetCellImages` 使用精確匹配，找不到圖片 → 返回 `null`
3. 結果：前端收到的資料顯示有圖片但沒有實際圖片資料

## 解決方案

統一兩個方法的位置檢查邏輯，**都使用精確匹配**：

### 修正內容

#### 檔案：`ExcelController.cs`
#### 方法：`DetectCellContentType`

**修改前**（第 195-211 行）：
```csharp
foreach (var drawing in worksheet.Drawings.Take(5)) // 只檢查前5個物件
{
    if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
    {
        if (picture.From != null)
        {
            var fromRow = picture.From.Row + 1;
            var fromCol = picture.From.Column + 1;
            
            // 寬鬆的位置檢查（± 1 容錯）
            if (fromRow >= cellStartRow - 1 && fromRow <= cellEndRow + 1 &&
                fromCol >= cellStartCol - 1 && fromCol <= cellEndCol + 1)
            {
                hasImages = true;
                break;
            }
        }
    }
}
```

**修改後**：
```csharp
foreach (var drawing in worksheet.Drawings.Take(100)) // 檢查更多物件以確保不會遺漏
{
    if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
    {
        if (picture.From != null)
        {
            var fromRow = picture.From.Row + 1;
            var fromCol = picture.From.Column + 1;
            
            // 精確的位置檢查（與 GetCellImages 一致）
            if (fromRow >= cellStartRow && fromRow <= cellEndRow &&
                fromCol >= cellStartCol && fromCol <= cellEndCol)
            {
                hasImages = true;
                break;
            }
        }
    }
}
```

### 關鍵改進

1. **統一匹配邏輯**：
   - 移除 `± 1` 容錯範圍
   - 使用與 `GetCellImages` 相同的精確匹配條件

2. **增加檢查數量**：
   - 從 `Take(5)` 改為 `Take(100)`
   - 確保不會因為檢查物件數量太少而遺漏圖片

3. **添加說明註解**：
   - 明確標註「與 GetCellImages 一致」
   - 便於未來維護時理解設計意圖

## 效果驗證

修正後的行為：

### 情況 1：圖片在儲存格範圍內
- `DetectCellContentType` → 檢測到圖片 ✅
- `GetCellImages` → 返回圖片資料 ✅
- **結果**：`DataType = "Image"` + `Images = [圖片資料]` ✅

### 情況 2：圖片不在儲存格範圍內
- `DetectCellContentType` → 未檢測到圖片 ✅
- `GetCellImages` → 未返回圖片資料 ✅
- **結果**：`DataType = "Empty"` + `Images = null` ✅

### 情況 3：合併儲存格（如 B5-M5）
- 圖片起始點在 B5-M5 範圍內 → 兩個方法都會正確檢測 ✅
- 圖片起始點在範圍外 → 兩個方法都不會檢測 ✅

## 技術細節

### 位置計算說明
```csharp
var fromRow = picture.From.Row + 1;  // EPPlus 使用 0-based，Excel 使用 1-based
var fromCol = picture.From.Column + 1;

// 精確匹配條件
fromRow >= cellStartRow && fromRow <= cellEndRow  // 行必須在範圍內
&&
fromCol >= cellStartCol && fromCol <= cellEndCol  // 列必須在範圍內
```

### 合併儲存格處理
對於合併儲存格（如 B5-M5）：
- `cellStartRow = 5, cellEndRow = 5`（只有一行）
- `cellStartCol = 2, cellEndCol = 13`（B=2 到 M=13）
- 只有當圖片的起始點在這個範圍內時才會被包含

## 相關文件

- **前次修正**：`測試資料.xlsx` 圖片位置匹配問題
  - 修正了 `GetCellImages` 方法的寬鬆匹配（± 5）
  - 本次修正確保 `DetectCellContentType` 與其保持一致

## 總結

通過統一兩個方法的位置檢查邏輯，解決了 `DataType` 和 `Images` 資料不一致的問題。現在：

✅ 檢測邏輯統一且精確
✅ 前端接收的資料保持一致
✅ 避免了誤判和資料缺失
✅ 程式碼更容易維護和理解

---
**修正日期**：2025-10-01  
**影響範圍**：`ExcelController.cs` - `DetectCellContentType` 方法  
**測試狀態**：已通過建置 ✅
