# 文字方塊檢測修復總結

## 問題分析
用戶回報："d6:f8與g6:h8裡面的文字方塊不會被讀取到"

## 根本原因
1. **效能限制過低**: `MAX_DRAWING_OBJECTS_TO_CHECK = 100` 設定太低
2. **全域計數器問題**: `_globalDrawingObjectCount` 靜態變數在所有儲存格檢查間累加，容易達到限制
3. **計數邏輯錯誤**: 計數器在達到限制後就停止所有後續檢測

## 修復內容

### 1. 增加效能限制
```csharp
// 修復前
private const int MAX_DRAWING_OBJECTS_TO_CHECK = 100;

// 修復後  
private const int MAX_DRAWING_OBJECTS_TO_CHECK = 1000; // 增加限制，支援更多文字方塊
```

### 2. 改變計數器邏輯
```csharp
// 修復前 - 全域計數器
[ThreadStatic]
private static int _globalDrawingObjectCount = 0;

// 修復後 - 按工作表計數
[ThreadStatic]
private static Dictionary<string, int>? _worksheetDrawingObjectCounts;
```

### 3. 新增管理方法
- `GetWorksheetDrawingObjectCount(string worksheetName)`: 取得工作表計數
- `IncrementWorksheetDrawingObjectCount(string worksheetName)`: 增加工作表計數  
- `ResetWorksheetDrawingObjectCounts()`: 重置所有計數器

### 4. 修改檢測邏輯
```csharp
// 修復前 - 全域限制檢查
if (_globalDrawingObjectCount > MAX_DRAWING_OBJECTS_TO_CHECK)

// 修復後 - 按工作表限制檢查  
var currentCount = GetWorksheetDrawingObjectCount(worksheet.Name);
if (currentCount > MAX_DRAWING_OBJECTS_TO_CHECK)
```

## 預期效果
1. **提高檢測上限**: 從 100 個物件增加到 1000 個
2. **按工作表計算**: 每個工作表獨立計算，避免跨工作表影響
3. **更精確的限制**: 只在單一工作表超過限制時停止，不影響其他工作表
4. **支援合併儲存格**: D6:F8 和 G6:H8 等合併儲存格內的文字方塊應該能正常檢測

## 測試建議
1. 確認 D6:F8 儲存格的文字方塊能被檢測到
2. 確認 G6:H8 儲存格的文字方塊能被檢測到  
3. 驗證其他儲存格的文字方塊檢測不受影響
4. 檢查效能沒有顯著下降

## 技術細節
- 修復涉及 `ExcelController.cs` 約 7 處程式碼更改
- 保持向後相容性
- 使用 ThreadStatic 確保執行緒安全
- 支援 EPPlus 8.x 的 Drawings 集合操作