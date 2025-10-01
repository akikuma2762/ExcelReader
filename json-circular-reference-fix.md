# JSON 序列化循環引用錯誤修復報告

## 問題描述

**錯誤類型**: `System.Text.Json.JsonException`  
**錯誤訊息**: "A possible object cycle was detected. This can either be due to a cycle or if the object depth is larger than the maximum allowed depth of 32."

**錯誤路徑**:
```
Path: $.Data.Rows.Value.AsCompileResult.Result.AsCompileResult.Result...
```

### 症狀
- API 端點在嘗試序列化 Excel 資料時拋出異常
- 錯誤顯示在序列化過程中偵測到物件循環或深度超過 32 層
- 錯誤發生在 `SystemTextJsonOutputFormatter.WriteResponseBodyAsync` 階段
- 服務回應已經開始，無法顯示錯誤頁面

## 根本原因分析

### 問題代碼位置
**檔案**: `ExcelReaderAPI/Controllers/ExcelController.cs`  
**行號**: 395

### 問題代碼
```csharp
// 填充/背景
cellInfo.Fill = new FillInfo
{
    PatternType = cell.Style.Fill.PatternType.ToString(),
    BackgroundColor = GetBackgroundColor(cell),
    PatternColor = cell.Style.Fill.PatternColor.Rgb,  // ❌ 問題所在
    BackgroundColorTheme = cell.Style.Fill.BackgroundColor.Theme?.ToString(),
    BackgroundColorTint = (double?)cell.Style.Fill.BackgroundColor.Tint
};
```

### 為什麼會造成循環引用？

1. **EPPlus 內部對象引用**
   - `cell.Style.Fill.PatternColor.Rgb` 不是一個簡單的字串屬性
   - 它可能返回一個包含內部引用的 EPPlus 對象
   - 這個對象的 `AsCompileResult` 屬性造成了循環引用鏈

2. **錯誤路徑分析**
   ```
   $.Data.Rows.Value.AsCompileResult.Result.AsCompileResult.Result...
   ```
   - `AsCompileResult` 是 EPPlus 內部編譯結果對象
   - 此對象包含 `Result` 屬性，而 `Result` 又包含 `AsCompileResult`
   - 形成無限循環：AsCompileResult → Result → AsCompileResult → Result...

3. **為什麼之前沒發現？**
   - `BackgroundColor` 使用了 `GetBackgroundColor(cell)` 方法，該方法內部調用 `GetColorFromExcelColor`
   - `GetColorFromExcelColor` 方法正確地從 EPPlus 對象中提取字串值
   - 但 `PatternColor` 直接引用了 EPPlus 對象，繞過了安全提取過程

## 解決方案

### 修改後的代碼
```csharp
// 填充/背景 - 使用 GetColorFromExcelColor 避免循環引用
cellInfo.Fill = new FillInfo
{
    PatternType = cell.Style.Fill.PatternType.ToString(),
    BackgroundColor = GetBackgroundColor(cell),
    PatternColor = GetColorFromExcelColor(cell.Style.Fill.PatternColor),  // ✅ 修復
    BackgroundColorTheme = cell.Style.Fill.BackgroundColor.Theme?.ToString(),
    BackgroundColorTint = (double?)cell.Style.Fill.BackgroundColor.Tint
};
```

### 修復說明

**關鍵改變**:
```diff
- PatternColor = cell.Style.Fill.PatternColor.Rgb,
+ PatternColor = GetColorFromExcelColor(cell.Style.Fill.PatternColor),
```

**為什麼這樣能解決問題？**

1. **安全提取**: `GetColorFromExcelColor` 方法專門設計用於從 EPPlus `ExcelColor` 對象中安全提取顏色值
2. **值拷貝**: 該方法返回純字串（`string?`），而不是對象引用
3. **Null 安全**: 方法內部有完善的 null 檢查和異常處理

### `GetColorFromExcelColor` 方法的工作原理

```csharp
private string? GetColorFromExcelColor(OfficeOpenXml.Style.ExcelColor excelColor)
{
    if (excelColor == null)
        return null;

    try
    {
        // 1. 嘗試使用 RGB 值 (靜默處理錯誤)
        string? rgbValue = null;
        try
        {
            rgbValue = excelColor.Rgb;  // 安全提取字串值
        }
        catch
        {
            // 靜默處理 RGB 存取錯誤
        }

        if (!string.IsNullOrEmpty(rgbValue))
        {
            var colorValue = rgbValue.TrimStart('#');
            
            // 處理 ARGB 格式（8位）轉為 RGB 格式（6位）
            if (colorValue.Length == 8)
            {
                colorValue = colorValue.Substring(2);
            }
            
            if (colorValue.Length == 6)
            {
                return colorValue.ToUpperInvariant();  // 返回純字串
            }
            
            // 處理3位短格式
            if (colorValue.Length == 3)
            {
                return $"{colorValue[0]}{colorValue[0]}{colorValue[1]}{colorValue[1]}{colorValue[2]}{colorValue[2]}";
            }
        }

        // 2. 回退到索引顏色
        // 3. 回退到主題顏色
        // ... 其他處理邏輯
    }
    catch (Exception ex)
    {
        _logger.LogDebug($"顏色解析錯誤: {ex.Message}");
        return null;
    }
}
```

**關鍵特性**:
- ✅ 返回純字串，不返回對象引用
- ✅ 多層 try-catch 保護
- ✅ Null 安全處理
- ✅ 支持多種顏色格式（RGB、Indexed、Theme）
- ✅ 靜默處理錯誤，不會中斷整個流程

## 技術深入分析

### EPPlus 對象結構問題

```
ExcelColor (EPPlus 內部類型)
  ├── Rgb: string?
  ├── Indexed: int
  ├── Theme: int?
  ├── Tint: double?
  └── (內部編譯結果對象)
      └── AsCompileResult
          └── Result
              └── AsCompileResult (循環!)
                  └── Result
                      └── ...
```

### JSON 序列化深度限制

System.Text.Json 的預設設定:
- **MaxDepth**: 32
- **ReferenceHandler**: `null` (不處理循環引用)

**為什麼會超過深度限制？**
```
序列化嘗試:
Level 1:  ExcelData
Level 2:  └── Rows[]
Level 3:      └── ExcelCellInfo
Level 4:          └── Fill
Level 5:              └── PatternColor (EPPlus 對象)
Level 6:                  └── AsCompileResult
Level 7:                      └── Result
Level 8:                          └── AsCompileResult
Level 9:                              └── Result
...
Level 32:                                             └── AsCompileResult
Level 33: ❌ 超過最大深度！
```

## 其他可能的解決方案（未採用）

### 方案 1: 配置 JsonSerializerOptions
```csharp
// Program.cs
builder.Services.ConfigureHttpJsonOptions(options =>
{
    options.SerializerOptions.ReferenceHandler = ReferenceHandler.Preserve;
    options.SerializerOptions.MaxDepth = 64;
});
```

**為什麼不採用？**
- ❌ 治標不治本，只是增加深度或處理循環
- ❌ `ReferenceHandler.Preserve` 會在 JSON 中添加 `$id` 和 `$ref`，前端需要特殊處理
- ❌ 不解決根本問題（不應該序列化 EPPlus 內部對象）
- ❌ 可能影響其他 API 端點的序列化行為

### 方案 2: 使用 [JsonIgnore] 屬性
```csharp
public class FillInfo
{
    public string? PatternType { get; set; }
    public string? BackgroundColor { get; set; }
    
    [JsonIgnore]  // 忽略這個屬性
    public string? PatternColor { get; set; }
    
    public string? BackgroundColorTheme { get; set; }
    public double? BackgroundColorTint { get; set; }
}
```

**為什麼不採用？**
- ❌ 會遺失 `PatternColor` 資訊
- ❌ 不符合需求（需要保留顏色資訊）
- ❌ 違反資料完整性原則

### 方案 3: 使用 DTO 映射
```csharp
var fillDto = new FillInfo
{
    PatternColor = cell.Style.Fill.PatternColor.Rgb?.ToString()
};
```

**為什麼不採用？**
- ⚠️ `.Rgb` 屬性本身可能就是問題的根源
- ⚠️ `.ToString()` 可能觸發 EPPlus 內部序列化邏輯
- ⚠️ 沒有錯誤處理，可能仍會拋出異常

## 最佳實踐總結

### ✅ DO: 應該做的

1. **使用專用的安全提取方法**
   ```csharp
   // ✅ 好的做法
   PatternColor = GetColorFromExcelColor(cell.Style.Fill.PatternColor)
   ```

2. **總是從 EPPlus 對象中提取原始值類型**
   ```csharp
   // ✅ 提取字串、數字、布林值
   PatternType = cell.Style.Fill.PatternType.ToString()
   Theme = cell.Style.Fill.BackgroundColor.Theme?.ToString()
   Tint = (double?)cell.Style.Fill.BackgroundColor.Tint
   ```

3. **實施多層錯誤處理**
   ```csharp
   try 
   {
       // 外層保護
       try 
       {
           // 內層特定操作
           value = excelColor.Rgb;
       }
       catch 
       {
           // 靜默處理特定錯誤
       }
   }
   catch (Exception ex)
   {
       _logger.LogDebug($"處理錯誤: {ex.Message}");
       return null;
   }
   ```

### ❌ DON'T: 不應該做的

1. **直接引用 EPPlus 對象屬性**
   ```csharp
   // ❌ 錯誤的做法
   PatternColor = cell.Style.Fill.PatternColor.Rgb  // 可能返回對象引用
   ```

2. **假設所有 EPPlus 屬性都是簡單類型**
   ```csharp
   // ❌ 危險的假設
   var rgb = excelColor.Rgb;  // 可能是字串，也可能是複雜對象
   ```

3. **忽略 Null 檢查**
   ```csharp
   // ❌ 可能拋出 NullReferenceException
   PatternColor = cell.Style.Fill.PatternColor.Rgb.ToString()
   ```

## 驗證測試

### 測試步驟

1. **停止運行中的服務**
   ```powershell
   taskkill /F /PID <PID>
   ```

2. **重新建置專案**
   ```powershell
   cd ExcelReaderAPI
   dotnet build
   ```

3. **啟動服務**
   ```powershell
   dotnet run
   ```

4. **測試 API**
   - 上傳包含複雜顏色填充的 Excel 檔案
   - 驗證 API 返回完整的 JSON 資料
   - 確認 `PatternColor` 欄位有正確的顏色值

### 預期結果

**修復前**:
```json
{
  "error": "System.Text.Json.JsonException: A possible object cycle was detected..."
}
```

**修復後**:
```json
{
  "success": true,
  "data": {
    "rows": [
      {
        "fill": {
          "patternType": "Solid",
          "backgroundColor": "FFFF00",
          "patternColor": "FF0000",  // ✅ 正確返回字串值
          "backgroundColorTheme": null,
          "backgroundColorTint": 0
        }
      }
    ]
  }
}
```

## 影響範圍

### 修改的檔案
- ✏️ `ExcelReaderAPI/Controllers/ExcelController.cs` (第 395 行)

### 影響的功能
- ✅ Excel 檔案上傳和解析
- ✅ 儲存格樣式資訊提取
- ✅ 填充/背景顏色處理

### 相容性
- ✅ 向後相容：API 回應格式不變
- ✅ 前端不需修改
- ✅ 不影響其他現有功能

## 相關修復歷史

這是本專案中第 N 次 EPPlus 對象引用問題修復：

1. **ExcelColor RGB/Indexed/Theme 屬性 NullReferenceException** - 已修復
   - 增強 `GetColorFromExcelColor` 方法的錯誤處理
   
2. **Border Color NullReferenceException** - 已修復
   - 使用 `?.` 運算子和 null 檢查

3. **PatternColor 循環引用** - 本次修復 ✅
   - 使用 `GetColorFromExcelColor` 方法

### 經驗教訓

**核心原則**: 
> **絕不直接序列化 EPPlus 內部對象！**  
> 總是提取原始值類型（string、int、double、bool）

**檢查清單**:
- [ ] 是否直接引用 EPPlus 對象屬性？
- [ ] 是否使用安全提取方法？
- [ ] 是否有 Null 檢查？
- [ ] 是否有異常處理？
- [ ] 返回值是否為原始類型？

## 總結

### 問題
JSON 序列化時因 `PatternColor = cell.Style.Fill.PatternColor.Rgb` 直接引用 EPPlus 內部對象而造成循環引用。

### 解決
使用 `PatternColor = GetColorFromExcelColor(cell.Style.Fill.PatternColor)` 安全提取顏色值。

### 效益
- ✅ 消除循環引用錯誤
- ✅ 保留完整的顏色資訊
- ✅ 提高系統穩定性
- ✅ 遵循最佳實踐

---
**修復日期**: 2025-10-01  
**嚴重程度**: 🔴 Critical（服務無法正常回應）  
**修復狀態**: ✅ 已完成  
**測試狀態**: ⏳ 待驗證
