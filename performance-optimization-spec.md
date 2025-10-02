# Excel 儲存格處理效能優化規格書

## 📋 文件資訊
- **版本**: 1.0
- **日期**: 2025-10-02
- **狀態**: 待審核
- **專案**: ExcelReader EPPlus 7.1.0

---

## 🎯 優化目標

### 主要問題
目前系統在處理大量儲存格時效能低下，主要瓶頸在於：

1. **重複遍歷 Drawings 集合** - 每個儲存格都遍歷一次所有繪圖物件
2. **DISPIMG 函數檢查無效** - EPPlus 7.1.0 無法存取 DISPIMG 圖片,檢查邏輯無用
3. **內容類型檢測冗餘** - `DetectCellContentType` 和 `GetCellImages` 做重複檢查
4. **沒有快取機制** - 圖片位置資訊未快取,每次都重新計算

### 效能指標
- **當前**: 1000 個儲存格約需 30-60 秒
- **目標**: 1000 個儲存格應在 3-5 秒內完成
- **改善率**: 提升 10-20 倍效能

---

## 🔍 效能瓶頸分析

### 1. 重複遍歷 worksheet.Drawings (最嚴重)

#### 問題描述
```csharp
// 每個儲存格都執行以下代碼:
private CellContentType DetectCellContentType(ExcelRange cell, ExcelWorksheet worksheet)
{
    foreach (var drawing in worksheet.Drawings.Take(100)) // ← 重複 N 次
    {
        if (drawing is ExcelPicture picture)
        {
            // 檢查位置...
        }
    }
}

private List<ImageInfo>? GetCellImages(ExcelWorksheet worksheet, ExcelRange cell)
{
    foreach (var drawing in worksheet.Drawings) // ← 又重複 N 次
    {
        // 處理圖片...
    }
}
```

#### 複雜度
- **當前**: O(N × M × D)
  - N = 儲存格數量 (例: 1000)
  - M = 繪圖物件數量 (例: 50)
  - D = 檢查次數 (DetectCellContentType + GetCellImages = 2)
  - **總操作**: 1000 × 50 × 2 = **100,000 次遍歷**

- **優化後**: O(D + N)
  - D = 一次性建立索引 (50 次)
  - N = 查詢索引 (1000 次)
  - **總操作**: 50 + 1000 = **1,050 次操作** (減少 99%)

### 2. DISPIMG 函數相關代碼 (次要)

#### 無效代碼位置
```csharp
// ExcelController.cs Line 858-897
// 2. 檢查 DISPIMG 函數
if (!string.IsNullOrEmpty(cell.Formula))
{
    if (formula.Contains("DISPIMG") || formula.Contains("_xlfn.DISPIMG"))
    {
        // 提取 ID, 查找圖片...
        // ❌ EPPlus 7.1.0 無法存取 DISPIMG 圖片
    }
}
```

#### 影響
- 每個含公式的儲存格都執行字串檢查
- 呼叫 `ExtractImageIdFromFormula`, `FindEmbeddedImageById` 等無效方法
- 增加約 5-10% 處理時間

### 3. 內容類型檢測冗餘 (中等)

#### 問題流程
```
CreateCellInfo()
  └─> DetectCellContentType()  ← 遍歷 Drawings 檢查圖片
       └─> foreach worksheet.Drawings.Take(100)
  
  └─> GetCellImages()          ← 又遍歷 Drawings 取得圖片
       └─> foreach worksheet.Drawings
```

#### 重複工作
- 同一個儲存格的圖片位置檢查執行兩次
- `DetectCellContentType` 只需要知道 "有沒有圖片"
- `GetCellImages` 需要完整圖片資訊
- 可合併為一次操作

---

## 💡 優化方案

### 方案 1: 圖片位置索引快取 (核心優化)

#### 實作策略
```csharp
// 新增類別: 工作表層級的圖片索引
private class WorksheetImageIndex
{
    // Key: "Row_Column" (例: "5_3" 代表 Row=5, Col=3)
    // Value: 該儲存格的所有圖片
    public Dictionary<string, List<ExcelPicture>> CellImageMap { get; set; }
    
    // 建構時一次性遍歷所有 Drawings
    public WorksheetImageIndex(ExcelWorksheet worksheet)
    {
        CellImageMap = new Dictionary<string, List<ExcelPicture>>();
        
        foreach (var drawing in worksheet.Drawings)
        {
            if (drawing is ExcelPicture picture && picture.From != null)
            {
                int fromRow = picture.From.Row + 1;
                int fromCol = picture.From.Column + 1;
                string key = $"{fromRow}_{fromCol}";
                
                if (!CellImageMap.ContainsKey(key))
                    CellImageMap[key] = new List<ExcelPicture>();
                
                CellImageMap[key].Add(picture);
            }
        }
    }
    
    // 快速查詢
    public List<ExcelPicture>? GetImagesAtCell(int row, int col)
    {
        string key = $"{row}_{col}";
        return CellImageMap.TryGetValue(key, out var images) ? images : null;
    }
}
```

#### 使用方式
```csharp
// 在 Upload 方法開始時建立索引
var imageIndex = new WorksheetImageIndex(worksheet);

// 在 CreateCellInfo 中使用索引
private ExcelCellInfo CreateCellInfo(ExcelRange cell, ExcelWorksheet worksheet, WorksheetImageIndex imageIndex)
{
    // ...
    
    // 快速查詢圖片 (O(1) 而非 O(M))
    var pictures = imageIndex.GetImagesAtCell(cell.Start.Row, cell.Start.Column);
    if (pictures != null && pictures.Any())
    {
        cellInfo.Images = ProcessImages(pictures, worksheet, cell);
    }
    
    // ...
}
```

#### 效能提升
- **Before**: 1000 cells × 50 drawings = 50,000 次遍歷
- **After**: 50 drawings (建索引) + 1000 cells (查詢) = 1,050 次操作
- **提升**: **98% 減少**

### 方案 2: 移除 DISPIMG 相關代碼 (清理優化)

#### 要移除的方法
```csharp
// 1. ExcelController.cs Line 1592-1611
private string? ExtractImageIdFromFormula(string formula)

// 2. ExcelController.cs Line 1618-1738
private ImageInfo? FindEmbeddedImageById(ExcelWorkbook workbook, string imageId)

// 3. ExcelController.cs Line 1744-2106
private ImageInfo? ParseOOXMLForImage(ExcelWorkbook workbook, string imageId)

// 4. ExcelController.cs Line 2113-2353
private string? GeneratePlaceholderImage()

// 5. GetCellImages 方法中 Line 858-897
// 2. 檢查 DISPIMG 函數 (整段移除)
```

#### 前端相關
```vue
// ExcelReader.vue Line 110-118, 517-530
// 移除 isPlaceholderImage 相關邏輯
```

#### 效能提升
- 每個含公式的儲存格節省 5-10ms
- 減少代碼約 700 行
- 簡化維護成本

### 方案 3: 合併內容類型檢測 (中度優化)

#### 當前問題
```csharp
// Step 1: 檢測內容類型
var contentType = DetectCellContentType(cell, worksheet); 
// ↑ 遍歷 Drawings 檢查是否有圖片

// Step 2: 獲取圖片
cellInfo.Images = GetCellImages(worksheet, rangeToCheck);
// ↑ 又遍歷 Drawings 取得圖片資料
```

#### 優化方案
```csharp
// 使用索引後,兩個方法都從索引查詢
private CellContentType DetectCellContentType(ExcelRange cell, WorksheetImageIndex imageIndex)
{
    var hasText = !string.IsNullOrEmpty(cell.Text) || !string.IsNullOrEmpty(cell.Formula);
    var hasImages = imageIndex.GetImagesAtCell(cell.Start.Row, cell.Start.Column) != null;
    
    // 判斷類型...
}

private List<ImageInfo>? GetCellImages(ExcelRange cell, WorksheetImageIndex imageIndex, ExcelWorksheet worksheet)
{
    var pictures = imageIndex.GetImagesAtCell(cell.Start.Row, cell.Start.Column);
    if (pictures == null) return null;
    
    // 處理圖片資料...
}
```

#### 效能提升
- 從 O(M) 降為 O(1) 查詢
- 消除重複遍歷

### 方案 4: 延遲載入圖片資料 (進階優化)

#### 策略
```csharp
// 第一階段: 只記錄圖片位置
cellInfo.Images = pictures.Select(p => new ImageInfo 
{
    Name = p.Name,
    // 不立即計算 Base64Data, Width, Height
    _lazyPicture = p // 保存引用
}).ToList();

// 第二階段: 前端按需載入
[HttpGet("image/{imageName}")]
public ActionResult GetImageData(string imageName)
{
    // 單獨 API 載入圖片資料
}
```

#### 效益
- 初次載入快 50-70%
- 圖片按需載入
- 減少記憶體佔用

---

## 📊 優化實施計畫

### Phase 1: 核心優化 (必須)
**預計時間**: 2-3 小時
**預期效果**: 10-15倍效能提升

#### 任務清單
- [x] 創建 `WorksheetImageIndex` 類別
- [ ] 修改 `Upload` 方法建立索引
- [ ] 修改 `DetectCellContentType` 使用索引
- [ ] 修改 `GetCellImages` 使用索引
- [ ] 修改 `CreateCellInfo` 傳遞索引參數
- [ ] 測試大檔案 (1000+ 儲存格)

#### 成功指標
- ✅ 1000 儲存格處理時間 < 5 秒
- ✅ 不遺漏任何圖片
- ✅ 向下相容現有功能

### Phase 2: 清理優化 (建議)
**預計時間**: 1-2 小時
**預期效果**: 5-10% 效能提升 + 代碼清理

#### 任務清單
- [ ] 移除 `ExtractImageIdFromFormula` 方法
- [ ] 移除 `FindEmbeddedImageById` 方法
- [ ] 移除 `ParseOOXMLForImage` 方法
- [ ] 移除 `GeneratePlaceholderImage` 方法
- [ ] 移除 `GetCellImages` 中的 DISPIMG 檢查邏輯
- [ ] 移除前端 `isPlaceholderImage` 相關代碼
- [ ] 更新文檔

#### 成功指標
- ✅ 移除約 700 行無效代碼
- ✅ 所有測試通過
- ✅ 無 DISPIMG 殘留提示

### Phase 3: 進階優化 (可選)
**預計時間**: 3-4 小時
**預期效果**: 額外 2-3倍提升 (大圖片場景)

#### 任務清單
- [ ] 實作圖片延遲載入 API
- [ ] 修改前端支援按需載入
- [ ] 添加圖片快取機制
- [ ] 實作分頁載入 (大工作表)

---

## 🧪 測試計畫

### 效能測試案例

#### 測試 1: 小檔案 (基準線)
- **內容**: 100 儲存格, 5 張圖片
- **當前時間**: ~3 秒
- **目標時間**: < 1 秒

#### 測試 2: 中檔案
- **內容**: 500 儲存格, 20 張圖片
- **當前時間**: ~15 秒
- **目標時間**: < 2 秒

#### 測試 3: 大檔案
- **內容**: 1000 儲存格, 50 張圖片
- **當前時間**: ~40 秒
- **目標時間**: < 5 秒

#### 測試 4: 無圖片檔案
- **內容**: 1000 儲存格, 0 張圖片
- **當前時間**: ~10 秒
- **目標時間**: < 2 秒

### 功能測試案例

#### 測試 A: 圖片完整性
- ✅ 所有圖片都被正確識別
- ✅ 圖片位置資訊正確
- ✅ 圖片尺寸計算正確
- ✅ Base64 資料完整

#### 測試 B: 邊界情況
- ✅ 工作表無圖片
- ✅ 儲存格無圖片
- ✅ 圖片跨多個儲存格
- ✅ 合併儲存格包含圖片

#### 測試 C: 向下相容
- ✅ 現有 API 回應格式不變
- ✅ 前端無需修改
- ✅ 所有現有功能正常

---

## 📝 程式碼變更摘要

### 新增類別
```csharp
// Controllers/ExcelController.cs
private class WorksheetImageIndex
{
    public Dictionary<string, List<ExcelPicture>> CellImageMap { get; set; }
    public WorksheetImageIndex(ExcelWorksheet worksheet) { /* ... */ }
    public List<ExcelPicture>? GetImagesAtCell(int row, int col) { /* ... */ }
}
```

### 方法簽名變更
```csharp
// Before
private ExcelCellInfo CreateCellInfo(ExcelRange cell, ExcelWorksheet worksheet)

// After
private ExcelCellInfo CreateCellInfo(ExcelRange cell, ExcelWorksheet worksheet, WorksheetImageIndex imageIndex)

// Before
private CellContentType DetectCellContentType(ExcelRange cell, ExcelWorksheet worksheet)

// After
private CellContentType DetectCellContentType(ExcelRange cell, WorksheetImageIndex imageIndex)

// Before
private List<ImageInfo>? GetCellImages(ExcelWorksheet worksheet, ExcelRange cell)

// After
private List<ImageInfo>? GetCellImages(ExcelRange cell, WorksheetImageIndex imageIndex, ExcelWorksheet worksheet)
```

### 移除方法 (Phase 2)
- `ExtractImageIdFromFormula` (Line 1592-1611)
- `FindEmbeddedImageById` (Line 1618-1738)
- `ParseOOXMLForImage` (Line 1744-2106)
- `GeneratePlaceholderImage` (Line 2113-2353)

### 移除代碼區塊
- `GetCellImages` 中的 DISPIMG 檢查 (Line 858-897)

---

## ⚠️ 風險評估

### 高風險
❌ **無**

### 中風險
⚠️ **索引建立失敗**
- **風險**: 某些特殊 Drawing 物件導致索引建立失敗
- **緩解**: Try-catch 保護, 失敗時回退到舊邏輯
- **機率**: 低 (5%)

### 低風險
⚠️ **記憶體增加**
- **風險**: 索引增加記憶體佔用
- **影響**: 每個圖片約 200 bytes, 100 張圖片 = 20KB
- **緩解**: 索引在請求結束後釋放
- **機率**: 可接受

---

## 📈 預期效益

### 效能提升
| 場景 | 當前時間 | 優化後 | 提升倍數 |
|------|---------|--------|---------|
| 100 儲存格 + 5 圖 | 3s | 0.3s | 10x |
| 500 儲存格 + 20 圖 | 15s | 1.5s | 10x |
| 1000 儲存格 + 50 圖 | 40s | 3s | 13x |
| 1000 儲存格 + 0 圖 | 10s | 1s | 10x |

### 代碼品質
- ✅ 移除 700 行無效代碼
- ✅ 降低維護成本
- ✅ 提升可讀性
- ✅ 減少 bug 風險

### 用戶體驗
- ✅ 大檔案處理不再卡頓
- ✅ 回應時間大幅縮短
- ✅ 支援更大的 Excel 檔案

---

## ✅ 審核檢查清單

### 技術審核
- [ ] 優化方案技術可行
- [ ] 效能提升預期合理
- [ ] 風險評估完整
- [ ] 測試計畫充分

### 業務審核
- [ ] 優先級排序正確
- [ ] 時間估算合理
- [ ] 資源分配充足
- [ ] 預期效益明確

### 決策
- [ ] **批准實施 Phase 1 (核心優化)**
- [ ] **批准實施 Phase 2 (清理優化)**
- [ ] 考慮實施 Phase 3 (進階優化)
- [ ] 需要修改規格

---

## 📌 附錄

### A. 相關文件
- `dispimg-improvement-report.md` - DISPIMG 限制說明
- `DISPIMG-Solutions-Report.md` - DISPIMG 解決方案
- `image-detection-consistency-fix.md` - 圖片檢測一致性修復

### B. 效能分析工具
```csharp
// 可添加效能監控
var stopwatch = System.Diagnostics.Stopwatch.StartNew();
// ... 操作 ...
stopwatch.Stop();
_logger.LogInformation($"操作耗時: {stopwatch.ElapsedMilliseconds}ms");
```

### C. 聯絡資訊
- **負責人**: [待填寫]
- **審核人**: [待填寫]
- **預計開始**: 2025-10-02
- **預計完成**: 2025-10-04

---

**文件狀態**: ✅ 待審核  
**下一步**: 技術主管審核並批准實施計畫
