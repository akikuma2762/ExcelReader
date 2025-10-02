# Phase 2: 移除 DISPIMG 相關代碼

## 🎯 目標
移除所有 WPS 專用的 DISPIMG 函數處理代碼,這些代碼已經確認無法正常工作且不再需要。

## 📋 需要移除的內容

### 1. GetCellImages 方法中的 DISPIMG 檢查區塊
- **位置**: Lines 1470-1512
- **說明**: 移除檢查和處理 DISPIMG 公式的整個 if 區塊
- **狀態**: ⏳ 待執行

### 2. ExtractImageIdFromFormula 方法
- **位置**: Lines 2206-2228  
- **說明**: 從 DISPIMG 公式中提取圖片 ID 的方法
- **狀態**: ⏳ 待執行

### 3. FindEmbeddedImageById 方法
- **位置**: Lines 2230-2298
- **說明**: 根據 ID 查找嵌入圖片的方法
- **狀態**: ⏳ 待執行

### 4. TryAdvancedImageSearch 及相關方法
- **位置**: Lines 2300-2980
- **包含方法**:
  - TryAdvancedImageSearch
  - TryDirectOoxmlImageSearch
  - DeepSearchWorksheetInternals
  - TryReflectionBasedImageSearch
  - TryImageCacheSearch
  - ExtractHiddenImageData
  - SearchObjectForImages
  - SearchHiddenSheets
  - TryGenerateImageFromId
  - CreateImageFromBase64
  - IsBase64String
  - TryFindImageInWorksheets
  - CheckAllPictureProperties
  - CreateImageInfoFromPicture
  - TryFindImageInVbaProject
  - TryFindBackgroundImage
  - TryDetailedDrawingSearch
  - IsPartialIdMatch
- **狀態**: ⏳ 待執行

### 5. LogAvailableDrawings 方法
- **位置**: Lines 2980-3056
- **說明**: 記錄所有可用繪圖物件的診斷方法
- **狀態**: ⏳ 待執行

### 6. CountDispimgFormulas 方法
- **位置**: Lines 3060-3092
- **說明**: 計算工作表中 DISPIMG 公式數量的方法
- **狀態**: ⏳ 待執行

### 7. GeneratePlaceholderImage 方法
- **位置**: Lines 3112-3210
- **說明**: 生成佔位符圖片的方法
- **狀態**: ⏳ 待執行

## 📝 執行策略

採用**自頂向下、小步快跑**的策略:

1. **步驟 1**: 移除 GetCellImages 中的 DISPIMG 檢查區塊 (最上層調用)
2. **步驟 2**: 移除 ExtractImageIdFromFormula 方法
3. **步驟 3**: 移除 FindEmbeddedImageById 及其依賴的所有方法 (一次性大刪除)
4. **步驟 4**: 移除 LogAvailableDrawings 和 CountDispimgFormulas
5. **步驟 5**: 移除 GeneratePlaceholderImage 方法
6. **步驟 6**: 編譯驗證,確保沒有殘留引用

## ⚠️ 注意事項

- 每次刪除後都要確保有足夠的上下文代碼來定位
- 使用 3-5 行的上下文來避免誤刪
- 刪除大塊代碼時,要確認起始和結束的方法簽名
- 每個步驟完成後進行編譯檢查

## 📊 預期結果

- 移除代碼行數: ~1200 行
- 移除方法數量: ~25 個
- 提升代碼可維護性
- 減少不必要的複雜度
