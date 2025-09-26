# Excel模型升級指南

## 概述

此次升級將Excel資料模型重構為基於EPPlus完整屬性的結構化設計，提供更詳細和完整的Excel格式化資訊。

## 主要變更

### 1. 後端模型重構 (ExcelReaderAPI/Models/ExcelData.cs)

#### 新增的結構化類型

- **`CellPosition`**: 儲存格位置資訊（行、列、地址）
- **`FontInfo`**: 完整的字體樣式資訊
- **`AlignmentInfo`**: 對齊方式詳細設定
- **`BorderInfo` 和 `BorderStyle`**: 邊框樣式資訊
- **`FillInfo`**: 填充和背景色彩資訊
- **`DimensionInfo`**: 尺寸和合併儲存格資訊
- **`CommentInfo`**: 註解資訊
- **`HyperlinkInfo`**: 超連結資訊
- **`CellMetadata`**: 儲存格中繼資料

#### ExcelCellInfo 升級

**新屬性結構:**
```csharp
public class ExcelCellInfo
{
    // 位置資訊
    public CellPosition Position { get; set; }

    // 基本資料
    public object? Value { get; set; }
    public string Text { get; set; }
    public string? Formula { get; set; }
    public string? FormulaR1C1 { get; set; }

    // 格式化
    public string? NumberFormat { get; set; }
    public int? NumberFormatId { get; set; }

    // 結構化樣式資訊
    public FontInfo Font { get; set; }
    public AlignmentInfo Alignment { get; set; }
    public BorderInfo Border { get; set; }
    public FillInfo Fill { get; set; }
    public DimensionInfo Dimensions { get; set; }

    // 進階功能
    public List<RichTextPart>? RichText { get; set; }
    public CommentInfo? Comment { get; set; }
    public HyperlinkInfo? Hyperlink { get; set; }
    public CellMetadata Metadata { get; set; }
}
```

**向後兼容屬性 (標記為過時):**
- `DisplayText` → 使用 `Text`
- `FormatCode` → 使用 `NumberFormat`
- `FontBold` → 使用 `Font.Bold`
- `FontSize` → 使用 `Font.Size`
- `FontName` → 使用 `Font.Name`
- `BackgroundColor` → 使用 `Fill.BackgroundColor`
- `FontColor` → 使用 `Font.Color`
- `TextAlign` → 使用 `Alignment.Horizontal`
- `ColumnWidth` → 使用 `Dimensions.ColumnWidth`
- `IsRichText` → 使用 `Metadata.IsRichText`
- `RowSpan`/`ColSpan` → 使用 `Dimensions.RowSpan`/`Dimensions.ColSpan`
- `IsMerged`/`IsMainMergedCell` → 使用 `Dimensions.IsMerged`/`Dimensions.IsMainMergedCell`

### 2. 前端類型系統重構

#### 新增 types 資料夾結構
```
src/
├── types/
│   ├── excel.ts     # 完整的Excel類型定義
│   └── index.ts     # 統一匯出
```

#### RichTextPart 屬性名稱變更
```typescript
// 舊屬性 → 新屬性
fontBold → bold
fontItalic → italic
fontUnderline → underLine
fontSize → size
fontColor → color
```

### 3. Controller 升級

#### 更新的 CreateCellInfo 方法
- 完整對應所有EPPlus屬性
- 結構化的資料組織
- 更詳細的樣式和格式資訊

#### Debug Endpoint 改進
- 返回結構化的 `DebugExcelData` 類型
- 更清晰的工作表資訊組織

## 遷移步驟

### 對於後端開發者

1. **更新屬性存取方式**:
   ```csharp
   // 舊方式
   var isBold = cell.FontBold;
   var width = cell.ColumnWidth;

   // 新方式
   var isBold = cell.Font.Bold;
   var width = cell.Dimensions.ColumnWidth;
   ```

2. **使用新的結構化屬性**:
   ```csharp
   // 存取邊框資訊
   var topBorderStyle = cell.Border.Top.Style;
   var topBorderColor = cell.Border.Top.Color;

   // 存取對齊資訊
   var horizontalAlign = cell.Alignment.Horizontal;
   var wrapText = cell.Alignment.WrapText;
   ```

### 對於前端開發者

1. **更新 import 語句**:
   ```typescript
   // 新方式
   import type { ExcelCellInfo, ExcelData, UploadResponse } from '@/types'
   ```

2. **更新屬性存取**:
   ```typescript
   // 舊方式
   const isBold = cell.fontBold
   const width = cell.columnWidth

   // 新方式
   const isBold = cell.font?.bold
   const width = cell.dimensions?.columnWidth
   ```

3. **更新 Rich Text 處理**:
   ```typescript
   // 舊方式
   if (part.fontBold) styles.push('font-weight: bold')

   // 新方式
   if (part.bold) styles.push('font-weight: bold')
   ```

## 向後兼容性

所有舊屬性仍然可用，但標記為 `@deprecated`。建議逐步遷移到新的屬性結構。

## 新功能優勢

### 1. 完整的EPPlus屬性支援
- 所有EPPlus提供的50+個屬性完整對應
- 詳細的樣式和格式資訊
- 完整的Rich Text支援

### 2. 結構化的資料組織
- 邏輯分群的屬性組織
- 更清晰的型別定義
- 更好的IntelliSense支援

### 3. 進階功能支援
- 註解 (Comments)
- 超連結 (Hyperlinks)
- 完整的邊框樣式
- 對齊和文字旋轉
- 背景填充模式

### 4. Debug 和開發支援
- 完整的EPPlus屬性debug endpoint
- 結構化的debug資料
- 更好的開發者體驗

## 測試和驗證

### 後端測試
```bash
cd ExcelReaderAPI
dotnet build
dotnet test  # 如果有測試的話
```

### 前端測試
```bash
cd ExcelReaderVue
npm run type-check
npm run build
```

## 未來發展

這次重構為未來的功能擴展奠定了基礎：

1. **更多Excel功能支援**: 條件格式化、資料驗證等
2. **進階匯出功能**: 保持所有格式的Excel匯出
3. **完整的工作表操作**: 插入、刪除、複製等
4. **即時協作**: 基於結構化資料的協作編輯

## 技術債務清理

此次升級同時清理了：

- 不一致的屬性命名
- 缺乏型別安全的資料結構
- 不完整的Excel功能支援
- 分散的型別定義

## 結論

此次升級大幅提升了系統的：

- **可維護性**: 清晰的型別結構和組織
- **擴展性**: 基於EPPlus完整API的設計
- **開發體驗**: 完整的TypeScript支援和IntelliSense
- **功能完整性**: 支援所有主要Excel格式化功能

建議開發團隊逐步採用新的API，並在下一個主要版本中移除舊的deprecated屬性。
