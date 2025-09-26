# 合併儲存格顯示修正報告

## 問題描述

### 🚨 原始問題
雖然已經正確設定了 `rowspan` 和 `colspan` 屬性，但在 HTML 表格中，被合併的儲存格仍然會顯示出來，導致：
- 表格佈局錯誤
- 重複顯示內容
- 視覺效果與原始 Excel 不符

### 📊 HTML 表格合併規則
在正確的 HTML 表格合併中：
```html
<!-- ✅ 正確的合併儲存格渲染 -->
<tr>
  <td colspan="2" rowspan="2">主儲存格</td>
  <!-- 被合併的儲存格不應該出現在 HTML 中 -->
  <td>其他儲存格</td>
</tr>
<tr>
  <!-- 第二行也不應該有被合併的儲存格 -->
  <td>其他儲存格</td>
</tr>
```

## 解決方案

### 1. 🔧 後端改進

#### 新增屬性
在 `ExcelCellInfo` 模型中新增：
```csharp
public bool IsMainMergedCell { get; set; }
```

#### 邏輯改進
```csharp
// 檢查是否為主儲存格（合併範圍的左上角）
cellInfo.IsMainMergedCell = (cell.Start.Row == mergedRange.Start.Row && 
                            cell.Start.Column == mergedRange.Start.Column);

if (cellInfo.IsMainMergedCell)
{
    // 只有主儲存格設定 rowspan 和 colspan
    cellInfo.RowSpan = mergedRange.Rows;
    cellInfo.ColSpan = mergedRange.Columns;
}
```

### 2. 🎨 前端改進

#### TypeScript 介面更新
```typescript
interface ExcelCellInfo {
  // ... 其他屬性
  isMainMergedCell?: boolean  // 新增屬性
}
```

#### 渲染邏輯
```typescript
const shouldRenderCell = (cell: ExcelCellInfo): boolean => {
  // 非合併儲存格：正常顯示
  if (!cell.isMerged) {
    return true
  }
  
  // 合併儲存格：只顯示主儲存格
  return cell.isMainMergedCell === true
}
```

#### 模板更新
```vue
<template v-for="(cell, cellIndex) in row" :key="cellIndex">
  <td v-if="shouldRenderCell(cell)" ...>
    <!-- 只渲染主儲存格或非合併儲存格 -->
  </td>
</template>
```

## 技術實作細節

### 🔍 儲存格狀態分析

| 儲存格類型 | `isMerged` | `isMainMergedCell` | `rowSpan/colSpan` | 是否渲染 |
|-----------|------------|-------------------|------------------|----------|
| 一般儲存格 | `false` | `false` | `1/1` | ✅ 是 |
| 主合併儲存格 | `true` | `true` | `實際範圍` | ✅ 是 |
| 被合併儲存格 | `true` | `false` | `1/1` | ❌ 否 |

### 📋 合併範圍範例

假設有一個 2x3 的合併儲存格（A1:C2）：

```
   A    B    C    D
1  [主儲存格---]  D1
2  [被合併----]  D2
3  A3   B3   C3   D3
```

**後端資料結構**：
- A1: `isMerged=true, isMainMergedCell=true, rowSpan=2, colSpan=3`
- B1: `isMerged=true, isMainMergedCell=false, rowSpan=1, colSpan=1`
- C1: `isMerged=true, isMainMergedCell=false, rowSpan=1, colSpan=1`
- A2: `isMerged=true, isMainMergedCell=false, rowSpan=1, colSpan=1`
- B2: `isMerged=true, isMainMergedCell=false, rowSpan=1, colSpan=1`
- C2: `isMerged=true, isMainMergedCell=false, rowSpan=1, colSpan=1`

**前端渲染結果**：
```html
<tr>
  <td colspan="3" rowspan="2">主儲存格內容</td>
  <!-- B1, C1 不渲染 -->
  <td>D1</td>
</tr>
<tr>
  <!-- A2, B2, C2 不渲染 -->
  <td>D2</td>
</tr>
```

## 向後相容性

### ✅ 完全相容
- 現有的非合併儲存格功能完全不受影響
- 新增屬性為可選項，不會破壞現有資料
- 所有其他功能（Rich Text、格式化、對齊等）正常運作

### 🔄 升級路徑
1. 後端自動處理所有合併邏輯
2. 前端自動過濾顯示內容
3. 無需手動調整現有資料

## 測試建議

### 📝 測試案例
1. **基本合併**: 2x2 合併儲存格
2. **複雜合併**: 不同大小的多個合併區域
3. **混合內容**: 合併儲存格 + Rich Text + 格式化
4. **邊界情況**: 表格邊緣的合併儲存格
5. **多工作表**: 不同工作表的合併儲存格

### ✅ 預期結果
- 表格佈局與原始 Excel 完全一致
- 合併儲存格正確跨越指定範圍
- 被合併區域不顯示重複內容
- 滑鼠懸停工具提示正確顯示合併資訊

這個修正解決了合併儲存格顯示的核心問題，讓網頁表格完美還原 Excel 的視覺效果！