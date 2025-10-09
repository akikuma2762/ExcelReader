# ExcelReaderVue - 貢獻指南

**版本:** 2.0.0  
**最後更新:** 2025年10月9日

---

## 如何貢獻

感謝您對 ExcelReaderVue 的興趣!

### 開發環境設定

#### 必要工具

- Node.js: ^20.19.0 或 >=22.12.0
- npm: 10.0.0 或更高
- Git: 2.30 或更高

#### 安裝步驟

```bash
# Clone 專案
git clone https://github.com/akikuma2762/ExcelReader.git
cd ExcelReader/ExcelReaderVue

# 安裝依賴
npm install

# 啟動開發伺服器
npm run dev
```

---

## 程式碼規範

### Vue 元件

```vue
<script setup lang="ts">
// 使用 Composition API + TypeScript
import { ref, computed } from 'vue'

const count = ref(0)
const doubleCount = computed(() => count.value * 2)
</script>

<template>
  <!-- 使用語義化 HTML -->
  <div class="component">
    <button @click="count++">Count: {{ count }}</button>
  </div>
</template>

<style scoped>
/* 使用 scoped 樣式 */
.component {
  /* styles */
}
</style>
```

### TypeScript

```typescript
// ✅ 明確的型別定義
const data = ref<ExcelData | null>(null)

// ✅ 介面定義
interface Props {
  title: string
  count?: number
}

// ❌ 避免使用 any
const data: any = {}
```

### 命名規範

- 元件: PascalCase (`ExcelReader.vue`)
- 函數: camelCase (`handleFileUpload`)
- 常數: UPPER_SNAKE_CASE (`API_BASE_URL`)
- 型別: PascalCase (`ExcelData`)

---

## 提交規範

使用 Conventional Commits:

```bash
feat: 新增功能
fix: 修復 Bug
docs: 文檔變更
style: 程式碼格式
refactor: 重構
perf: 效能優化
test: 測試
chore: 建置/工具
```

範例:

```bash
git commit -m "feat(upload): add drag and drop support"
git commit -m "fix(style): correct cell border rendering"
git commit -m "docs(readme): update installation guide"
```

---

## Pull Request 流程

1. Fork 專案
2. 建立功能分支 (`git checkout -b feature/amazing-feature`)
3. 提交變更 (`git commit -m 'feat: add amazing feature'`)
4. 推送分支 (`git push origin feature/amazing-feature`)
5. 開啟 Pull Request

---

## 測試

### 執行測試

```bash
# 單元測試
npm run test:unit

# E2E 測試
npm run test:e2e

# 型別檢查
npm run type-check

# Lint 檢查
npm run lint
```

### 撰寫測試

```typescript
import { describe, it, expect } from 'vitest'
import { mount } from '@vue/test-utils'
import ExcelReader from '@/components/ExcelReader.vue'

describe('ExcelReader', () => {
  it('should render correctly', () => {
    const wrapper = mount(ExcelReader)
    expect(wrapper.find('h1').text()).toBe('Excel 讀取器')
  })
})
```

---

## 問題回報

使用 GitHub Issues 回報問題,請包含:

- 問題描述
- 重現步驟
- 預期行為
- 實際行為
- 環境資訊 (瀏覽器、Node.js 版本等)
- 截圖 (如適用)

---

**聯絡方式:** support@excelreader.com  
**文檔維護者:** ExcelReader Team
