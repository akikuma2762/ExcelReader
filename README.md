# Vue + .NET Core Excel Reader

這是一個使用Vue.js前端和.NET Core Web API後端的Excel讀取應用程式，使用EPPlus套件讀取Excel檔案並轉換為JSON格式。

## 專案結構

```
VUE_EPPLUS/
├── ExcelReaderAPI/          # .NET Core Web API 後端
│   ├── Controllers/         # API控制器
│   ├── Models/             # 資料模型
│   └── ...
└── ExcelReaderVue/         # Vue.js 前端
    ├── src/
    │   ├── components/     # Vue組件
    │   └── ...
    └── ...
```

## 功能特色

- 📁 支援Excel檔案上傳（.xlsx, .xls）
- 🖱️ 拖拽上傳功能
- 📊 Excel資料表格顯示
- 📄 JSON格式資料預覽
- 🔧 範例資料測試
- 📱 響應式設計

## 技術棧

### 後端
- .NET 9.0
- ASP.NET Core Web API
- EPPlus 8.2.0（Excel處理）
- Swagger/OpenAPI（API文檔）

### 前端
- Vue 3
- TypeScript
- Vite
- Axios（HTTP客戶端）

## 安裝與運行

### 前置要求
- .NET 9.0 SDK
- Node.js 18+
- npm

### 後端API設定

1. 進入API專案目錄：
```bash
cd ExcelReaderAPI
```

2. 還原NuGet套件：
```bash
dotnet restore
```

3. 啟動API服務：
```bash
dotnet run
```

API將在 `https://localhost:7254` 運行，Swagger文檔位於 `https://localhost:7254/swagger`

### 前端設定

1. 進入Vue專案目錄：
```bash
cd ExcelReaderVue
```

2. 安裝npm依賴：
```bash
npm install
```

3. 啟動開發伺服器：
```bash
npm run dev
```

前端應用將在 `http://localhost:5173` 運行

## API端點

### POST /api/excel/upload
上傳Excel檔案並解析為JSON

**請求**：
- Content-Type: multipart/form-data
- Body: Excel檔案

**回應**：
```json
{
  "success": true,
  "message": "成功讀取 Excel 檔案，共 3 筆資料",
  "data": {
    "fileName": "example.xlsx",
    "totalRows": 4,
    "totalColumns": 4,
    "headers": [["姓名", "年齡", "部門", "薪資"]],
    "rows": [
      ["張三", 30, "資訊部", 50000],
      ["李四", 25, "人事部", 45000],
      ["王五", 35, "財務部", 55000]
    ]
  }
}
```

### GET /api/excel/sample
取得範例資料

## 使用方式

1. 啟動後端API和前端應用
2. 在瀏覽器中開啟前端應用
3. 選擇或拖拽Excel檔案到上傳區域
4. 查看解析後的表格資料
5. 點擊"顯示JSON"按鈕查看原始JSON資料

## 支援的Excel格式

- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)

## 注意事項

- EPPlus使用非商業授權
- 僅讀取Excel檔案的第一個工作表
- 第一行被視為標題行
- 支援文字、數字和日期資料類型

## 開發

### 後端開發
```bash
# 監視模式運行
dotnet watch run

# 建置專案
dotnet build

# 執行測試
dotnet test
```

### 前端開發
```bash
# 開發模式
npm run dev

# 建置生產版本
npm run build

# 預覽建置結果
npm run preview

# 程式碼檢查
npm run lint

# 程式碼格式化
npm run format
```

## 疑難排解

### CORS錯誤
確保後端API已正確設定CORS策略以允許前端域名。

### EPPlus授權錯誤
如果遇到EPPlus授權問題，請確認使用的是非商業用途，或購買商業授權。

### 檔案上傳大小限制
預設檔案上傳大小有限制，如需要上傳大檔案，請調整後端設定。