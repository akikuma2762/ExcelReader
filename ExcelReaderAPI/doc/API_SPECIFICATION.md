# ExcelReaderAPI - API 規格文件

**版本:** 2.0.0  
**最後更新:** 2025年10月9日  
**Base URL:** `http://localhost:5000/api/excel`

---

## 📋 目錄

- [API 端點總覽](#api-端點總覽)
- [通用規範](#通用規範)
- [認證機制](#認證機制)
- [端點詳細說明](#端點詳細說明)
- [資料模型](#資料模型)
- [錯誤處理](#錯誤處理)
- [範例代碼](#範例代碼)

---

## API 端點總覽

| 方法 | 端點 | 描述 | 認證 |
|------|------|------|------|
| POST | `/api/excel/upload` | 上傳並解析 Excel 檔案 | 無 |
| GET | `/api/excel/sample` | 獲取範例資料 | 無 |
| GET | `/api/excel/test-smart-detection` | 測試智能內容檢測 | 無 |
| POST | `/api/excel/debug-raw-data` | 調試原始資料 (開發用) | 無 |

---

## 通用規範

### HTTP Headers

#### 請求 Headers

```http
Content-Type: multipart/form-data  # 檔案上傳時使用
Accept: application/json            # 接受 JSON 響應
```

#### 響應 Headers

```http
Content-Type: application/json; charset=utf-8
Access-Control-Allow-Origin: *     # CORS (如已配置)
```

### HTTP 狀態碼

| 狀態碼 | 說明 | 使用情境 |
|--------|------|---------|
| 200 | OK | 請求成功 |
| 400 | Bad Request | 請求參數錯誤或檔案格式不正確 |
| 404 | Not Found | 資源不存在 |
| 500 | Internal Server Error | 伺服器處理錯誤 |

### 檔案限制

- **檔案大小:** 最大 100MB
- **檔案類型:** `.xlsx`, `.xls`
- **工作表數量:** 無限制
- **儲存格數量:** 建議 <100,000 個 (效能考量)

---

## 認證機制

**當前版本:** 無需認證 (開發/測試環境)

**生產環境建議:** 實作 JWT Bearer Token 或 API Key 認證

```http
# 未來版本可能需要
Authorization: Bearer {your-jwt-token}
# 或
X-API-Key: {your-api-key}
```

---

## 端點詳細說明

### 1. 上傳並解析 Excel 檔案

解析 Excel 檔案為 JSON 格式,包含完整的儲存格資訊、圖片、樣式等。

#### 請求

```http
POST /api/excel/upload
Content-Type: multipart/form-data
```

**參數:**

| 參數名 | 類型 | 必填 | 描述 |
|--------|------|------|------|
| file | File | ✅ | Excel 檔案 (.xlsx 或 .xls) |

**範例請求:**

```bash
curl -X POST "http://localhost:5000/api/excel/upload" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@/path/to/your/file.xlsx"
```

#### 響應

**成功響應 (200 OK):**

```json
{
  "success": true,
  "message": "檔案上傳成功",
  "data": {
    "fileName": "test.xlsx",
    "fileSize": 123456,
    "worksheets": [
      {
        "name": "Sheet1",
        "index": 0,
        "rowCount": 100,
        "columnCount": 10,
        "cells": [
          {
            "position": {
              "row": 1,
              "column": 1,
              "address": "A1"
            },
            "value": "Hello World",
            "text": "Hello World",
            "dataType": "Text",
            "font": { /* ... */ },
            "alignment": { /* ... */ },
            "border": { /* ... */ },
            "fill": { /* ... */ },
            "dimensions": { /* ... */ },
            "images": [ /* ... */ ],
            "floatingObjects": [ /* ... */ ]
          }
        ],
        "mergedCells": ["A1:B2", "C3:D4"],
        "metadata": { /* ... */ }
      }
    ],
    "totalWorksheets": 1,
    "processingTime": "1.234s"
  }
}
```

**錯誤響應 (400 Bad Request):**

```json
{
  "success": false,
  "message": "檔案格式不正確或檔案損壞",
  "error": {
    "code": "INVALID_FILE_FORMAT",
    "details": "The uploaded file is not a valid Excel file"
  }
}
```

**錯誤響應 (500 Internal Server Error):**

```json
{
  "success": false,
  "message": "處理檔案時發生錯誤",
  "error": {
    "code": "PROCESSING_ERROR",
    "details": "Error message details"
  }
}
```

---

### 2. 獲取範例資料

返回範例 Excel 資料結構,用於前端開發和測試。

#### 請求

```http
GET /api/excel/sample
```

**參數:** 無

**範例請求:**

```bash
curl -X GET "http://localhost:5000/api/excel/sample" \
  -H "accept: application/json"
```

#### 響應

**成功響應 (200 OK):**

```json
{
  "fileName": "sample.xlsx",
  "fileSize": 0,
  "worksheets": [
    {
      "name": "Sample Sheet",
      "index": 0,
      "rowCount": 3,
      "columnCount": 3,
      "cells": [
        {
          "position": {
            "row": 1,
            "column": 1,
            "address": "A1"
          },
          "value": "Name",
          "text": "Name",
          "dataType": "Text",
          "font": {
            "name": "Calibri",
            "size": 11,
            "bold": true,
            "color": "000000"
          }
        }
      ]
    }
  ]
}
```

---

### 3. 測試智能內容檢測

測試 API 的智能內容檢測功能,返回檢測能力說明。

#### 請求

```http
GET /api/excel/test-smart-detection
```

**參數:** 無

**範例請求:**

```bash
curl -X GET "http://localhost:5000/api/excel/test-smart-detection" \
  -H "accept: application/json"
```

#### 響應

**成功響應 (200 OK):**

```json
{
  "feature": "Smart Content Detection",
  "version": "2.0",
  "capabilities": [
    {
      "name": "Empty Cell Detection",
      "description": "快速跳過空白儲存格",
      "enabled": true
    },
    {
      "name": "Image-Only Cell Optimization",
      "description": "僅圖片儲存格使用最小化樣式處理",
      "enabled": true
    },
    {
      "name": "Text Cell Full Processing",
      "description": "文字儲存格完整樣式解析",
      "enabled": true
    },
    {
      "name": "Mixed Content Handling",
      "description": "混合內容智能處理",
      "enabled": true
    }
  ],
  "performanceGain": "約 30-50% 處理速度提升"
}
```

---

### 4. 調試原始資料 (開發用)

返回 Excel 檔案的原始資料結構,用於開發和調試。

#### 請求

```http
POST /api/excel/debug-raw-data
Content-Type: multipart/form-data
```

**參數:**

| 參數名 | 類型 | 必填 | 描述 |
|--------|------|------|------|
| file | File | ✅ | Excel 檔案 (.xlsx 或 .xls) |

**範例請求:**

```bash
curl -X POST "http://localhost:5000/api/excel/debug-raw-data" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@test.xlsx"
```

#### 響應

**成功響應 (200 OK):**

```json
{
  "fileName": "test.xlsx",
  "worksheets": [
    {
      "name": "Sheet1",
      "rawData": [
        ["A1 Value", "B1 Value", "C1 Value"],
        ["A2 Value", "B2 Value", "C2 Value"]
      ],
      "dimensions": {
        "rows": 2,
        "columns": 3
      }
    }
  ]
}
```

---

## 資料模型

### ExcelData

完整的 Excel 檔案資料結構。

```typescript
interface ExcelData {
  fileName: string;           // 檔案名稱
  fileSize: number;          // 檔案大小 (bytes)
  worksheets: Worksheet[];   // 工作表陣列
  totalWorksheets: number;   // 工作表總數
  processingTime?: string;   // 處理時間
}
```

### Worksheet

工作表資料結構。

```typescript
interface Worksheet {
  name: string;              // 工作表名稱
  index: number;            // 工作表索引 (0-based)
  rowCount: number;         // 列數
  columnCount: number;      // 欄數
  cells: ExcelCellInfo[];   // 儲存格陣列
  mergedCells: string[];    // 合併儲存格地址陣列 (如 ["A1:B2"])
  metadata?: WorksheetMetadata;  // 工作表元資料
}
```

### ExcelCellInfo

儲存格完整資訊。

```typescript
interface ExcelCellInfo {
  // 位置資訊
  position: CellPosition;
  
  // 內容資訊
  value: any;               // 原始值
  text: string;            // 顯示文字
  formula?: string;        // 公式
  formulaR1C1?: string;   // R1C1 格式公式
  dataType: string;       // 資料類型: Text/Number/DateTime/Boolean/Empty
  valueType?: string;     // .NET 類型名稱
  
  // 樣式資訊
  font: FontInfo;
  alignment: AlignmentInfo;
  border: BorderInfo;
  fill: FillInfo;
  numberFormat?: string;
  numberFormatId?: number;
  
  // 尺寸與合併
  dimensions: DimensionInfo;
  
  // 圖片與浮動物件
  images?: ImageInfo[];
  floatingObjects?: FloatingObjectInfo[];
  
  // Rich Text
  richText?: RichTextPart[];
  
  // 註解與超連結
  comment?: CommentInfo;
  hyperlink?: HyperlinkInfo;
  
  // 元資料
  metadata?: CellMetadata;
}
```

### CellPosition

儲存格位置資訊。

```typescript
interface CellPosition {
  row: number;        // 列號 (1-based)
  column: number;     // 欄號 (1-based)
  address: string;    // 地址 (如 "A1")
}
```

### FontInfo

字體資訊。

```typescript
interface FontInfo {
  name: string;           // 字體名稱 (如 "Calibri")
  size: number;          // 字體大小 (pt)
  bold: boolean;         // 粗體
  italic: boolean;       // 斜體
  underLine: string;     // 底線 ("None", "Single", "Double")
  strike: boolean;       // 刪除線
  color?: string;        // 顏色 (HEX, 如 "FF0000")
  colorTheme?: string;   // 主題顏色索引
  colorTint?: number;    // 色調 (-1.0 to 1.0)
  charset?: number;      // 字符集
  scheme?: string;       // 字體方案
  family?: number;       // 字體家族
}
```

### AlignmentInfo

對齊資訊。

```typescript
interface AlignmentInfo {
  horizontal: string;     // 水平對齊: Left/Center/Right/Justify
  vertical: string;       // 垂直對齊: Top/Center/Bottom
  wrapText: boolean;      // 自動換行
  indent: number;         // 縮排級別
  readingOrder: string;   // 閱讀順序
  textRotation: number;   // 文字旋轉角度
  shrinkToFit: boolean;   // 縮小以適應
}
```

### BorderInfo

邊框資訊。

```typescript
interface BorderInfo {
  top: BorderStyle;
  bottom: BorderStyle;
  left: BorderStyle;
  right: BorderStyle;
  diagonal: BorderStyle;
  diagonalUp: boolean;
  diagonalDown: boolean;
}

interface BorderStyle {
  style: string;    // None/Thin/Medium/Thick/Double/Dotted/Dashed
  color?: string;   // HEX 顏色
}
```

### FillInfo

填充資訊。

```typescript
interface FillInfo {
  patternType: string;           // None/Solid/Gray125/etc.
  backgroundColor?: string;       // 背景色 (HEX)
  patternColor?: string;         // 圖案色 (HEX)
  backgroundColorTheme?: string; // 主題顏色
  backgroundColorTint?: number;  // 色調
}
```

### DimensionInfo

儲存格尺寸與合併資訊。

```typescript
interface DimensionInfo {
  columnWidth: number;        // 欄寬
  rowHeight: number;          // 列高
  isMerged: boolean;          // 是否為合併儲存格
  isMainMergedCell?: boolean; // 是否為合併範圍的主儲存格
  rowSpan?: number;           // 合併列數
  colSpan?: number;           // 合併欄數
  mergedRangeAddress?: string; // 合併範圍地址 (如 "A1:B2")
}
```

### ImageInfo

圖片資訊。

```typescript
interface ImageInfo {
  name: string;              // 圖片名稱
  description?: string;      // 描述
  imageType: string;         // PNG/JPEG/GIF/BMP/EMF
  width: number;            // 顯示寬度 (px)
  height: number;           // 顯示高度 (px)
  originalWidth?: number;    // 原始寬度 (px)
  originalHeight?: number;   // 原始高度 (px)
  left: number;             // 左偏移 (px)
  top: number;              // 上偏移 (px)
  base64Data: string;       // Base64 圖片資料
  fileName?: string;         // 檔案名稱
  fileSize: number;         // 檔案大小 (bytes)
  anchorCell: CellPosition; // 錨點儲存格
  hyperlinkAddress?: string; // 超連結
  isInCellPicture?: boolean; // 是否為 In-Cell 圖片
  altText?: string;          // 替代文字
  excelWidthCm?: number;     // Excel 顯示寬度 (cm)
  excelHeightCm?: number;    // Excel 顯示高度 (cm)
  scaleFactor?: number;      // 縮放比例
  isScaled?: boolean;        // 是否縮放
  scaleMethod?: string;      // 縮放方法說明
}
```

### FloatingObjectInfo

浮動物件資訊 (文字方塊、圖形等)。

```typescript
interface FloatingObjectInfo {
  name: string;              // 物件名稱
  description?: string;      // 描述
  objectType: string;        // Shape/TextBox/Chart/Table
  width: number;            // 寬度
  height: number;           // 高度
  left: number;             // 左偏移
  top: number;              // 上偏移
  text?: string;            // 文字內容
  anchorCell: CellPosition; // 錨點儲存格
  fromCell: CellPosition;   // 起始儲存格
  toCell: CellPosition;     // 結束儲存格
  isFloating: boolean;      // 是否為浮動物件
  style?: string;           // 樣式資訊
  hyperlinkAddress?: string; // 超連結
}
```

### RichTextPart

富文本片段。

```typescript
interface RichTextPart {
  text: string;           // 文字內容
  bold: boolean;          // 粗體
  italic: boolean;        // 斜體
  underLine: boolean;     // 底線
  strike: boolean;        // 刪除線
  size: number;          // 字體大小
  fontName: string;      // 字體名稱
  color?: string;        // 顏色 (HEX)
  verticalAlign: string; // 垂直對齊
}
```

### CommentInfo

註解資訊。

```typescript
interface CommentInfo {
  text: string;      // 註解文字
  author: string;    // 作者
  autoFit: boolean;  // 自動調整大小
  visible: boolean;  // 是否可見
}
```

### HyperlinkInfo

超連結資訊。

```typescript
interface HyperlinkInfo {
  absoluteUri: string;    // 絕對 URI
  originalString: string; // 原始字串
  isAbsoluteUri: boolean; // 是否為絕對 URI
}
```

---

## 錯誤處理

### 錯誤響應格式

所有錯誤響應遵循統一格式:

```json
{
  "success": false,
  "message": "人類可讀的錯誤訊息",
  "error": {
    "code": "ERROR_CODE",
    "details": "詳細錯誤資訊",
    "timestamp": "2025-10-09T10:30:00Z"
  }
}
```

### 錯誤碼列表

| 錯誤碼 | HTTP 狀態 | 說明 | 解決方案 |
|--------|----------|------|---------|
| `INVALID_FILE_FORMAT` | 400 | 檔案格式不正確 | 確認檔案為 .xlsx 或 .xls |
| `FILE_TOO_LARGE` | 400 | 檔案超過大小限制 | 減小檔案大小或壓縮 |
| `CORRUPTED_FILE` | 400 | 檔案損壞 | 修復檔案或使用備份 |
| `PROCESSING_ERROR` | 500 | 處理過程錯誤 | 檢查檔案內容,聯繫支援 |
| `OUT_OF_MEMORY` | 500 | 記憶體不足 | 減小檔案大小或增加伺服器記憶體 |
| `TIMEOUT` | 504 | 處理超時 | 減小檔案大小或增加超時限制 |

### 錯誤處理最佳實踐

**客戶端處理:**

```javascript
try {
  const response = await fetch('/api/excel/upload', {
    method: 'POST',
    body: formData
  });
  
  const result = await response.json();
  
  if (result.success) {
    // 處理成功結果
    console.log('Data:', result.data);
  } else {
    // 處理錯誤
    console.error('Error:', result.error.code);
    alert(result.message);
  }
} catch (error) {
  // 處理網路錯誤
  console.error('Network error:', error);
}
```

---

## 範例代碼

### JavaScript/TypeScript

#### 使用 Fetch API 上傳檔案

```javascript
async function uploadExcelFile(file) {
  const formData = new FormData();
  formData.append('file', file);
  
  try {
    const response = await fetch('http://localhost:5000/api/excel/upload', {
      method: 'POST',
      body: formData,
      headers: {
        'Accept': 'application/json'
      }
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    
    if (result.success) {
      console.log('檔案上傳成功!');
      console.log('工作表數量:', result.data.totalWorksheets);
      console.log('處理時間:', result.data.processingTime);
      return result.data;
    } else {
      console.error('上傳失敗:', result.message);
      throw new Error(result.message);
    }
  } catch (error) {
    console.error('上傳錯誤:', error);
    throw error;
  }
}

// 使用範例
const fileInput = document.querySelector('input[type="file"]');
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (file) {
    try {
      const data = await uploadExcelFile(file);
      // 處理返回的資料
      displayExcelData(data);
    } catch (error) {
      alert('上傳失敗: ' + error.message);
    }
  }
});
```

#### 使用 Axios

```javascript
import axios from 'axios';

async function uploadExcelFileWithAxios(file) {
  const formData = new FormData();
  formData.append('file', file);
  
  try {
    const response = await axios.post(
      'http://localhost:5000/api/excel/upload',
      formData,
      {
        headers: {
          'Content-Type': 'multipart/form-data'
        },
        onUploadProgress: (progressEvent) => {
          const percentCompleted = Math.round(
            (progressEvent.loaded * 100) / progressEvent.total
          );
          console.log(`上傳進度: ${percentCompleted}%`);
        }
      }
    );
    
    if (response.data.success) {
      return response.data.data;
    } else {
      throw new Error(response.data.message);
    }
  } catch (error) {
    if (error.response) {
      // 伺服器返回錯誤響應
      console.error('伺服器錯誤:', error.response.data);
      throw new Error(error.response.data.message);
    } else if (error.request) {
      // 請求已發送但沒有收到響應
      console.error('網路錯誤:', error.request);
      throw new Error('網路連線錯誤');
    } else {
      // 其他錯誤
      console.error('錯誤:', error.message);
      throw error;
    }
  }
}
```

### C# / .NET

#### 使用 HttpClient

```csharp
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;

public class ExcelApiClient
{
    private readonly HttpClient _httpClient;
    private const string BaseUrl = "http://localhost:5000/api/excel";
    
    public ExcelApiClient()
    {
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri(BaseUrl)
        };
    }
    
    public async Task<ExcelData> UploadExcelFileAsync(string filePath)
    {
        using var content = new MultipartFormDataContent();
        
        // 讀取檔案
        var fileBytes = await File.ReadAllBytesAsync(filePath);
        var fileContent = new ByteArrayContent(fileBytes);
        fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("multipart/form-data");
        
        content.Add(fileContent, "file", Path.GetFileName(filePath));
        
        // 發送請求
        var response = await _httpClient.PostAsync("/upload", content);
        
        // 檢查響應
        if (response.IsSuccessStatusCode)
        {
            var json = await response.Content.ReadAsStringAsync();
            var result = JsonConvert.DeserializeObject<ApiResponse>(json);
            
            if (result.Success)
            {
                return result.Data;
            }
            else
            {
                throw new Exception($"API 錯誤: {result.Message}");
            }
        }
        else
        {
            throw new HttpRequestException($"HTTP 錯誤: {response.StatusCode}");
        }
    }
    
    public class ApiResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public ExcelData Data { get; set; }
    }
}

// 使用範例
var client = new ExcelApiClient();
try
{
    var data = await client.UploadExcelFileAsync("path/to/file.xlsx");
    Console.WriteLine($"成功讀取 {data.TotalWorksheets} 個工作表");
    Console.WriteLine($"處理時間: {data.ProcessingTime}");
}
catch (Exception ex)
{
    Console.WriteLine($"錯誤: {ex.Message}");
}
```

### Python

#### 使用 requests

```python
import requests
import json

def upload_excel_file(file_path, api_url="http://localhost:5000/api/excel/upload"):
    """
    上傳 Excel 檔案到 API
    
    Args:
        file_path: Excel 檔案路徑
        api_url: API 端點 URL
        
    Returns:
        dict: API 響應資料
    """
    try:
        with open(file_path, 'rb') as file:
            files = {'file': (file_path, file, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            
            response = requests.post(api_url, files=files)
            response.raise_for_status()
            
            result = response.json()
            
            if result['success']:
                print(f"上傳成功!")
                print(f"工作表數量: {result['data']['totalWorksheets']}")
                print(f"處理時間: {result['data']['processingTime']}")
                return result['data']
            else:
                print(f"上傳失敗: {result['message']}")
                return None
                
    except requests.exceptions.RequestException as e:
        print(f"請求錯誤: {e}")
        return None
    except Exception as e:
        print(f"錯誤: {e}")
        return None

# 使用範例
if __name__ == "__main__":
    data = upload_excel_file("test.xlsx")
    if data:
        # 處理資料
        for worksheet in data['worksheets']:
            print(f"\n工作表: {worksheet['name']}")
            print(f"儲存格數量: {len(worksheet['cells'])}")
```

---

## 效能優化建議

### 客戶端優化

1. **檔案壓縮:** 上傳前壓縮大型 Excel 檔案
2. **分片上傳:** 對於超大檔案,實作分片上傳
3. **快取結果:** 快取已處理的檔案結果
4. **進度顯示:** 實作上傳和處理進度條

### 伺服器優化

1. **非同步處理:** 大檔案使用佇列系統非同步處理
2. **結果快取:** 使用 Redis 快取處理結果
3. **負載平衡:** 部署多個實例處理並發請求
4. **資源限制:** 設定合理的超時和記憶體限制

---

## 版本歷史

### v2.0.0 (2025-10-09)
- ✅ 完整 Service Layer 架構
- ✅ 智能內容檢測
- ✅ 索引快取系統
- ✅ EPPlus 8.x 支援

### v1.0.0 (Initial)
- 基本 Excel 讀取功能
- 儲存格資訊解析

---

**文檔維護者:** [Your Name]  
**最後審核:** 2025年10月9日  
**狀態:** ✅ 當前版本
