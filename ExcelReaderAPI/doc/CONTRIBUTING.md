# 貢獻指南 (Contributing Guide)

感謝您對 ExcelReaderAPI 專案的興趣!本文檔將指導您如何為專案做出貢獻。

---

## 📋 目錄

- [行為準則](#行為準則)
- [如何開始](#如何開始)
- [開發環境設定](#開發環境設定)
- [分支策略](#分支策略)
- [程式碼規範](#程式碼規範)
- [提交規範](#提交規範)
- [Pull Request 流程](#pull-request-流程)
- [測試要求](#測試要求)
- [文檔撰寫](#文檔撰寫)
- [問題回報](#問題回報)
- [功能建議](#功能建議)

---

## 行為準則

### 我們的承諾

為了營造開放和友善的環境,我們承諾:

- ✅ 尊重不同的觀點和經驗
- ✅ 優雅地接受建設性的批評
- ✅ 專注於對社群最有利的事情
- ✅ 對其他社群成員表現同理心

### 不可接受的行為

- ❌ 使用性別化語言或意象,以及不受歡迎的性關注
- ❌ 嘲諷、侮辱性評論,以及人身或政治攻擊
- ❌ 公開或私下騷擾
- ❌ 未經明確許可,發布他人的私人資訊

---

## 如何開始

### 1. 尋找可以貢獻的地方

- 🔍 瀏覽 [Issues](https://github.com/akikuma2762/ExcelReader/issues) 頁面
- 🏷️ 尋找標記為 `good first issue` 的問題
- 🆘 尋找標記為 `help wanted` 的問題
- 💡 查看 [Project Board](https://github.com/akikuma2762/ExcelReader/projects) 了解開發路線圖

### 2. 認領問題

在開始工作前,請在 Issue 中留言表示您想處理此問題,避免重複工作。

```markdown
我想處理這個問題,預計在本週末完成。
```

### 3. Fork 專案

點擊 GitHub 頁面右上角的 **Fork** 按鈕,將專案 Fork 到您的帳號。

---

## 開發環境設定

### 必要工具

| 工具 | 版本要求 | 用途 |
|------|---------|------|
| **.NET SDK** | 9.0 或更高 | 編譯和執行專案 |
| **Visual Studio** | 2022 或更高 | IDE (推薦) |
| **VS Code** | 最新版 | 輕量級編輯器 (可選) |
| **Git** | 2.30 或更高 | 版本控制 |

### 安裝步驟

#### 1. Clone 您的 Fork

```bash
git clone https://github.com/YOUR_USERNAME/ExcelReader.git
cd ExcelReader
```

#### 2. 添加上游遠端

```bash
git remote add upstream https://github.com/akikuma2762/ExcelReader.git
```

#### 3. 安裝依賴

```bash
cd ExcelReaderAPI
dotnet restore
```

#### 4. 驗證建置

```bash
dotnet build
```

#### 5. 執行測試

```bash
dotnet test
```

#### 6. 啟動開發伺服器

```bash
dotnet run
```

專案將在 `http://localhost:5000` 啟動。

### 推薦的 VS Code 擴展

- **C# Dev Kit** - C# 語言支援
- **C# Extensions** - C# 程式碼片段
- **GitLens** - Git 增強工具
- **REST Client** - API 測試
- **EditorConfig** - 程式碼格式化

### Visual Studio 設定

1. 開啟 `ExcelReaderAPI.sln`
2. 設定 `ExcelReaderAPI` 為啟動專案
3. 按 `F5` 開始偵錯

---

## 分支策略

我們使用 **Git Flow** 分支模型:

```
main (生產環境)
    │
    ├── develop (開發分支)
    │       │
    │       ├── feature/xxx (功能分支)
    │       ├── bugfix/xxx (錯誤修復分支)
    │       └── hotfix/xxx (緊急修復分支)
    │
    └── release/x.x.x (發布分支)
```

### 分支命名規範

| 分支類型 | 命名格式 | 範例 |
|---------|---------|------|
| **功能** | `feature/<issue-number>-<short-description>` | `feature/123-add-csv-support` |
| **錯誤修復** | `bugfix/<issue-number>-<short-description>` | `bugfix/456-fix-memory-leak` |
| **緊急修復** | `hotfix/<issue-number>-<short-description>` | `hotfix/789-fix-crash` |
| **發布** | `release/v<version>` | `release/v2.1.0` |

### 建立功能分支

```bash
# 確保在最新的 develop 分支
git checkout develop
git pull upstream develop

# 建立新的功能分支
git checkout -b feature/123-add-csv-support
```

---

## 程式碼規範

### C# 程式碼風格

遵循 [Microsoft C# Coding Conventions](https://docs.microsoft.com/en-us/dotnet/csharp/fundamentals/coding-style/coding-conventions)

#### 命名規範

```csharp
// ✅ 正確
public class ExcelProcessingService { }  // PascalCase for classes
public interface IExcelService { }       // I prefix + PascalCase for interfaces
public void ProcessFile() { }            // PascalCase for methods
private string _fileName;                // _camelCase for private fields
public string FileName { get; set; }     // PascalCase for properties
const int MaxFileSize = 100;             // PascalCase for constants

// ❌ 錯誤
public class excelProcessingService { }  // 小寫開頭
public interface ExcelService { }        // 缺少 I 前綴
private string fileName;                 // 缺少底線前綴
```

#### 程式碼組織

```csharp
public class ExcelProcessingService : IExcelProcessingService
{
    // 1. 常數
    private const int MaxRetries = 3;
    
    // 2. 私有欄位
    private readonly IExcelCellService _cellService;
    private readonly ILogger<ExcelProcessingService> _logger;
    
    // 3. 建構子
    public ExcelProcessingService(
        IExcelCellService cellService,
        ILogger<ExcelProcessingService> logger)
    {
        _cellService = cellService;
        _logger = logger;
    }
    
    // 4. 公開方法
    public async Task<ExcelData> ProcessExcelFileAsync(Stream stream, string fileName)
    {
        // Implementation
    }
    
    // 5. 私有方法
    private ExcelCellInfo CreateCellInfo(ExcelRange cell)
    {
        // Implementation
    }
}
```

#### 註解規範

```csharp
/// <summary>
/// 處理 Excel 檔案並轉換為 JSON 格式
/// </summary>
/// <param name="stream">Excel 檔案流</param>
/// <param name="fileName">檔案名稱</param>
/// <returns>包含完整 Excel 資料的物件</returns>
/// <exception cref="ArgumentNullException">當 stream 為 null 時拋出</exception>
public async Task<ExcelData> ProcessExcelFileAsync(Stream stream, string fileName)
{
    // 驗證參數
    if (stream == null)
        throw new ArgumentNullException(nameof(stream));
    
    // TODO: 實作公式計算功能
    // FIXME: 主題顏色計算可能不正確
    
    // 載入 Excel 套件
    using var package = new ExcelPackage(stream);
    
    // ... 其他程式碼
}
```

### 效能考量

#### ✅ 做

```csharp
// 使用 StringBuilder 拼接字串
var sb = new StringBuilder();
foreach (var item in items)
{
    sb.Append(item);
}

// 使用 using 確保資源釋放
using var package = new ExcelPackage(stream);

// 使用索引快速查找
var imageDict = images.ToDictionary(i => i.Name);
var image = imageDict[name]; // O(1)
```

#### ❌ 不要

```csharp
// 不要用 + 拼接大量字串
string result = "";
foreach (var item in items)
{
    result += item; // 效能差
}

// 不要忘記釋放資源
var package = new ExcelPackage(stream);
// 處理...
// 忘記 Dispose()

// 不要用 List.Find 重複查找
foreach (var name in names)
{
    var image = images.FirstOrDefault(i => i.Name == name); // O(n)
}
```

### 錯誤處理

```csharp
// ✅ 正確的錯誤處理
public async Task<ExcelData> ProcessExcelFileAsync(Stream stream, string fileName)
{
    try
    {
        // 驗證參數
        ValidateParameters(stream, fileName);
        
        // 業務邏輯
        return await ProcessAsync(stream, fileName);
    }
    catch (ArgumentException ex)
    {
        _logger.LogWarning(ex, "Invalid parameters: {FileName}", fileName);
        throw; // 重新拋出讓 Controller 處理
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "Failed to process Excel file: {FileName}", fileName);
        throw new ExcelProcessingException("Failed to process Excel file", ex);
    }
}
```

---

## 提交規範

### Commit Message 格式

使用 [Conventional Commits](https://www.conventionalcommits.org/) 格式:

```
<type>(<scope>): <subject>

<body>

<footer>
```

#### Type 類型

| Type | 說明 | 範例 |
|------|------|------|
| **feat** | 新功能 | `feat(service): add CSV export support` |
| **fix** | 錯誤修復 | `fix(controller): fix memory leak in upload` |
| **docs** | 文檔變更 | `docs(readme): update installation guide` |
| **style** | 程式碼格式 | `style(service): format code with EditorConfig` |
| **refactor** | 重構 | `refactor(service): extract color service` |
| **perf** | 效能優化 | `perf(index): add image index cache` |
| **test** | 測試 | `test(service): add unit tests for cell service` |
| **chore** | 建置/工具 | `chore(deps): update EPPlus to 8.1.0` |

#### Scope 範圍

- `controller` - Controller 層
- `service` - Service 層
- `model` - 資料模型
- `config` - 配置
- `deps` - 依賴套件
- `docs` - 文檔

#### 範例

```bash
# 簡單提交
git commit -m "feat(service): add formula evaluation support"

# 完整提交
git commit -m "fix(controller): fix memory leak in file upload

- Add using statement to ensure ExcelPackage disposal
- Implement GC.Collect() after large file processing
- Add memory usage logging

Fixes #456"
```

### 好的 Commit Message

```
✅ feat(service): add support for CSV file format
✅ fix(controller): prevent memory leak in upload endpoint
✅ docs(api): update API specification for v2.1
✅ perf(index): optimize image index creation (7x faster)
✅ refactor(service): extract color parsing to separate service
```

### 不好的 Commit Message

```
❌ update code
❌ fix bug
❌ WIP
❌ asdfasdf
❌ final version
```

---

## Pull Request 流程

### 1. 準備您的變更

```bash
# 確保程式碼最新
git checkout develop
git pull upstream develop

# 切換到您的功能分支
git checkout feature/123-add-csv-support

# 合併最新的 develop 分支
git merge develop

# 解決衝突 (如有)
# ...

# 執行測試
dotnet test

# 建置專案
dotnet build
```

### 2. 推送到您的 Fork

```bash
git push origin feature/123-add-csv-support
```

### 3. 建立 Pull Request

1. 前往 GitHub 上您的 Fork
2. 點擊 **Compare & pull request** 按鈕
3. 選擇 **base: develop** ← **compare: feature/123-add-csv-support**
4. 填寫 PR 描述 (使用下方範本)

### PR 描述範本

```markdown
## 描述
簡短描述此 PR 的目的和內容。

## 相關 Issue
Closes #123

## 變更類型
- [ ] 新功能
- [ ] 錯誤修復
- [ ] 重構
- [ ] 效能優化
- [ ] 文檔更新
- [ ] 測試

## 變更內容
- 新增 CSV 檔案解析支援
- 實作 `ICsvProcessor` 介面
- 新增單元測試

## 測試
- [ ] 單元測試通過
- [ ] 整合測試通過
- [ ] 手動測試完成

## 檢查清單
- [ ] 程式碼遵循專案的程式碼規範
- [ ] 已撰寫/更新測試
- [ ] 已更新相關文檔
- [ ] 所有測試通過
- [ ] 無編譯警告
- [ ] PR 標題符合 Conventional Commits 規範

## 截圖 (如適用)
[新增截圖或 GIF]

## 額外說明
[其他需要說明的內容]
```

### 4. Code Review

- 維護者會審查您的程式碼
- 根據反饋修改程式碼
- 推送更新 (會自動更新 PR)

```bash
# 根據反饋修改程式碼
# ...

# 提交變更
git add .
git commit -m "refactor: address code review comments"
git push origin feature/123-add-csv-support
```

### 5. 合併

- 當 PR 被批准後,維護者會合併您的變更
- 您的貢獻將出現在下一個版本中!

---

## 測試要求

### 測試覆蓋率目標

- **最低要求**: 70%
- **建議目標**: 80%+
- **核心邏輯**: 90%+

### 單元測試

每個新功能都應包含單元測試:

```csharp
public class ExcelCellServiceTests
{
    private readonly Mock<IExcelColorService> _mockColorService;
    private readonly ExcelCellService _service;
    
    public ExcelCellServiceTests()
    {
        _mockColorService = new Mock<IExcelColorService>();
        _service = new ExcelCellService(_mockColorService.Object);
    }
    
    [Fact]
    public void FindMergedRange_WhenCellIsMerged_ReturnsRange()
    {
        // Arrange
        var worksheet = CreateMockWorksheet();
        worksheet.Cells["A1:B2"].Merge = true;
        
        // Act
        var result = _service.FindMergedRange(worksheet, 1, 1);
        
        // Assert
        Assert.NotNull(result);
        Assert.Equal("A1:B2", result.Address);
    }
    
    [Theory]
    [InlineData(1, 1, true)]  // 合併範圍內
    [InlineData(3, 3, false)] // 合併範圍外
    public void FindMergedRange_VariousPositions_ReturnsCorrectResult(
        int row, int col, bool shouldBeMerged)
    {
        // Arrange
        var worksheet = CreateMockWorksheet();
        worksheet.Cells["A1:B2"].Merge = true;
        
        // Act
        var result = _service.FindMergedRange(worksheet, row, col);
        
        // Assert
        if (shouldBeMerged)
            Assert.NotNull(result);
        else
            Assert.Null(result);
    }
}
```

### 整合測試

```csharp
public class ExcelProcessingIntegrationTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly HttpClient _client;
    
    [Fact]
    public async Task UploadExcel_ValidFile_ReturnsSuccess()
    {
        // Arrange
        var content = new MultipartFormDataContent();
        var fileContent = new ByteArrayContent(File.ReadAllBytes("test.xlsx"));
        fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        content.Add(fileContent, "file", "test.xlsx");
        
        // Act
        var response = await _client.PostAsync("/api/excel/upload", content);
        
        // Assert
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync();
        var result = JsonSerializer.Deserialize<ApiResponse>(json);
        Assert.True(result.Success);
    }
}
```

### 執行測試

```bash
# 執行所有測試
dotnet test

# 執行特定測試
dotnet test --filter "FullyQualifiedName~ExcelCellServiceTests"

# 產生覆蓋率報告
dotnet test /p:CollectCoverage=true /p:CoverletOutputFormat=opencover
```

---

## 文檔撰寫

### 需要更新的文檔

當您的變更影響以下內容時,請更新相關文檔:

| 文檔 | 更新時機 |
|------|---------|
| **README.md** | 安裝步驟、快速開始、功能列表變更 |
| **API_SPECIFICATION.md** | API 端點、請求/響應格式變更 |
| **ARCHITECTURE.md** | 架構設計、Service 變更 |
| **CHANGELOG.md** | 每次 PR 合併 |

### 程式碼文檔

使用 XML 註解為公開 API 撰寫文檔:

```csharp
/// <summary>
/// 處理跨越多個儲存格的圖片
/// </summary>
/// <param name="worksheet">工作表</param>
/// <param name="picture">圖片物件</param>
/// <param name="cellDictionary">儲存格字典</param>
/// <param name="imageIndex">圖片索引</param>
/// <remarks>
/// 此方法會計算圖片覆蓋的所有儲存格,並將圖片資訊加入到每個儲存格中。
/// 使用索引可以避免重複遍歷,大幅提升效能。
/// </remarks>
/// <example>
/// <code>
/// var imageIndex = new WorksheetImageIndex(worksheet);
/// _cellService.ProcessImageCrossCells(worksheet, picture, cellDict, imageIndex);
/// </code>
/// </example>
public void ProcessImageCrossCells(
    ExcelWorksheet worksheet,
    ExcelPicture picture,
    Dictionary<string, ExcelCellInfo> cellDictionary,
    WorksheetImageIndex imageIndex)
{
    // Implementation
}
```

---

## 問題回報

### 回報 Bug

使用 [Bug Report 範本](https://github.com/akikuma2762/ExcelReader/issues/new?template=bug_report.md):

```markdown
**描述 Bug**
清楚且簡潔地描述 Bug。

**重現步驟**
1. 上傳包含合併儲存格的 Excel 檔案
2. 查看 JSON 輸出
3. 發現合併範圍不正確

**預期行為**
合併範圍應該是 "A1:B2"

**實際行為**
合併範圍顯示為 "A1:A1"

**截圖**
[新增截圖]

**環境**
- OS: Windows 11
- .NET 版本: 9.0
- EPPlus 版本: 8.1.0
- 專案版本: 2.0.0

**額外資訊**
[其他相關資訊]

**測試檔案**
[附上可重現問題的 Excel 檔案]
```

---

## 功能建議

### 提出新功能

使用 [Feature Request 範本](https://github.com/akikuma2762/ExcelReader/issues/new?template=feature_request.md):

```markdown
**功能描述**
希望 API 能支援 CSV 檔案格式。

**使用場景**
我們的系統需要處理大量的 CSV 資料,希望能用相同的 API 處理。

**建議解決方案**
1. 新增 `ICsvProcessor` 介面
2. 實作 CSV 解析邏輯
3. 在 Controller 中支援 .csv 檔案上傳

**替代方案**
可以先將 CSV 轉換為 Excel 格式再上傳。

**優先級**
- [ ] 高 - 阻塞性需求
- [x] 中 - 重要但不緊急
- [ ] 低 - Nice to have

**願意貢獻**
- [x] 我願意提交 PR 實作此功能
- [ ] 我只是提出建議
```

---

## 社群支援

### 獲得幫助

- 💬 [Discussions](https://github.com/akikuma2762/ExcelReader/discussions) - 提問和討論
- 📧 Email: support@excelreader.com
- 💼 [Stack Overflow](https://stackoverflow.com/questions/tagged/excelreaderapi) - 技術問題

### 保持更新

- ⭐ Star 專案以獲得更新通知
- 👀 Watch 專案以接收所有活動通知
- 🔔 訂閱 [Release Notes](https://github.com/akikuma2762/ExcelReader/releases)

---

## 貢獻者認可

所有貢獻者都會被列在:

- 📜 [CONTRIBUTORS.md](CONTRIBUTORS.md) 文件中
- 🏆 GitHub Contributors 頁面
- 📝 Release Notes 中的 "Contributors" 區塊

感謝所有貢獻者! 🎉

---

## 授權

貢獻此專案,即表示您同意您的貢獻將使用與專案相同的授權條款。

---

**最後更新:** 2025年10月9日  
**維護者:** ExcelReader Team

有問題嗎?歡迎在 [Discussions](https://github.com/akikuma2762/ExcelReader/discussions) 提出!
