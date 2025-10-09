using OfficeOpenXml;
using ExcelReaderAPI.Models;
using ExcelReaderAPI.Services.Interfaces;

namespace ExcelReaderAPI.Services
{
    /// <summary>
    /// Excel 範例產生服務實作
    /// </summary>
    public class ExcelSampleService : IExcelSampleService
    {
        private readonly ILogger<ExcelSampleService> _logger;

        public ExcelSampleService(ILogger<ExcelSampleService> logger)
        {
            _logger = logger;
        }

        public ExcelData GetSampleData()
        {
            try
            {
                var sampleData = new ExcelData
                {
                    FileName = "範例檔案.xlsx",
                    WorksheetName = "範例工作表",
                    Headers = new object[][]
                    {
                        new object[] { "姓名", "年齡", "職業" }
                    },
                    Rows = new object[][]
                    {
                        new object[] { "張三", 25, "工程師" },
                        new object[] { "李四", 30, "設計師" },
                        new object[] { "王五", 28, "分析師" }
                    },
                    TotalRows = 3,
                    TotalColumns = 3,
                    AvailableWorksheets = new List<string> { "範例工作表" },
                    WorksheetInfo = new WorksheetInfo
                    {
                        Name = "範例工作表",
                        TotalRows = 4, // 包含標題行
                        TotalColumns = 3,
                        DefaultColWidth = 12.0,
                        DefaultRowHeight = 15.0
                    }
                };

                _logger.LogInformation("✅ 產生範例資料成功");
                return sampleData;
            }
            catch (Exception ex)
            {
                _logger.LogError($"❌ 產生範例資料失敗: {ex.Message}");
                throw;
            }
        }

        public byte[] GenerateSampleExcel()
        {
            try
            {
                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("範例工作表");

                // 設置標題行
                worksheet.Cells[1, 1].Value = "姓名";
                worksheet.Cells[1, 2].Value = "年齡";
                worksheet.Cells[1, 3].Value = "職業";
                worksheet.Cells[1, 4].Value = "薪資";

                // 設置標題行樣式
                using (var titleRange = worksheet.Cells[1, 1, 1, 4])
                {
                    titleRange.Style.Font.Bold = true;
                    titleRange.Style.Font.Size = 12;
                    titleRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    titleRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    titleRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                }

                // 添加範例資料
                var sampleData = new object[,]
                {
                    { "張三", 25, "軟體工程師", 50000 },
                    { "李四", 30, "UI/UX 設計師", 45000 },
                    { "王五", 28, "專案經理", 60000 },
                    { "趙六", 35, "資深工程師", 70000 },
                    { "錢七", 26, "前端工程師", 48000 }
                };

                // 填入資料
                for (int i = 0; i < sampleData.GetLength(0); i++)
                {
                    for (int j = 0; j < sampleData.GetLength(1); j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = sampleData[i, j];
                    }
                }

                // 設置資料區域樣式
                using (var dataRange = worksheet.Cells[2, 1, 6, 4])
                {
                    dataRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // 自動調整欄寬
                worksheet.Cells.AutoFitColumns();

                // 添加一些格式化
                worksheet.Cells[2, 4, 6, 4].Style.Numberformat.Format = "#,##0";

                var fileBytes = package.GetAsByteArray();
                
                _logger.LogInformation($"✅ 產生範例 Excel 檔案成功，大小: {fileBytes.Length} bytes");
                return fileBytes;
            }
            catch (Exception ex)
            {
                _logger.LogError($"❌ 產生範例 Excel 檔案失敗: {ex.Message}");
                throw;
            }
        }


    }
}