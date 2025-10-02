using OfficeOpenXml;
using System.IO;

namespace ExcelReaderAPI.Utils
{
    public static class ExcelSampleGenerator
    {
        static ExcelSampleGenerator()
        {
            // 設定EPPlus授權（非商業用途）- EPPlus 8.x 新 API
            ExcelPackage.License.SetNonCommercialPersonal("dek");
        }

        public static byte[] GenerateSampleExcel()
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("員工資料");

            // 設定標題
            worksheet.Cells[1, 1].Value = "姓名";
            worksheet.Cells[1, 2].Value = "年齡";
            worksheet.Cells[1, 3].Value = "部門";
            worksheet.Cells[1, 4].Value = "薪資";
            worksheet.Cells[1, 5].Value = "入職日期";

            // 樣本資料
            var sampleData = new object[,]
            {
                { "張三", 30, "資訊部", 50000, DateTime.Parse("2020-01-15") },
                { "李四", 25, "人事部", 45000, DateTime.Parse("2021-03-20") },
                { "王五", 35, "財務部", 55000, DateTime.Parse("2019-05-10") },
                { "趙六", 28, "行銷部", 48000, DateTime.Parse("2022-07-01") },
                { "錢七", 32, "研發部", 60000, DateTime.Parse("2018-12-05") },
                { "孫八", 29, "客服部", 42000, DateTime.Parse("2021-09-15") },
                { "周九", 31, "業務部", 52000, DateTime.Parse("2020-11-20") }
            };

            // 填入資料
            for (int i = 0; i < sampleData.GetLength(0); i++)
            {
                for (int j = 0; j < sampleData.GetLength(1); j++)
                {
                    worksheet.Cells[i + 2, j + 1].Value = sampleData[i, j];
                }
            }

            // 設定日期格式
            worksheet.Column(5).Style.Numberformat.Format = "yyyy-mm-dd";

            // 自動調整欄寬
            worksheet.Cells.AutoFitColumns();

            return package.GetAsByteArray();
        }
    }
}