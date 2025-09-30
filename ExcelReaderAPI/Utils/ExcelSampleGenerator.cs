using ClosedXML.Excel;
using System.IO;

namespace ExcelReaderAPI.Utils
{
    public static class ExcelSampleGenerator
    {
        public static byte[] GenerateSampleExcel()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("員工資料");

            // 設定標題
            worksheet.Cell("A1").Value = "姓名";
            worksheet.Cell("B1").Value = "年齡";
            worksheet.Cell("C1").Value = "部門";
            worksheet.Cell("D1").Value = "薪資";
            worksheet.Cell("E1").Value = "入職日期";

            // 設定標題樣式
            var headerRange = worksheet.Range("A1:E1");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            // 員工資料
            var employees = new[]
            {
                new { Name = "張三", Age = 30, Department = "資訊部", Salary = 50000, StartDate = new DateTime(2020, 1, 15) },
                new { Name = "李四", Age = 25, Department = "人事部", Salary = 45000, StartDate = new DateTime(2021, 3, 20) },
                new { Name = "王五", Age = 35, Department = "財務部", Salary = 55000, StartDate = new DateTime(2019, 5, 10) },
                new { Name = "趙六", Age = 28, Department = "行銷部", Salary = 48000, StartDate = new DateTime(2022, 7, 1) },
                new { Name = "錢七", Age = 32, Department = "研發部", Salary = 60000, StartDate = new DateTime(2018, 12, 5) },
                new { Name = "孫八", Age = 29, Department = "客服部", Salary = 42000, StartDate = new DateTime(2021, 9, 15) },
                new { Name = "周九", Age = 31, Department = "業務部", Salary = 52000, StartDate = new DateTime(2020, 11, 20) }
            };

            // 填入資料並測試 Rich Text 和顏色
            for (int i = 0; i < employees.Length; i++)
            {
                var row = i + 2;
                var emp = employees[i];

                // 姓名欄位 - 測試 Rich Text 顏色功能
                var nameCell = worksheet.Cell(row, 1);
                
                // 為特定字元設定顏色，測試 Rich Text
                if (emp.Name == "張三")
                {
                    var richText = nameCell.GetRichText();
                    richText.AddText("張").SetFontColor(XLColor.Black);
                    richText.AddText("三").SetFontColor(XLColor.Red); // 重點：測試紅色字體
                }
                else if (emp.Name == "李四")
                {
                    var richText = nameCell.GetRichText();
                    richText.AddText("李").SetFontColor(XLColor.Blue);
                    richText.AddText("四").SetFontColor(XLColor.Green);
                }
                else
                {
                    nameCell.Value = emp.Name;
                }

                worksheet.Cell(row, 2).Value = emp.Age;
                worksheet.Cell(row, 3).Value = emp.Department;
                worksheet.Cell(row, 4).Value = emp.Salary;
                worksheet.Cell(row, 5).Value = emp.StartDate;

                // 設定薪資格式
                worksheet.Cell(row, 4).Style.NumberFormat.Format = "#,##0";
                
                // 設定日期格式
                worksheet.Cell(row, 5).Style.NumberFormat.Format = "yyyy/mm/dd";

                // 交替行顏色
                if (i % 2 == 0)
                {
                    worksheet.Range(row, 1, row, 5).Style.Fill.BackgroundColor = XLColor.AliceBlue;
                }
            }

            // 設定邊框
            var dataRange = worksheet.Range($"A1:E{employees.Length + 1}");
            dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // 自動調整欄寬
            worksheet.Columns().AdjustToContents();

            // 凍結標題行
            worksheet.SheetView.FreezeRows(1);

            // 新增測試用的合併儲存格
            var mergedCell = worksheet.Range("G1:H2");
            mergedCell.Merge();
            mergedCell.Value = "合併儲存格測試";
            mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            mergedCell.Style.Fill.BackgroundColor = XLColor.Yellow;
            mergedCell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

            // 新增顏色測試區域
            worksheet.Cell("J1").Value = "顏色測試";
            worksheet.Cell("J1").Style.Font.FontColor = XLColor.Red;
            worksheet.Cell("J1").Style.Fill.BackgroundColor = XLColor.LightYellow;

            worksheet.Cell("J2").Value = "主題顏色";
            worksheet.Cell("J2").Style.Font.FontColor = XLColor.Blue;
            worksheet.Cell("J2").Style.Fill.BackgroundColor = XLColor.LightGray;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }
    }
}