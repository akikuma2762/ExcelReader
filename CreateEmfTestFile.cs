using OfficeOpenXml;
using System;
using System.IO;

// 創建包含EMF圖片的Excel檔案來測試轉換功能
class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        // 讀取EMF圖片數據
        var emfBase64 = File.ReadAllText("圖片編碼.txt").Trim();
        var emfBytes = Convert.FromBase64String(emfBase64);
        
        Console.WriteLine($"EMF數據長度: {emfBytes.Length} bytes");
        
        // 創建新的Excel檔案
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("EMF測試");
        
        // 添加一些測試數據
        worksheet.Cells["A1"].Value = "EMF圖片測試";
        worksheet.Cells["A2"].Value = "下方應該有一張EMF圖片";
        
        // 嘗試添加EMF圖片到Excel
        try
        {
            using var stream = new MemoryStream(emfBytes);
            var picture = worksheet.Drawings.AddPicture("EMF測試圖片", stream);
            picture.SetPosition(2, 0, 1, 0); // 從B3開始
            picture.SetSize(400, 300);
            
            Console.WriteLine("EMF圖片已添加到Excel");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"添加EMF圖片失敗: {ex.Message}");
        }
        
        // 保存檔案
        var fileName = "EMF測試檔案.xlsx";
        File.WriteAllBytes(fileName, package.GetAsByteArray());
        
        Console.WriteLine($"Excel檔案已保存: {fileName}");
        Console.WriteLine("現在可以使用此檔案測試EMF轉PNG功能");
    }
}