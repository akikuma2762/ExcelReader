using ExcelReaderAPI.Models;

namespace ExcelReaderAPI.Services.Interfaces
{
    /// <summary>
    /// Excel 範例資料服務接口
    /// </summary>
    public interface IExcelSampleService
    {
        /// <summary>
        /// 取得範例資料
        /// </summary>
        ExcelData GetSampleData();

        /// <summary>
        /// 產生範例 Excel 檔案
        /// </summary>
        byte[] GenerateSampleExcel();
    }
}
