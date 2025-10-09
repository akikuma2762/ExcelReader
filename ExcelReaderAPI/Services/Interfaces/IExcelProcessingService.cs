using ExcelReaderAPI.Models;
using OfficeOpenXml;
using ExcelReaderAPI.Models.Caches;
using ExcelReaderAPI.Models.Enums;

namespace ExcelReaderAPI.Services.Interfaces
{
    /// <summary>
    /// Excel 處理服務介面 - 核心儲存格和工作表處理邏輯
    /// </summary>
    public interface IExcelProcessingService
    {
        /// <summary>
        /// 創建儲存格資訊 (使用快取優化)
        /// </summary>
        ExcelCellInfo CreateCellInfo(
            ExcelRange cell,
            ExcelWorksheet worksheet,
            WorksheetImageIndex? imageIndex,
            ColorCache colorCache,
            MergedCellIndex mergedCellIndex);

        /// <summary>
        /// 創建儲存格資訊 (簡化版)
        /// </summary>
        ExcelCellInfo CreateCellInfo(ExcelRange cell, ExcelWorksheet worksheet);

        /// <summary>
        /// 偵測儲存格內容類型 (使用快取)
        /// </summary>
        CellContentType DetectCellContentType(ExcelRange cell, WorksheetImageIndex? imageIndex);

        /// <summary>
        /// 偵測儲存格內容類型 (不使用快取)
        /// </summary>
        CellContentType DetectCellContentType(ExcelRange cell, ExcelWorksheet worksheet);

        /// <summary>
        /// 取得原始儲存格資料陣列
        /// </summary>
        object[,] GetRawCellData(ExcelWorksheet worksheet, int maxRows, int maxCols);

        /// <summary>
        /// 創建預設字型資訊
        /// </summary>
        FontInfo CreateDefaultFontInfo();

        /// <summary>
        /// 創建預設對齊資訊
        /// </summary>
        AlignmentInfo CreateDefaultAlignmentInfo();

        /// <summary>
        /// 創建預設邊框資訊
        /// </summary>
        BorderInfo CreateDefaultBorderInfo();

        /// <summary>
        /// 創建預設填充資訊
        /// </summary>
        FillInfo CreateDefaultFillInfo();

        /// <summary>
        /// 安全取得儲存格值
        /// </summary>
        string? GetSafeValue(ExcelRange cell);
    }
}
