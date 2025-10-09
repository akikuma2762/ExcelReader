using OfficeOpenXml.Style;

namespace ExcelReaderAPI.Services.Interfaces
{
    /// <summary>
    /// Excel 顏色處理服務介面 - 顏色轉換、主題顏色、Tint 效果
    /// </summary>
    public interface IExcelColorService
    {
        /// <summary>
        /// 取得背景顏色
        /// </summary>
        string? GetBackgroundColor(ExcelFill fill);

        /// <summary>
        /// 從 ExcelColor 取得顏色字串
        /// </summary>
        string? GetColorFromExcelColor(ExcelColor? color);

        /// <summary>
        /// 取得索引顏色
        /// </summary>
        string? GetIndexedColor(int? index);

        /// <summary>
        /// 取得主題顏色
        /// </summary>
        string? GetThemeColor(int? theme, double? tint);

        /// <summary>
        /// 套用 Tint 效果
        /// </summary>
        string ApplyTint(string hexColor, double tint);
    }
}
