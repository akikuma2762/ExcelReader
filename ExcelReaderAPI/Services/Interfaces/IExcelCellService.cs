using ExcelReaderAPI.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace ExcelReaderAPI.Services.Interfaces
{
    /// <summary>
    /// Excel 儲存格服務介面 - 浮動物件、合併儲存格處理
    /// </summary>
    public interface IExcelCellService
    {
        /// <summary>
        /// 取得儲存格的浮動物件
        /// </summary>
        List<FloatingObjectInfo>? GetCellFloatingObjects(ExcelWorksheet worksheet, ExcelRange cell);

        /// <summary>
        /// 取得繪圖物件類型
        /// </summary>
        string GetDrawingObjectType(ExcelDrawing drawing);

        /// <summary>
        /// 從繪圖物件提取文字
        /// </summary>
        string? ExtractTextFromDrawing(ExcelDrawing drawing);

        /// <summary>
        /// 從繪圖物件提取樣式
        /// </summary>
        string? ExtractStyleFromDrawing(ExcelDrawing drawing);

        /// <summary>
        /// 從繪圖物件提取超連結
        /// </summary>
        string? ExtractHyperlinkFromDrawing(ExcelDrawing drawing);

        /// <summary>
        /// 查找合併儲存格範圍 (返回字串地址)
        /// </summary>
        string? FindMergedRange(ExcelWorksheet worksheet, ExcelRange cell);

        /// <summary>
        /// 查找合併儲存格範圍 (返回 ExcelRange 物件,與 Controller 一致)
        /// </summary>
        ExcelRange? FindMergedRange(ExcelWorksheet worksheet, int row, int column);

        /// <summary>
        /// 取得合併儲存格的邊框
        /// </summary>
        BorderInfo? GetMergedCellBorder(ExcelWorksheet worksheet, string mergedRange);

        /// <summary>
        /// 設定儲存格合併資訊 (從已存在的合併範圍讀取)
        /// </summary>
        void SetCellMergedInfo(ExcelCellInfo cellInfo, ExcelWorksheet worksheet, ExcelRange cell);

        /// <summary>
        /// 設定儲存格合併資訊 (自動設定合併,與 Controller 一致)
        /// </summary>
        void SetCellMergedInfo(ExcelCellInfo cellInfo, int fromRow, int fromCol, int toRow, int toCol);

        /// <summary>
        /// 合併浮動物件文字 (批次處理)
        /// </summary>
        string MergeFloatingObjectText(string? cellText, List<FloatingObjectInfo>? floatingObjects);

        /// <summary>
        /// 合併浮動物件文字到儲存格 (單一處理,與 Controller 一致)
        /// </summary>
        void MergeFloatingObjectText(ExcelCellInfo cellInfo, string? floatingObjectText, string cellAddress);

        /// <summary>
        /// 在繪圖物件中查找圖片 (按座標)
        /// </summary>
        ExcelPicture? FindPictureInDrawings(ExcelWorksheet worksheet, int row, int column);

        /// <summary>
        /// 在繪圖物件中查找圖片 (按名稱,與 Controller 一致)
        /// </summary>
        ExcelPicture? FindPictureInDrawings(ExcelWorksheet worksheet, string imageName);

        /// <summary>
        /// 處理跨儲存格圖片 (與 Controller 一致)
        /// </summary>
        void ProcessImageCrossCells(ExcelCellInfo cellInfo, ExcelRange cell, ExcelWorksheet worksheet);

        /// <summary>
        /// 處理跨儲存格浮動物件 (與 Controller 一致)
        /// </summary>
        void ProcessFloatingObjectCrossCells(ExcelCellInfo cellInfo, ExcelRange cell);

        /// <summary>
        /// 取得文字對齊方式
        /// </summary>
        string GetTextAlign(OfficeOpenXml.Style.ExcelHorizontalAlignment alignment);

        /// <summary>
        /// 取得欄寬
        /// </summary>
        double GetColumnWidth(ExcelWorksheet worksheet, int column);

        /// <summary>
        /// 取得欄名稱
        /// </summary>
        string GetColumnName(int column);
    }
}
