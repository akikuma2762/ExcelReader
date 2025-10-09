using ExcelReaderAPI.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using ExcelReaderAPI.Models.Caches;

namespace ExcelReaderAPI.Services.Interfaces
{
    /// <summary>
    /// Excel 圖片處理服務介面 - 圖片轉換、查找、尺寸計算
    /// </summary>
    public interface IExcelImageService
    {
        /// <summary>
        /// 取得儲存格的圖片 (使用索引優化)
        /// </summary>
        List<ImageInfo>? GetCellImages(ExcelRange cell, WorksheetImageIndex imageIndex, ExcelWorksheet worksheet);

        /// <summary>
        /// 取得儲存格的圖片 (不使用索引)
        /// </summary>
        List<ImageInfo>? GetCellImages(ExcelWorksheet worksheet, ExcelRange cell);

        /// <summary>
        /// 轉換 EMF 圖片為 PNG
        /// </summary>
        byte[]? ConvertEmfToPng(byte[] emfData, int width, int height);

        /// <summary>
        /// 轉換圖片為 Base64
        /// </summary>
        string ConvertImageToBase64(byte[] imageData, string imageType);

        /// <summary>
        /// 從圖片物件取得圖片類型
        /// </summary>
        string GetImageTypeFromPicture(ExcelPicture picture);

        /// <summary>
        /// 取得圖片實際尺寸
        /// </summary>
        (int width, int height) GetActualImageDimensions(byte[] imageData, string imageType);

        /// <summary>
        /// 分析圖片資料取得尺寸
        /// </summary>
        (int width, int height)? AnalyzeImageDataDimensions(byte[] imageData);

        /// <summary>
        /// 分析 JPEG 圖片尺寸
        /// </summary>
        (int width, int height)? AnalyzeJpegDimensions(byte[] data);

        /// <summary>
        /// 檢查是否為 EMF 格式
        /// </summary>
        bool IsEmfFormat(byte[] data);

        /// <summary>
        /// 取得圖片格式
        /// </summary>
        string? GetImageFormat(byte[] imageData);

        /// <summary>
        /// 創建 EMF 佔位圖片
        /// </summary>
        string CreateEmfPlaceholderPng(int width, int height, string emfInfo);

        /// <summary>
        /// 創建 EMF 錯誤圖片
        /// </summary>
        string CreateEmfErrorPng();

        /// <summary>
        /// 根據 ID 查找嵌入圖片
        /// </summary>
        ExcelPicture? FindEmbeddedImageById(ExcelWorksheet worksheet, string imageId);

        /// <summary>
        /// 嘗試進階圖片搜尋
        /// </summary>
        ExcelPicture? TryAdvancedImageSearch(ExcelWorksheet worksheet, ExcelRange cell, string? searchImageId = null);

        /// <summary>
        /// 嘗試在所有工作表中查找圖片
        /// </summary>
        ExcelPicture? TryFindImageInWorksheets(ExcelWorksheet currentWorksheet, string imageId);

        /// <summary>
        /// 檢查所有圖片屬性
        /// </summary>
        void CheckAllPictureProperties(ExcelPicture picture, string context);

        /// <summary>
        /// 從圖片物件創建圖片資訊
        /// </summary>
        ImageInfo CreateImageInfoFromPicture(ExcelPicture picture, ExcelWorksheet worksheet);

        /// <summary>
        /// 嘗試在 VBA 專案中查找圖片
        /// </summary>
        ExcelPicture? TryFindImageInVbaProject(ExcelWorksheet worksheet, string imageId);

        /// <summary>
        /// 嘗試查找背景圖片
        /// </summary>
        ExcelPicture? TryFindBackgroundImage(ExcelWorksheet worksheet);

        /// <summary>
        /// 嘗試詳細繪圖物件搜尋
        /// </summary>
        ExcelPicture? TryDetailedDrawingSearch(ExcelWorksheet worksheet, ExcelRange cell);

        /// <summary>
        /// 檢查是否為部分 ID 匹配
        /// </summary>
        bool IsPartialIdMatch(string? pictureId, string searchId);

        /// <summary>
        /// 記錄可用的繪圖物件
        /// </summary>
        void LogAvailableDrawings(ExcelWorksheet worksheet, string context);

        /// <summary>
        /// 從名稱取得圖片類型
        /// </summary>
        string GetImageTypeFromName(string? name);

        /// <summary>
        /// 從檔案名稱取得圖片類型
        /// </summary>
        string GetImageTypeFromFileName(string fileName);

        /// <summary>
        /// 取得圖片類型 (綜合判斷)
        /// </summary>
        string GetImageType(ExcelPicture picture);

        /// <summary>
        /// 從 URI 取得圖片類型
        /// </summary>
        string? GetImageTypeFromUri(Uri? uri);

        /// <summary>
        /// 取得圖片檔案大小
        /// </summary>
        long GetImageFileSize(byte[] imageData);

        /// <summary>
        /// 產生佔位圖片
        /// </summary>
        string GeneratePlaceholderImage(int width, int height, string text);

        /// <summary>
        /// 檢查是否為 Base64 字串
        /// </summary>
        bool IsBase64String(string str);

        /// <summary>
        /// 取得儲存格像素尺寸
        /// </summary>
        (int width, int height) GetCellPixelDimensions(ExcelRange cell);

        /// <summary>
        /// 縮放圖片至儲存格尺寸
        /// </summary>
        (int width, int height) ScaleImageToCell(int imageWidth, int imageHeight, int cellWidth, int cellHeight);
    }
}
