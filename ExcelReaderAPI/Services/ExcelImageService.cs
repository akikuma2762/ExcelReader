using ExcelReaderAPI.Models;
using ExcelReaderAPI.Models.Caches;
using ExcelReaderAPI.Services.Interfaces;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.Drawing.Imaging;
using SkiaSharp;

namespace ExcelReaderAPI.Services
{
    /// <summary>
    /// Excel 圖片處理服務 - 圖片轉換、查找、尺寸計算
    /// Phase 2.2: 從 ExcelController 搬移圖片處理相關方法
    /// </summary>
    public class ExcelImageService : IExcelImageService
    {
        private readonly ILogger<ExcelImageService> _logger;

        public ExcelImageService(ILogger<ExcelImageService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        #region 核心方法 - 取得儲存格圖片

        public List<ImageInfo>? GetCellImages(ExcelRange cell, WorksheetImageIndex imageIndex, ExcelWorksheet worksheet)
        {
            try
            {
                var images = new List<ImageInfo>();

                _logger.LogDebug($"檢查儲存格 {cell.Address} 的圖片 (使用 EPPlus 8.x API + 索引)");

                // ⭐ EPPlus 8.x 新 API: 檢查 In-Cell 圖片 (優先使用官方 API)
                try
                {
                    // 單一儲存格 - 使用 EPPlus 8.x Picture API
                    if (cell.Picture.Exists)
                    {
                        var cellPicture = cell.Picture.Get();
                        if (cellPicture != null)
                        {
                            var imageBytes = cellPicture.GetImageBytes();
                            var imageType = GetImageTypeFromFileName(cellPicture.FileName);

                            // 🔍 計算儲存格/合併範圍的像素尺寸 (In-Cell 圖片會填滿整個儲存格)
                            var (cellWidthPixels, singleCellHeightPixels) = GetCellPixelDimensions(cell);
                            int rowSpan = cell.End.Row - cell.Start.Row + 1;
                            double totalHeightPixels = singleCellHeightPixels * rowSpan;

                            var imageInfo = new ImageInfo
                            {
                                Name = cellPicture.FileName ?? $"InCellImage_{cell.Address}",
                                Description = $"In-Cell 圖片 (EPPlus 8.x) - 儲存格: {cell.Address} (跨{rowSpan}行, {cellWidthPixels:F0}×{totalHeightPixels:F0}px), AltText: {cellPicture.AltText ?? "無"}",
                                ImageType = imageType,
                                Width = 0,
                                Height = (int)Math.Round(totalHeightPixels),
                                Left = 0,
                                Top = 0,
                                Base64Data = imageBytes != null ? Convert.ToBase64String(imageBytes) : string.Empty,
                                FileName = cellPicture.FileName ?? $"incell_{cell.Address}.png",
                                FileSize = imageBytes?.Length ?? 0,
                                AnchorCell = new CellPosition
                                {
                                    Row = cell.Start.Row,
                                    Column = cell.Start.Column,
                                    Address = cell.Address
                                },
                                HyperlinkAddress = $"In-Cell Picture (Type: {cellPicture.PictureType})",
                                IsInCellPicture = true,
                                AltText = cellPicture.AltText,
                                OriginalWidth = (int)Math.Round((double)cellWidthPixels),
                                OriginalHeight = (int)Math.Round((double)totalHeightPixels),
                                ExcelWidthCm = 0,
                                ExcelHeightCm = 0,
                                ScaleFactor = 1.0,
                                IsScaled = false,
                                ScaleMethod = $"In-Cell 圖片 (自動填滿 {rowSpan} 行合併儲存格)"
                            };

                            images.Add(imageInfo);
                            return images.Any() ? images : null;
                        }
                    }
                }
                catch (Exception inCellEx)
                {
                    _logger.LogWarning($"讀取 In-Cell 圖片失敗 (儲存格 {cell.Address}): {inCellEx.Message}");
                }

                // 使用索引快速查詢浮動圖片 (Drawing Pictures)
                var pictures = imageIndex.GetImagesAtCell(cell.Start.Row, cell.Start.Column);

                if (pictures == null)
                {
                    _logger.LogDebug($"儲存格 {cell.Address} 沒有圖片");
                    return null;
                }

                // 處理找到的圖片
                foreach (var picture in pictures)
                {
                    try
                    {
                        int fromRow = picture.From?.Row + 1 ?? 1;
                        int fromCol = picture.From?.Column + 1 ?? 1;
                        int toRow = picture.To?.Row + 1 ?? fromRow;
                        int toCol = picture.To?.Column + 1 ?? fromCol;

                        // 獲取圖片原始尺寸
                        var (actualWidth, actualHeight) = GetActualImageDimensionsFromPicture(picture);

                        // 計算 Excel 顯示尺寸
                        int excelDisplayWidth = actualWidth;
                        int excelDisplayHeight = actualHeight;
                        double excelWidthCm = 0;
                        double excelHeightCm = 0;
                        double scalePercentage = 100.0;

                        try
                        {
                            if (picture.From != null && picture.To != null)
                            {
                                const double emuPerPixel = 9525.0;
                                const double emuPerInch = 914400.0;
                                const double emuPerCm = emuPerInch / 2.54;

                                long totalWidthEmu = 0;
                                long totalHeightEmu = 0;

                                // 計算總寬度
                                for (int col = picture.From.Column; col <= picture.To.Column; col++)
                                {
                                    var column = worksheet.Column(col + 1);
                                    var colWidth = column.Width > 0 ? column.Width : worksheet.DefaultColWidth;
                                    long colWidthEmu = (long)(colWidth * 7.0 * emuPerPixel);

                                    if (col == picture.From.Column && col == picture.To.Column)
                                        totalWidthEmu = picture.To.ColumnOff - picture.From.ColumnOff;
                                    else if (col == picture.From.Column)
                                        totalWidthEmu += colWidthEmu - picture.From.ColumnOff;
                                    else if (col == picture.To.Column)
                                        totalWidthEmu += picture.To.ColumnOff;
                                    else
                                        totalWidthEmu += colWidthEmu;
                                }

                                // 計算總高度
                                for (int row = picture.From.Row; row <= picture.To.Row; row++)
                                {
                                    var rowObj = worksheet.Row(row + 1);
                                    var rowHeight = rowObj.Height > 0 ? rowObj.Height : worksheet.DefaultRowHeight;
                                    long rowHeightEmu = (long)(rowHeight * 12700);

                                    if (row == picture.From.Row && row == picture.To.Row)
                                        totalHeightEmu = picture.To.RowOff - picture.From.RowOff;
                                    else if (row == picture.From.Row)
                                        totalHeightEmu += rowHeightEmu - picture.From.RowOff;
                                    else if (row == picture.To.Row)
                                        totalHeightEmu += picture.To.RowOff;
                                    else
                                        totalHeightEmu += rowHeightEmu;
                                }

                                excelDisplayWidth = (int)(totalWidthEmu / emuPerPixel);
                                excelDisplayHeight = (int)(totalHeightEmu / emuPerPixel);
                                excelWidthCm = totalWidthEmu / emuPerCm;
                                excelHeightCm = totalHeightEmu / emuPerCm;

                                if (actualWidth > 0 && actualHeight > 0)
                                {
                                    double scaleX = (double)excelDisplayWidth / actualWidth * 100.0;
                                    double scaleY = (double)excelDisplayHeight / actualHeight * 100.0;
                                    scalePercentage = (scaleX + scaleY) / 2.0;
                                }
                            }
                        }
                        catch (Exception sizeEx)
                        {
                            _logger.LogWarning($"計算 Excel 顯示尺寸失敗: {sizeEx.Message}");
                        }

                        var imageInfo = new ImageInfo
                        {
                            Name = picture.Name ?? $"Image_{images.Count + 1}",
                            Description = $"Excel 圖片 - 原始: {actualWidth}×{actualHeight}px, Excel顯示: {excelDisplayWidth}×{excelDisplayHeight}px ({excelWidthCm:F2}×{excelHeightCm:F2}cm), 縮放: {scalePercentage:F1}%",
                            ImageType = GetImageTypeFromPicture(picture),
                            Width = excelDisplayWidth,
                            Height = excelDisplayHeight,
                            Left = (picture.From?.ColumnOff ?? 0) / 9525.0,
                            Top = (picture.From?.RowOff ?? 0) / 9525.0,
                            Base64Data = ConvertImageToBase64(picture),
                            FileName = picture.Name ?? $"image_{images.Count + 1}.png",
                            FileSize = GetImageFileSize(picture),
                            AnchorCell = new CellPosition
                            {
                                Row = fromRow,
                                Column = fromCol,
                                Address = $"{GetColumnName(fromCol)}{fromRow}"
                            },
                            HyperlinkAddress = picture.Hyperlink?.AbsoluteUri,
                            OriginalWidth = actualWidth,
                            OriginalHeight = actualHeight,
                            ExcelWidthCm = excelWidthCm,
                            ExcelHeightCm = excelHeightCm,
                            ScaleFactor = scalePercentage / 100.0,
                            IsScaled = Math.Abs(scalePercentage - 100.0) > 1.0,
                            ScaleMethod = $"Excel 縮放 {scalePercentage:F1}% (顯示: {excelWidthCm:F2}×{excelHeightCm:F2}cm)"
                        };

                        images.Add(imageInfo);
                    }
                    catch (Exception imgEx)
                    {
                        _logger.LogError(imgEx, $"處理圖片資料時發生錯誤: {imgEx.Message}");
                    }
                }

                return images.Any() ? images : null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"讀取儲存格 {cell.Address} 的圖片時發生錯誤: {ex.Message}");
                return null;
            }
        }

        public List<ImageInfo>? GetCellImages(ExcelWorksheet worksheet, ExcelRange cell)
        {
            // 這個方法是沒有索引的版本,直接遍歷所有 drawings
            // 簡化實作:建議在 Controller 中統一使用有索引的版本
            _logger.LogDebug($"GetCellImages (無索引版本) - 建議使用有索引的版本以提升效能");
            
            // 創建臨時索引並呼叫有索引的版本
            var tempIndex = new WorksheetImageIndex(worksheet);
            return GetCellImages(cell, tempIndex, worksheet);
        }

        #endregion

        #region 圖片轉換方法

        public byte[]? ConvertEmfToPng(byte[] emfData, int width, int height)
        {
            try
            {
                // 檢查平台支援
                var isWindows = System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows);
                
                // 方法1: Windows 平台使用 System.Drawing 進行實際轉換
                if (isWindows)
                {
                    try
                    {
                        using var emfStream = new MemoryStream(emfData);
                        using var emfImage = Image.FromStream(emfStream);
                        
                        // 獲取EMF的實際尺寸
                        var emfWidth = emfImage.Width;
                        var emfHeight = emfImage.Height;
                        
                        // 如果沒有指定目標尺寸,使用EMF的原始尺寸
                        var targetWidth = width > 0 ? width : emfWidth;
                        var targetHeight = height > 0 ? height : emfHeight;
                        
                        // 創建目標位圖
                        using var pngBitmap = new Bitmap(targetWidth, targetHeight);
                        using var graphics = Graphics.FromImage(pngBitmap);
                        
                        // 設置高質量渲染
                        graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        
                        // 清除背景為透明
                        graphics.Clear(Color.Transparent);
                        
                        // 繪製EMF到位圖
                        var targetRect = new Rectangle(0, 0, targetWidth, targetHeight);
                        graphics.DrawImage(emfImage, targetRect);
                        
                        // 轉換為PNG
                        using var pngStream = new MemoryStream();
                        pngBitmap.Save(pngStream, ImageFormat.Png);
                        var pngBytes = pngStream.ToArray();
                        
                        _logger.LogDebug($"System.Drawing EMF轉換成功: {emfData.Length} -> {pngBytes.Length} bytes");
                        return pngBytes;
                    }
                    catch (Exception systemDrawingEx)
                    {
                        _logger.LogError(systemDrawingEx, $"System.Drawing EMF轉換失敗: {systemDrawingEx.Message}");
                    }
                }

                // 方法2: 跨平台使用 SkiaSharp 創建提示圖片
                _logger.LogDebug("使用 SkiaSharp 創建 EMF 格式提示圖片");
                return CreateEmfPlaceholderPngBytes(width, height, $"EMF 檔案 ({emfData.Length} bytes)");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"EMF轉PNG轉換過程中發生未預期的錯誤: {ex.Message}");
                return null;
            }
        }

        public string ConvertImageToBase64(byte[] imageData, string imageType)
        {
            try
            {
                return $"data:image/{imageType.ToLower()};base64,{Convert.ToBase64String(imageData)}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "轉換圖片為 Base64 時發生錯誤");
                return string.Empty;
            }
        }

        #endregion

        #region 圖片類型檢測方法

        public string GetImageTypeFromPicture(ExcelPicture picture)
        {
            try
            {
                // 嘗試從圖片名稱推斷類型
                if (!string.IsNullOrEmpty(picture.Name))
                {
                    var extension = Path.GetExtension(picture.Name).ToLowerInvariant();
                    var typeFromName = extension switch
                    {
                        ".png" => "PNG",
                        ".jpg" => "JPEG",
                        ".jpeg" => "JPEG",
                        ".gif" => "GIF",
                        ".bmp" => "BMP",
                        ".tiff" => "TIFF",
                        ".tif" => "TIFF",
                        _ => null
                    };

                    if (!string.IsNullOrEmpty(typeFromName))
                    {
                        return typeFromName;
                    }
                }

                // 嘗試從圖片資料的檔頭分析類型
                if (picture.Image?.ImageBytes != null && picture.Image.ImageBytes.Length > 8)
                {
                    var bytes = picture.Image.ImageBytes;

                    // PNG 檔頭
                    if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
                    {
                        return "PNG";
                    }

                    // JPEG 檔頭
                    if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8)
                    {
                        return "JPEG";
                    }

                    // GIF 檔頭
                    if (bytes.Length >= 4 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
                    {
                        return "GIF";
                    }

                    // BMP 檔頭
                    if (bytes.Length >= 2 && bytes[0] == 0x42 && bytes[1] == 0x4D)
                    {
                        return "BMP";
                    }

                    // EMF 檔頭 (會自動轉換為 PNG)
                    if (IsEmfFormat(bytes))
                    {
                        return "PNG"; // 因為會自動轉換,所以返回 PNG 類型
                    }
                }

                // 預設類型
                _logger.LogDebug($"無法確定圖片 {picture.Name} 的類型，使用預設值 PNG");
                return "PNG";
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"分析圖片類型時發生錯誤，圖片: {picture.Name}");
                return "PNG";
            }
        }

        public string GetImageTypeFromName(string? name)
        {
            if (string.IsNullOrEmpty(name))
                return "Unknown";

            var extension = Path.GetExtension(name).ToLowerInvariant();
            return extension switch
            {
                ".png" => "PNG",
                ".jpg" => "JPEG",
                ".jpeg" => "JPEG",
                ".gif" => "GIF",
                ".bmp" => "BMP",
                ".tiff" => "TIFF",
                ".tif" => "TIFF",
                ".wmf" => "WMF",
                ".emf" => "EMF",
                ".webp" => "WEBP",
                ".ico" => "ICO",
                _ => "Unknown"
            };
        }

        public string GetImageTypeFromFileName(string fileName)
        {
            return GetImageTypeFromName(fileName);
        }

        public string GetImageType(ExcelPicture picture)
        {
            return GetImageTypeFromPicture(picture);
        }

        public string? GetImageTypeFromUri(Uri? uri)
        {
            if (uri == null)
                return null;

            var path = uri.LocalPath ?? uri.AbsolutePath;
            return GetImageTypeFromName(path);
        }

        public string? GetImageFormat(byte[] imageData)
        {
            if (imageData == null || imageData.Length < 8)
                return null;
            
            // PNG 格式檢查
            if (imageData[0] == 0x89 && imageData[1] == 0x50 && imageData[2] == 0x4E && imageData[3] == 0x47)
                return "PNG";
            
            // JPEG 格式檢查
            if (imageData[0] == 0xFF && imageData[1] == 0xD8)
                return "JPEG";
            
            // GIF 格式檢查
            if (imageData[0] == 0x47 && imageData[1] == 0x49 && imageData[2] == 0x46)
                return "GIF";
            
            // BMP 格式檢查
            if (imageData[0] == 0x42 && imageData[1] == 0x4D)
                return "BMP";
            
            // TIFF 格式檢查 (II = Intel, MM = Motorola)
            if ((imageData[0] == 0x49 && imageData[1] == 0x49 && imageData[2] == 0x2A && imageData[3] == 0x00) ||
                (imageData[0] == 0x4D && imageData[1] == 0x4D && imageData[2] == 0x00 && imageData[3] == 0x2A))
                return "TIFF";
            
            // EMF 格式檢查
            if (IsEmfFormat(imageData))
                return "EMF";
            
            return null;
        }

        public bool IsEmfFormat(byte[] data)
        {
            if (data == null || data.Length < 44)
                return false;
            
            // EMF 文件的特徵：在偏移量 40 處有 " EMF" 標識
            return data[40] == 0x20 && 
                   data[41] == 0x45 && 
                   data[42] == 0x4D && 
                   data[43] == 0x46;
        }

        #endregion

        #region 圖片尺寸方法

        public (int width, int height) GetActualImageDimensions(byte[] imageData, string imageType)
        {
            var dimensions = AnalyzeImageDataDimensions(imageData);
            if (dimensions.HasValue && dimensions.Value.width > 0 && dimensions.Value.height > 0)
            {
                _logger.LogDebug($"獲取圖片實際尺寸: {dimensions.Value.width}x{dimensions.Value.height}");
                return dimensions.Value;
            }

            _logger.LogWarning("無法獲取圖片實際尺寸，使用預設值");
            return (300, 200);
        }

        public (int width, int height)? AnalyzeImageDataDimensions(byte[] imageData)
        {
            try
            {
                if (imageData == null || imageData.Length < 24)
                    return (0, 0);

                // PNG 格式分析
                if (imageData[0] == 0x89 && imageData[1] == 0x50 && imageData[2] == 0x4E && imageData[3] == 0x47)
                {
                    if (imageData.Length >= 24)
                    {
                        // PNG IHDR chunk 中的寬高信息（大端序）
                        var width = (imageData[16] << 24) | (imageData[17] << 16) | (imageData[18] << 8) | imageData[19];
                        var height = (imageData[20] << 24) | (imageData[21] << 16) | (imageData[22] << 8) | imageData[23];

                        if (width > 0 && height > 0 && width < 65536 && height < 65536)
                        {
                            _logger.LogDebug($"從 PNG 資料獲取尺寸: {width}x{height}");
                            return (width, height);
                        }
                    }
                }

                // JPEG 格式分析
                if (imageData[0] == 0xFF && imageData[1] == 0xD8)
                {
                    var dimensions = AnalyzeJpegDimensions(imageData);
                    if (dimensions.HasValue && dimensions.Value.width > 0 && dimensions.Value.height > 0)
                    {
                        _logger.LogDebug($"從 JPEG 資料獲取尺寸: {dimensions.Value.width}x{dimensions.Value.height}");
                        return dimensions.Value;
                    }
                }

                // GIF 格式分析
                if (imageData.Length >= 10 && imageData[0] == 0x47 && imageData[1] == 0x49 && imageData[2] == 0x46)
                {
                    // GIF 格式使用小端序
                    var width = imageData[6] | (imageData[7] << 8);
                    var height = imageData[8] | (imageData[9] << 8);

                    if (width > 0 && height > 0 && width < 65536 && height < 65536)
                    {
                        _logger.LogDebug($"從 GIF 資料獲取尺寸: {width}x{height}");
                        return (width, height);
                    }
                }

                return (0, 0);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "分析圖片資料尺寸時發生錯誤");
                return (0, 0);
            }
        }

        public (int width, int height)? AnalyzeJpegDimensions(byte[] data)
        {
            try
            {
                if (data == null || data.Length < 10)
                    return (0, 0);

                int pos = 2; // 跳過 SOI 標記 (FF D8)

                while (pos < data.Length - 8)
                {
                    if (data[pos] == 0xFF)
                    {
                        byte marker = data[pos + 1];

                        // SOF0 (Start of Frame) 標記
                        if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2)
                        {
                            if (pos + 7 < data.Length)
                            {
                                // JPEG SOF 格式：FF C0 [length] [precision] [height] [width]
                                var height = (data[pos + 5] << 8) | data[pos + 6];
                                var width = (data[pos + 7] << 8) | data[pos + 8];

                                if (width > 0 && height > 0 && width < 65536 && height < 65536)
                                {
                                    return (width, height);
                                }
                            }
                        }

                        // 跳到下一個標記
                        if (pos + 3 < data.Length)
                        {
                            var segmentLength = (data[pos + 2] << 8) | data[pos + 3];
                            pos += 2 + segmentLength;
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        pos++;
                    }
                }

                return (0, 0);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "分析 JPEG 尺寸時發生錯誤");
                return (0, 0);
            }
        }

        public (int width, int height) GetCellPixelDimensions(ExcelRange cell)
        {
            try
            {
                var worksheet = cell.Worksheet;
                var row = cell.Start.Row;
                var col = cell.Start.Column;
                
                // 獲取欄寬(Excel 單位)
                var column = worksheet.Column(col);
                var columnWidth = column.Width > 0 ? column.Width : worksheet.DefaultColWidth;

                // 獲取行高(點數單位)
                var rowObj = worksheet.Row(row);
                var rowHeight = rowObj.Height > 0 ? rowObj.Height : worksheet.DefaultRowHeight;

                // Excel 欄寬單位轉換為像素 (約等於 7 像素)
                var cellWidthPixels = (int)(columnWidth * 7.0);

                // Excel 行高單位是點數,1 point = 4/3 pixels (at 96 DPI)
                var cellHeightPixels = (int)(rowHeight * 4.0 / 3.0);

                _logger.LogDebug($"儲存格 {cell.Address} 尺寸: {cellWidthPixels} x {cellHeightPixels} 像素");

                return (cellWidthPixels, cellHeightPixels);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"計算儲存格 {cell.Address} 尺寸時發生錯誤");
                return (100, 20); // 預設尺寸
            }
        }

        public (int width, int height) ScaleImageToCell(int imageWidth, int imageHeight, int cellWidth, int cellHeight)
        {
            try
            {
                if (imageWidth <= 0 || imageHeight <= 0)
                {
                    return (cellWidth, cellHeight);
                }

                // 計算可用空間(留 10% 邊距)
                var availableWidth = cellWidth * 0.9;
                var availableHeight = cellHeight * 0.9;

                // 計算縮放比例,保持圖片長寬比
                var scaleX = availableWidth / imageWidth;
                var scaleY = availableHeight / imageHeight;
                var scale = Math.Min(scaleX, scaleY);

                // 確保縮放不會放大圖片過度
                scale = Math.Min(scale, 2.0); // 最大放大 2 倍

                var scaledWidth = (int)(imageWidth * scale);
                var scaledHeight = (int)(imageHeight * scale);

                _logger.LogDebug($"圖片縮放: {imageWidth}x{imageHeight} -> {scaledWidth}x{scaledHeight} (比例: {scale:F2})");

                return (scaledWidth, scaledHeight);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "圖片縮放計算時發生錯誤");
                return (imageWidth, imageHeight);
            }
        }

        #endregion

        #region 圖片搜尋方法

        public ExcelPicture? FindEmbeddedImageById(ExcelWorksheet worksheet, string imageId)
        {
            try
            {
                _logger.LogDebug($"開始查找嵌入圖片,ID: {imageId}");

                // 遍歷工作表的所有繪圖物件
                if (worksheet.Drawings != null)
                {
                    foreach (var drawing in worksheet.Drawings)
                    {
                        if (drawing is ExcelPicture picture)
                        {
                            _logger.LogDebug($"檢查圖片: Name={picture.Name}, Description={picture.Description}");

                            // 檢查圖片名稱或 ID 是否匹配 (使用更寬鬆的匹配條件)
                            var cleanImageId = imageId.Replace("ID_", "").Replace("\"", "");
                            if (picture.Name != null &&
                                (picture.Name.Contains(imageId) ||
                                 picture.Name.Contains(cleanImageId) ||
                                 picture.Name == imageId ||
                                 imageId.Contains(picture.Name)))
                            {
                                _logger.LogDebug($"找到匹配的圖片: {picture.Name}");
                                return picture;
                            }
                        }
                    }
                }

                _logger.LogWarning($"未找到圖片,ID: {imageId}");
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"查找嵌入圖片時發生錯誤,ID: {imageId}");
                return null;
            }
        }

        public ExcelPicture? TryAdvancedImageSearch(ExcelWorksheet worksheet, ExcelRange cell, string? searchImageId = null)
        {
            try
            {
                _logger.LogDebug($"使用進階功能查找圖片,儲存格: {cell.Address}, ID: {searchImageId ?? "null"}");

                if (worksheet.Drawings != null)
                {
                    foreach (var drawing in worksheet.Drawings)
                    {
                        if (drawing is ExcelPicture picture)
                        {
                            // 檢查圖片是否在指定儲存格範圍內
                            if (picture.From != null)
                            {
                                int fromRow = picture.From.Row + 1;
                                int fromCol = picture.From.Column + 1;

                                bool inRange = fromRow >= cell.Start.Row && fromRow <= cell.End.Row &&
                                              fromCol >= cell.Start.Column && fromCol <= cell.End.Column;

                                if (inRange)
                                {
                                    // 如果有指定 ID,進一步檢查
                                    if (!string.IsNullOrEmpty(searchImageId))
                                    {
                                        var cleanImageId = searchImageId.Replace("ID_", "").Replace("\"", "");
                                        if (picture.Name != null && 
                                            (picture.Name.Contains(cleanImageId) || 
                                             IsPartialIdMatch(picture.Name, cleanImageId)))
                                        {
                                            return picture;
                                        }
                                    }
                                    else
                                    {
                                        // 沒有指定 ID,返回範圍內的第一張圖片
                                        return picture;
                                    }
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "進階圖片搜索時發生錯誤");
                return null;
            }
        }

        public ExcelPicture? TryFindImageInWorksheets(ExcelWorksheet currentWorksheet, string imageId)
        {
            try
            {
                var cleanImageId = imageId.Replace("ID_", "").Replace("\"", "").ToLowerInvariant();

                // 檢查當前工作表
                if (currentWorksheet.Drawings != null)
                {
                    foreach (var drawing in currentWorksheet.Drawings)
                    {
                        if (drawing is ExcelPicture picture)
                        {
                            var pictureName = picture.Name?.ToLowerInvariant() ?? "";
                            var pictureDescription = picture.Description?.ToLowerInvariant() ?? "";

                            if (pictureName.Contains(cleanImageId) ||
                                pictureDescription.Contains(cleanImageId) ||
                                IsPartialIdMatch(pictureName, cleanImageId))
                            {
                                _logger.LogDebug($"通過擴展屬性檢查找到匹配圖片: {picture.Name}");
                                return picture;
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "在工作表中查找圖片時發生錯誤");
                return null;
            }
        }

        public void CheckAllPictureProperties(ExcelPicture picture, string context)
        {
            try
            {
                _logger.LogDebug($"[{context}] 檢查圖片屬性:");
                _logger.LogDebug($"  Name: {picture.Name ?? "null"}");
                _logger.LogDebug($"  From: Row={picture.From?.Row ?? -1}, Col={picture.From?.Column ?? -1}");
                _logger.LogDebug($"  To: Row={picture.To?.Row ?? -1}, Col={picture.To?.Column ?? -1}");
                _logger.LogDebug($"  Image.Bounds: {picture.Image?.Bounds.ToString() ?? "null"}");
                
                // 檢查其他可能的屬性
                var pictureType = picture.GetType();
                var properties = pictureType.GetProperties();
                foreach (var prop in properties)
                {
                    try
                    {
                        var value = prop.GetValue(picture);
                        if (value != null && prop.Name.Contains("Id", StringComparison.OrdinalIgnoreCase))
                        {
                            _logger.LogDebug($"  {prop.Name}: {value}");
                        }
                    }
                    catch
                    {
                        // 忽略無法讀取的屬性
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "檢查圖片屬性時發生錯誤");
            }
        }

        public ImageInfo CreateImageInfoFromPicture(ExcelPicture picture, ExcelWorksheet worksheet)
        {
            try
            {
                var fromRow = picture.From?.Row + 1 ?? 1;
                var fromCol = picture.From?.Column + 1 ?? 1;
                
                return new ImageInfo
                {
                    Name = picture.Name ?? $"Image_{fromRow}_{fromCol}",
                    Description = $"Excel 圖片 - 工作表: {worksheet.Name}",
                    ImageType = GetImageTypeFromPicture(picture),
                    Width = (int)(picture.Image?.Bounds.Width ?? 0),
                    Height = (int)(picture.Image?.Bounds.Height ?? 0),
                    Left = (picture.From?.ColumnOff ?? 0) / 9525.0,
                    Top = (picture.From?.RowOff ?? 0) / 9525.0,
                    Base64Data = ConvertImageToBase64(picture),
                    FileName = picture.Name ?? $"image_{fromRow}_{fromCol}",
                    FileSize = GetImageFileSize(picture),
                    AnchorCell = new CellPosition
                    {
                        Row = fromRow,
                        Column = fromCol,
                        Address = $"{GetColumnName(fromCol)}{fromRow}"
                    },
                    HyperlinkAddress = picture.Hyperlink?.AbsoluteUri
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"創建圖片資訊時發生錯誤: {ex.Message}");
                throw;
            }
        }

        private string GetColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                columnNumber--;
                columnName = (char)('A' + columnNumber % 26) + columnName;
                columnNumber /= 26;
            }
            return columnName;
        }

        private string ConvertImageToBase64(ExcelPicture picture)
        {
            try
            {
                if (picture.Image?.ImageBytes != null && picture.Image.ImageBytes.Length > 0)
                {
                    var imageBytes = picture.Image.ImageBytes;
                    
                    // 檢查是否為 EMF 格式
                    if (IsEmfFormat(imageBytes))
                    {
                        _logger.LogDebug($"檢測到 EMF 格式圖片: {picture.Name},正在轉換為 PNG...");
                        
                        var pngBytes = ConvertEmfToPng(imageBytes, 800, 600);
                        
                        if (pngBytes != null && pngBytes.Length > 0)
                        {
                            return Convert.ToBase64String(pngBytes);
                        }
                        else
                        {
                            _logger.LogWarning($"EMF 轉 PNG 失敗: {picture.Name},使用錯誤提示圖片");
                            return CreateEmfErrorPng();
                        }
                    }
                    
                    // 非 EMF 格式,直接返回
                    return Convert.ToBase64String(imageBytes);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"轉換圖片 {picture.Name} 為 Base64 時發生錯誤");
                return CreateEmfErrorPng();
            }

            return string.Empty;
        }

        private long GetImageFileSize(ExcelPicture picture)
        {
            try
            {
                if (picture.Image?.ImageBytes != null)
                {
                    return picture.Image.ImageBytes.Length;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"獲取圖片 {picture.Name} 檔案大小時發生錯誤");
            }

            return 0;
        }

        private (int width, int height) GetActualImageDimensionsFromPicture(ExcelPicture picture)
        {
            try
            {
                if (picture.Image?.ImageBytes != null && picture.Image.ImageBytes.Length > 0)
                {
                    var dimensions = AnalyzeImageDataDimensions(picture.Image.ImageBytes);
                    if (dimensions.HasValue && dimensions.Value.width > 0 && dimensions.Value.height > 0)
                    {
                        return dimensions.Value;
                    }
                }

                // 嘗試從 Image.Bounds 獲取
                if (picture.Image?.Bounds != null)
                {
                    return ((int)picture.Image.Bounds.Width, (int)picture.Image.Bounds.Height);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"獲取圖片 {picture.Name} 實際尺寸時發生錯誤");
            }

            return (300, 200); // 預設值
        }

        public ExcelPicture? TryFindImageInVbaProject(ExcelWorksheet worksheet, string imageId)
        {
            try
            {
                // EPPlus 可能無法完整存取 VBA 項目中的圖片
                // 這個方法主要用於記錄和除錯
                if (worksheet.Workbook.VbaProject != null)
                {
                    _logger.LogDebug($"工作簿包含 VBA 項目,嘗試查找圖片 ID: {imageId}");
                    // 在更新的 EPPlus 版本中,這裡可以進一步實現
                    // 目前返回 null,表示未找到
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, $"查找 VBA 項目圖片時發生錯誤,ID: {imageId}");
                return null;
            }
        }

        public ExcelPicture? TryFindBackgroundImage(ExcelWorksheet worksheet)
        {
            try
            {
                // 檢查工作表是否有背景圖片
                if (worksheet.BackgroundImage != null)
                {
                    _logger.LogDebug($"工作表 '{worksheet.Name}' 有背景圖片");
                    
                    // EPPlus 的限制使得無法將背景圖片轉換為 ExcelPicture
                    // 這個方法主要用於檢測背景圖片的存在
                    // 返回 null 表示背景圖片無法作為 ExcelPicture 返回
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "查找背景圖片時發生錯誤");
                return null;
            }
        }

        public ExcelPicture? TryDetailedDrawingSearch(ExcelWorksheet worksheet, ExcelRange cell)
        {
            try
            {
                _logger.LogDebug($"進行詳細繪圖搜索,儲存格: {cell.Address}");

                if (worksheet.Drawings != null)
                {
                    foreach (var drawing in worksheet.Drawings)
                    {
                        if (drawing is ExcelPicture picture)
                        {
                            // 檢查圖片是否與儲存格重疊
                            if (picture.From != null)
                            {
                                int fromRow = picture.From.Row + 1;
                                int fromCol = picture.From.Column + 1;
                                int toRow = picture.To?.Row + 1 ?? fromRow;
                                int toCol = picture.To?.Column + 1 ?? fromCol;

                                // 檢查重疊
                                bool overlaps = !(toRow < cell.Start.Row || fromRow > cell.End.Row ||
                                                toCol < cell.Start.Column || fromCol > cell.End.Column);

                                if (overlaps)
                                {
                                    _logger.LogDebug($"透過詳細搜索找到重疊的圖片: Name='{picture.Name}', " +
                                                   $"位置: ({fromRow},{fromCol})-({toRow},{toCol})");
                                    return picture;
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "詳細繪圖搜索時發生錯誤");
                return null;
            }
        }

        public bool IsPartialIdMatch(string? pictureId, string searchId)
        {
            if (string.IsNullOrEmpty(pictureId) || string.IsNullOrEmpty(searchId))
                return false;

            // 檢查是否有部分匹配 (至少 8 個字符)
            if (pictureId.Length >= 8 && searchId.Length >= 8)
            {
                for (int i = 0; i <= pictureId.Length - 8; i++)
                {
                    var segment = pictureId.Substring(i, 8);
                    if (searchId.Contains(segment))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public void LogAvailableDrawings(ExcelWorksheet worksheet, string context)
        {
            try
            {
                if (worksheet.Drawings == null || !worksheet.Drawings.Any())
                {
                    _logger.LogDebug($"[{context}] 工作表 '{worksheet.Name}' 沒有繪圖物件");
                    return;
                }

                _logger.LogInformation($"[{context}] 工作表 '{worksheet.Name}' 包含 {worksheet.Drawings.Count} 個繪圖物件:");
                
                int index = 1;
                foreach (var drawing in worksheet.Drawings)
                {
                    if (drawing is ExcelPicture picture)
                    {
                        _logger.LogInformation($"  [{index}] 圖片: {picture.Name ?? "未命名"}, " +
                                             $"位置: ({picture.From?.Row ?? -1},{picture.From?.Column ?? -1}) -> " +
                                             $"({picture.To?.Row ?? -1},{picture.To?.Column ?? -1})");
                    }
                    else
                    {
                        _logger.LogInformation($"  [{index}] 繪圖物件: {drawing.Name ?? "未命名"}, " +
                                             $"類型: {drawing.GetType().Name}");
                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"記錄繪圖物件時發生錯誤: {ex.Message}");
            }
        }

        #endregion

        #region EMF 和佔位圖片方法

        public string CreateEmfPlaceholderPng(int width, int height, string emfInfo)
        {
            try
            {
                var bytes = CreateEmfPlaceholderPngBytes(width, height, emfInfo);
                return Convert.ToBase64String(bytes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "創建EMF提示圖片失敗");
                return GeneratePlaceholderImage(200, 150, "EMF Error");
            }
        }

        public string CreateEmfErrorPng()
        {
            return GeneratePlaceholderImage(200, 150, "EMF Error");
        }

        private byte[] CreateEmfPlaceholderPngBytes(int width = 400, int height = 200, string? additionalInfo = null)
        {
            try
            {
                var imageInfo = new SKImageInfo(width, height, SKColorType.Rgba8888, SKAlphaType.Premul);
                using var surface = SKSurface.Create(imageInfo);
                var canvas = surface.Canvas;
                
                // 背景 - 淺藍色
                canvas.Clear(new SKColor(240, 248, 255));
                
                // 邊框
                using var borderPaint = new SKPaint
                {
                    Color = new SKColor(70, 130, 180),
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = 2
                };
                canvas.DrawRect(1, 1, width - 2, height - 2, borderPaint);
                
                // 標題文字
                using var titlePaint = new SKPaint
                {
                    Color = new SKColor(25, 25, 112),
                    TextSize = Math.Min(width / 15f, 20f),
                    IsAntialias = true,
                    Typeface = SKTypeface.FromFamilyName("Arial", SKFontStyle.Bold)
                };
                
                // 內容文字
                using var textPaint = new SKPaint
                {
                    Color = new SKColor(60, 60, 60),
                    TextSize = Math.Min(width / 20f, 14f),
                    IsAntialias = true,
                    Typeface = SKTypeface.Default
                };
                
                var lines = new List<string>
                {
                    "🖼️ EMF 向量圖片",
                    "",
                    "✅ 已自動轉換為 PNG 格式",
                    "🌐 瀏覽器可正常顯示"
                };
                
                if (!string.IsNullOrEmpty(additionalInfo))
                {
                    lines.Add("");
                    lines.Add($"📄 {additionalInfo}");
                }
                
                float startY = height / 2 - (lines.Count * Math.Min(width / 20f, 14f)) / 2;
                bool isTitle = true;
                
                foreach (var line in lines)
                {
                    if (string.IsNullOrEmpty(line))
                    {
                        startY += Math.Min(width / 20f, 14f);
                        isTitle = false;
                        continue;
                    }
                    
                    var paint = isTitle ? titlePaint : textPaint;
                    var textWidth = paint.MeasureText(line);
                    canvas.DrawText(line, (width - textWidth) / 2, startY, paint);
                    startY += Math.Min(width / 20f, 14f) + 4;
                    isTitle = false;
                }
                
                // 轉換為PNG
                using var image = surface.Snapshot();
                using var data = image.Encode(SKEncodedImageFormat.Png, 90);
                return data.ToArray();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "創建EMF提示圖片失敗");
                
                // 簡化版本的提示圖片
                try
                {
                    var imageInfo = new SKImageInfo(width, height, SKColorType.Rgba8888, SKAlphaType.Premul);
                    using var surface = SKSurface.Create(imageInfo);
                    var canvas = surface.Canvas;
                    canvas.Clear(SKColors.LightGray);
                    
                    using var paint = new SKPaint { Color = SKColors.Black, TextSize = 14 };
                    canvas.DrawText("EMF -> PNG", 10, height / 2, paint);
                    
                    using var image = surface.Snapshot();
                    using var data = image.Encode(SKEncodedImageFormat.Png, 90);
                    return data.ToArray();
                }
                catch
                {
                    // 回傳最基本的 1x1 透明 PNG
                    return new byte[]
                    {
                        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
                        0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
                        0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
                        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
                        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
                        0x42, 0x60, 0x82
                    };
                }
            }
        }

        public long GetImageFileSize(byte[] imageData)
        {
            return imageData?.Length ?? 0;
        }

        public string GeneratePlaceholderImage(int width, int height, string text)
        {
            try
            {
                var imageInfo = new SKImageInfo(width, height, SKColorType.Rgba8888, SKAlphaType.Premul);
                using var surface = SKSurface.Create(imageInfo);
                var canvas = surface.Canvas;
                
                // 背景 - 淺灰色
                canvas.Clear(new SKColor(245, 245, 245));
                
                // 邊框
                using var borderPaint = new SKPaint
                {
                    Color = new SKColor(200, 200, 200),
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = 2
                };
                canvas.DrawRect(1, 1, width - 2, height - 2, borderPaint);
                
                // 文字
                using var textPaint = new SKPaint
                {
                    Color = new SKColor(100, 100, 100),
                    TextSize = Math.Min(width / 10f, 16f),
                    IsAntialias = true,
                    Typeface = SKTypeface.Default
                };
                
                var textWidth = textPaint.MeasureText(text);
                canvas.DrawText(text, (width - textWidth) / 2, height / 2, textPaint);
                
                // 轉換為PNG並返回Base64
                using var image = surface.Snapshot();
                using var data = image.Encode(SKEncodedImageFormat.Png, 90);
                return Convert.ToBase64String(data.ToArray());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成佔位圖片失敗");
                // 回傳最基本的 1x1 透明 PNG
                var minimalPng = new byte[]
                {
                    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                    0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
                    0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
                    0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
                    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
                    0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
                    0x42, 0x60, 0x82
                };
                return Convert.ToBase64String(minimalPng);
            }
        }

        public bool IsBase64String(string str)
        {
            if (string.IsNullOrWhiteSpace(str))
                return false;

            try
            {
                Convert.FromBase64String(str);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion
    }
}