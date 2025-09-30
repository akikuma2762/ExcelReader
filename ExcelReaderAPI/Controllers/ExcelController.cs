using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using ExcelReaderAPI.Models;
using ExcelReaderAPI.Utils;
using System.Data;
using System.IO.Packaging;
using System.IO.Compression;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace ExcelReaderAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }

        static ExcelController()
        {
            // 設定EPPlus授權（非商業用途）
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private ExcelRange? FindMergedRange(ExcelWorksheet worksheet, int row, int column)
        {
            // 檢查所有合併範圍，找到包含指定儲存格的範圍
            foreach (var mergedRange in worksheet.MergedCells)
            {
                var range = worksheet.Cells[mergedRange];
                if (row >= range.Start.Row && row <= range.End.Row &&
                    column >= range.Start.Column && column <= range.End.Column)
                {
                    return range;
                }
            }
            return null;
        }

        private string? GetTextAlign(OfficeOpenXml.Style.ExcelHorizontalAlignment alignment)
        {
            return alignment switch
            {
                OfficeOpenXml.Style.ExcelHorizontalAlignment.Left => "left",
                OfficeOpenXml.Style.ExcelHorizontalAlignment.Center => "center",
                OfficeOpenXml.Style.ExcelHorizontalAlignment.Right => "right",
                OfficeOpenXml.Style.ExcelHorizontalAlignment.Justify => "justify",
                OfficeOpenXml.Style.ExcelHorizontalAlignment.Fill => "left",
                OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous => "center",
                OfficeOpenXml.Style.ExcelHorizontalAlignment.Distributed => "justify",
                _ => null
            };
        }

        private double GetColumnWidth(ExcelWorksheet worksheet, int columnIndex)
        {
            // 取得該欄的寬度，若未設定則使用預設寬度
            var column = worksheet.Column(columnIndex);
            if (column.Width > 0)
            {
                return column.Width;
            }
            else
            {
                // 使用預設欄寬
                return worksheet.DefaultColWidth;
            }
        }

        private ExcelCellInfo CreateCellInfo(ExcelRange cell, ExcelWorksheet worksheet)
        {
            if (cell == null || worksheet == null)
                throw new ArgumentNullException("Cell or worksheet cannot be null");

            var cellInfo = new ExcelCellInfo();

            try
            {
                // 位置資訊
                cellInfo.Position = new CellPosition
                {
                    Row = cell.Start.Row,
                    Column = cell.Start.Column,
                    Address = cell.Address ?? $"{GetColumnName(cell.Start.Column)}{cell.Start.Row}"
                };

            // 基本值和顯示
            cellInfo.Value = cell.Value;
            cellInfo.Text = cell.Text;
            cellInfo.Formula = cell.Formula;
            cellInfo.FormulaR1C1 = cell.FormulaR1C1;

            // 資料類型
            cellInfo.ValueType = cell.Value?.GetType().Name;
            if (cell.Value == null)
            {
                cellInfo.DataType = "Empty";
            }
            else if (cell.Value is DateTime)
            {
                cellInfo.DataType = "DateTime";
            }
            else if (cell.Value is double || cell.Value is float || cell.Value is decimal)
            {
                cellInfo.DataType = "Number";
            }
            else if (cell.Value is int || cell.Value is long || cell.Value is short)
            {
                cellInfo.DataType = "Integer";
            }
            else if (cell.Value is bool)
            {
                cellInfo.DataType = "Boolean";
            }
            else
            {
                cellInfo.DataType = "Text";
            }

            // 格式化
            cellInfo.NumberFormat = cell.Style.Numberformat.Format;
            cellInfo.NumberFormatId = cell.Style.Numberformat.NumFmtID;

            // 字體樣式
            cellInfo.Font = new FontInfo
            {
                Name = cell.Style.Font.Name,
                Size = cell.Style.Font.Size,
                Bold = cell.Style.Font.Bold,
                Italic = cell.Style.Font.Italic,
                UnderLine = cell.Style.Font.UnderLine.ToString(),
                Strike = cell.Style.Font.Strike,
                Color = GetColorFromExcelColor(cell.Style.Font.Color),
                ColorTheme = cell.Style.Font.Color.Theme?.ToString(),
                ColorTint = (double?)cell.Style.Font.Color.Tint,
                Charset = cell.Style.Font.Charset,
                Scheme = cell.Style.Font.Scheme?.ToString(),
                Family = cell.Style.Font.Family
            };

            // 對齊方式
            cellInfo.Alignment = new AlignmentInfo
            {
                Horizontal = cell.Style.HorizontalAlignment.ToString(),
                Vertical = cell.Style.VerticalAlignment.ToString(),
                WrapText = cell.Style.WrapText,
                Indent = cell.Style.Indent,
                ReadingOrder = cell.Style.ReadingOrder.ToString(),
                TextRotation = cell.Style.TextRotation,
                ShrinkToFit = cell.Style.ShrinkToFit
            };

            // 邊框
            // 邊框設定 - 使用增強的顏色處理
            cellInfo.Border = new BorderInfo
            {
                Top = new BorderStyle 
                { 
                    Style = cell.Style.Border.Top.Style.ToString(), 
                    Color = GetColorFromExcelColor(cell.Style.Border.Top.Color)
                },
                Bottom = new BorderStyle 
                { 
                    Style = cell.Style.Border.Bottom.Style.ToString(), 
                    Color = GetColorFromExcelColor(cell.Style.Border.Bottom.Color)
                },
                Left = new BorderStyle 
                { 
                    Style = cell.Style.Border.Left.Style.ToString(), 
                    Color = GetColorFromExcelColor(cell.Style.Border.Left.Color)
                },
                Right = new BorderStyle 
                { 
                    Style = cell.Style.Border.Right.Style.ToString(), 
                    Color = GetColorFromExcelColor(cell.Style.Border.Right.Color)
                },
                Diagonal = new BorderStyle 
                { 
                    Style = cell.Style.Border.Diagonal.Style.ToString(), 
                    Color = GetColorFromExcelColor(cell.Style.Border.Diagonal.Color)
                },
                DiagonalUp = cell.Style.Border.DiagonalUp,
                DiagonalDown = cell.Style.Border.DiagonalDown
            };

            // 填充/背景
            cellInfo.Fill = new FillInfo
            {
                PatternType = cell.Style.Fill.PatternType.ToString(),
                BackgroundColor = GetBackgroundColor(cell),
                PatternColor = cell.Style.Fill.PatternColor.Rgb,
                BackgroundColorTheme = cell.Style.Fill.BackgroundColor.Theme?.ToString(),
                BackgroundColorTint = (double?)cell.Style.Fill.BackgroundColor.Tint
            };

            // 尺寸和合併
            var column = worksheet.Column(cell.Start.Column);
            cellInfo.Dimensions = new DimensionInfo
            {
                ColumnWidth = column.Width > 0 ? column.Width : worksheet.DefaultColWidth,
                RowHeight = worksheet.Row(cell.Start.Row).Height,
                IsMerged = cell.Merge
            };

            // 檢查是否為合併儲存格
            if (cell.Merge)
            {
                var mergedRange = FindMergedRange(worksheet, cell.Start.Row, cell.Start.Column);
                if (mergedRange != null)
                {
                    cellInfo.Dimensions.MergedRangeAddress = mergedRange.Address;
                    cellInfo.Dimensions.IsMainMergedCell = (cell.Start.Row == mergedRange.Start.Row && 
                                                           cell.Start.Column == mergedRange.Start.Column);
                    
                    if (cellInfo.Dimensions.IsMainMergedCell == true)
                    {
                        cellInfo.Dimensions.RowSpan = mergedRange.Rows;
                        cellInfo.Dimensions.ColSpan = mergedRange.Columns;
                        
                        // 對於主合併儲存格，使用整個合併範圍的邊框
                        cellInfo.Border = GetMergedCellBorder(worksheet, mergedRange, cell);
                    }
                    else
                    {
                        cellInfo.Dimensions.RowSpan = 1;
                        cellInfo.Dimensions.ColSpan = 1;
                    }
                }
            }

            // Rich Text
            if (cell.IsRichText && cell.RichText != null && cell.RichText.Count > 0)
            {
                cellInfo.RichText = new List<RichTextPart>();
                
                for (int i = 0; i < cell.RichText.Count; i++)
                {
                    var richTextPart = cell.RichText[i];
                    
                    // 修正第一個 Rich Text 部分的格式問題
                    // EPPlus 的第一個 Rich Text 部分經常缺少格式資訊，需要從儲存格樣式繼承
                    var bold = richTextPart.Bold;
                    var italic = richTextPart.Italic;
                    var size = richTextPart.Size;
                    var fontName = richTextPart.FontName;
                    var color = richTextPart.Color;
                    
                    // 如果第一個 Rich Text 部分沒有格式資訊，從儲存格樣式繼承
                    if (i == 0)
                    {
                        if (size == 0 || string.IsNullOrEmpty(fontName) || (!bold && !italic))
                        {
                            size = size == 0 ? cell.Style.Font.Size : size;
                            fontName = string.IsNullOrEmpty(fontName) ? cell.Style.Font.Name : fontName;
                            
                            // 只有當 Rich Text 部分沒有設定格式時才繼承
                            if (!richTextPart.Bold && cell.Style.Font.Bold)
                                bold = true;
                            if (!richTextPart.Italic && cell.Style.Font.Italic)
                                italic = true;
                        }
                    }
                    
                    cellInfo.RichText.Add(new RichTextPart
                    {
                        Text = richTextPart.Text,
                        Bold = bold,
                        Italic = italic,
                        UnderLine = richTextPart.UnderLine,
                        Strike = richTextPart.Strike,
                        Size = size,
                        FontName = fontName,
                        Color = richTextPart.Color.IsEmpty ? null : $"#{richTextPart.Color.R:X2}{richTextPart.Color.G:X2}{richTextPart.Color.B:X2}",
                        VerticalAlign = richTextPart.VerticalAlign.ToString()
                    });
                }
            }

            // 註解
            if (cell.Comment != null)
            {
                cellInfo.Comment = new CommentInfo
                {
                    Text = cell.Comment.Text,
                    Author = cell.Comment.Author,
                    AutoFit = cell.Comment.AutoFit,
                    Visible = cell.Comment.Visible
                };
            }

            // 超連結
            if (cell.Hyperlink != null)
            {
                cellInfo.Hyperlink = new HyperlinkInfo
                {
                    AbsoluteUri = cell.Hyperlink.AbsoluteUri,
                    OriginalString = cell.Hyperlink.OriginalString,
                    IsAbsoluteUri = cell.Hyperlink.IsAbsoluteUri
                };
            }

            // 圖片
            cellInfo.Images = GetCellImages(worksheet, cell);

            // 中繼資料
            cellInfo.Metadata = new CellMetadata
            {
                HasFormula = !string.IsNullOrEmpty(cell.Formula),
                IsRichText = cell.IsRichText,
                StyleId = cell.StyleID,
                StyleName = cell.StyleName,
                Rows = cell.Rows,
                Columns = cell.Columns,
                Start = new CellPosition 
                { 
                    Row = cell.Start.Row, 
                    Column = cell.Start.Column, 
                    Address = cell.Start.Address 
                },
                End = new CellPosition 
                { 
                    Row = cell.End.Row, 
                    Column = cell.End.Column, 
                    Address = cell.End.Address 
                }
            };

            return cellInfo;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"讀取儲存格 {cell?.Address ?? "未知位置"} 時發生錯誤");
                
                // 返回基本的儲存格資訊，避免整個處理中斷
                return new ExcelCellInfo
                {
                    Position = new CellPosition
                    {
                        Row = cell?.Start.Row ?? 0,
                        Column = cell?.Start.Column ?? 0,
                        Address = cell?.Address ?? "未知"
                    },
                    Value = null,
                    Text = "",
                    DataType = "Error",
                    Font = new FontInfo { Color = "000000" }
                };
            }
        }

        /// <summary>
        /// 獲取指定儲存格範圍內的所有圖片 (修復版)
        /// </summary>
        private List<ImageInfo>? GetCellImages(ExcelWorksheet worksheet, ExcelRange cell)
        {
            try
            {
                var images = new List<ImageInfo>();
                
                // 儲存格的邊界
                var cellStartRow = cell.Start.Row;
                var cellEndRow = cell.End.Row;
                var cellStartCol = cell.Start.Column;
                var cellEndCol = cell.End.Column;

                _logger.LogDebug($"檢查儲存格 {cell.Address} 的圖片，範圍: Row {cellStartRow}-{cellEndRow}, Col {cellStartCol}-{cellEndCol}");

                // 1. 檢查所有工作表中的圖片 (採用寬鬆匹配策略)
                if (worksheet.Drawings != null && worksheet.Drawings.Any())
                {
                    _logger.LogInformation($"工作表 '{worksheet.Name}' 包含 {worksheet.Drawings.Count} 個繪圖物件");
                    
                    foreach (var drawing in worksheet.Drawings)
                    {
                        try
                        {
                            if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
                            {
                                // 安全獲取圖片位置 - 修復 NullReference 問題
                                int fromRow = 1, fromCol = 1, toRow = 1, toCol = 1;
                                
                                if (picture.From != null)
                                {
                                    fromRow = picture.From.Row + 1;
                                    fromCol = picture.From.Column + 1;
                                }
                                
                                if (picture.To != null)
                                {
                                    toRow = picture.To.Row + 1;
                                    toCol = picture.To.Column + 1;
                                }
                                else
                                {
                                    toRow = fromRow;
                                    toCol = fromCol;
                                }

                                _logger.LogInformation($"發現圖片: '{picture.Name ?? "未命名"}' 位置: Row {fromRow}-{toRow}, Col {fromCol}-{toCol}");

                                // 寬鬆的位置檢查 - 如果沒有找到任何圖片，就包含所有圖片
                                bool shouldInclude = images.Count == 0 || 
                                                   (fromRow >= cellStartRow - 10 && fromRow <= cellEndRow + 10 &&
                                                    fromCol >= cellStartCol - 10 && fromCol <= cellEndCol + 10);

                                if (shouldInclude)
                                {
                                    try
                                    {
                                        // 獲取圖片實際尺寸
                                        var (actualWidth, actualHeight) = GetActualImageDimensions(picture);
                                        
                                        // 獲取錨點儲存格的像素尺寸
                                        var (cellWidth, cellHeight) = GetCellPixelDimensions(worksheet, fromRow, fromCol);
                                        
                                        // 等比例縮放圖片以適應儲存格
                                        var (scaledWidth, scaledHeight) = ScaleImageToCell(actualWidth, actualHeight, cellWidth, cellHeight);
                                        
                                        // 計算縮放比例
                                        var scaleFactor = actualWidth > 0 ? (double)scaledWidth / actualWidth : 1.0;
                                        var isScaled = Math.Abs(scaleFactor - 1.0) > 0.01; // 允許 1% 誤差
                                        
                                        var imageInfo = new ImageInfo
                                        {
                                            Name = picture.Name ?? $"Image_{images.Count + 1}",
                                            Description = isScaled 
                                                ? $"Excel 檔案中的圖片 (已縮放至儲存格尺寸)"
                                                : $"Excel 檔案中的圖片",
                                            ImageType = GetImageTypeFromPicture(picture),
                                            Width = scaledWidth,
                                            Height = scaledHeight,
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
                                            
                                            // 縮放相關資訊
                                            OriginalWidth = actualWidth,
                                            OriginalHeight = actualHeight,
                                            ScaleFactor = scaleFactor,
                                            IsScaled = isScaled,
                                            ScaleMethod = isScaled 
                                                ? $"儲存格適應縮放 (儲存格: {cellWidth:F0}x{cellHeight:F0}px)"
                                                : "無縮放"
                                        };

                                        images.Add(imageInfo);
                                        _logger.LogInformation($"成功解析圖片: {imageInfo.Name}, 大小: {imageInfo.FileSize} bytes");
                                    }
                                    catch (Exception imgEx)
                                    {
                                        _logger.LogError(imgEx, $"處理圖片資料時發生錯誤: {imgEx.Message}");
                                    }
                                }
                            }
                            else
                            {
                                _logger.LogDebug($"跳過非圖片繪圖物件: {drawing.GetType().Name}");
                            }
                        }
                        catch (Exception drawEx)
                        {
                            _logger.LogError(drawEx, $"處理繪圖物件時發生錯誤: {drawEx.Message}");
                        }
                    }
                }
                else
                {
                    _logger.LogDebug($"工作表 '{worksheet.Name}' 沒有繪圖物件");
                }

                // 2. 檢查 DISPIMG 函數
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    var formula = cell.Formula;
                    _logger.LogDebug($"檢查儲存格 {cell.Address} 的公式: {formula}");
                    
                    if (formula.Contains("DISPIMG") || formula.Contains("_xlfn.DISPIMG"))
                    {
                        _logger.LogInformation($"在儲存格 {cell.Address} 中找到 DISPIMG 函數: {formula}");
                        
                        var imageId = ExtractImageIdFromFormula(formula);
                        if (!string.IsNullOrEmpty(imageId))
                        {
                            var embeddedImage = FindEmbeddedImageById(worksheet.Workbook, imageId);
                            if (embeddedImage != null)
                            {
                                images.Add(embeddedImage);
                                _logger.LogInformation($"成功找到 DISPIMG 嵌入圖片: {embeddedImage.Name}");
                            }
                            else
                            {
                                var placeholderImage = new ImageInfo
                                {
                                    Name = $"DISPIMG_{imageId}",
                                    Description = $"DISPIMG 函數引用的圖片 (ID: {imageId}) - 無法存取實際圖片資料",
                                    ImageType = "PNG",
                                    Width = 100,
                                    Height = 100,
                                    Left = 0,
                                    Top = 0,
                                    Base64Data = GeneratePlaceholderImage(),
                                    FileName = $"dispimg_{imageId.Replace("ID_", "").Replace("\"", "")}.png",
                                    FileSize = 0,
                                    AnchorCell = new CellPosition { Row = cellStartRow, Column = cellStartCol, Address = cell.Address },
                                    HyperlinkAddress = $"原始公式: {formula}"
                                };
                                images.Add(placeholderImage);
                                _logger.LogInformation($"已創建 DISPIMG 佔位符，ID: {imageId}");
                            }
                        }
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

        /// <summary>
        /// 計算儲存格的實際像素尺寸
        /// </summary>
        private (double width, double height) GetCellPixelDimensions(ExcelWorksheet worksheet, int row, int col)
        {
            try
            {
                // 獲取欄寬（Excel 單位）
                var column = worksheet.Column(col);
                var columnWidth = column.Width > 0 ? column.Width : worksheet.DefaultColWidth;
                
                // 獲取行高（點數單位）
                var rowObj = worksheet.Row(row);
                var rowHeight = rowObj.Height > 0 ? rowObj.Height : worksheet.DefaultRowHeight;
                
                // Excel 欄寬單位轉換為像素
                // Excel 欄寬單位是基於預設字型的字符寬度，約等於 7 像素
                var cellWidthPixels = columnWidth * 7.0;
                
                // Excel 行高單位是點數（points），1 point = 4/3 pixels (at 96 DPI)
                var cellHeightPixels = rowHeight * 4.0 / 3.0;
                
                _logger.LogDebug($"儲存格 {GetColumnName(col)}{row} 尺寸: {cellWidthPixels:F1} x {cellHeightPixels:F1} 像素");
                
                return (cellWidthPixels, cellHeightPixels);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"計算儲存格 {GetColumnName(col)}{row} 尺寸時發生錯誤");
                return (100.0, 20.0); // 預設尺寸
            }
        }

        /// <summary>
        /// 根據儲存格尺寸等比例縮放圖片
        /// </summary>
        private (int scaledWidth, int scaledHeight) ScaleImageToCell(int originalWidth, int originalHeight, double cellWidth, double cellHeight, double scaleFactor = 0.9)
        {
            try
            {
                if (originalWidth <= 0 || originalHeight <= 0)
                {
                    return ((int)(cellWidth * scaleFactor), (int)(cellHeight * scaleFactor));
                }

                // 計算可用空間（留 10% 邊距）
                var availableWidth = cellWidth * scaleFactor;
                var availableHeight = cellHeight * scaleFactor;

                // 計算縮放比例，保持圖片長寬比
                var scaleX = availableWidth / originalWidth;
                var scaleY = availableHeight / originalHeight;
                var scale = Math.Min(scaleX, scaleY);

                // 確保縮放不會放大圖片過度
                scale = Math.Min(scale, 2.0); // 最大放大 2 倍
                
                var scaledWidth = (int)(originalWidth * scale);
                var scaledHeight = (int)(originalHeight * scale);

                _logger.LogDebug($"圖片縮放: {originalWidth}x{originalHeight} -> {scaledWidth}x{scaledHeight} (比例: {scale:F2})");

                return (scaledWidth, scaledHeight);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "圖片縮放計算時發生錯誤");
                return (originalWidth, originalHeight);
            }
        }

        /// <summary>
        /// 獲取圖片的實際尺寸（像素）
        /// </summary>
        private (int width, int height) GetActualImageDimensions(OfficeOpenXml.Drawing.ExcelPicture picture)
        {
            try
            {
                // 方法 1: 從圖片的 Image 屬性獲取
                if (picture.Image?.Bounds != null)
                {
                    var boundsWidth = (int)picture.Image.Bounds.Width;
                    var boundsHeight = (int)picture.Image.Bounds.Height;
                    
                    if (boundsWidth > 0 && boundsHeight > 0)
                    {
                        _logger.LogDebug($"圖片 {picture.Name} 從 Bounds 獲取尺寸: {boundsWidth}x{boundsHeight}");
                        return (boundsWidth, boundsHeight);
                    }
                }

                // 方法 2: 從圖片位置計算尺寸（EMU 單位轉像素）
                if (picture.From != null && picture.To != null)
                {
                    // EPPlus 使用 EMU (English Metric Units)，1 inch = 914400 EMU
                    // 假設 96 DPI (dots per inch)
                    const double emuPerPixel = 9525.0; // 914400 / 96
                    
                    var widthEmu = picture.To.ColumnOff - picture.From.ColumnOff;
                    var heightEmu = picture.To.RowOff - picture.From.RowOff;
                    
                    var calculatedWidth = Math.Max(1, (int)(widthEmu / emuPerPixel));
                    var calculatedHeight = Math.Max(1, (int)(heightEmu / emuPerPixel));
                    
                    if (calculatedWidth > 0 && calculatedHeight > 0)
                    {
                        _logger.LogDebug($"圖片 {picture.Name} 從位置計算尺寸: {calculatedWidth}x{calculatedHeight}");
                        return (calculatedWidth, calculatedHeight);
                    }
                }

                // 方法 3: 從圖片資料分析實際尺寸
                if (picture.Image?.ImageBytes != null && picture.Image.ImageBytes.Length > 0)
                {
                    var (dataWidth, dataHeight) = AnalyzeImageDataDimensions(picture.Image.ImageBytes);
                    if (dataWidth > 0 && dataHeight > 0)
                    {
                        _logger.LogDebug($"圖片 {picture.Name} 從資料分析尺寸: {dataWidth}x{dataHeight}");
                        return (dataWidth, dataHeight);
                    }
                }

                // 方法 4: 檢查圖片的其他屬性
                if (picture.Image != null)
                {
                    // 嘗試獲取其他可能的尺寸屬性
                    var imageType = picture.Image.GetType();
                    var widthProp = imageType.GetProperty("Width");
                    var heightProp = imageType.GetProperty("Height");
                    
                    if (widthProp != null && heightProp != null)
                    {
                        var propWidth = widthProp.GetValue(picture.Image);
                        var propHeight = heightProp.GetValue(picture.Image);
                        
                        if (propWidth is int w && propHeight is int h && w > 0 && h > 0)
                        {
                            _logger.LogDebug($"圖片 {picture.Name} 從屬性獲取尺寸: {w}x{h}");
                            return (w, h);
                        }
                    }
                }

                _logger.LogWarning($"無法獲取圖片 {picture.Name} 的實際尺寸，使用預設值");
                return (300, 200); // 預設尺寸
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"獲取圖片 {picture.Name} 尺寸時發生錯誤");
                return (300, 200); // 預設尺寸
            }
        }

        /// <summary>
        /// 從圖片二進位資料分析實際尺寸
        /// </summary>
        private (int width, int height) AnalyzeImageDataDimensions(byte[] imageData)
        {
            try
            {
                if (imageData.Length < 24) return (0, 0);

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
                    if (dimensions.width > 0 && dimensions.height > 0)
                    {
                        _logger.LogDebug($"從 JPEG 資料獲取尺寸: {dimensions.width}x{dimensions.height}");
                        return dimensions;
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

        /// <summary>
        /// 分析 JPEG 圖片尺寸
        /// </summary>
        private (int width, int height) AnalyzeJpegDimensions(byte[] jpegData)
        {
            try
            {
                int pos = 2; // 跳過 SOI 標記 (FF D8)
                
                while (pos < jpegData.Length - 8)
                {
                    if (jpegData[pos] == 0xFF)
                    {
                        byte marker = jpegData[pos + 1];
                        
                        // SOF0 (Start of Frame) 標記
                        if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2)
                        {
                            if (pos + 7 < jpegData.Length)
                            {
                                // JPEG SOF 格式：FF C0 [length] [precision] [height] [width]
                                var height = (jpegData[pos + 5] << 8) | jpegData[pos + 6];
                                var width = (jpegData[pos + 7] << 8) | jpegData[pos + 8];
                                
                                if (width > 0 && height > 0 && width < 65536 && height < 65536)
                                {
                                    return (width, height);
                                }
                            }
                        }
                        
                        // 跳到下一個標記
                        if (pos + 3 < jpegData.Length)
                        {
                            var segmentLength = (jpegData[pos + 2] << 8) | jpegData[pos + 3];
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

        /// <summary>
        /// 從圖片名稱獲取圖片格式類型
        /// </summary>
        private string GetImageTypeFromName(string? imageName)
        {
            if (string.IsNullOrEmpty(imageName))
                return "Unknown";

            var extension = Path.GetExtension(imageName).ToLowerInvariant();
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
                _ => "Unknown"
            };
        }

        /// <summary>
        /// 獲取圖片格式類型
        /// </summary>
        private string GetImageType(OfficeOpenXml.Drawing.ePictureType imageFormat)
        {
            return imageFormat switch
            {
                OfficeOpenXml.Drawing.ePictureType.Png => "PNG",
                OfficeOpenXml.Drawing.ePictureType.Jpg => "JPEG",
                OfficeOpenXml.Drawing.ePictureType.Gif => "GIF",
                OfficeOpenXml.Drawing.ePictureType.Bmp => "BMP",
                OfficeOpenXml.Drawing.ePictureType.Tif => "TIFF",
                OfficeOpenXml.Drawing.ePictureType.Wmf => "WMF",
                OfficeOpenXml.Drawing.ePictureType.Emf => "EMF",
                _ => "Unknown"
            };
        }

        /// <summary>
        /// 從 ExcelPicture 物件獲取圖片類型 (適用於 Google Sheets 檔案)
        /// </summary>
        private string GetImageTypeFromPicture(OfficeOpenXml.Drawing.ExcelPicture picture)
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
                    
                    // PNG 檔頭: 89 50 4E 47 0D 0A 1A 0A
                    if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
                    {
                        return "PNG";
                    }
                    
                    // JPEG 檔頭: FF D8
                    if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xD8)
                    {
                        return "JPEG";
                    }
                    
                    // GIF 檔頭: 47 49 46 38
                    if (bytes.Length >= 4 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
                    {
                        return "GIF";
                    }
                    
                    // BMP 檔頭: 42 4D
                    if (bytes.Length >= 2 && bytes[0] == 0x42 && bytes[1] == 0x4D)
                    {
                        return "BMP";
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

        /// <summary>
        /// 獲取圖片檔案大小
        /// </summary>
        private long GetImageFileSize(OfficeOpenXml.Drawing.ExcelPicture picture)
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

        /// <summary>
        /// 將圖片轉換為 Base64 字串
        /// </summary>
        private string ConvertImageToBase64(OfficeOpenXml.Drawing.ExcelPicture picture)
        {
            try
            {
                if (picture.Image?.ImageBytes != null && picture.Image.ImageBytes.Length > 0)
                {
                    return Convert.ToBase64String(picture.Image.ImageBytes);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"轉換圖片 {picture.Name} 為 Base64 時發生錯誤");
            }
            
            return string.Empty;
        }

        /// <summary>
        /// WPS專用 功能目前無效
        /// 修正需參考 https://blog.csdn.net/m0_59983333/article/details/138164606
        /// 從 DISPIMG 公式中提取圖片 ID 
        /// </summary>
        private string? ExtractImageIdFromFormula(string formula)
        {
            try
            {
                // 匹配 DISPIMG("ID_...", 1) 格式
                var pattern = @"DISPIMG\s*\(\s*""([^""]+)""\s*,\s*\d+\s*\)";
                var match = System.Text.RegularExpressions.Regex.Match(formula, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"提取圖片 ID 時發生錯誤，公式: {formula}");
            }
            
            return null;
        }

        /// <summary>
        /// 根據 ID 在工作簿中查找嵌入的圖片 (支援 EPPlus 7.1.0)
        /// </summary>
        private ImageInfo? FindEmbeddedImageById(ExcelWorkbook workbook, string imageId)
        {
            try
            {
                _logger.LogInformation($"開始查找嵌入圖片，ID: {imageId}");
                
                // 方法 1: 遍歷所有工作表的所有繪圖物件
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (worksheet.Drawings != null)
                    {
                        foreach (var drawing in worksheet.Drawings)
                        {
                            if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
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
                                    _logger.LogInformation($"找到匹配的圖片: {picture.Name}");
                                    return new ImageInfo
                                    {
                                        Name = picture.Name,
                                        Description = picture.Description ?? "",
                                        ImageType = GetImageTypeFromName(picture.Name),
                                        Width = (int)(picture.Image?.Bounds.Width ?? 0),
                                        Height = (int)(picture.Image?.Bounds.Height ?? 0),
                                        Left = picture.From.ColumnOff / 9525.0,
                                        Top = picture.From.RowOff / 9525.0,
                                        Base64Data = ConvertImageToBase64(picture),
                                        FileName = picture.Name,
                                        FileSize = GetImageFileSize(picture),
                                        AnchorCell = new CellPosition
                                        {
                                            Row = picture.From.Row + 1,
                                            Column = picture.From.Column + 1,
                                            Address = $"{GetColumnName(picture.From.Column + 1)}{picture.From.Row + 1}"
                                        }
                                    };
                                }
                            }
                        }
                    }
                }

                // 方法 2: 使用 EPPlus 7.1.0 進階功能查找圖片
                var foundImage = TryAdvancedImageSearch(workbook, imageId);
                if (foundImage != null)
                {
                    _logger.LogInformation($"通過進階搜索找到圖片: {imageId}");
                    return foundImage;
                }

                _logger.LogWarning($"未找到圖片，ID: {imageId}。嘗試列出所有可用的繪圖物件...");
                LogAvailableDrawings(workbook);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"查找嵌入圖片時發生錯誤，ID: {imageId}");
            }
            
            return null;
        }

        /// <summary>
        /// 使用 EPPlus 7.1.0 進階功能查找圖片
        /// </summary>
        private ImageInfo? TryAdvancedImageSearch(ExcelWorkbook workbook, string imageId)
        {
            try
            {
                _logger.LogInformation($"使用 EPPlus 7.1.0 進階功能查找圖片，ID: {imageId}");
                
                // 方法 1: 直接解析 OOXML 包結構 (新增)
                var ooxmlImage = TryDirectOoxmlImageSearch(workbook, imageId);
                if (ooxmlImage != null)
                {
                    return ooxmlImage;
                }
                
                // 方法 2: 嘗試透過 VBA 項目查找圖片
                var vbaImage = TryFindImageInVbaProject(workbook, imageId);
                if (vbaImage != null)
                {
                    return vbaImage;
                }
                
                // 方法 3: 搜索所有工作表中的背景圖片
                var backgroundImage = TryFindBackgroundImage(workbook, imageId);
                if (backgroundImage != null)
                {
                    return backgroundImage;
                }
                
                // 方法 4: 檢查所有繪圖物件的更多屬性 (EPPlus 7.1.0 增強)
                var detailedImage = TryDetailedDrawingSearch(workbook, imageId);
                if (detailedImage != null)
                {
                    return detailedImage;
                }
                
                // 方法 5: 嘗試透過工作表的其他圖片相關屬性
                var worksheetImage = TryFindImageInWorksheets(workbook, imageId);
                if (worksheetImage != null)
                {
                    return worksheetImage;
                }
                
                _logger.LogDebug($"EPPlus 7.1.0 所有進階方法都無法找到圖片，ID: {imageId}");
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"進階圖片搜索時發生錯誤，ID: {imageId}");
            }
            
            return null;
        }

        /// <summary>
        /// 直接解析 OOXML 包結構來查找 DISPIMG 圖片
        /// </summary>
        private ImageInfo? TryDirectOoxmlImageSearch(ExcelWorkbook workbook, string imageId)
        {
            try
            {
                _logger.LogInformation($"嘗試直接解析 OOXML ZIP 結構查找 DISPIMG 圖片，ID: {imageId}");
                
                // 由於無法直接存取 EPPlus 的底層 Stream，我們改用另一種方法
                // 透過檢查工作簿的內部結構來尋找圖片資料
                
                var cleanImageId = imageId.Replace("ID_", "").Replace("\"", "").ToLowerInvariant();
                _logger.LogDebug($"清理後的圖片 ID: {cleanImageId}");
                
                // 方法 1: 深度搜索所有工作表的內部資料
                var deepSearchResult = DeepSearchWorksheetInternals(workbook, cleanImageId, imageId);
                if (deepSearchResult != null)
                {
                    return deepSearchResult;
                }
                
                // 方法 2: 嘗試透過 Workbook 的 _package 屬性（反射方式）
                var reflectionResult = TryReflectionBasedImageSearch(workbook, cleanImageId, imageId);
                if (reflectionResult != null)
                {
                    return reflectionResult;
                }
                
                // 方法 3: 模擬 DISPIMG 功能，檢查是否有相關的圖片快取
                var cacheResult = TryImageCacheSearch(workbook, cleanImageId, imageId);
                if (cacheResult != null)
                {
                    return cacheResult;
                }
                
                _logger.LogDebug($"OOXML 直接解析未找到圖片，ID: {imageId}");
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"OOXML 直接解析時發生錯誤，ID: {imageId}");
                return null;
            }
        }

        /// <summary>
        /// 深度搜索工作表內部結構
        /// </summary>
        private ImageInfo? DeepSearchWorksheetInternals(ExcelWorkbook workbook, string cleanImageId, string originalImageId)
        {
            try
            {
                _logger.LogDebug("進行工作表內部深度搜索...");
                
                foreach (var worksheet in workbook.Worksheets)
                {
                    // 檢查工作表的所有儲存格是否有相關的圖片資料
                    for (int row = 1; row <= worksheet.Dimension?.End.Row; row++)
                    {
                        for (int col = 1; col <= worksheet.Dimension?.End.Column; col++)
                        {
                            var cell = worksheet.Cells[row, col];
                            
                            // 檢查儲存格公式是否包含我們要找的 ID
                            if (!string.IsNullOrEmpty(cell.Formula) && 
                                cell.Formula.Contains(originalImageId))
                            {
                                // 檢查這個位置是否有任何隱藏的圖片資料
                                var hiddenImage = ExtractHiddenImageData(worksheet, row, col, cleanImageId, originalImageId);
                                if (hiddenImage != null)
                                {
                                    return hiddenImage;
                                }
                            }
                        }
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "深度搜索工作表內部時發生錯誤");
                return null;
            }
        }

        /// <summary>
        /// 嘗試使用反射存取更深層的資料結構
        /// </summary>
        private ImageInfo? TryReflectionBasedImageSearch(ExcelWorkbook workbook, string cleanImageId, string originalImageId)
        {
            try
            {
                _logger.LogDebug("嘗試反射方式存取圖片資料...");
                
                // 嘗試透過反射存取 EPPlus 的內部屬性
                var workbookType = workbook.GetType();
                var properties = workbookType.GetProperties(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                
                foreach (var prop in properties)
                {
                    if (prop.Name.ToLowerInvariant().Contains("package") || 
                        prop.Name.ToLowerInvariant().Contains("stream"))
                    {
                        _logger.LogDebug($"檢查內部屬性: {prop.Name}");
                        
                        try
                        {
                            var value = prop.GetValue(workbook);
                            if (value != null)
                            {
                                // 嘗試從這個物件中搜索圖片
                                var reflectionImage = SearchObjectForImages(value, cleanImageId, originalImageId);
                                if (reflectionImage != null)
                                {
                                    return reflectionImage;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogDebug(ex, $"存取屬性 {prop.Name} 時發生錯誤");
                        }
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "反射搜索時發生錯誤");
                return null;
            }
        }

        /// <summary>
        /// 嘗試從圖片快取中搜索
        /// </summary>
        private ImageInfo? TryImageCacheSearch(ExcelWorkbook workbook, string cleanImageId, string originalImageId)
        {
            try
            {
                _logger.LogDebug("搜索可能的圖片快取...");
                
                // 在實際的 DISPIMG 場景中，圖片資料可能被快取在其他地方
                // 這裡我們嘗試一些替代方法來獲取圖片資料
                
                // 方法 1: 檢查是否有任何隱藏的工作表包含圖片資料
                var hiddenSheetImage = SearchHiddenSheets(workbook, cleanImageId, originalImageId);
                if (hiddenSheetImage != null)
                {
                    return hiddenSheetImage;
                }
                
                // 方法 2: 嘗試創建一個基於 ID 的圖片查找
                var generatedImage = TryGenerateImageFromId(cleanImageId, originalImageId);
                if (generatedImage != null)
                {
                    return generatedImage;
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "圖片快取搜索時發生錯誤");
                return null;
            }
        }

        /// <summary>
        /// 提取隱藏的圖片資料
        /// </summary>
        private ImageInfo? ExtractHiddenImageData(ExcelWorksheet worksheet, int row, int col, string cleanImageId, string originalImageId)
        {
            try
            {
                // 檢查這個位置是否有任何特殊的屬性或註解
                var cell = worksheet.Cells[row, col];
                
                // 檢查註解
                if (cell.Comment != null)
                {
                    var commentText = cell.Comment.Text;
                    if (!string.IsNullOrEmpty(commentText) && IsBase64String(commentText))
                    {
                        _logger.LogInformation($"在儲存格 {cell.Address} 的註解中找到可能的圖片資料");
                        
                        return new ImageInfo
                        {
                            Name = $"Hidden_DISPIMG_{originalImageId}",
                            Description = $"從儲存格註解提取的圖片資料",
                            ImageType = "PNG",
                            Width = 100,
                            Height = 100,
                            Left = 0,
                            Top = 0,
                            Base64Data = commentText,
                            FileName = $"hidden_dispimg_{cleanImageId}.png",
                            FileSize = Convert.FromBase64String(commentText).Length,
                            AnchorCell = new CellPosition 
                            { 
                                Row = row, 
                                Column = col, 
                                Address = cell.Address 
                            },
                            HyperlinkAddress = "從儲存格註解提取"
                        };
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, $"提取隱藏圖片資料時發生錯誤，位置: {row},{col}");
                return null;
            }
        }

        /// <summary>
        /// 從物件中搜索圖片
        /// </summary>
        private ImageInfo? SearchObjectForImages(object obj, string cleanImageId, string originalImageId)
        {
            try
            {
                // 這是一個簡化的實現，在真實場景中可能需要更複雜的邏輯
                var objType = obj.GetType();
                _logger.LogDebug($"搜索物件類型: {objType.Name}");
                
                // 如果物件有 ToString 方法且內容包含我們的 ID
                var stringRepresentation = obj.ToString();
                if (!string.IsNullOrEmpty(stringRepresentation) && 
                    (stringRepresentation.Contains(cleanImageId) || stringRepresentation.Contains(originalImageId)))
                {
                    _logger.LogDebug($"物件內容匹配 ID: {objType.Name}");
                    
                    // 檢查是否可能包含圖片資料
                    if (IsBase64String(stringRepresentation))
                    {
                        return CreateImageFromBase64(stringRepresentation, originalImageId, "反射搜索結果");
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "搜索物件時發生錯誤");
                return null;
            }
        }

        /// <summary>
        /// 搜索隱藏的工作表
        /// </summary>
        private ImageInfo? SearchHiddenSheets(ExcelWorkbook workbook, string cleanImageId, string originalImageId)
        {
            try
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (worksheet.Hidden != eWorkSheetHidden.Visible)
                    {
                        _logger.LogDebug($"檢查隱藏工作表: {worksheet.Name}");
                        
                        // 在隱藏工作表中搜索圖片
                        for (int row = 1; row <= Math.Min(worksheet.Dimension?.End.Row ?? 0, 100); row++)
                        {
                            for (int col = 1; col <= Math.Min(worksheet.Dimension?.End.Column ?? 0, 20); col++)
                            {
                                var cell = worksheet.Cells[row, col];
                                var cellText = cell.Text;
                                
                                if (!string.IsNullOrEmpty(cellText) &&
                                    (cellText.Contains(cleanImageId) || cellText.Contains(originalImageId)) &&
                                    IsBase64String(cellText))
                                {
                                    _logger.LogInformation($"在隱藏工作表 {worksheet.Name} 的儲存格 {cell.Address} 找到圖片資料");
                                    
                                    return CreateImageFromBase64(cellText, originalImageId, $"隱藏工作表 {worksheet.Name}");
                                }
                            }
                        }
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "搜索隱藏工作表時發生錯誤");
                return null;
            }
        }

        /// <summary>
        /// 嘗試根據 ID 生成圖片
        /// </summary>
        private ImageInfo? TryGenerateImageFromId(string cleanImageId, string originalImageId)
        {
            try
            {
                // 在某些情況下，我們可能需要根據 ID 生成或查找對應的圖片
                // 這裡是一個佔位符實現，在實際應用中可能需要連接到外部資源
                
                _logger.LogDebug($"嘗試根據 ID 生成圖片: {cleanImageId}");
                
                // 檢查 ID 是否符合特定的模式，可能包含圖片資源的提示
                if (cleanImageId.Length >= 32) // 看起來像 GUID
                {
                    // 在真實場景中，這裡可能會查詢資料庫、快取或外部 API
                    _logger.LogDebug("ID 看起來像 GUID，但無法從中直接提取圖片資料");
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "根據 ID 生成圖片時發生錯誤");
                return null;
            }
        }

        /// <summary>
        /// 從 Base64 字串創建 ImageInfo
        /// </summary>
        private ImageInfo CreateImageFromBase64(string base64Data, string originalImageId, string source)
        {
            return new ImageInfo
            {
                Name = $"Found_DISPIMG_{originalImageId}",
                Description = $"從 {source} 找到的 DISPIMG 圖片",
                ImageType = "PNG",
                Width = 100,
                Height = 100,
                Left = 0,
                Top = 0,
                Base64Data = base64Data,
                FileName = $"found_dispimg_{originalImageId.Replace("ID_", "").ToLowerInvariant()}.png",
                FileSize = Convert.FromBase64String(base64Data).Length,
                AnchorCell = new CellPosition { Row = 1, Column = 1, Address = "A1" },
                HyperlinkAddress = source
            };
        }

        /// <summary>
        /// 檢查字串是否為有效的 base64
        /// </summary>
        private bool IsBase64String(string str)
        {
            if (string.IsNullOrEmpty(str) || str.Length < 10)
                return false;
                
            try
            {
                var base64Regex = new Regex(@"^[A-Za-z0-9+/]*={0,2}$");
                return base64Regex.IsMatch(str) && str.Length % 4 == 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 嘗試在工作表中查找圖片 (EPPlus 7.1.0 專用)
        /// </summary>
        private ImageInfo? TryFindImageInWorksheets(ExcelWorkbook workbook, string imageId)
        {
            try
            {
                var cleanImageId = imageId.Replace("ID_", "").Replace("\"", "").ToLowerInvariant();
                
                foreach (var worksheet in workbook.Worksheets)
                {
                    // 檢查工作表是否有任何隱藏的圖片屬性
                    if (worksheet.Drawings != null)
                    {
                        foreach (var drawing in worksheet.Drawings)
                        {
                            // EPPlus 7.1.0 可能有更多的圖片類型
                            if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
                            {
                                // 檢查圖片的所有可能屬性
                                if (CheckAllPictureProperties(picture, cleanImageId, imageId))
                                {
                                    _logger.LogInformation($"通過擴展屬性檢查找到匹配圖片: {picture.Name}");
                                    
                                    return CreateImageInfoFromPicture(picture, imageId);
                                }
                            }
                            else
                            {
                                // 檢查其他類型的繪圖物件
                                _logger.LogDebug($"檢查非圖片繪圖物件: {drawing.GetType().Name}");
                            }
                        }
                    }
                    
                    // EPPlus 7.1.0 可能有其他方式存取圖片
                    // 這裡可以添加更多特定於新版本的搜索方法
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "在工作表中查找圖片時發生錯誤");
                return null;
            }
        }

        /// <summary>
        /// 檢查圖片的所有屬性以尋找匹配
        /// </summary>
        private bool CheckAllPictureProperties(OfficeOpenXml.Drawing.ExcelPicture picture, string cleanImageId, string originalImageId)
        {
            try
            {
                // 檢查基本屬性
                var name = picture.Name?.ToLowerInvariant() ?? "";
                var description = picture.Description?.ToLowerInvariant() ?? "";
                
                // EPPlus 7.1.0 可能有額外的屬性可以檢查
                // 這裡可以添加更多屬性檢查
                
                return name.Contains(cleanImageId) || 
                       name.Contains(originalImageId.ToLowerInvariant()) ||
                       description.Contains(cleanImageId) ||
                       IsPartialIdMatch(cleanImageId, name) ||
                       IsPartialIdMatch(cleanImageId, description);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "檢查圖片屬性時發生錯誤");
                return false;
            }
        }

        /// <summary>
        /// 從 ExcelPicture 創建 ImageInfo
        /// </summary>
        private ImageInfo CreateImageInfoFromPicture(OfficeOpenXml.Drawing.ExcelPicture picture, string originalImageId)
        {
            return new ImageInfo
            {
                Name = picture.Name ?? $"EPPlus7_Found_{originalImageId}",
                Description = $"通過 EPPlus 7.1.0 擴展搜索找到 (原始 ID: {originalImageId})",
                ImageType = GetImageTypeFromName(picture.Name ?? ""),
                Width = (int)(picture.Image?.Bounds.Width ?? 0),
                Height = (int)(picture.Image?.Bounds.Height ?? 0),
                Left = picture.From.ColumnOff / 9525.0,
                Top = picture.From.RowOff / 9525.0,
                Base64Data = ConvertImageToBase64(picture),
                FileName = picture.Name ?? $"epplus7_found_{originalImageId}",
                FileSize = GetImageFileSize(picture),
                AnchorCell = new CellPosition
                {
                    Row = picture.From.Row + 1,
                    Column = picture.From.Column + 1,
                    Address = $"{GetColumnName(picture.From.Column + 1)}{picture.From.Row + 1}"
                },
                HyperlinkAddress = $"EPPlus 7.1.0 擴展搜索結果"
            };
        }

        /// <summary>
        /// 嘗試從 VBA 項目中查找圖片
        /// </summary>
        private ImageInfo? TryFindImageInVbaProject(ExcelWorkbook workbook, string imageId)
        {
            try
            {
                // EPPlus 4.5.3 可能無法存取 VBA 項目中的圖片
                // 但我們可以嘗試檢查是否有相關的 VBA 模組
                if (workbook.VbaProject != null)
                {
                    _logger.LogDebug($"工作簿包含 VBA 項目，嘗試查找圖片 ID: {imageId}");
                    // 在更新的 EPPlus 版本中，這裡可以進一步實現
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, $"查找 VBA 項目圖片時發生錯誤，ID: {imageId}");
                return null;
            }
        }

        /// <summary>
        /// 嘗試查找工作表背景圖片
        /// </summary>
        private ImageInfo? TryFindBackgroundImage(ExcelWorkbook workbook, string imageId)
        {
            try
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    // 檢查工作表是否有背景圖片
                    if (worksheet.BackgroundImage != null)
                    {
                        _logger.LogDebug($"工作表 '{worksheet.Name}' 有背景圖片");
                        
                        // 這裡可以進一步檢查背景圖片是否與我們要找的 ID 相關
                        // EPPlus 4.5.3 的限制使得這個功能可能無法完全實現
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, $"查找背景圖片時發生錯誤，ID: {imageId}");
                return null;
            }
        }

        /// <summary>
        /// 詳細搜索繪圖物件，包括更多屬性和可能的關聯
        /// </summary>
        private ImageInfo? TryDetailedDrawingSearch(ExcelWorkbook workbook, string imageId)
        {
            try
            {
                var cleanImageId = imageId.Replace("ID_", "").Replace("\"", "").ToLowerInvariant();
                _logger.LogDebug($"進行詳細繪圖搜索，清理後的 ID: {cleanImageId}");
                
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (worksheet.Drawings != null)
                    {
                        foreach (var drawing in worksheet.Drawings)
                        {
                            if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
                            {
                                // 檢查更多可能的匹配條件
                                var pictureName = picture.Name?.ToLowerInvariant() ?? "";
                                var pictureDescription = picture.Description?.ToLowerInvariant() ?? "";
                                
                                // 嘗試各種匹配模式
                                if (pictureName.Contains(cleanImageId) || 
                                    pictureDescription.Contains(cleanImageId) ||
                                    cleanImageId.Contains(pictureName) ||
                                    IsPartialIdMatch(cleanImageId, pictureName))
                                {
                                    _logger.LogInformation($"透過詳細搜索找到可能匹配的圖片: Name='{picture.Name}', Description='{picture.Description}'");
                                    
                                    return new ImageInfo
                                    {
                                        Name = picture.Name ?? $"Found_{cleanImageId}",
                                        Description = $"透過詳細搜索找到的圖片 (原始 ID: {imageId})",
                                        ImageType = GetImageTypeFromName(picture.Name ?? ""),
                                        Width = (int)(picture.Image?.Bounds.Width ?? 0),
                                        Height = (int)(picture.Image?.Bounds.Height ?? 0),
                                        Left = picture.From.ColumnOff / 9525.0,
                                        Top = picture.From.RowOff / 9525.0,
                                        Base64Data = ConvertImageToBase64(picture),
                                        FileName = picture.Name ?? $"detailed_search_{cleanImageId}",
                                        FileSize = GetImageFileSize(picture),
                                        AnchorCell = new CellPosition
                                        {
                                            Row = picture.From.Row + 1,
                                            Column = picture.From.Column + 1,
                                            Address = $"{GetColumnName(picture.From.Column + 1)}{picture.From.Row + 1}"
                                        }
                                    };
                                }
                            }
                        }
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"詳細繪圖搜索時發生錯誤，ID: {imageId}");
                return null;
            }
        }

        /// <summary>
        /// 檢查部分 ID 匹配 (用於處理可能的 ID 變形)
        /// </summary>
        private bool IsPartialIdMatch(string cleanId, string pictureName)
        {
            if (string.IsNullOrEmpty(cleanId) || string.IsNullOrEmpty(pictureName))
                return false;
                
            // 檢查是否有部分匹配 (至少 8 個字符)
            if (cleanId.Length >= 8 && pictureName.Length >= 8)
            {
                for (int i = 0; i <= cleanId.Length - 8; i++)
                {
                    var segment = cleanId.Substring(i, 8);
                    if (pictureName.Contains(segment))
                    {
                        return true;
                    }
                }
            }
            
            return false;
        }

        /// <summary>
        /// 記錄所有可用的繪圖物件資訊 (用於除錯)
        /// </summary>
        private void LogAvailableDrawings(ExcelWorkbook workbook)
        {
            try
            {
                _logger.LogInformation("=================== Excel 文件診斷報告 ===================");
                
                // 統計總體資訊
                int totalDrawings = 0;
                int totalPictures = 0;
                
                foreach (var worksheet in workbook.Worksheets)
                {
                    _logger.LogInformation($"📊 工作表分析: '{worksheet.Name}'");
                    
                    if (worksheet.Drawings != null && worksheet.Drawings.Any())
                    {
                        totalDrawings += worksheet.Drawings.Count;
                        _logger.LogInformation($"  🎨 繪圖物件數量: {worksheet.Drawings.Count}");
                        
                        for (int i = 0; i < worksheet.Drawings.Count; i++)
                        {
                            var drawing = worksheet.Drawings[i];
                            if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
                            {
                                totalPictures++;
                                _logger.LogInformation($"  📷 圖片 #{i + 1}:");
                                _logger.LogInformation($"    - Name: '{picture.Name ?? "未命名"}'");
                                _logger.LogInformation($"    - Description: '{picture.Description ?? "無描述"}'");
                                _logger.LogInformation($"    - Position: Row {picture.From.Row + 1}, Col {picture.From.Column + 1}");
                                _logger.LogInformation($"    - Size: {picture.Image?.Bounds.Width ?? 0} x {picture.Image?.Bounds.Height ?? 0}");
                                
                                // 嘗試獲取更多屬性
                                try
                                {
                                    var imageData = ConvertImageToBase64(picture);
                                    var dataSize = string.IsNullOrEmpty(imageData) ? 0 : imageData.Length;
                                    _logger.LogInformation($"    - Base64 資料長度: {dataSize} 字符");
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogWarning($"    - 無法獲取圖片資料: {ex.Message}");
                                }
                            }
                            else
                            {
                                _logger.LogInformation($"  🔧 其他繪圖物件 #{i + 1}:");
                                _logger.LogInformation($"    - Type: {drawing.GetType().Name}");
                                _logger.LogInformation($"    - Name: '{drawing.Name ?? "未命名"}'");
                            }
                        }
                    }
                    else
                    {
                        _logger.LogInformation($"  ❌ 無繪圖物件");
                    }
                    
                    // 檢查工作表中是否有 DISPIMG 公式
                    var dispimgCount = CountDispimgFormulas(worksheet);
                    if (dispimgCount > 0)
                    {
                        _logger.LogInformation($"  🔍 發現 {dispimgCount} 個 DISPIMG 公式");
                    }
                }
                
                // 總體統計
                _logger.LogInformation($"=================== 總體統計 ===================");
                _logger.LogInformation($"📈 總工作表數: {workbook.Worksheets.Count}");
                _logger.LogInformation($"📈 總繪圖物件數: {totalDrawings}");
                _logger.LogInformation($"📈 總圖片數: {totalPictures}");
                _logger.LogInformation($"=================== 診斷完成 ===================");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "列出繪圖物件時發生錯誤");
            }
        }

        /// <summary>
        /// 計算工作表中 DISPIMG 公式的數量
        /// </summary>
        private int CountDispimgFormulas(ExcelWorksheet worksheet)
        {
            int count = 0;
            try
            {
                var dimension = worksheet.Dimension;
                if (dimension != null)
                {
                    for (int row = 1; row <= dimension.End.Row; row++)
                    {
                        for (int col = 1; col <= dimension.End.Column; col++)
                        {
                            var cell = worksheet.Cells[row, col];
                            if (!string.IsNullOrEmpty(cell.Formula) && 
                                (cell.Formula.Contains("DISPIMG") || cell.Formula.Contains("_xlfn.DISPIMG")))
                            {
                                count++;
                                var imageId = ExtractImageIdFromFormula(cell.Formula);
                                _logger.LogInformation($"    📍 {cell.Address}: ID={imageId}, 公式={cell.Formula}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "統計 DISPIMG 公式時發生錯誤");
            }
            return count;
        }

        /// <summary>
        /// 從 URI 中獲取圖片類型
        /// </summary>
        private string GetImageTypeFromUri(string uri)
        {
            var extension = Path.GetExtension(uri)?.ToLowerInvariant();
            return extension switch
            {
                ".png" => "PNG",
                ".jpg" => "JPEG",
                ".jpeg" => "JPEG",
                ".gif" => "GIF",
                ".bmp" => "BMP",
                _ => "Unknown"
            };
        }

        /// <summary>
        /// 生成佔位符圖片的 Base64 資料
        /// </summary>
        private string GeneratePlaceholderImage()
        {
            try
            {
                // 創建一個 100x100 的灰色佔位符圖片，帶有 "圖片未找到" 的視覺提示
                // 使用更大的尺寸和更明顯的佔位符設計
                var pngBytes = new byte[]
                {
                    // 完整的 100x100 灰色 PNG 圖片
                    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG 簽名
                    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk 標頭
                    0x00, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00, 0x64, // 100x100 像素
                    0x08, 0x02, 0x00, 0x00, 0x00, 0xFF, 0x80, 0x02, // 8-bit RGB
                    0x03, 0x00, 0x00, 0x00, 0x18, 0x50, 0x4C, 0x54, // PLTE chunk
                    0x45, 0xC0, 0xC0, 0xC0, 0xE0, 0xE0, 0xE0, 0xF0, // 調色盤 (灰色系)
                    0xF0, 0xF0, 0xFF, 0xFF, 0xFF, 0x80, 0x80, 0x80,
                    0x60, 0x60, 0x60, 0x40, 0x40, 0x40, 0x20, 0x20,
                    0x20, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89
                };
                
                // 為了簡化，我們使用一個固定的小尺寸佔位符
                // 實際的完整 100x100 PNG 會很大，這裡用一個簡化版本
                var simplePlaceholder = new byte[]
                {
                    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG 標頭
                    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
                    0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x20, // 32x32 像素
                    0x08, 0x06, 0x00, 0x00, 0x00, 0x73, 0x7A, 0x7A, // 8-bit RGBA
                    0xF4, 0x00, 0x00, 0x00, 0x19, 0x74, 0x45, 0x58, // tEXt chunk
                    0x74, 0x43, 0x6F, 0x6D, 0x6D, 0x65, 0x6E, 0x74, // "Comment"
                    0x00, 0x49, 0x6D, 0x61, 0x67, 0x65, 0x20, 0x6E, // "Image n"
                    0x6F, 0x74, 0x20, 0x66, 0x6F, 0x75, 0x6E, 0x64, // "ot found"
                    0xC9, 0x38, 0x29, 0xCB, 0x00, 0x00, 0x00, 0x3E, // 圖片資料開始
                    0x49, 0x44, 0x41, 0x54, 0x58, 0x85, 0xED, 0xD0, // IDAT chunk
                    0x31, 0x01, 0x00, 0x00, 0x08, 0x03, 0xA0, 0xF5, // 壓縮的圖片資料
                    0x53, 0xE0, 0x00, 0x02, 0x00, 0x00, 0x40, 0x00, // (32x32 灰色方塊)
                    0x00, 0x10, 0x00, 0x00, 0x04, 0x00, 0x00, 0x01,
                    0x00, 0x00, 0x00, 0x40, 0x00, 0x00, 0x10, 0x00,
                    0x00, 0x04, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,
                    0x40, 0x00, 0x00, 0x10, 0x00, 0x00, 0x04, 0x00,
                    0x00, 0x01, 0x00, 0x00, 0x00, 0x40, 0x8A, 0x0D,
                    0x8C, 0x08, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, // IEND chunk
                    0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82
                };
                
                return Convert.ToBase64String(simplePlaceholder);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "生成佔位符圖片時發生錯誤");
                
                // 如果生成失敗，返回最小的透明圖片
                var fallbackBytes = new byte[]
                {
                    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG 標頭
                    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
                    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1 像素
                    0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, // 8-bit RGBA
                    0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, // IDAT chunk
                    0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, // 透明像素資料
                    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
                    0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, // IEND chunk
                    0x42, 0x60, 0x82
                };
                
                return Convert.ToBase64String(fallbackBytes);
            }
        }

        /// <summary>
        /// 將欄位編號轉換為 Excel 欄位名稱 (1 -> A, 2 -> B, 27 -> AA, etc.)
        /// </summary>
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

        /// <summary>
        /// 獲取合併儲存格的邊框 (考慮整個合併範圍的外邊界)
        /// </summary>
        private BorderInfo GetMergedCellBorder(ExcelWorksheet worksheet, ExcelRange mergedRange, ExcelRange currentCell)
        {
            var border = new BorderInfo();
            
            // 獲取合併範圍的邊界
            int topRow = mergedRange.Start.Row;
            int bottomRow = mergedRange.End.Row;
            int leftCol = mergedRange.Start.Column;
            int rightCol = mergedRange.End.Column;
            
            // 上邊框：來自合併範圍頂部的儲存格
            var topCell = worksheet.Cells[topRow, currentCell.Start.Column];
            border.Top = new BorderStyle 
            { 
                Style = topCell.Style.Border.Top.Style.ToString(), 
                Color = GetColorFromExcelColor(topCell.Style.Border.Top.Color)
            };
            
            // 下邊框：來自合併範圍底部的儲存格
            var bottomCell = worksheet.Cells[bottomRow, currentCell.Start.Column];
            border.Bottom = new BorderStyle 
            { 
                Style = bottomCell.Style.Border.Bottom.Style.ToString(), 
                Color = GetColorFromExcelColor(bottomCell.Style.Border.Bottom.Color)
            };
            
            // 左邊框：來自合併範圍左側的儲存格
            var leftCell = worksheet.Cells[currentCell.Start.Row, leftCol];
            border.Left = new BorderStyle 
            { 
                Style = leftCell.Style.Border.Left.Style.ToString(), 
                Color = GetColorFromExcelColor(leftCell.Style.Border.Left.Color)
            };
            
            // 右邊框：來自合併範圍右側的儲存格
            var rightCell = worksheet.Cells[currentCell.Start.Row, rightCol];
            border.Right = new BorderStyle 
            { 
                Style = rightCell.Style.Border.Right.Style.ToString(), 
                Color = GetColorFromExcelColor(rightCell.Style.Border.Right.Color)
            };
            
            // 對角線邊框使用當前儲存格的設定
            border.Diagonal = new BorderStyle 
            { 
                Style = currentCell.Style.Border.Diagonal.Style.ToString(), 
                Color = GetColorFromExcelColor(currentCell.Style.Border.Diagonal.Color)
            };
            border.DiagonalUp = currentCell.Style.Border.DiagonalUp;
            border.DiagonalDown = currentCell.Style.Border.DiagonalDown;
            
            return border;
        }

        /// <summary>
        /// 獲取儲存格的背景色
        /// </summary>
        private string? GetBackgroundColor(ExcelRange cell)
        {
            var fill = cell.Style.Fill;
            
            // 調試：顯示完整的顏色資訊
            _logger.LogInformation($"Cell {cell.Address} - PatternType: {fill.PatternType}, " +
                $"BackgroundColor[Rgb: '{fill.BackgroundColor.Rgb}', Theme: {fill.BackgroundColor.Theme}, Tint: {fill.BackgroundColor.Tint}, Indexed: {fill.BackgroundColor.Indexed}], " +
                $"PatternColor[Rgb: '{fill.PatternColor.Rgb}', Theme: {fill.PatternColor.Theme}, Tint: {fill.PatternColor.Tint}, Indexed: {fill.PatternColor.Indexed}]");
            
            // 檢查填充類型，只有 Solid 或 Pattern 類型才有背景色
            if (fill.PatternType == OfficeOpenXml.Style.ExcelFillStyle.Solid)
            {
                // Solid 填充：使用背景色
                return GetColorFromExcelColor(fill.BackgroundColor);
            }
            else if (fill.PatternType != OfficeOpenXml.Style.ExcelFillStyle.None)
            {
                // Pattern 填充：優先使用 BackgroundColor，其次使用 PatternColor
                return GetColorFromExcelColor(fill.BackgroundColor) ?? 
                       GetColorFromExcelColor(fill.PatternColor);
            }
            
            return null;
        }

        /// <summary>
        /// 從 EPPlus ExcelColor 物件提取顏色值
        /// </summary>
        private string? GetColorFromExcelColor(OfficeOpenXml.Style.ExcelColor excelColor)
        {
            if (excelColor == null)
                return null;

            try
            {
                // 1. 優先使用 RGB 值
                string? rgbValue = null;
                try
                {
                    rgbValue = excelColor.Rgb;
                }
                catch (Exception ex)
                {
                    _logger.LogDebug($"無法存取 RGB 值: {ex.Message}");
                }

                if (!string.IsNullOrEmpty(rgbValue))
                {
                    var colorValue = rgbValue.TrimStart('#');
                    
                    // 處理 ARGB 格式（8位）轉為 RGB 格式（6位）
                    if (colorValue.Length == 8)
                    {
                        // ARGB 格式：前2位是Alpha，後6位是RGB
                        colorValue = colorValue.Substring(2);
                    }
                    
                    if (colorValue.Length == 6)
                    {
                        return colorValue.ToUpperInvariant();
                    }
                    
                    // 處理3位短格式（例如：F00 -> FF0000）
                    if (colorValue.Length == 3)
                    {
                        return $"{colorValue[0]}{colorValue[0]}{colorValue[1]}{colorValue[1]}{colorValue[2]}{colorValue[2]}";
                    }
                }

                // 2. 嘗試使用索引顏色
                if (excelColor.Indexed >= 0)
                {
                    return GetIndexedColor(excelColor.Indexed);
                }

                // 3. 嘗試使用主題顏色
                if (excelColor.Theme != null)
                {
                    var themeValue = (int)excelColor.Theme;
                    var tintValue = (double)excelColor.Tint;
                    return GetThemeColor(themeValue, tintValue);
                }

                // 4. 嘗試自動顏色
                if (excelColor.Auto == true)
                {
                    return "000000"; // 預設黑色
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "解析顏色時發生錯誤");
                return null;
            }
        }

        /// <summary>
        /// 獲取 Excel 索引顏色對應的 RGB 值
        /// </summary>
        private string? GetIndexedColor(int colorIndex)
        {
            // Excel 標準索引顏色對應表（使用 Excel 2016+ 標準色彩）
            var indexedColors = new Dictionary<int, string>
            {
                // Excel 自動色彩和系統色彩 (0-7)
                { 0, "000000" },  // Automatic / Black
                { 1, "FFFFFF" },  // White
                { 2, "FF0000" },  // Red
                { 3, "00FF00" },  // Bright Green
                { 4, "0000FF" },  // Blue
                { 5, "FFFF00" },  // Yellow
                { 6, "FF00FF" },  // Magenta
                { 7, "00FFFF" },  // Cyan
                
                // Excel 標準色彩 (8-15) - 重複定義確保相容性
                { 8, "000000" },  // Black
                { 9, "FFFFFF" },  // White
                { 10, "FF0000" }, // Red
                { 11, "00FF00" }, // Bright Green
                { 12, "0000FF" }, // Blue
                { 13, "FFFF00" }, // Yellow
                { 14, "FF00FF" }, // Magenta
                { 15, "00FFFF" }, // Cyan
                
                // Excel 標準調色板 (16-31)
                { 16, "800000" }, // Dark Red (Maroon)
                { 17, "008000" }, // Green
                { 18, "000080" }, // Dark Blue (Navy)
                { 19, "808000" }, // Dark Yellow (Olive)
                { 20, "800080" }, // Purple
                { 21, "008080" }, // Dark Cyan (Teal)
                { 22, "C0C0C0" }, // Light Gray (Silver)
                { 23, "808080" }, // Gray
                
                // Excel 擴展色彩 (24-39)
                { 24, "9999FF" }, // Periwinkle
                { 25, "993366" }, // Plum
                { 26, "FFFFCC" }, // Ivory
                { 27, "CCFFFF" }, // Light Turquoise
                { 28, "660066" }, // Dark Purple
                { 29, "FF8080" }, // Coral
                { 30, "0066CC" }, // Ocean Blue
                { 31, "CCCCFF" }, // Ice Blue
                
                // Excel 標準色彩擴展 (32-39)
                { 32, "000080" }, // Dark Blue
                { 33, "FF00FF" }, // Pink
                { 34, "FFFF00" }, // Yellow
                { 35, "00FFFF" }, // Turquoise
                { 36, "800080" }, // Violet
                { 37, "800000" }, // Dark Red
                { 38, "008080" }, // Teal
                { 39, "0000FF" }, // Blue
                
                // Excel 淺色系列 (40-47)
                { 40, "00CCFF" }, // Sky Blue
                { 41, "CCFFFF" }, // Light Turquoise
                { 42, "CCFFCC" }, // Light Green
                { 43, "FFFF99" }, // Light Yellow
                { 44, "99CCFF" }, // Pale Blue
                { 45, "FF99CC" }, // Rose
                { 46, "CC99FF" }, // Lavender
                { 47, "FFCC99" }, // Peach
                
                // Excel 亮色系列 (48-55)
                { 48, "3366FF" }, // Light Blue
                { 49, "33CCCC" }, // Aqua
                { 50, "99CC00" }, // Lime
                { 51, "FFCC00" }, // Gold
                { 52, "FF9900" }, // Orange
                { 53, "FF6600" }, // Orange Red
                { 54, "666699" }, // Blue Gray
                { 55, "969696" }, // Gray 40%
                
                // Excel 深色系列 (56-63)
                { 56, "003366" }, // Dark Teal
                { 57, "339966" }, // Sea Green
                { 58, "003300" }, // Dark Green
                { 59, "333300" }, // Dark Olive
                { 60, "964B00" }, // Brown (咖啡色)
                { 61, "993366" }, // Dark Rose
                { 62, "333399" }, // Indigo
                { 63, "333333" }  // Gray 80%
            };
            
            return indexedColors.ContainsKey(colorIndex) ? indexedColors[colorIndex] : null;
        }

        /// <summary>
        /// 獲取 Excel 主題顏色對應的 RGB 值
        /// </summary>
        private string? GetThemeColor(int themeIndex, double tint)
        {
            // Excel 標準主題顏色對應表（Office 預設主題）
            var themeColors = new Dictionary<int, string>
            {
                { 0, "FFFFFF" },  // Background 1 / Light 1
                { 1, "000000" },  // Text 1 / Dark 1
                { 2, "E7E6E6" },  // Background 2 / Light 2
                { 3, "44546A" },  // Text 2 / Dark 2
                { 4, "5B9BD5" },  // Accent 1
                { 5, "70AD47" },  // Accent 2
                { 6, "A5A5A5" },  // Accent 3
                { 7, "FFC000" },  // Accent 4
                { 8, "4472C4" },  // Accent 5
                { 9, "264478" },  // Accent 6
                { 10, "0563C1" }, // Hyperlink
                { 11, "954F72" }  // Followed Hyperlink
            };
            
            if (!themeColors.ContainsKey(themeIndex))
            {
                return null;
            }
            
            var baseColor = themeColors[themeIndex];
            
            // 如果有 Tint 值，需要調整顏色亮度
            if (Math.Abs(tint) > 0.001)
            {
                return ApplyTint(baseColor, tint);
            }
            
            return baseColor;
        }

        /// <summary>
        /// 對顏色應用 Tint 效果
        /// </summary>
        private string ApplyTint(string hexColor, double tint)
        {
            if (hexColor.Length != 6) return hexColor;
            
            try
            {
                var r = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                var g = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                var b = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                
                if (tint < 0)
                {
                    // Tint < 0: 變暗
                    r = (int)(r * (1 + tint));
                    g = (int)(g * (1 + tint));
                    b = (int)(b * (1 + tint));
                }
                else
                {
                    // Tint > 0: 變亮
                    r = (int)(r + (255 - r) * tint);
                    g = (int)(g + (255 - g) * tint);
                    b = (int)(b + (255 - b) * tint);
                }
                
                // 確保值在 0-255 範圍內
                r = Math.Max(0, Math.Min(255, r));
                g = Math.Max(0, Math.Min(255, g));
                b = Math.Max(0, Math.Min(255, b));
                
                return $"{r:X2}{g:X2}{b:X2}";
            }
            catch
            {
                return hexColor;
            }
        }

        [HttpPost("upload")]
        public async Task<ActionResult<UploadResponse>> UploadExcel(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest(new UploadResponse
                    {
                        Success = false,
                        Message = "未選擇檔案或檔案為空"
                    });
                }

                // 檢查檔案格式
                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(file.FileName).ToLower();
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new UploadResponse
                    {
                        Success = false,
                        Message = "僅支援 Excel 檔案格式 (.xlsx, .xls)"
                    });
                }

                var excelData = new ExcelData
                {
                    FileName = file.FileName
                };

                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using var package = new ExcelPackage(stream);
                
                // 取得所有工作表名稱
                excelData.AvailableWorksheets = package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
                
                var worksheet = package.Workbook.Worksheets[0]; // 使用第一個工作表
                excelData.WorksheetName = worksheet.Name;

                if (worksheet.Dimension == null)
                {
                    return BadRequest(new UploadResponse
                    {
                        Success = false,
                        Message = "Excel 檔案為空或無有效資料"
                    });
                }

                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;

                excelData.TotalRows = rowCount;
                excelData.TotalColumns = colCount;

                // 生成 Excel 欄位標頭 (A, B, C, D...) 包含寬度資訊
                var columnHeaders = new List<object>();
                for (int col = 1; col <= colCount; col++)
                {
                    var column = worksheet.Column(col);
                    var width = column.Width > 0 ? column.Width : worksheet.DefaultColWidth;
                    
                    columnHeaders.Add(new 
                    {
                        Name = GetColumnName(col),
                        Width = width,
                        Index = col
                    });
                }

                // 讀取第一行內容作為內容標頭，保留格式信息
                var contentHeaders = new List<object>();
                for (int col = 1; col <= colCount; col++)
                {
                    var headerCell = worksheet.Cells[1, col];
                    contentHeaders.Add(CreateCellInfo(headerCell, worksheet));
                }
                
                // 提供兩種標頭：Excel 欄位標頭和內容標頭
                excelData.Headers = new[] { columnHeaders.ToArray(), contentHeaders.ToArray() };

                // 讀取資料行，保留原始格式（包含Rich Text）
                var rows = new List<object[]>();
                for (int row = 1; row <= rowCount; row++) // 從第一行開始（包含所有行）
                {
                    var rowData = new List<object>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        rowData.Add(CreateCellInfo(cell, worksheet));
                    }
                    rows.Add(rowData.ToArray());
                }

                excelData.Rows = rows.ToArray();

                _logger.LogInformation($"成功讀取 Excel 檔案: {file.FileName}, 行數: {rowCount}, 欄數: {colCount}");

                return Ok(new UploadResponse
                {
                    Success = true,
                    Message = $"成功讀取 Excel 檔案，共 {rowCount - 1} 筆資料",
                    Data = excelData
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "讀取 Excel 檔案時發生錯誤");
                return StatusCode(500, new UploadResponse
                {
                    Success = false,
                    Message = $"讀取檔案時發生錯誤: {ex.Message}"
                });
            }
        }

        [HttpGet("sample")]
        public ActionResult<ExcelData> GetSampleData()
        {
            // 提供範例資料供前端測試
            var sampleData = new ExcelData
            {
                FileName = "範例資料.xlsx",
                TotalRows = 8,
                TotalColumns = 5,
                Headers = new[] { new[] { "姓名", "年齡", "部門", "薪資", "入職日期" } },
                Rows = new object[][]
                {
                    new object[] { "張三", 30, "資訊部", 50000, "2020-01-15" },
                    new object[] { "李四", 25, "人事部", 45000, "2021-03-20" },
                    new object[] { "王五", 35, "財務部", 55000, "2019-05-10" },
                    new object[] { "趙六", 28, "行銷部", 48000, "2022-07-01" },
                    new object[] { "錢七", 32, "研發部", 60000, "2018-12-05" },
                    new object[] { "孫八", 29, "客服部", 42000, "2021-09-15" },
                    new object[] { "周九", 31, "業務部", 52000, "2020-11-20" }
                }
            };

            return Ok(sampleData);
        }

        [HttpPost("upload-worksheet")]
        public async Task<ActionResult<UploadResponse>> UploadExcelWorksheet(IFormFile file, [FromQuery] string? worksheetName = null, [FromQuery] int worksheetIndex = 0)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest(new UploadResponse
                    {
                        Success = false,
                        Message = "未選擇檔案或檔案為空"
                    });
                }

                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(file.FileName).ToLower();
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new UploadResponse
                    {
                        Success = false,
                        Message = "僅支援 Excel 檔案格式 (.xlsx, .xls)"
                    });
                }

                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using var package = new ExcelPackage(stream);
                var excelData = new ExcelData
                {
                    FileName = file.FileName,
                    AvailableWorksheets = package.Workbook.Worksheets.Select(ws => ws.Name).ToList()
                };

                // 選擇工作表
                ExcelWorksheet worksheet;
                if (!string.IsNullOrEmpty(worksheetName))
                {
                    worksheet = package.Workbook.Worksheets[worksheetName];
                    if (worksheet == null)
                    {
                        return BadRequest(new UploadResponse
                        {
                            Success = false,
                            Message = $"找不到名為 '{worksheetName}' 的工作表"
                        });
                    }
                }
                else
                {
                    if (worksheetIndex >= package.Workbook.Worksheets.Count)
                    {
                        return BadRequest(new UploadResponse
                        {
                            Success = false,
                            Message = $"工作表索引 {worksheetIndex} 超出範圍"
                        });
                    }
                    worksheet = package.Workbook.Worksheets[worksheetIndex];
                }

                excelData.WorksheetName = worksheet.Name;

                if (worksheet.Dimension == null)
                {
                    return BadRequest(new UploadResponse
                    {
                        Success = false,
                        Message = "選擇的工作表為空或無有效資料"
                    });
                }

                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;
                excelData.TotalRows = rowCount;
                excelData.TotalColumns = colCount;

                // 生成 Excel 欄位標頭 (A, B, C, D...) 包含寬度資訊
                var columnHeaders = new List<object>();
                for (int col = 1; col <= colCount; col++)
                {
                    var column = worksheet.Column(col);
                    var width = column.Width > 0 ? column.Width : worksheet.DefaultColWidth;
                    
                    columnHeaders.Add(new 
                    {
                        Name = GetColumnName(col),
                        Width = width,
                        Index = col
                    });
                }

                // 讀取第一行內容作為內容標頭，保留格式信息
                var contentHeaders = new List<object>();
                for (int col = 1; col <= colCount; col++)
                {
                    var headerCell = worksheet.Cells[1, col];
                    contentHeaders.Add(CreateCellInfo(headerCell, worksheet));
                }
                
                // 提供兩種標頭：Excel 欄位標頭和內容標頭
                excelData.Headers = new[] { columnHeaders.ToArray(), contentHeaders.ToArray() };

                var rows = new List<object[]>();
                for (int row = 1; row <= rowCount; row++) // 從第一行開始（包含所有行）
                {
                    var rowData = new List<object>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        rowData.Add(CreateCellInfo(cell, worksheet));
                    }
                    rows.Add(rowData.ToArray());
                }
                excelData.Rows = rows.ToArray();

                return Ok(new UploadResponse
                {
                    Success = true,
                    Message = $"成功讀取工作表 '{worksheet.Name}'，共 {rowCount - 1} 筆資料",
                    Data = excelData
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "讀取 Excel 工作表時發生錯誤");
                return StatusCode(500, new UploadResponse
                {
                    Success = false,
                    Message = $"讀取檔案時發生錯誤: {ex.Message}"
                });
            }
        }

        [HttpGet("download-sample")]
        public IActionResult DownloadSampleExcel()
        {
            try
            {
                var fileBytes = ExcelSampleGenerator.GenerateSampleExcel();
                return File(fileBytes, 
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "範例員工資料.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "產生範例Excel檔案時發生錯誤");
                return StatusCode(500, new UploadResponse
                {
                    Success = false,
                    Message = $"產生範例檔案時發生錯誤: {ex.Message}"
                });
            }
        }

        [HttpPost("debug-raw-data")]
        public async Task<ActionResult> DebugRawExcelData(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest("未選擇檔案或檔案為空");
                }

                var allowedExtensions = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(file.FileName).ToLower();
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest("僅支援 Excel 檔案格式 (.xlsx, .xls)");
                }

                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using var package = new ExcelPackage(stream);
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension == null)
                {
                    return BadRequest("Excel 檔案為空或無有效資料");
                }

                var debugData = new DebugExcelData
                {
                    FileName = file.FileName,
                    WorksheetInfo = new WorksheetInfo
                    {
                        Name = worksheet.Name,
                        TotalRows = worksheet.Dimension.Rows,
                        TotalColumns = worksheet.Dimension.Columns,
                        DefaultColWidth = worksheet.DefaultColWidth,
                        DefaultRowHeight = worksheet.DefaultRowHeight
                    },
                    SampleCells = GetRawCellData(worksheet, Math.Min(5, worksheet.Dimension.Rows), Math.Min(5, worksheet.Dimension.Columns)),
                    AllWorksheets = package.Workbook.Worksheets.Select(ws => new
                    {
                        Name = ws.Name,
                        Index = ws.Index,
                        State = ws.Hidden.ToString()
                    }).Cast<object>().ToList()
                };

                return Ok(debugData);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "讀取 Excel 原始資料時發生錯誤");
                return StatusCode(500, $"讀取檔案時發生錯誤: {ex.Message}");
            }
        }

        private object[,] GetRawCellData(ExcelWorksheet worksheet, int maxRows, int maxCols)
        {
            var cells = new object[maxRows, maxCols];
            
            for (int row = 1; row <= maxRows; row++)
            {
                for (int col = 1; col <= maxCols; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var column = worksheet.Column(col);
                    
                    cells[row - 1, col - 1] = new
                    {
                        Position = new { Row = row, Column = col, Address = cell.Address },
                        
                        // 基本值和顯示
                        Value = cell.Value,
                        Text = cell.Text,
                        Formula = cell.Formula,
                        FormulaR1C1 = cell.FormulaR1C1,
                        
                        // 資料類型
                        ValueType = cell.Value?.GetType().Name,
                        
                        // 格式化
                        NumberFormat = cell.Style.Numberformat.Format,
                        NumberFormatId = cell.Style.Numberformat.NumFmtID,
                        
                        // 字體樣式
                        Font = new
                        {
                            Name = cell.Style.Font.Name,
                            Size = cell.Style.Font.Size,
                            Bold = cell.Style.Font.Bold,
                            Italic = cell.Style.Font.Italic,
                            Underline = cell.Style.Font.UnderLine,
                            Strike = cell.Style.Font.Strike,
                            Color = GetColorFromExcelColor(cell.Style.Font.Color),
                            ColorTheme = cell.Style.Font.Color.Theme?.ToString(),
                            ColorTint = cell.Style.Font.Color.Tint,
                            Charset = cell.Style.Font.Charset,
                            Scheme = cell.Style.Font.Scheme?.ToString(),
                            Family = cell.Style.Font.Family
                        },
                        
                        // 對齊方式
                        Alignment = new
                        {
                            Horizontal = cell.Style.HorizontalAlignment.ToString(),
                            Vertical = cell.Style.VerticalAlignment.ToString(),
                            WrapText = cell.Style.WrapText,
                            Indent = cell.Style.Indent,
                            ReadingOrder = cell.Style.ReadingOrder.ToString(),
                            TextRotation = cell.Style.TextRotation,
                            ShrinkToFit = cell.Style.ShrinkToFit
                        },
                        
                        // 邊框
                        Border = new
                        {
                            Top = new { Style = cell.Style.Border.Top.Style.ToString(), Color = cell.Style.Border.Top.Color.Rgb },
                            Bottom = new { Style = cell.Style.Border.Bottom.Style.ToString(), Color = cell.Style.Border.Bottom.Color.Rgb },
                            Left = new { Style = cell.Style.Border.Left.Style.ToString(), Color = cell.Style.Border.Left.Color.Rgb },
                            Right = new { Style = cell.Style.Border.Right.Style.ToString(), Color = cell.Style.Border.Right.Color.Rgb },
                            Diagonal = new { Style = cell.Style.Border.Diagonal.Style.ToString(), Color = cell.Style.Border.Diagonal.Color.Rgb },
                            DiagonalUp = cell.Style.Border.DiagonalUp,
                            DiagonalDown = cell.Style.Border.DiagonalDown
                        },
                        
                        // 填充/背景
                        Fill = new
                        {
                            PatternType = cell.Style.Fill.PatternType.ToString(),
                            BackgroundColor = cell.Style.Fill.BackgroundColor.Rgb,
                            PatternColor = cell.Style.Fill.PatternColor.Rgb,
                            BackgroundColorTheme = cell.Style.Fill.BackgroundColor.Theme?.ToString(),
                            BackgroundColorTint = cell.Style.Fill.BackgroundColor.Tint
                        },
                        
                        // 尺寸和合併
                        Dimensions = new
                        {
                            ColumnWidth = column.Width > 0 ? column.Width : worksheet.DefaultColWidth,
                            RowHeight = worksheet.Row(row).Height,
                            IsMerged = cell.Merge,
                            MergedRangeAddress = cell.Merge ? FindMergedRange(worksheet, row, col)?.Address : null
                        },
                        
                        // Rich Text
                        RichText = cell.IsRichText ? cell.RichText?.Select(rt => new
                        {
                            Text = rt.Text,
                            Bold = rt.Bold,
                            Italic = rt.Italic,
                            UnderLine = rt.UnderLine,
                            Strike = rt.Strike,
                            Size = rt.Size,
                            FontName = rt.FontName,
                            Color = rt.Color.IsEmpty ? null : $"#{rt.Color.R:X2}{rt.Color.G:X2}{rt.Color.B:X2}",
                            VerticalAlign = rt.VerticalAlign.ToString()
                        }).ToList() : null,
                        
                        // 註解
                        Comment = cell.Comment != null ? new
                        {
                            Text = cell.Comment.Text,
                            Author = cell.Comment.Author,
                            AutoFit = cell.Comment.AutoFit,
                            Visible = cell.Comment.Visible
                        } : null,
                        
                        // 超連結
                        Hyperlink = cell.Hyperlink != null ? new
                        {
                            AbsoluteUri = cell.Hyperlink.AbsoluteUri,
                            OriginalString = cell.Hyperlink.OriginalString,
                            IsAbsoluteUri = cell.Hyperlink.IsAbsoluteUri
                        } : null,
                        
                        // 其他屬性
                        Metadata = new
                        {
                            HasFormula = !string.IsNullOrEmpty(cell.Formula),
                            IsRichText = cell.IsRichText,
                            StyleId = cell.StyleID,
                            StyleName = cell.StyleName,
                            Rows = cell.Rows,
                            Columns = cell.Columns,
                            Start = new { Row = cell.Start.Row, Column = cell.Start.Column, Address = cell.Start.Address },
                            End = new { Row = cell.End.Row, Column = cell.End.Column, Address = cell.End.Address }
                        }
                    };
                }
            }
            
            return cells;
        }
    }
}