using ExcelReaderAPI.Models;
using ExcelReaderAPI.Models.Caches;
using ExcelReaderAPI.Models.Enums;
using ExcelReaderAPI.Services.Interfaces;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace ExcelReaderAPI.Services
{
    /// <summary>
    /// Excel 處理服務 - 核心儲存格和工作表處理邏輯
    /// Phase 2.1: 從 ExcelController 搬移 10 個方法,保持邏輯完全不變
    /// </summary>
    public class ExcelProcessingService : IExcelProcessingService
    {
        private readonly IExcelImageService _imageService;
        private readonly IExcelCellService _cellService;
        private readonly IExcelColorService _colorService;
        private readonly ILogger<ExcelProcessingService> _logger;

        public ExcelProcessingService(
            IExcelImageService imageService,
            IExcelCellService cellService,
            IExcelColorService colorService,
            ILogger<ExcelProcessingService> logger)
        {
            _imageService = imageService ?? throw new ArgumentNullException(nameof(imageService));
            _cellService = cellService ?? throw new ArgumentNullException(nameof(cellService));
            _colorService = colorService ?? throw new ArgumentNullException(nameof(colorService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        #region CreateCellInfo Methods

        /// <summary>
        /// 創建儲存格資訊 (使用快取優化版本)
        /// 從 ExcelController.CreateCellInfo 搬移,保持邏輯完全不變
        /// </summary>
        public ExcelCellInfo CreateCellInfo(
            ExcelRange cell,
            ExcelWorksheet worksheet,
            WorksheetImageIndex? imageIndex,
            ColorCache colorCache,
            MergedCellIndex mergedCellIndex)
        {
            // 從 ExcelController.CreateCellInfo 完整搬移 (370行)
            // ⚠️ 重要: 保持所有邏輯完全不變
            
            var cellInfo = new ExcelCellInfo();

            try
            {
                // 智能內容檢測:先判斷儲存格的主要內容類型 (使用索引)
                var contentType = DetectCellContentType(cell, imageIndex);

                // 位置資訊（所有類型都需要）
                cellInfo.Position = new CellPosition
                {
                    Row = cell.Start.Row,
                    Column = cell.Start.Column,
                    Address = cell.Address ?? $"{GetColumnName(cell.Start.Column)}{cell.Start.Row}"
                };

                // 基本值和顯示（所有類型都需要）
                cellInfo.Value = GetSafeValue(cell);
                cellInfo.Text = cell.Text;
                cellInfo.Formula = cell.Formula;
                cellInfo.FormulaR1C1 = cell.FormulaR1C1;

                // 資料類型（所有類型都需要）
                cellInfo.ValueType = cell.Value?.GetType().Name;
                if (cell.Value == null)
                {
                    cellInfo.DataType = contentType == CellContentType.ImageOnly ? "Image" : "Empty";
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

                // 根據內容類型決定是否處理樣式資訊
                if (contentType == CellContentType.ImageOnly)
                {
                    cellInfo.Font = CreateDefaultFontInfo();
                    cellInfo.Alignment = CreateDefaultAlignmentInfo();
                    cellInfo.Border = CreateDefaultBorderInfo();
                    cellInfo.Fill = CreateDefaultFillInfo();

                    try
                    {
                        cellInfo.NumberFormat = cell.Style.Numberformat.Format;
                        cellInfo.NumberFormatId = cell.Style.Numberformat.NumFmtID;
                    }
                    catch
                    {
                        cellInfo.NumberFormat = "";
                        cellInfo.NumberFormatId = 0;
                    }
                }
                else
                {
                    // 完整樣式處理 (使用快取)
                    cellInfo.NumberFormat = cell.Style.Numberformat.Format;
                    cellInfo.NumberFormatId = cell.Style.Numberformat.NumFmtID;

                    cellInfo.Font = new FontInfo
                    {
                        Name = cell.Style.Font.Name,
                        Size = cell.Style.Font.Size,
                        Bold = cell.Style.Font.Bold,
                        Italic = cell.Style.Font.Italic,
                        UnderLine = cell.Style.Font.UnderLine.ToString(),
                        Strike = cell.Style.Font.Strike,
                        Color = _colorService.GetColorFromExcelColor(cell.Style.Font.Color),
                        ColorTheme = cell.Style.Font.Color.Theme?.ToString(),
                        ColorTint = (double?)cell.Style.Font.Color.Tint,
                        Charset = cell.Style.Font.Charset,
                        Scheme = cell.Style.Font.Scheme?.ToString(),
                        Family = cell.Style.Font.Family
                    };

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

                    try
                    {
                        cellInfo.Border = new BorderInfo
                        {
                            Top = new BorderStyle
                            {
                                Style = cell.Style.Border?.Top?.Style.ToString() ?? "None",
                                Color = cell.Style.Border?.Top?.Color != null ? _colorService.GetColorFromExcelColor(cell.Style.Border.Top.Color) : null
                            },
                            Bottom = new BorderStyle
                            {
                                Style = cell.Style.Border?.Bottom?.Style.ToString() ?? "None",
                                Color = cell.Style.Border?.Bottom?.Color != null ? _colorService.GetColorFromExcelColor(cell.Style.Border.Bottom.Color) : null
                            },
                            Left = new BorderStyle
                            {
                                Style = cell.Style.Border?.Left?.Style.ToString() ?? "None",
                                Color = cell.Style.Border?.Left?.Color != null ? _colorService.GetColorFromExcelColor(cell.Style.Border.Left.Color) : null
                            },
                            Right = new BorderStyle
                            {
                                Style = cell.Style.Border?.Right?.Style.ToString() ?? "None",
                                Color = cell.Style.Border?.Right?.Color != null ? _colorService.GetColorFromExcelColor(cell.Style.Border.Right.Color) : null
                            },
                            Diagonal = new BorderStyle
                            {
                                Style = cell.Style.Border?.Diagonal?.Style.ToString() ?? "None",
                                Color = cell.Style.Border?.Diagonal?.Color != null ? _colorService.GetColorFromExcelColor(cell.Style.Border.Diagonal.Color) : null
                            },
                            DiagonalUp = cell.Style.Border?.DiagonalUp ?? false,
                            DiagonalDown = cell.Style.Border?.DiagonalDown ?? false
                        };
                    }
                    catch (Exception borderEx)
                    {
                        _logger.LogDebug($"儲存格 {cell.Address} 邊框處理時發生錯誤: {borderEx.Message}，使用預設邊框");
                        cellInfo.Border = CreateDefaultBorderInfo();
                    }

                    cellInfo.Fill = new FillInfo
                    {
                        PatternType = cell.Style.Fill.PatternType.ToString(),
                        BackgroundColor = _colorService.GetBackgroundColor(cell.Style.Fill),
                        PatternColor = _colorService.GetColorFromExcelColor(cell.Style.Fill.PatternColor),
                        BackgroundColorTheme = cell.Style.Fill.BackgroundColor.Theme?.ToString(),
                        BackgroundColorTint = (double?)cell.Style.Fill.BackgroundColor.Tint
                    };
                }

                // 尺寸和合併
                var column = worksheet.Column(cell.Start.Column);
                cellInfo.Dimensions = new DimensionInfo
                {
                    ColumnWidth = column.Width > 0 ? column.Width : worksheet.DefaultColWidth,
                    RowHeight = worksheet.Row(cell.Start.Row).Height,
                    IsMerged = cell.Merge
                };

                // 合併儲存格處理 (使用快取索引)
                if (cell.Merge)
                {
                    ExcelRange? mergedRange = null;

                    // 優先使用索引查詢
                    if (mergedCellIndex != null)
                    {
                        var mergeAddress = mergedCellIndex.GetMergeRange(cell.Start.Row, cell.Start.Column);
                        if (mergeAddress != null)
                        {
                            mergedRange = worksheet.Cells[mergeAddress];
                        }
                    }
                    else
                    {
                        // 回退到原始查詢方式
                        var address = _cellService.FindMergedRange(worksheet, cell);
                        if (address != null)
                        {
                            mergedRange = worksheet.Cells[address];
                        }
                    }

                    if (mergedRange != null)
                    {
                        cellInfo.Dimensions.MergedRangeAddress = mergedRange.Address;
                        cellInfo.Dimensions.IsMainMergedCell = (cell.Start.Row == mergedRange.Start.Row &&
                                                               cell.Start.Column == mergedRange.Start.Column);

                        if (cellInfo.Dimensions.IsMainMergedCell == true)
                        {
                            cellInfo.Dimensions.RowSpan = mergedRange.Rows;
                            cellInfo.Dimensions.ColSpan = mergedRange.Columns;
                            // mergedRange.Address 不會為 null,因為我們已經檢查過 mergedRange != null
                            cellInfo.Border = _cellService.GetMergedCellBorder(worksheet, mergedRange.Address!) ?? CreateDefaultBorderInfo();
                        }
                        else
                        {
                            cellInfo.Dimensions.RowSpan = 1;
                            cellInfo.Dimensions.ColSpan = 1;
                        }
                    }
                }

                // Rich Text 處理
                if (cell.IsRichText && cell.RichText != null && cell.RichText.Count > 0)
                {
                    cellInfo.RichText = new List<RichTextPart>();
                    for (int i = 0; i < cell.RichText.Count; i++)
                    {
                        var richTextPart = cell.RichText[i];
                        var bold = richTextPart.Bold;
                        var italic = richTextPart.Italic;
                        var size = richTextPart.Size;
                        var fontName = richTextPart.FontName;

                        if (i == 0)
                        {
                            if (size == 0 || string.IsNullOrEmpty(fontName) || (!bold && !italic))
                            {
                                size = size == 0 ? cell.Style.Font.Size : size;
                                fontName = string.IsNullOrEmpty(fontName) ? cell.Style.Font.Name : fontName;
                                if (!richTextPart.Bold && cell.Style.Font.Bold) bold = true;
                                if (!richTextPart.Italic && cell.Style.Font.Italic) italic = true;
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

                // 圖片 - 使用索引版本
                ExcelRange rangeToCheck = cell;
                if (cell.Merge)
                {
                    ExcelRange? mergedRange = null;

                    // 優先使用索引查詢
                    if (mergedCellIndex != null)
                    {
                        var mergeAddress = mergedCellIndex.GetMergeRange(cell.Start.Row, cell.Start.Column);
                        if (mergeAddress != null)
                        {
                            mergedRange = worksheet.Cells[mergeAddress];
                        }
                    }
                    else
                    {
                        var address = _cellService.FindMergedRange(worksheet, cell);
                        if (address != null)
                        {
                            mergedRange = worksheet.Cells[address];
                        }
                    }

                    if (mergedRange != null)
                    {
                        rangeToCheck = mergedRange;
                    }
                }

                // ⭐ 修復: 只在合併儲存格的主要儲存格 (左上角) 中查找圖片
                if (cell.Merge && cellInfo.Dimensions?.IsMainMergedCell != true)
                {
                    // 如果是合併儲存格但不是主要儲存格，則不查找圖片
                    cellInfo.Images = null;
                    _logger.LogDebug($"儲存格 {cell.Address} 是合併儲存格的次要儲存格，跳過圖片檢查");
                }
                else
                {
                    // 使用 imageService 的方法
                    cellInfo.Images = _imageService.GetCellImages(rangeToCheck, imageIndex!, worksheet);
                }

                // 圖片跨儲存格處理 - 使用注入的 Service 方法
                _cellService.ProcessImageCrossCells(cellInfo, cell, worksheet);

                // 浮動物件 - ⭐ 修復: 只在合併儲存格的主要儲存格中查找浮動物件
                if (cell.Merge && cellInfo.Dimensions?.IsMainMergedCell != true)
                {
                    // 如果是合併儲存格但不是主要儲存格，則不查找浮動物件
                    cellInfo.FloatingObjects = null;
                    _logger.LogDebug($"儲存格 {cell.Address} 是合併儲存格的次要儲存格，跳過浮動物件檢查");
                }
                else
                {
                    cellInfo.FloatingObjects = _cellService.GetCellFloatingObjects(worksheet, rangeToCheck);
                }

                // 浮動物件跨儲存格處理 - 使用注入的 Service 方法
                _cellService.ProcessFloatingObjectCrossCells(cellInfo, cell);

                // 合併浮動物件的文字到儲存格文字中 (使用 DRY 方法)
                cellInfo.Text = _cellService.MergeFloatingObjectText(cellInfo.Text, cellInfo.FloatingObjects);

                return cellInfo;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"創建儲存格 {cell.Address} 資訊時發生錯誤: {ex.Message}");
                
                // 返回最基本的資訊
                cellInfo.Position = new CellPosition
                {
                    Row = cell.Start.Row,
                    Column = cell.Start.Column,
                    Address = cell.Address ?? $"{GetColumnName(cell.Start.Column)}{cell.Start.Row}"
                };
                cellInfo.Value = cell.Value;
                cellInfo.Text = cell.Text;
                cellInfo.DataType = "Error";
                
                return cellInfo;
            }
        }

        /// <summary>
        /// 創建儲存格資訊 (簡化版,不使用快取)
        /// 從 ExcelController.CreateCellInfo 搬移
        /// </summary>
        public ExcelCellInfo CreateCellInfo(ExcelRange cell, ExcelWorksheet worksheet)
        {
            // 為了保持向後相容,創建臨時索引並調用優化版本
            var imageIndex = new WorksheetImageIndex(worksheet);
            var colorCache = new ColorCache();
            var mergedCellIndex = new MergedCellIndex(worksheet);
            return CreateCellInfo(cell, worksheet, imageIndex, colorCache, mergedCellIndex);
        }

        #endregion

        #region DetectCellContentType Methods

        /// <summary>
        /// 偵測儲存格內容類型 (使用索引優化)
        /// 從 ExcelController.DetectCellContentType 搬移
        /// </summary>
        public CellContentType DetectCellContentType(ExcelRange cell, WorksheetImageIndex? imageIndex)
        {
            try
            {
                // 檢查是否有文字內容
                var hasText = !string.IsNullOrEmpty(cell.Text) || !string.IsNullOrEmpty(cell.Formula);

                // ⭐ EPPlus 8.x: 優先檢查 In-Cell 圖片
                bool hasInCellPicture = false;
                try
                {
                    // 只有單一儲存格才能檢查 In-Cell Picture
                    hasInCellPicture = cell.Picture.Exists;
                }
                catch
                {
                    // 忽略 Picture API 錯誤
                }

                // 使用索引快速檢查是否有浮動圖片 (Drawing Pictures) - O(1) 複雜度
                var hasDrawingImages = imageIndex?.HasImagesAtCell(cell.Start.Row, cell.Start.Column) ?? false;

                // 合併判斷：In-Cell 圖片或浮動圖片
                var hasImages = hasInCellPicture || hasDrawingImages;

                // 判斷內容類型
                if (!hasText && !hasImages)
                    return CellContentType.Empty;
                else if (hasText && !hasImages)
                    return CellContentType.TextOnly;
                else if (!hasText && hasImages)
                    return CellContentType.ImageOnly;
                else
                    return CellContentType.Mixed;
            }
            catch (Exception ex)
            {
                _logger.LogDebug($"檢測儲存格 {cell.Address} 內容類型時發生錯誤: {ex.Message}");
                return CellContentType.Mixed; // 預設為混合類型以確保完整處理
            }
        }

        /// <summary>
        /// 偵測儲存格內容類型 (不使用索引)
        /// 從 ExcelController.DetectCellContentType 搬移
        /// </summary>
        public CellContentType DetectCellContentType(ExcelRange cell, ExcelWorksheet worksheet)
        {
            try
            {
                // 檢查是否有文字內容
                var hasText = !string.IsNullOrEmpty(cell.Text) || !string.IsNullOrEmpty(cell.Formula);

                // ⭐ EPPlus 8.x: 優先檢查 In-Cell 圖片
                bool hasInCellPicture = false;
                try
                {
                    // 只有單一儲存格才能檢查 In-Cell Picture
                    if (cell.Start.Row == cell.End.Row && cell.Start.Column == cell.End.Column)
                    {
                        hasInCellPicture = cell.Picture.Exists;
                    }
                }
                catch
                {
                    // 忽略 Picture API 錯誤
                }

                // 快速檢查是否有浮動圖片（僅檢查位置，不做詳細處理）
                var hasDrawingImages = false;

                if (worksheet.Drawings != null && worksheet.Drawings.Any())
                {
                    var cellStartRow = cell.Start.Row;
                    var cellEndRow = cell.End.Row;
                    var cellStartCol = cell.Start.Column;
                    var cellEndCol = cell.End.Column;

                    foreach (var drawing in worksheet.Drawings.Take(100)) // 檢查更多物件以確保不會遺漏
                    {
                        if (drawing is ExcelPicture picture)
                        {
                            if (picture.From != null)
                            {
                                var fromRow = picture.From.Row + 1;
                                var fromCol = picture.From.Column + 1;

                                // 精確的位置檢查（與 GetCellImages 一致）
                                if (fromRow >= cellStartRow && fromRow <= cellEndRow &&
                                    fromCol >= cellStartCol && fromCol <= cellEndCol)
                                {
                                    hasDrawingImages = true;
                                    break;
                                }
                            }
                        }
                    }
                }

                // 合併判斷：In-Cell 圖片或浮動圖片
                var hasImages = hasInCellPicture || hasDrawingImages;

                // 判斷內容類型
                if (!hasText && !hasImages)
                    return CellContentType.Empty;
                else if (hasText && !hasImages)
                    return CellContentType.TextOnly;
                else if (!hasText && hasImages)
                    return CellContentType.ImageOnly;
                else
                    return CellContentType.Mixed;
            }
            catch (Exception ex)
            {
                _logger.LogDebug($"檢測儲存格 {cell.Address} 內容類型時發生錯誤: {ex.Message}");
                return CellContentType.Mixed; // 預設為混合類型以確保完整處理
            }
        }

        #endregion

        #region Data and Default Methods

        /// <summary>
        /// 取得原始儲存格資料陣列
        /// 從 ExcelController.GetRawCellData 搬移 (135行)
        /// </summary>
        public object[,] GetRawCellData(ExcelWorksheet worksheet, int maxRows, int maxCols)
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
                            Color = _colorService.GetColorFromExcelColor(cell.Style.Font.Color),
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

                        // 邊框 - 使用 ColorService
                        Border = new
                        {
                            Top = new { Style = cell.Style.Border.Top.Style.ToString(), Color = _colorService.GetColorFromExcelColor(cell.Style.Border.Top.Color) },
                            Bottom = new { Style = cell.Style.Border.Bottom.Style.ToString(), Color = _colorService.GetColorFromExcelColor(cell.Style.Border.Bottom.Color) },
                            Left = new { Style = cell.Style.Border.Left.Style.ToString(), Color = _colorService.GetColorFromExcelColor(cell.Style.Border.Left.Color) },
                            Right = new { Style = cell.Style.Border.Right.Style.ToString(), Color = _colorService.GetColorFromExcelColor(cell.Style.Border.Right.Color) },
                            Diagonal = new { Style = cell.Style.Border.Diagonal.Style.ToString(), Color = _colorService.GetColorFromExcelColor(cell.Style.Border.Diagonal.Color) },
                            DiagonalUp = cell.Style.Border.DiagonalUp,
                            DiagonalDown = cell.Style.Border.DiagonalDown
                        },

                        // 填充/背景 - 使用 ColorService
                        Fill = new
                        {
                            PatternType = cell.Style.Fill.PatternType.ToString(),
                            BackgroundColor = _colorService.GetColorFromExcelColor(cell.Style.Fill.BackgroundColor),
                            PatternColor = _colorService.GetColorFromExcelColor(cell.Style.Fill.PatternColor),
                            BackgroundColorTheme = cell.Style.Fill.BackgroundColor.Theme?.ToString(),
                            BackgroundColorTint = cell.Style.Fill.BackgroundColor.Tint
                        },

                        // 尺寸和合併
                        Dimensions = new
                        {
                            ColumnWidth = column.Width > 0 ? column.Width : worksheet.DefaultColWidth,
                            RowHeight = worksheet.Row(row).Height,
                            IsMerged = cell.Merge,
                            MergedRangeAddress = cell.Merge ? FindMergedRangeAddress(worksheet, row, col) : null
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

        /// <summary>
        /// 創建預設字型資訊
        /// 從 ExcelController.CreateDefaultFontInfo 搬移
        /// </summary>
        public FontInfo CreateDefaultFontInfo()
        {
            return new FontInfo
            {
                Name = "Calibri",
                Size = 11,
                Bold = false,
                Italic = false,
                UnderLine = "None",
                Strike = false,
                Color = "000000",
                ColorTheme = null,
                ColorTint = null,
                Charset = 1,
                Scheme = null,
                Family = 2
            };
        }

        /// <summary>
        /// 創建預設對齊資訊
        /// 從 ExcelController.CreateDefaultAlignmentInfo 搬移
        /// </summary>
        public AlignmentInfo CreateDefaultAlignmentInfo()
        {
            return new AlignmentInfo
            {
                Horizontal = "General",
                Vertical = "Bottom",
                WrapText = false,
                Indent = 0,
                ReadingOrder = "ContextDependent",
                TextRotation = 0,
                ShrinkToFit = false
            };
        }

        /// <summary>
        /// 創建預設邊框資訊
        /// 從 ExcelController.CreateDefaultBorderInfo 搬移
        /// </summary>
        public BorderInfo CreateDefaultBorderInfo()
        {
            var defaultBorderStyle = new BorderStyle { Style = "None", Color = null };
            return new BorderInfo
            {
                Top = defaultBorderStyle,
                Bottom = defaultBorderStyle,
                Left = defaultBorderStyle,
                Right = defaultBorderStyle,
                Diagonal = defaultBorderStyle,
                DiagonalUp = false,
                DiagonalDown = false
            };
        }

        /// <summary>
        /// 創建預設填充資訊
        /// 從 ExcelController.CreateDefaultFillInfo 搬移
        /// </summary>
        public FillInfo CreateDefaultFillInfo()
        {
            return new FillInfo
            {
                PatternType = "None",
                BackgroundColor = null,
                PatternColor = null,
                BackgroundColorTheme = null,
                BackgroundColorTint = null
            };
        }

        /// <summary>
        /// 安全取得儲存格值
        /// 從 ExcelController.GetSafeValue 搬移
        /// </summary>
        public string? GetSafeValue(ExcelRange cell)
        {
            var value = cell?.Value;
            if (value == null)
                return null;

            try
            {
                // 獲取值的類型
                var valueType = value.GetType();

                // 如果是基本類型（string, int, double, bool, DateTime 等），直接返回
                if (valueType.IsPrimitive || value is string || value is DateTime || value is decimal)
                {
                    return value.ToString();
                }

                var typeName = valueType.FullName ?? valueType.Name;

                // 🚀 特別處理: 檢測 EPPlus 圖片相關類型 (In-Cell Images)
                if (typeName.Contains("Picture", StringComparison.OrdinalIgnoreCase) ||
                    typeName.Contains("Image", StringComparison.OrdinalIgnoreCase) ||
                    typeName.Contains("Drawing", StringComparison.OrdinalIgnoreCase) ||
                    typeName.Contains("ExcelPicture", StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogDebug($"檢測到 In-Cell 圖片類型 {typeName}，返回 null (圖片資訊將在 Images 屬性中提供)");
                    return null; // 返回 null，圖片資訊會在 cellInfo.Images 中處理
                }

                // 如果類型名稱包含 "Compile" 或 "Result"（EPPlus 內部類型），嘗試轉換為字串
                if (typeName.Contains("Compile", StringComparison.OrdinalIgnoreCase) ||
                    typeName.Contains("Result", StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogWarning($"檢測到 EPPlus 內部類型 {typeName}，轉換為字串以避免循環引用");
                    return value.ToString();
                }

                // 對於其他複雜類型，也轉換為字串
                _logger.LogDebug($"將複雜類型 {typeName} 轉換為字串");
                return value.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"轉換 cell.Value 時發生錯誤: {ex.Message}，返回 null");
                return null;
            }
        }

        #endregion

        #region Helper Methods (Private)

        /// <summary>
        /// 取得欄名稱 (Excel 字母格式)
        /// </summary>
        private string GetColumnName(int column)
        {
            string columnName = "";
            while (column > 0)
            {
                int modulo = (column - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                column = (column - modulo) / 26;
            }
            return columnName;
        }

        /// <summary>
        /// 查找合併儲存格範圍地址
        /// </summary>
        private string? FindMergedRangeAddress(ExcelWorksheet worksheet, int row, int column)
        {
            foreach (var mergedRange in worksheet.MergedCells)
            {
                var range = worksheet.Cells[mergedRange];
                if (row >= range.Start.Row && row <= range.End.Row &&
                    column >= range.Start.Column && column <= range.End.Column)
                {
                    return range.Address;
                }
            }
            return null;
        }

        #endregion
    }
}
