using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using ExcelReaderAPI.Models;
using ExcelReaderAPI.Utils;
using System.Data;

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
            var cellInfo = new ExcelCellInfo();

            // 位置資訊
            cellInfo.Position = new CellPosition
            {
                Row = cell.Start.Row,
                Column = cell.Start.Column,
                Address = cell.Address
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
                Color = cell.Style.Font.Color.Rgb,
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

            // 1. 優先使用 RGB 值
            if (!string.IsNullOrEmpty(excelColor.Rgb))
            {
                var colorValue = excelColor.Rgb.TrimStart('#');
                
                // 處理 ARGB 格式（8位）轉為 RGB 格式（6位）
                if (colorValue.Length == 8)
                {
                    colorValue = colorValue.Substring(2);
                }
                
                if (colorValue.Length == 6)
                {
                    return colorValue;
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
                            Color = cell.Style.Font.Color.Rgb,
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