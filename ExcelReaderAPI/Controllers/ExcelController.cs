using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
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

        private IXLRange? FindMergedRange(IXLWorksheet worksheet, int row, int column)
        {
            // ClosedXML 提供更簡單的合併儲存格檢查
            var cell = worksheet.Cell(row, column);
            return cell.IsMerged() ? cell.MergedRange() : null;
        }

        private string? GetTextAlign(XLAlignmentHorizontalValues alignment)
        {
            return alignment switch
            {
                XLAlignmentHorizontalValues.Left => "left",
                XLAlignmentHorizontalValues.Center => "center",
                XLAlignmentHorizontalValues.Right => "right",
                XLAlignmentHorizontalValues.Justify => "justify",
                XLAlignmentHorizontalValues.Fill => "left",
                XLAlignmentHorizontalValues.CenterContinuous => "center",
                XLAlignmentHorizontalValues.Distributed => "justify",
                _ => null
            };
        }

        private double GetColumnWidth(IXLWorksheet worksheet, int columnIndex)
        {
            return worksheet.Column(columnIndex).Width;
        }

        private ExcelCellInfo CreateCellInfo(IXLCell cell, IXLWorksheet worksheet)
        {
            if (cell == null || worksheet == null)
                throw new ArgumentNullException("Cell or worksheet cannot be null");

            var cellInfo = new ExcelCellInfo();

            try
            {
                // 位置資訊
                cellInfo.Position = new CellPosition
                {
                    Row = cell.Address.RowNumber,
                    Column = cell.Address.ColumnNumber,
                    Address = cell.Address.ToString() ?? ""
                };

                // 基本值和顯示
                cellInfo.Value = cell.Value.IsBlank ? null : cell.Value.ToString();
                cellInfo.Text = cell.GetFormattedString();
                cellInfo.Formula = cell.HasFormula ? cell.FormulaA1 : string.Empty;
                cellInfo.FormulaR1C1 = cell.HasFormula ? cell.FormulaR1C1 : string.Empty;

                // 資料類型
                cellInfo.ValueType = cell.Value.Type.ToString();
                if (cell.Value.IsBlank)
                {
                    cellInfo.DataType = "Empty";
                }
                else if (cell.Value.IsDateTime)
                {
                    cellInfo.DataType = "DateTime";
                }
                else if (cell.Value.IsNumber)
                {
                    cellInfo.DataType = "Number";
                }
                else if (cell.Value.IsBoolean)
                {
                    cellInfo.DataType = "Boolean";
                }
                else
                {
                    cellInfo.DataType = "Text";
                }

                // 格式化
                cellInfo.NumberFormat = cell.Style.NumberFormat.Format;
                cellInfo.NumberFormatId = cell.Style.NumberFormat.NumberFormatId;

                // 字體樣式 - ClosedXML 的顏色處理更直觀
                cellInfo.Font = new FontInfo
                {
                    Name = cell.Style.Font.FontName,
                    Size = (float)cell.Style.Font.FontSize,
                    Bold = cell.Style.Font.Bold,
                    Italic = cell.Style.Font.Italic,
                    UnderLine = cell.Style.Font.Underline.ToString(),
                    Strike = cell.Style.Font.Strikethrough,
                    Color = GetColorFromXLColor(cell.Style.Font.FontColor),
                    ColorTheme = null, // ClosedXML 處理主題顏色的方式不同
                    ColorTint = null,
                    Scheme = null,
                    Family = (int?)cell.Style.Font.FontFamilyNumbering
                };

                // 對齊方式
                cellInfo.Alignment = new AlignmentInfo
                {
                    Horizontal = cell.Style.Alignment.Horizontal.ToString(),
                    Vertical = cell.Style.Alignment.Vertical.ToString(),
                    WrapText = cell.Style.Alignment.WrapText,
                    Indent = cell.Style.Alignment.Indent,
                    ReadingOrder = cell.Style.Alignment.ReadingOrder.ToString(),
                    TextRotation = cell.Style.Alignment.TextRotation,
                    ShrinkToFit = cell.Style.Alignment.ShrinkToFit
                };

                // 邊框
                cellInfo.Border = new BorderInfo
                {
                    Top = new BorderStyle 
                    { 
                        Style = cell.Style.Border.TopBorder.ToString(), 
                        Color = GetColorFromXLColor(cell.Style.Border.TopBorderColor)
                    },
                    Bottom = new BorderStyle 
                    { 
                        Style = cell.Style.Border.BottomBorder.ToString(), 
                        Color = GetColorFromXLColor(cell.Style.Border.BottomBorderColor)
                    },
                    Left = new BorderStyle 
                    { 
                        Style = cell.Style.Border.LeftBorder.ToString(), 
                        Color = GetColorFromXLColor(cell.Style.Border.LeftBorderColor)
                    },
                    Right = new BorderStyle 
                    { 
                        Style = cell.Style.Border.RightBorder.ToString(), 
                        Color = GetColorFromXLColor(cell.Style.Border.RightBorderColor)
                    },
                    Diagonal = new BorderStyle 
                    { 
                        Style = cell.Style.Border.DiagonalBorder.ToString(), 
                        Color = GetColorFromXLColor(cell.Style.Border.DiagonalBorderColor)
                    },
                    DiagonalUp = cell.Style.Border.DiagonalUp,
                    DiagonalDown = cell.Style.Border.DiagonalDown
                };

                // 填充/背景
                cellInfo.Fill = new FillInfo
                {
                    PatternType = cell.Style.Fill.PatternType.ToString(),
                    BackgroundColor = GetColorFromXLColor(cell.Style.Fill.BackgroundColor),
                    PatternColor = GetColorFromXLColor(cell.Style.Fill.PatternColor),
                    BackgroundColorTheme = null,
                    BackgroundColorTint = null
                };

                // 尺寸和合併
                cellInfo.Dimensions = new DimensionInfo
                {
                    ColumnWidth = worksheet.Column(cell.Address.ColumnNumber).Width,
                    RowHeight = worksheet.Row(cell.Address.RowNumber).Height,
                    IsMerged = cell.IsMerged()
                };

                // 檢查是否為合併儲存格
                if (cell.IsMerged())
                {
                    var mergedRange = cell.MergedRange();
                    cellInfo.Dimensions.MergedRangeAddress = mergedRange.RangeAddress.ToString();
                    cellInfo.Dimensions.IsMainMergedCell = (cell.Address.RowNumber == mergedRange.FirstCell().Address.RowNumber && 
                                                           cell.Address.ColumnNumber == mergedRange.FirstCell().Address.ColumnNumber);
                    
                    if (cellInfo.Dimensions.IsMainMergedCell == true)
                    {
                        cellInfo.Dimensions.RowSpan = mergedRange.RowCount();
                        cellInfo.Dimensions.ColSpan = mergedRange.ColumnCount();
                        
                        // 對於主合併儲存格，使用整個合併範圍的邊框
                        cellInfo.Border = GetMergedCellBorder(worksheet, mergedRange, cell);
                    }
                    else
                    {
                        cellInfo.Dimensions.RowSpan = 1;
                        cellInfo.Dimensions.ColSpan = 1;
                    }
                }

                // Rich Text - ClosedXML 的 Rich Text 處理更準確
                if (cell.HasRichText)
                {
                    cellInfo.RichText = new List<RichTextPart>();
                    
                    foreach (var richTextRun in cell.GetRichText())
                    {
                        cellInfo.RichText.Add(new RichTextPart
                        {
                            Text = richTextRun.Text,
                            Bold = richTextRun.Bold,
                            Italic = richTextRun.Italic,
                            UnderLine = richTextRun.Underline != XLFontUnderlineValues.None,
                            Strike = richTextRun.Strikethrough,
                            Size = (float)richTextRun.FontSize,
                            FontName = richTextRun.FontName,
                            Color = GetColorFromXLColor(richTextRun.FontColor), // ClosedXML 的顏色處理更準確
                            VerticalAlign = richTextRun.VerticalAlignment.ToString()
                        });
                    }
                }

                // 註解
                if (cell.HasComment)
                {
                    cellInfo.Comment = new CommentInfo
                    {
                        Text = cell.GetComment().Text,
                        Author = cell.GetComment().Author,
                        AutoFit = false, // ClosedXML 中需要不同的處理方式
                        Visible = cell.GetComment().Visible
                    };
                }

                // 超連結
                if (cell.HasHyperlink)
                {
                    cellInfo.Hyperlink = new HyperlinkInfo
                    {
                        AbsoluteUri = cell.GetHyperlink().ExternalAddress?.ToString(),
                        OriginalString = cell.GetHyperlink().ExternalAddress?.ToString(),
                        IsAbsoluteUri = cell.GetHyperlink().IsExternal
                    };
                }

                // 中繼資料
                cellInfo.Metadata = new CellMetadata
                {
                    HasFormula = cell.HasFormula,
                    IsRichText = cell.HasRichText,
                    StyleId = 0, // ClosedXML 不直接暴露 StyleId
                    StyleName = string.Empty,
                    Rows = 1,
                    Columns = 1,
                    Start = new CellPosition 
                    { 
                        Row = cell.Address.RowNumber, 
                        Column = cell.Address.ColumnNumber, 
                        Address = cell.Address.ToString() ?? ""
                    },
                    End = new CellPosition 
                    { 
                        Row = cell.Address.RowNumber, 
                        Column = cell.Address.ColumnNumber, 
                        Address = cell.Address.ToString() ?? ""
                    }
                };

                return cellInfo;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"讀取儲存格 {cell?.Address?.ToString() ?? "未知位置"} 時發生錯誤");
                
                // 返回基本的儲存格資訊，避免整個處理中斷
                return new ExcelCellInfo
                {
                    Position = new CellPosition
                    {
                        Row = cell?.Address?.RowNumber ?? 0,
                        Column = cell?.Address?.ColumnNumber ?? 0,
                        Address = cell?.Address?.ToString() ?? "未知"
                    },
                    Value = null,
                    Text = "",
                    DataType = "Error",
                    Font = new FontInfo { Color = "000000" }
                };
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
                columnNumber--; // 轉換為 0-based 索引
                columnName = (char)('A' + (columnNumber % 26)) + columnName;
                columnNumber /= 26;
            }
            return columnName;
        }

        /// <summary>
        /// 獲取合併儲存格的邊框 (考慮整個合併範圍的外邊界)
        /// </summary>
        private BorderInfo GetMergedCellBorder(IXLWorksheet worksheet, IXLRange mergedRange, IXLCell currentCell)
        {
            var border = new BorderInfo();
            
            // 獲取合併範圍的邊界
            int topRow = mergedRange.FirstCell().Address.RowNumber;
            int bottomRow = mergedRange.LastCell().Address.RowNumber;
            int leftCol = mergedRange.FirstCell().Address.ColumnNumber;
            int rightCol = mergedRange.LastCell().Address.ColumnNumber;
            
            // 上邊框：來自合併範圍頂部的儲存格
            var topCell = worksheet.Cell(topRow, currentCell.Address.ColumnNumber);
            border.Top = new BorderStyle 
            { 
                Style = topCell.Style.Border.TopBorder.ToString(), 
                Color = GetColorFromXLColor(topCell.Style.Border.TopBorderColor)
            };
            
            // 下邊框：來自合併範圍底部的儲存格
            var bottomCell = worksheet.Cell(bottomRow, currentCell.Address.ColumnNumber);
            border.Bottom = new BorderStyle 
            { 
                Style = bottomCell.Style.Border.BottomBorder.ToString(), 
                Color = GetColorFromXLColor(bottomCell.Style.Border.BottomBorderColor)
            };
            
            // 左邊框：來自合併範圍左側的儲存格
            var leftCell = worksheet.Cell(currentCell.Address.RowNumber, leftCol);
            border.Left = new BorderStyle 
            { 
                Style = leftCell.Style.Border.LeftBorder.ToString(), 
                Color = GetColorFromXLColor(leftCell.Style.Border.LeftBorderColor)
            };
            
            // 右邊框：來自合併範圍右側的儲存格
            var rightCell = worksheet.Cell(currentCell.Address.RowNumber, rightCol);
            border.Right = new BorderStyle 
            { 
                Style = rightCell.Style.Border.RightBorder.ToString(), 
                Color = GetColorFromXLColor(rightCell.Style.Border.RightBorderColor)
            };
            
            // 對角線邊框使用當前儲存格的設定
            border.Diagonal = new BorderStyle 
            { 
                Style = currentCell.Style.Border.DiagonalBorder.ToString(), 
                Color = GetColorFromXLColor(currentCell.Style.Border.DiagonalBorderColor)
            };
            border.DiagonalUp = currentCell.Style.Border.DiagonalUp;
            border.DiagonalDown = currentCell.Style.Border.DiagonalDown;
            
            return border;
        }

        /// <summary>
        /// 從 ClosedXML XLColor 物件提取顏色值
        /// ClosedXML 的顏色處理比 EPPlus 更直觀和準確
        /// </summary>
        private string? GetColorFromXLColor(XLColor xlColor)
        {
            if (xlColor == null)
                return null;

            try
            {
                // 檢查顏色類型，避免主題顏色轉換錯誤
                switch (xlColor.ColorType)
                {
                    case XLColorType.Color:
                        // 直接的顏色值
                        var color = xlColor.Color;
                        if (color.A == 0)
                            return null;
                        return $"rgb({color.R},{color.G},{color.B})";
                        
                    case XLColorType.Theme:
                        // 主題顏色，使用索引和色調
                        return GetThemeColorRgb(xlColor.ThemeColor, xlColor.ThemeTint);
                        
                    case XLColorType.Indexed:
                        // 索引顏色
                        return GetIndexedColorRgb(xlColor.Indexed);
                        
                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"處理顏色時發生錯誤，顏色類型: {xlColor.ColorType}");
                return null;
            }
        }

        /// <summary>
        /// 獲取主題顏色的 RGB 值
        /// </summary>
        private string? GetThemeColorRgb(XLThemeColor themeColor, double tint)
        {
            // Office 標準主題顏色對應表
            var themeColors = new Dictionary<XLThemeColor, (int R, int G, int B)>
            {
                { XLThemeColor.Background1, (255, 255, 255) },  // White
                { XLThemeColor.Text1, (0, 0, 0) },              // Black
                { XLThemeColor.Background2, (231, 230, 230) },  // Light Gray
                { XLThemeColor.Text2, (68, 84, 106) },          // Dark Gray
                { XLThemeColor.Accent1, (91, 155, 213) },       // Blue
                { XLThemeColor.Accent2, (112, 173, 71) },       // Green
                { XLThemeColor.Accent3, (165, 165, 165) },      // Gray
                { XLThemeColor.Accent4, (255, 192, 0) },        // Orange
                { XLThemeColor.Accent5, (68, 114, 196) },       // Dark Blue
                { XLThemeColor.Accent6, (38, 68, 120) },        // Navy Blue
                { XLThemeColor.Hyperlink, (5, 99, 193) },       // Hyperlink Blue
                { XLThemeColor.FollowedHyperlink, (149, 79, 114) } // Followed Hyperlink Purple
            };

            if (!themeColors.ContainsKey(themeColor))
            {
                return "rgb(0,0,0)"; // 預設黑色
            }

            var baseColor = themeColors[themeColor];
            
            // 應用色調調整
            if (Math.Abs(tint) > 0.001)
            {
                var adjustedColor = ApplyTintToRgb(baseColor, tint);
                return $"rgb({adjustedColor.R},{adjustedColor.G},{adjustedColor.B})";
            }
            
            return $"rgb({baseColor.R},{baseColor.G},{baseColor.B})";
        }

        /// <summary>
        /// 獲取索引顏色的 RGB 值
        /// </summary>
        private string? GetIndexedColorRgb(int colorIndex)
        {
            // Excel 標準索引顏色對應表
            var indexedColors = new Dictionary<int, (int R, int G, int B)>
            {
                { 0, (0, 0, 0) },        // Black
                { 1, (255, 255, 255) },  // White
                { 2, (255, 0, 0) },      // Red
                { 3, (0, 255, 0) },      // Bright Green
                { 4, (0, 0, 255) },      // Blue
                { 5, (255, 255, 0) },    // Yellow
                { 6, (255, 0, 255) },    // Magenta
                { 7, (0, 255, 255) },    // Cyan
                { 8, (0, 0, 0) },        // Black
                { 9, (255, 255, 255) },  // White
                { 10, (255, 0, 0) },     // Red
                { 16, (128, 0, 0) },     // Dark Red
                { 17, (0, 128, 0) },     // Green
                { 18, (0, 0, 128) },     // Dark Blue
                { 19, (128, 128, 0) },   // Dark Yellow
                { 20, (128, 0, 128) },   // Purple
                { 21, (0, 128, 128) },   // Dark Cyan
                { 22, (192, 192, 192) }, // Light Gray
                { 23, (128, 128, 128) }  // Gray
            };
            
            if (indexedColors.ContainsKey(colorIndex))
            {
                var color = indexedColors[colorIndex];
                return $"rgb({color.R},{color.G},{color.B})";
            }
            
            return "rgb(0,0,0)"; // 預設黑色
        }

        /// <summary>
        /// 對 RGB 顏色應用色調效果
        /// </summary>
        private (int R, int G, int B) ApplyTintToRgb((int R, int G, int B) baseColor, double tint)
        {
            int r, g, b;
            
            if (tint < 0)
            {
                // Tint < 0: 變暗
                r = (int)(baseColor.R * (1 + tint));
                g = (int)(baseColor.G * (1 + tint));
                b = (int)(baseColor.B * (1 + tint));
            }
            else
            {
                // Tint > 0: 變亮
                r = (int)(baseColor.R + (255 - baseColor.R) * tint);
                g = (int)(baseColor.G + (255 - baseColor.G) * tint);
                b = (int)(baseColor.B + (255 - baseColor.B) * tint);
            }
            
            // 確保值在 0-255 範圍內
            r = Math.Max(0, Math.Min(255, r));
            g = Math.Max(0, Math.Min(255, g));
            b = Math.Max(0, Math.Min(255, b));
            
            return (r, g, b);
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

                using var workbook = new XLWorkbook(stream);
                
                // 取得所有工作表名稱
                excelData.AvailableWorksheets = workbook.Worksheets.Select(ws => ws.Name).ToList();
                
                var worksheet = workbook.Worksheet(1); // 使用第一個工作表
                excelData.WorksheetName = worksheet.Name;

                // ClosedXML 使用 RangeUsed 來獲取有效範圍
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                {
                    return BadRequest(new UploadResponse
                    {
                        Success = false,
                        Message = "Excel 檔案為空或無有效資料"
                    });
                }

                var rowCount = usedRange.RowCount();
                var colCount = usedRange.ColumnCount();

                excelData.TotalRows = rowCount;
                excelData.TotalColumns = colCount;

                // 生成 Excel 欄位標頭 (A, B, C, D...) 包含寬度資訊 - 使用 ClosedXML 原生方法
                var columnHeaders = new List<object>();
                var firstColumnNumber = usedRange.FirstColumn().ColumnNumber();
                var lastColumnNumber = usedRange.LastColumn().ColumnNumber();
                
                for (int col = firstColumnNumber; col <= lastColumnNumber; col++)
                {
                    var column = worksheet.Column(col);
                    var width = column.Width;
                    var cell = worksheet.Cell(1, col);
                    
                    columnHeaders.Add(new 
                    {
                        Name = cell.Address.ColumnLetter, // 使用 ClosedXML 原生的欄位字母
                        Width = width,
                        Index = col
                    });
                }

                // 讀取第一行內容作為內容標頭，保留格式信息
                var contentHeaders = new List<object>();
                for (int col = firstColumnNumber; col <= lastColumnNumber; col++)
                {
                    var headerCell = worksheet.Cell(1, col);
                    contentHeaders.Add(CreateCellInfo(headerCell, worksheet));
                }
                
                // 提供兩種標頭：Excel 欄位標頭和內容標頭
                excelData.Headers = new[] { columnHeaders.ToArray(), contentHeaders.ToArray() };

                // 讀取資料行，保留原始格式（包含Rich Text）
                var rows = new List<object[]>();
                for (int row = 1; row <= rowCount; row++)
                {
                    var rowData = new List<object>();
                    for (int col = firstColumnNumber; col <= lastColumnNumber; col++)
                    {
                        var cell = worksheet.Cell(row, col);
                        rowData.Add(CreateCellInfo(cell, worksheet));
                    }
                    rows.Add(rowData.ToArray());
                }

                excelData.Rows = rows.ToArray();

                _logger.LogInformation($"成功讀取 Excel 檔案 (ClosedXML): {file.FileName}, 行數: {rowCount}, 欄數: {colCount}");

                return Ok(new UploadResponse
                {
                    Success = true,
                    Message = $"成功讀取 Excel 檔案 (ClosedXML)，共 {rowCount - 1} 筆資料",
                    Data = excelData
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "讀取 Excel 檔案時發生錯誤 (ClosedXML)");
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
                FileName = "範例資料-ClosedXML.xlsx",
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
        public async Task<ActionResult<UploadResponse>> UploadExcelWorksheet(IFormFile file, [FromQuery] string? worksheetName = null, [FromQuery] int worksheetIndex = 1)
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

                using var workbook = new XLWorkbook(stream);
                var excelData = new ExcelData
                {
                    FileName = file.FileName,
                    AvailableWorksheets = workbook.Worksheets.Select(ws => ws.Name).ToList()
                };

                // 選擇工作表
                IXLWorksheet? worksheet = null;
                if (!string.IsNullOrEmpty(worksheetName))
                {
                    worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
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
                    if (worksheetIndex > workbook.Worksheets.Count)
                    {
                        return BadRequest(new UploadResponse
                        {
                            Success = false,
                            Message = $"工作表索引 {worksheetIndex} 超出範圍"
                        });
                    }
                    worksheet = workbook.Worksheet(worksheetIndex);
                }

                excelData.WorksheetName = worksheet.Name;

                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                {
                    return BadRequest(new UploadResponse
                    {
                        Success = false,
                        Message = "選擇的工作表為空或無有效資料"
                    });
                }

                var rowCount = usedRange.RowCount();
                var colCount = usedRange.ColumnCount();
                excelData.TotalRows = rowCount;
                excelData.TotalColumns = colCount;

                // 生成 Excel 欄位標頭 (A, B, C, D...) 包含寬度資訊 - 使用 ClosedXML 原生方法
                var columnHeaders = new List<object>();
                var firstColumnNumber = usedRange.FirstColumn().ColumnNumber();
                var lastColumnNumber = usedRange.LastColumn().ColumnNumber();
                
                for (int col = firstColumnNumber; col <= lastColumnNumber; col++)
                {
                    var column = worksheet.Column(col);
                    var width = column.Width;
                    var cell = worksheet.Cell(1, col);
                    
                    columnHeaders.Add(new 
                    {
                        Name = cell.Address.ColumnLetter, // 使用 ClosedXML 原生的欄位字母
                        Width = width,
                        Index = col
                    });
                }

                // 讀取第一行內容作為內容標頭，保留格式信息
                var contentHeaders = new List<object>();
                for (int col = firstColumnNumber; col <= lastColumnNumber; col++)
                {
                    var headerCell = worksheet.Cell(1, col);
                    contentHeaders.Add(CreateCellInfo(headerCell, worksheet));
                }
                
                // 提供兩種標頭：Excel 欄位標頭和內容標頭
                excelData.Headers = new[] { columnHeaders.ToArray(), contentHeaders.ToArray() };

                var rows = new List<object[]>();
                for (int row = 1; row <= rowCount; row++)
                {
                    var rowData = new List<object>();
                    for (int col = firstColumnNumber; col <= lastColumnNumber; col++)
                    {
                        var cell = worksheet.Cell(row, col);
                        rowData.Add(CreateCellInfo(cell, worksheet));
                    }
                    rows.Add(rowData.ToArray());
                }
                excelData.Rows = rows.ToArray();

                return Ok(new UploadResponse
                {
                    Success = true,
                    Message = $"成功讀取工作表 '{worksheet.Name}' (ClosedXML)，共 {rowCount - 1} 筆資料",
                    Data = excelData
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "讀取 Excel 工作表時發生錯誤 (ClosedXML)");
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
                // 使用 ClosedXML 創建範例 Excel
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("範例資料");

                // 設定標頭
                worksheet.Cell("A1").Value = "姓名";
                worksheet.Cell("B1").Value = "年齡";
                worksheet.Cell("C1").Value = "部門";
                worksheet.Cell("D1").Value = "薪資";
                worksheet.Cell("E1").Value = "入職日期";

                // 設定標頭樣式
                var headerRange = worksheet.Range("A1:E1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;

                // 新增資料並測試 Rich Text
                worksheet.Cell("A2").Value = "張三";
                worksheet.Cell("A2").GetRichText().AddText("三").SetFontColor(XLColor.Red); // 測試 Rich Text 顏色
                
                worksheet.Cell("B2").Value = 30;
                worksheet.Cell("C2").Value = "資訊部";
                worksheet.Cell("D2").Value = 50000;
                worksheet.Cell("E2").Value = DateTime.Parse("2020-01-15");

                // 更多測試資料
                worksheet.Cell("A3").Value = "李四";
                worksheet.Cell("B3").Value = 25;
                worksheet.Cell("C3").Value = "人事部";
                worksheet.Cell("D3").Value = 45000;
                worksheet.Cell("E3").Value = DateTime.Parse("2021-03-20");

                // 自動調整欄寬
                worksheet.Columns().AdjustToContents();

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                var fileBytes = stream.ToArray();

                return File(fileBytes, 
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "範例員工資料-ClosedXML.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "產生範例Excel檔案時發生錯誤 (ClosedXML)");
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

                using var workbook = new XLWorkbook(stream);
                var worksheet = workbook.Worksheet(1);

                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                {
                    return BadRequest("Excel 檔案為空或無有效資料");
                }

                var debugData = new DebugExcelData
                {
                    FileName = file.FileName,
                    WorksheetInfo = new WorksheetInfo
                    {
                        Name = worksheet.Name,
                        TotalRows = usedRange.RowCount(),
                        TotalColumns = usedRange.ColumnCount(),
                        DefaultColWidth = 8.43, // ClosedXML 的預設值
                        DefaultRowHeight = 15
                    },
                    SampleCells = GetRawCellDataClosedXML(worksheet, Math.Min(5, usedRange.RowCount()), Math.Min(5, usedRange.ColumnCount())),
                    AllWorksheets = workbook.Worksheets.Select(ws => new
                    {
                        Name = ws.Name,
                        Index = ws.Position,
                        State = ws.Visibility.ToString()
                    }).Cast<object>().ToList()
                };

                return Ok(debugData);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "讀取 Excel 原始資料時發生錯誤 (ClosedXML)");
                return StatusCode(500, $"讀取檔案時發生錯誤: {ex.Message}");
            }
        }

        private object[,] GetRawCellDataClosedXML(IXLWorksheet worksheet, int maxRows, int maxCols)
        {
            var cells = new object[maxRows, maxCols];
            
            for (int row = 1; row <= maxRows; row++)
            {
                for (int col = 1; col <= maxCols; col++)
                {
                    var cell = worksheet.Cell(row, col);
                    var column = worksheet.Column(col);
                    
                    cells[row - 1, col - 1] = new
                    {
                        Position = new { Row = row, Column = col, Address = cell.Address.ToString() },
                        
                        // 基本值和顯示
                        Value = cell.Value.IsBlank ? null : cell.Value.ToString(),
                        Text = cell.GetFormattedString(),
                        Formula = cell.HasFormula ? cell.FormulaA1 : string.Empty,
                        FormulaR1C1 = cell.HasFormula ? cell.FormulaR1C1 : string.Empty,
                        
                        // 資料類型
                        ValueType = cell.Value.Type.ToString(),
                        
                        // 格式化
                        NumberFormat = cell.Style.NumberFormat.Format,
                        NumberFormatId = cell.Style.NumberFormat.NumberFormatId,
                        
                        // 字體樣式
                        Font = new
                        {
                            Name = cell.Style.Font.FontName,
                            Size = cell.Style.Font.FontSize,
                            Bold = cell.Style.Font.Bold,
                            Italic = cell.Style.Font.Italic,
                            Underline = cell.Style.Font.Underline,
                            Strike = cell.Style.Font.Strikethrough,
                            Color = GetColorFromXLColor(cell.Style.Font.FontColor),
                            Family = cell.Style.Font.FontFamilyNumbering
                        },
                        
                        // 對齊方式
                        Alignment = new
                        {
                            Horizontal = cell.Style.Alignment.Horizontal.ToString(),
                            Vertical = cell.Style.Alignment.Vertical.ToString(),
                            WrapText = cell.Style.Alignment.WrapText,
                            Indent = cell.Style.Alignment.Indent,
                            ReadingOrder = cell.Style.Alignment.ReadingOrder.ToString(),
                            TextRotation = cell.Style.Alignment.TextRotation,
                            ShrinkToFit = cell.Style.Alignment.ShrinkToFit
                        },
                        
                        // 邊框
                        Border = new
                        {
                            Top = new { Style = cell.Style.Border.TopBorder.ToString(), Color = GetColorFromXLColor(cell.Style.Border.TopBorderColor) },
                            Bottom = new { Style = cell.Style.Border.BottomBorder.ToString(), Color = GetColorFromXLColor(cell.Style.Border.BottomBorderColor) },
                            Left = new { Style = cell.Style.Border.LeftBorder.ToString(), Color = GetColorFromXLColor(cell.Style.Border.LeftBorderColor) },
                            Right = new { Style = cell.Style.Border.RightBorder.ToString(), Color = GetColorFromXLColor(cell.Style.Border.RightBorderColor) },
                            Diagonal = new { Style = cell.Style.Border.DiagonalBorder.ToString(), Color = GetColorFromXLColor(cell.Style.Border.DiagonalBorderColor) },
                            DiagonalUp = cell.Style.Border.DiagonalUp,
                            DiagonalDown = cell.Style.Border.DiagonalDown
                        },
                        
                        // 填充/背景
                        Fill = new
                        {
                            PatternType = cell.Style.Fill.PatternType.ToString(),
                            BackgroundColor = GetColorFromXLColor(cell.Style.Fill.BackgroundColor),
                            PatternColor = GetColorFromXLColor(cell.Style.Fill.PatternColor)
                        },
                        
                        // 尺寸和合併
                        Dimensions = new
                        {
                            ColumnWidth = column.Width,
                            RowHeight = worksheet.Row(row).Height,
                            IsMerged = cell.IsMerged(),
                            MergedRangeAddress = cell.IsMerged() ? cell.MergedRange().RangeAddress.ToString() : null
                        },
                        
                        // Rich Text - ClosedXML 的 Rich Text 處理更準確
                        RichText = cell.HasRichText ? cell.GetRichText().Select(rt => new
                        {
                            Text = rt.Text,
                            Bold = rt.Bold,
                            Italic = rt.Italic,
                            UnderLine = rt.Underline.ToString(),
                            Strike = rt.Strikethrough,
                            Size = rt.FontSize,
                            FontName = rt.FontName,
                            Color = GetColorFromXLColor(rt.FontColor), // 這裡應該能正確獲取顏色
                            VerticalAlign = rt.VerticalAlignment.ToString()
                        }).ToList() : null,
                        
                        // 註解
                        Comment = cell.HasComment ? new
                        {
                            Text = cell.GetComment().Text,
                            Author = cell.GetComment().Author,
                            Visible = cell.GetComment().Visible
                        } : null,
                        
                        // 超連結
                        Hyperlink = cell.HasHyperlink ? new
                        {
                            ExternalAddress = cell.GetHyperlink().ExternalAddress?.ToString(),
                            IsExternal = cell.GetHyperlink().IsExternal
                        } : null,
                        
                        // 其他屬性
                        Metadata = new
                        {
                            HasFormula = cell.HasFormula,
                            IsRichText = cell.HasRichText,
                            HasComment = cell.HasComment,
                            HasHyperlink = cell.HasHyperlink
                        }
                    };
                }
            }
            
            return cells;
        }
    }
}