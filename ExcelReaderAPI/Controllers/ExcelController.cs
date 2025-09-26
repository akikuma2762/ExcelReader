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
            cellInfo.Border = new BorderInfo
            {
                Top = new BorderStyle 
                { 
                    Style = cell.Style.Border.Top.Style.ToString(), 
                    Color = cell.Style.Border.Top.Color.Rgb 
                },
                Bottom = new BorderStyle 
                { 
                    Style = cell.Style.Border.Bottom.Style.ToString(), 
                    Color = cell.Style.Border.Bottom.Color.Rgb 
                },
                Left = new BorderStyle 
                { 
                    Style = cell.Style.Border.Left.Style.ToString(), 
                    Color = cell.Style.Border.Left.Color.Rgb 
                },
                Right = new BorderStyle 
                { 
                    Style = cell.Style.Border.Right.Style.ToString(), 
                    Color = cell.Style.Border.Right.Color.Rgb 
                },
                Diagonal = new BorderStyle 
                { 
                    Style = cell.Style.Border.Diagonal.Style.ToString(), 
                    Color = cell.Style.Border.Diagonal.Color.Rgb 
                },
                DiagonalUp = cell.Style.Border.DiagonalUp,
                DiagonalDown = cell.Style.Border.DiagonalDown
            };

            // 填充/背景
            cellInfo.Fill = new FillInfo
            {
                PatternType = cell.Style.Fill.PatternType.ToString(),
                BackgroundColor = cell.Style.Fill.BackgroundColor.Rgb,
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
                
                foreach (var richTextPart in cell.RichText)
                {
                    cellInfo.RichText.Add(new RichTextPart
                    {
                        Text = richTextPart.Text,
                        Bold = richTextPart.Bold,
                        Italic = richTextPart.Italic,
                        UnderLine = richTextPart.UnderLine,
                        Strike = richTextPart.Strike,
                        Size = richTextPart.Size,
                        FontName = richTextPart.FontName,
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

                // 生成 Excel 欄位標頭 (A, B, C, D...)
                var columnHeaders = new List<object>();
                for (int col = 1; col <= colCount; col++)
                {
                    columnHeaders.Add(GetColumnName(col));
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

                // 生成 Excel 欄位標頭 (A, B, C, D...)
                var columnHeaders = new List<object>();
                for (int col = 1; col <= colCount; col++)
                {
                    columnHeaders.Add(GetColumnName(col));
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