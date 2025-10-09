using ExcelReaderAPI.Models;
using OfficeOpenXml;
using ExcelReaderAPI.Services.Interfaces;
using OfficeOpenXml.Drawing;
using Microsoft.Extensions.Logging;
using OfficeOpenXml.Style;

namespace ExcelReaderAPI.Services
{
    public class ExcelCellService : IExcelCellService
    {
        private readonly ILogger<ExcelCellService> _logger;
        private readonly IExcelColorService _colorService;

        [ThreadStatic]
        private static Dictionary<string, int>? _worksheetDrawingObjectCounts;

        private const int MAX_DRAWING_OBJECTS_TO_CHECK = 999999;

        public ExcelCellService(
            ILogger<ExcelCellService> logger,
            IExcelColorService colorService)
        {
            _logger = logger;
            _colorService = colorService;
        }

        #region Floating Objects Methods

        /// <summary>
        /// 取得儲存格的浮動物件（文字方塊、圖形等，但不包含圖片）
        /// 原始: ExcelController.GetCellFloatingObjects (行 1462-1640)
        /// </summary>
        public List<FloatingObjectInfo>? GetCellFloatingObjects(ExcelWorksheet worksheet, ExcelRange cell)
        {
            try
            {
                var floatingObjects = new List<FloatingObjectInfo>();

                // 儲存格的邊界 (支援合併儲存格範圍)
                var cellStartRow = cell.Start.Row;
                var cellEndRow = cell.End.Row;
                var cellStartCol = cell.Start.Column;
                var cellEndCol = cell.End.Column;

                _logger.LogDebug($"檢查儲存格 {cell.Address} 的浮動物件，範圍: Row {cellStartRow}-{cellEndRow}, Col {cellStartCol}-{cellEndCol}");

                // 安全檢查：如果已經檢查太多物件，直接跳過這個儲存格
                var currentCount = GetWorksheetDrawingObjectCount(worksheet.Name);
                if (currentCount > MAX_DRAWING_OBJECTS_TO_CHECK)
                {
                    _logger.LogDebug($"儲存格 {cell.Address} 跳過浮動物件檢查 - 已達到檢查限制 ({currentCount})");
                    return null;
                }

                // 檢查所有工作表中的繪圖物件（排除圖片）
                if (worksheet.Drawings != null && worksheet.Drawings.Any())
                {
                    currentCount = GetWorksheetDrawingObjectCount(worksheet.Name);
                    _logger.LogDebug($"工作表 '{worksheet.Name}' 包含 {worksheet.Drawings.Count} 個繪圖物件 (已檢查: {currentCount})");

                    foreach (var drawing in worksheet.Drawings)
                    {
                        // 安全檢查：防止處理過多物件
                        currentCount = IncrementWorksheetDrawingObjectCount(worksheet.Name);
                        if (currentCount > MAX_DRAWING_OBJECTS_TO_CHECK)
                        {
                            _logger.LogWarning($"工作表 '{worksheet.Name}' 已檢查 {currentCount} 個繪圖物件，停止進一步檢查以避免效能問題");
                            return floatingObjects.Any() ? floatingObjects : null;
                        }

                        try
                        {
                            // 排除圖片，只處理其他類型的繪圖物件
                            if (drawing is ExcelPicture)
                            {
                                continue; // 跳過圖片，已在 GetCellImages 中處理
                            }

                            // 安全獲取物件位置
                            int fromRow = 1, fromCol = 1, toRow = 1, toCol = 1;

                            if (drawing.From != null)
                            {
                                fromRow = drawing.From.Row + 1;
                                fromCol = drawing.From.Column + 1;
                            }

                            if (drawing.To != null)
                            {
                                toRow = drawing.To.Row + 1;
                                toCol = drawing.To.Column + 1;
                            }
                            else
                            {
                                toRow = fromRow;
                                toCol = fromCol;
                            }

                            // ⭐ 新邏輯: 解決合併儲存格與浮動物件範圍不一致的問題
                            // 檢查浮動物件是否與儲存格範圍有交集
                            bool hasOverlap = !(toRow < cellStartRow || fromRow > cellEndRow ||
                                               toCol < cellStartCol || fromCol > cellEndCol);

                            // ⭐ 關鍵修復: 改進錨點檢查邏輯，解決合併儲存格導致的文字方塊檢測問題
                            // 檢查浮動物件是否應該歸屬於當前儲存格
                            bool isAnchorCell = false;

                            // 情況1: 浮動物件的起始位置在當前儲存格範圍內
                            bool floatingStartsInCell = (fromRow >= cellStartRow && fromRow <= cellEndRow &&
                                                        fromCol >= cellStartCol && fromCol <= cellEndCol);

                            // 情況2: 當前儲存格是浮動物件覆蓋範圍中的第一個儲存格（左上角優先原則）
                            bool isCellTopLeftOfFloating = (cellStartRow <= fromRow && cellStartCol <= fromCol);

                            // 情況3: 對於合併儲存格，檢查是否為合併範圍的主儲存格
                            bool isMergedCellAnchor = (cellStartRow == cellEndRow && cellStartCol == cellEndCol) || // 單一儲存格
                                                     (cell.Merge && cellStartRow == cell.Start.Row && cellStartCol == cell.Start.Column); // 合併儲存格的主儲存格

                            // 根據不同情況判斷是否為錨點
                            if (floatingStartsInCell && isMergedCellAnchor)
                            {
                                isAnchorCell = true; // 浮動物件在儲存格內且該儲存格是主儲存格
                            }
                            else if (!cell.Merge && floatingStartsInCell)
                            {
                                isAnchorCell = true; // 非合併儲存格且浮動物件在其內
                            }
                            else if (cell.Merge && cellStartRow == fromRow && cellStartCol == fromCol)
                            {
                                isAnchorCell = true; // 合併儲存格且位置完全匹配
                            }

                            // ⭐ 最終決定: 浮動物件需要有交集且符合錨點條件
                            bool shouldInclude = hasOverlap && isAnchorCell;

                            if (shouldInclude)
                            {
                                try
                                {
                                    var floatingObjectInfo = new FloatingObjectInfo
                                    {
                                        Name = drawing.Name ?? $"FloatingObject_{floatingObjects.Count + 1}",
                                        Description = $"Excel 檔案中的浮動物件 ({drawing.GetType().Name})",
                                        ObjectType = GetDrawingObjectType(drawing),
                                        Width = (int)(drawing.To?.Column - drawing.From?.Column ?? 100),
                                        Height = (int)(drawing.To?.Row - drawing.From?.Row ?? 20),
                                        Left = (drawing.From?.ColumnOff ?? 0) / 9525.0,
                                        Top = (drawing.From?.RowOff ?? 0) / 9525.0,
                                        Text = ExtractTextFromDrawing(drawing),
                                        AnchorCell = new CellPosition
                                        {
                                            Row = fromRow,
                                            Column = fromCol,
                                            Address = $"{GetColumnName(fromCol)}{fromRow}"
                                        },
                                        FromCell = new CellPosition
                                        {
                                            Row = fromRow,
                                            Column = fromCol,
                                            Address = $"{GetColumnName(fromCol)}{fromRow}"
                                        },
                                        ToCell = new CellPosition
                                        {
                                            Row = toRow,
                                            Column = toCol,
                                            Address = $"{GetColumnName(toCol)}{toRow}"
                                        },
                                        IsFloating = true,
                                        Style = ExtractStyleFromDrawing(drawing),
                                        HyperlinkAddress = ExtractHyperlinkFromDrawing(drawing)
                                    };

                                    floatingObjects.Add(floatingObjectInfo);
                                }
                                catch (Exception objEx)
                                {
                                    _logger.LogError(objEx, $"處理浮動物件資料時發生錯誤: {objEx.Message}");
                                }
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

                return floatingObjects.Any() ? floatingObjects : null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"讀取儲存格 {cell.Address} 的浮動物件時發生錯誤: {ex.Message}");
                return null;
            }
        }

        public string GetDrawingObjectType(OfficeOpenXml.Drawing.ExcelDrawing drawing)
        {
            var typeName = drawing.GetType().Name;

            return typeName switch
            {
                "ExcelShape" => "Shape",
                "ExcelTextBox" => "TextBox",
                "ExcelChart" => "Chart",
                "ExcelTable" => "Table",
                "ExcelPicture" => "Picture",
                _ => typeName.Replace("Excel", "")
            };
        }

        /// <summary>
        /// 從繪圖物件中提取文字內容
        /// </summary>
        public string? ExtractTextFromDrawing(OfficeOpenXml.Drawing.ExcelDrawing drawing)
        {
            try
            {
                // 使用反射嘗試獲取文字屬性
                var textProperty = drawing.GetType().GetProperty("Text");
                if (textProperty != null)
                {
                    return textProperty.GetValue(drawing)?.ToString();
                }

                // 嘗試其他可能的文字屬性
                var richTextProperty = drawing.GetType().GetProperty("RichText");
                if (richTextProperty != null)
                {
                    var richText = richTextProperty.GetValue(drawing);
                    return richText?.ToString();
                }

                // 如果是 TextBox,嘗試特殊處理
                if (drawing.GetType().Name.Contains("TextBox"))
                {
                    // EPPlus 中 TextBox 的文字可能存儲在不同的屬性中
                    var contentProperty = drawing.GetType().GetProperty("Content");
                    if (contentProperty != null)
                    {
                        return contentProperty.GetValue(drawing)?.ToString();
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"提取繪圖物件文字時發生錯誤: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 從繪圖物件中提取樣式資訊
        /// </summary>
        public string? ExtractStyleFromDrawing(OfficeOpenXml.Drawing.ExcelDrawing drawing)
        {
            try
            {
                var styles = new List<string>();

                // 使用反射嘗試獲取樣式屬性
                var styleProperties = new[] { "Fill", "Border", "Font", "TextAlignment", "Style" };

                foreach (var propName in styleProperties)
                {
                    var property = drawing.GetType().GetProperty(propName);
                    if (property != null)
                    {
                        var value = property.GetValue(drawing);
                        if (value != null)
                        {
                            styles.Add($"{propName}: {value}");
                        }
                    }
                }

                return styles.Any() ? string.Join("; ", styles) : null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"提取繪圖物件樣式時發生錯誤: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 從繪圖物件中提取超連結
        /// </summary>
        public string? ExtractHyperlinkFromDrawing(OfficeOpenXml.Drawing.ExcelDrawing drawing)
        {
            try
            {
                var hyperlinkProperty = drawing.GetType().GetProperty("Hyperlink");
                if (hyperlinkProperty != null)
                {
                    var hyperlink = hyperlinkProperty.GetValue(drawing);
                    return hyperlink?.ToString();
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"提取繪圖物件超連結時發生錯誤: {ex.Message}");
                return null;
            }
        }

        public string GetColumnName(int columnNumber)
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

        private int GetWorksheetDrawingObjectCount(string worksheetName)
        {
            _worksheetDrawingObjectCounts ??= new Dictionary<string, int>();
            return _worksheetDrawingObjectCounts.TryGetValue(worksheetName, out var count) ? count : 0;
        }

        /// <summary>
        /// 增加工作表繪圖物件計數器
        /// </summary>
        private int IncrementWorksheetDrawingObjectCount(string worksheetName)
        {
            _worksheetDrawingObjectCounts ??= new Dictionary<string, int>();
            var count = GetWorksheetDrawingObjectCount(worksheetName) + 1;
            _worksheetDrawingObjectCounts[worksheetName] = count;
            return count;
        }

        /// <summary>
        /// 查找合併儲存格範圍 (返回字串地址)
        /// </summary>
        public string? FindMergedRange(ExcelWorksheet worksheet, ExcelRange cell)
        {
            // 檢查所有合併範圍,找到包含指定儲存格的範圍
            foreach (var mergedRange in worksheet.MergedCells)
            {
                var range = worksheet.Cells[mergedRange];
                if (cell.Start.Row >= range.Start.Row && cell.Start.Row <= range.End.Row &&
                    cell.Start.Column >= range.Start.Column && cell.Start.Column <= range.End.Column)
                {
                    return mergedRange;
                }
            }
            return null;
        }

        /// <summary>
        /// 查找合併儲存格範圍 (返回 ExcelRange 物件,與 Controller 一致)
        /// 原始: ExcelController.FindMergedRange (行 337-350)
        /// </summary>
        public ExcelRange? FindMergedRange(ExcelWorksheet worksheet, int row, int column)
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

        public string GetTextAlign(OfficeOpenXml.Style.ExcelHorizontalAlignment alignment)
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
                _ => "left"
            };
        }

        public double GetColumnWidth(ExcelWorksheet worksheet, int column)
        {
            // 取得該欄的寬度,若未設定則使用預設寬度
            var columnObj = worksheet.Column(column);
            if (columnObj.Width > 0)
            {
                return columnObj.Width;
            }
            else
            {
                // 使用預設欄寬
                return worksheet.DefaultColWidth;
            }
        }

        public BorderInfo? GetMergedCellBorder(ExcelWorksheet worksheet, string mergedRange)
        {
            if (string.IsNullOrEmpty(mergedRange))
                return null;

            var range = worksheet.Cells[mergedRange];
            var border = new BorderInfo();

            // 獲取合併範圍的邊界
            int topRow = range.Start.Row;
            int bottomRow = range.End.Row;
            int leftCol = range.Start.Column;
            int rightCol = range.End.Column;

            // 上邊框:來自合併範圍頂部的儲存格
            var topCell = worksheet.Cells[topRow, leftCol];
            border.Top = new BorderStyle
            {
                Style = topCell.Style.Border.Top.Style.ToString(),
                Color = _colorService.GetColorFromExcelColor(topCell.Style.Border.Top.Color)
            };

            // 下邊框:來自合併範圍底部的儲存格
            var bottomCell = worksheet.Cells[bottomRow, leftCol];
            border.Bottom = new BorderStyle
            {
                Style = bottomCell.Style.Border.Bottom.Style.ToString(),
                Color = _colorService.GetColorFromExcelColor(bottomCell.Style.Border.Bottom.Color)
            };

            // 左邊框:來自合併範圍左側的儲存格
            var leftCell = worksheet.Cells[topRow, leftCol];
            border.Left = new BorderStyle
            {
                Style = leftCell.Style.Border.Left.Style.ToString(),
                Color = _colorService.GetColorFromExcelColor(leftCell.Style.Border.Left.Color)
            };

            // 右邊框:來自合併範圍右側的儲存格
            var rightCell = worksheet.Cells[topRow, rightCol];
            border.Right = new BorderStyle
            {
                Style = rightCell.Style.Border.Right.Style.ToString(),
                Color = _colorService.GetColorFromExcelColor(rightCell.Style.Border.Right.Color)
            };

            // 對角線邊框
            border.Diagonal = new BorderStyle
            {
                Style = topCell.Style.Border.Diagonal.Style.ToString(),
                Color = _colorService.GetColorFromExcelColor(topCell.Style.Border.Diagonal.Color)
            };
            border.DiagonalUp = topCell.Style.Border.DiagonalUp;
            border.DiagonalDown = topCell.Style.Border.DiagonalDown;

            return border;
        }

        public void SetCellMergedInfo(ExcelCellInfo cellInfo, ExcelWorksheet worksheet, ExcelRange cell)
        {
            var mergedRange = FindMergedRange(worksheet, cell);
            if (string.IsNullOrEmpty(mergedRange))
                return;

            var range = worksheet.Cells[mergedRange];
            int fromRow = range.Start.Row;
            int fromCol = range.Start.Column;
            int toRow = range.End.Row;
            int toCol = range.End.Column;

            int rowSpan = toRow - fromRow + 1;
            int colSpan = toCol - fromCol + 1;

            cellInfo.Dimensions.IsMerged = true;
            cellInfo.Dimensions.IsMainMergedCell = (cell.Start.Row == fromRow && cell.Start.Column == fromCol);
            cellInfo.Dimensions.RowSpan = rowSpan;
            cellInfo.Dimensions.ColSpan = colSpan;
            cellInfo.Dimensions.MergedRangeAddress = mergedRange;
        }

        /// <summary>
        /// 設定儲存格的合併資訊 (方法重載 - 用於自動設定合併,與 Controller 一致)
        /// 原始: ExcelController.SetCellMergedInfo (行 140-153)
        /// </summary>
        public void SetCellMergedInfo(ExcelCellInfo cellInfo, int fromRow, int fromCol, int toRow, int toCol)
        {
            int rowSpan = toRow - fromRow + 1;
            int colSpan = toCol - fromCol + 1;

            cellInfo.Dimensions.IsMerged = true;
            cellInfo.Dimensions.IsMainMergedCell = true;
            cellInfo.Dimensions.RowSpan = rowSpan;
            cellInfo.Dimensions.ColSpan = colSpan;
            cellInfo.Dimensions.MergedRangeAddress =
                $"{GetColumnName(fromCol)}{fromRow}:{GetColumnName(toCol)}{toRow}";
        }

        public string MergeFloatingObjectText(string? cellText, List<FloatingObjectInfo>? floatingObjects)
        {
            if (floatingObjects == null || !floatingObjects.Any())
                return cellText ?? string.Empty;

            var result = cellText ?? string.Empty;
            foreach (var obj in floatingObjects)
            {
                if (!string.IsNullOrEmpty(obj.Text))
                {
                    if (!string.IsNullOrEmpty(result))
                    {
                        result += "\n" + obj.Text;
                    }
                    else
                    {
                        result = obj.Text;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 合併浮動物件的文字內容到儲存格文字中 (方法重載 - 與 Controller 一致)
        /// 原始: ExcelController.MergeFloatingObjectText (行 153-168)
        /// </summary>
        public void MergeFloatingObjectText(ExcelCellInfo cellInfo, string? floatingObjectText, string cellAddress)
        {
            if (string.IsNullOrEmpty(floatingObjectText))
                return;

            if (!string.IsNullOrEmpty(cellInfo.Text))
            {
                // 如果原本有文字,則換行加入
                cellInfo.Text += "\n" + floatingObjectText;
            }
            else
            {
                // 如果原本沒有文字,直接設定
                cellInfo.Text = floatingObjectText;
            }

            //_logger.LogInformation($"✅ 已將浮動物件文字合併到儲存格 {cellAddress}: '{floatingObjectText}'");
        }

        public OfficeOpenXml.Drawing.ExcelPicture? FindPictureInDrawings(ExcelWorksheet worksheet, int row, int column)
        {
            if (worksheet.Drawings == null)
                return null;

            return worksheet.Drawings
                .OfType<OfficeOpenXml.Drawing.ExcelPicture>()
                .FirstOrDefault(p => p.From.Row == row - 1 && p.From.Column == column - 1);
        }

        /// <summary>
        /// 按圖片名稱在工作表的繪圖集合中查找圖片 (方法重載 - 與 Controller 一致)
        /// 原始: ExcelController.FindPictureInDrawings (行 178-187)
        /// </summary>
        public OfficeOpenXml.Drawing.ExcelPicture? FindPictureInDrawings(ExcelWorksheet worksheet, string imageName)
        {
            if (worksheet.Drawings == null || string.IsNullOrEmpty(imageName))
                return null;

            return worksheet.Drawings
                .FirstOrDefault(d => d is OfficeOpenXml.Drawing.ExcelPicture p && p.Name == imageName)
                as OfficeOpenXml.Drawing.ExcelPicture;
        }

        /// <summary>
        /// 處理圖片跨儲存格邏輯 (檢查圖片是否跨越多個儲存格並自動設定合併)
        /// ⭐ 修復: 考慮已存在的合併儲存格範圍
        /// 完整從 ExcelController.ProcessImageCrossCells (行 193-258) 搬移,確保邏輯 100% 一致
        /// </summary>
        public void ProcessImageCrossCells(ExcelCellInfo cellInfo, ExcelRange cell, ExcelWorksheet worksheet)
        {
            if (cellInfo.Images == null || !cellInfo.Images.Any())
                return;
            
            if (cell.Address.Contains("H2"))
                Console.WriteLine("");
            
            foreach (var image in cellInfo.Images)
            {
                var fromRow = image.AnchorCell?.Row ?? cell.Start.Row;
                var fromCol = image.AnchorCell?.Column ?? cell.Start.Column;

                var picture = FindPictureInDrawings(worksheet, image.Name);

                if (picture != null)
                {
                    int toRow = picture.To?.Row + 1 ?? fromRow;
                    int toCol = picture.To?.Column + 1 ?? fromCol;

                    // ⭐ 關鍵修復: 檢查儲存格是否已經合併
                    if (cellInfo.Dimensions?.IsMerged == true && !string.IsNullOrEmpty(cellInfo.Dimensions.MergedRangeAddress))
                    {
                        // 如果儲存格已經合併，檢查圖片是否完全在合併範圍內
                        var mergedRange = cellInfo.Dimensions.MergedRangeAddress;
                        //_logger.LogInformation($"⚠️  儲存格 {cell.Address} 已合併 ({mergedRange})，圖片 '{image.Name}' 範圍: {GetColumnName(fromCol)}{fromRow}:{GetColumnName(toCol)}{toRow}");

                        // 解析合併範圍
                        var rangeParts = mergedRange.Split(':');
                        if (rangeParts.Length == 2)
                        {
                            // 提取合併範圍的行列信息
                            var mergedFromRow = cell.Start.Row;
                            var mergedFromCol = cell.Start.Column;
                            var mergedToRow = cell.End.Row;
                            var mergedToCol = cell.End.Column;

                            // 檢查圖片是否超出合併範圍
                            bool imageExceedsMerged = (toRow > mergedToRow || toCol > mergedToCol ||
                                                      fromRow < mergedFromRow || fromCol < mergedFromCol);

                            if (imageExceedsMerged)
                            {
                                _logger.LogWarning($"⚠️  圖片 '{image.Name}' 範圍 ({GetColumnName(fromCol)}{fromRow}:{GetColumnName(toCol)}{toRow}) " +
                                                 $"超出或不完全符合已存在的合併範圍 ({mergedRange})，跳過自動合併");
                            }
                            else
                            {
                                //_logger.LogInformation($"✅ 圖片 '{image.Name}' 完全在已存在的合併範圍內");
                            }
                        }
                    }
                    else if (toRow > fromRow || toCol > fromCol)
                    {
                        // 原始邏輯：儲存格未合併時，根據圖片範圍自動設定合併
                        int rowSpan = toRow - fromRow + 1;
                        int colSpan = toCol - fromCol + 1;

                        //_logger.LogInformation($"圖片 '{image.Name}' 跨越 {rowSpan} 行 x {colSpan} 欄，自動設定合併儲存格");

                        SetCellMergedInfo(cellInfo, fromRow, fromCol, toRow, toCol);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// 處理浮動物件跨儲存格邏輯 (包含文字合併)
        /// ⭐ 修復: 考慮已存在的合併儲存格範圍
        /// 完整從 ExcelController.ProcessFloatingObjectCrossCells (行 260-335) 搬移,確保邏輯 100% 一致
        /// </summary>
        public void ProcessFloatingObjectCrossCells(ExcelCellInfo cellInfo, ExcelRange cell)
        {
            if (cellInfo.FloatingObjects == null || !cellInfo.FloatingObjects.Any())
                return;

            foreach (var floatingObj in cellInfo.FloatingObjects)
            {
                var fromRow = floatingObj.FromCell?.Row ?? cell.Start.Row;
                var fromCol = floatingObj.FromCell?.Column ?? cell.Start.Column;
                var toRow = floatingObj.ToCell?.Row ?? fromRow;
                var toCol = floatingObj.ToCell?.Column ?? fromCol;

                // ⭐ 關鍵修復: 檢查儲存格是否已經合併
                if (cellInfo.Dimensions?.IsMerged == true && !string.IsNullOrEmpty(cellInfo.Dimensions.MergedRangeAddress))
                {
                    // 如果儲存格已經合併，檢查浮動物件是否完全在合併範圍內
                    var mergedRange = cellInfo.Dimensions.MergedRangeAddress;
                    //_logger.LogInformation($"⚠️  儲存格 {cell.Address} 已合併 ({mergedRange})，浮動物件 '{floatingObj.Name}' 範圍: {GetColumnName(fromCol)}{fromRow}:{GetColumnName(toCol)}{toRow}");

                    // 解析合併範圍
                    var rangeParts = mergedRange.Split(':');
                    if (rangeParts.Length == 2)
                    {
                        // 簡單解析 (假設格式如 "E2:G9")
                        var startCell = rangeParts[0];
                        var endCell = rangeParts[1];

                        // 提取行列信息 (簡化版本)
                        var mergedFromRow = cell.Start.Row;
                        var mergedFromCol = cell.Start.Column;
                        var mergedToRow = cell.End.Row;
                        var mergedToCol = cell.End.Column;

                        // 檢查浮動物件是否超出合併範圍
                        bool floatingExceedsMerged = (toRow > mergedToRow || toCol > mergedToCol ||
                                                     fromRow < mergedFromRow || fromCol < mergedFromCol);

                        if (floatingExceedsMerged)
                        {
                            _logger.LogWarning($"⚠️  浮動物件 '{floatingObj.Name}' 範圍 ({GetColumnName(fromCol)}{fromRow}:{GetColumnName(toCol)}{toRow}) " +
                                             $"超出或不完全符合已存在的合併範圍 ({mergedRange})，跳過自動合併");
                        }
                        else
                        {
                            //_logger.LogInformation($"✅ 浮動物件 '{floatingObj.Name}' 完全在已存在的合併範圍內，合併文字內容");
                        }
                    }

                    // 無論如何都要合併文字內容
                    MergeFloatingObjectText(cellInfo, floatingObj.Text, cell.Address);
                }
                else if (toRow > fromRow || toCol > fromCol)
                {
                    // 原始邏輯：儲存格未合併時，根據浮動物件範圍自動設定合併
                    int rowSpan = toRow - fromRow + 1;
                    int colSpan = toCol - fromCol + 1;

                    //_logger.LogInformation($"浮動物件 '{floatingObj.Name}' (類型: {floatingObj.ObjectType}) 跨越 {rowSpan} 行 x {colSpan} 欄，自動設定合併儲存格");

                    SetCellMergedInfo(cellInfo, fromRow, fromCol, toRow, toCol);
                    MergeFloatingObjectText(cellInfo, floatingObj.Text, cell.Address);

                    break; // 只需要設定一次
                }
                else
                {
                    // 單一儲存格的浮動物件，只合併文字內容
                    MergeFloatingObjectText(cellInfo, floatingObj.Text, cell.Address);
                }
            }
        }

        #endregion
    }
}