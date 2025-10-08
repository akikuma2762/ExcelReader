using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using ExcelReaderAPI.Models;
using ExcelReaderAPI.Utils;
using System.Data;
using System.IO.Packaging;
using System.IO.Compression;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;
using SkiaSharp;

namespace ExcelReaderAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;

        // 安全機制：防止無窮迴圈的常數
        private const int MAX_SEARCH_OPERATIONS = 1000;
        private const int MAX_DRAWING_OBJECTS_TO_CHECK = 999999; // 增加限制，支援更多文字方塊
        private const int MAX_CELLS_TO_SEARCH = 5000;

        // 功能開關
        private const bool ENABLE_FLOATING_OBJECTS_CHECK = true; // ✅ 啟用浮動物件檢查 (用於檢測文字方塊)
        private const bool ENABLE_CELL_IMAGES_CHECK = true; // 保持圖片檢查啟用

        // 日誌開關 - 用於效能優化
        private const bool ENABLE_VERBOSE_LOGGING = false; // 詳細日誌 (包含每個儲存格的處理日誌)
        private const bool ENABLE_DEBUG_LOGGING = false; // 調試日誌 (包含圖片檢查、內容類型檢測等)
        private const bool ENABLE_PERFORMANCE_LOGGING = true; // 效能日誌 (關鍵節點的耗時統計)

        // 請求層級的計數器 - 使用 ThreadStatic 避免併發問題
        [ThreadStatic]
        private static Dictionary<string, int>? _worksheetDrawingObjectCounts;
        [ThreadStatic]
        private static DateTime _requestStartTime = DateTime.MinValue;

        /// <summary>
        /// 取得或初始化工作表繪圖物件計數器
        /// </summary>
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
        /// 重置工作表繪圖物件計數器
        /// </summary>
        private void ResetWorksheetDrawingObjectCounts()
        {
            _worksheetDrawingObjectCounts?.Clear();
        }

        /// <summary>
        /// 工作表圖片位置索引 - 用於效能優化
        /// 一次性建立索引,避免每個儲存格都遍歷所有 Drawings
        /// 複雜度: 建立 O(D), 查詢 O(1), D = Drawings 數量
        /// </summary>
        private class WorksheetImageIndex
        {
            // Key: "Row_Column" (例: "5_3" 代表 Row=5, Col=3)
            // Value: 該儲存格起始位置的所有圖片
            private readonly Dictionary<string, List<OfficeOpenXml.Drawing.ExcelPicture>> _cellImageMap;

            public WorksheetImageIndex(ExcelWorksheet worksheet)
            {
                _cellImageMap = new Dictionary<string, List<OfficeOpenXml.Drawing.ExcelPicture>>();

                if (worksheet.Drawings == null || !worksheet.Drawings.Any())
                    return;

                // 一次性遍歷所有繪圖物件建立索引
                foreach (var drawing in worksheet.Drawings)
                {
                    if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture && picture.From != null)
                    {
                        int fromRow = picture.From.Row + 1; // EPPlus 使用 0-based, 轉為 1-based
                        int fromCol = picture.From.Column + 1;
                        string key = $"{fromRow}_{fromCol}";

                        if (!_cellImageMap.ContainsKey(key))
                            _cellImageMap[key] = new List<OfficeOpenXml.Drawing.ExcelPicture>();

                        _cellImageMap[key].Add(picture);
                    }
                }
            }

            /// <summary>
            /// 快速查詢指定儲存格的圖片 - O(1) 複雜度
            /// </summary>
            public List<OfficeOpenXml.Drawing.ExcelPicture>? GetImagesAtCell(int row, int col)
            {
                string key = $"{row}_{col}";
                return _cellImageMap.TryGetValue(key, out var images) && images.Any() ? images : null;
            }

            /// <summary>
            /// 檢查指定儲存格是否有圖片 - O(1) 複雜度
            /// </summary>
            public bool HasImagesAtCell(int row, int col)
            {
                string key = $"{row}_{col}";
                return _cellImageMap.ContainsKey(key) && _cellImageMap[key].Any();
            }

            /// <summary>
            /// 取得總圖片數量
            /// </summary>
            public int TotalImageCount => _cellImageMap.Values.Sum(list => list.Count);
        }

        /// <summary>
        /// 樣式快取 - 避免重複創建相同的樣式物件 (執行緒安全)
        /// Phase 3.2: 使用 ConcurrentDictionary 支援並行處理
        /// 複雜度: O(1) 查詢, 大幅減少 GC 壓力
        /// </summary>
        private class StyleCache
        {
            private readonly System.Collections.Concurrent.ConcurrentDictionary<string, FontInfo> _fontCache = new();
            private readonly System.Collections.Concurrent.ConcurrentDictionary<string, BorderInfo> _borderCache = new();
            private readonly System.Collections.Concurrent.ConcurrentDictionary<string, FillInfo> _fillCache = new();

            public string GetFontCacheKey(ExcelRange cell)
            {
                return GetFontKey(cell.Style.Font, cell.Style.Fill, cell.Style.Font.Color);
            }

            public void CacheFont(string key, FontInfo fontInfo)
            {
                _fontCache[key] = fontInfo;
            }

            public FontInfo? GetCachedFont(string key)
            {
                _fontCache.TryGetValue(key, out var fontInfo);
                return fontInfo;
            }

            public FillInfo GetOrCreateFill(ExcelRange cell)
            {
                var key = GetFillKey(cell.Style.Fill);
                if (!_fillCache.TryGetValue(key, out var fillInfo))
                {
                    fillInfo = new FillInfo
                    {
                        PatternType = cell.Style.Fill.PatternType.ToString(),
                        BackgroundColor = GetBackgroundColor(cell),
                        PatternColor = GetColorFromExcelColor(cell.Style.Fill.PatternColor)
                    };
                    _fillCache[key] = fillInfo;
                }
                return fillInfo;
            }

            private string GetFontKey(OfficeOpenXml.Style.ExcelFont font, OfficeOpenXml.Style.ExcelFill fill, OfficeOpenXml.Style.ExcelColor color)
            {
                return $"{font.Name}|{font.Size}|{font.Bold}|{font.Italic}|{font.UnderLine}|{font.Strike}|{color.Rgb ?? color.Theme.ToString()}";
            }

            private string GetFillKey(OfficeOpenXml.Style.ExcelFill fill)
            {
                return $"{fill.PatternType}|{fill.BackgroundColor.Rgb}|{fill.BackgroundColor.Theme}|{fill.PatternColor.Rgb}";
            }

            // 這些方法需要訪問 ExcelController 的方法,稍後會調整
            private string? GetColorFromExcelColor(OfficeOpenXml.Style.ExcelColor excelColor)
            {
                // 佔位符,稍後實作
                return null;
            }

            private string? GetBackgroundColor(ExcelRange cell)
            {
                // 佔位符,稍後實作
                return null;
            }
        }

        /// <summary>
        /// 顏色轉換快取 - 避免重複轉換相同顏色 (執行緒安全)
        /// Phase 3.2: 使用 ConcurrentDictionary 支援並行處理
        /// </summary>
        private class ColorCache
        {
            private readonly System.Collections.Concurrent.ConcurrentDictionary<string, string?> _cache = new();

            public string GetCacheKey(OfficeOpenXml.Style.ExcelColor color)
            {
                if (color == null) return "null";
                return $"{color.Rgb}|{color.Theme}|{color.Tint}|{color.Indexed}";
            }

            public void CacheColor(string key, string? color)
            {
                _cache[key] = color;
            }

            public bool TryGetCachedColor(string key, out string? color)
            {
                return _cache.TryGetValue(key, out color);
            }
        }

        /// <summary>
        /// 合併儲存格索引 - 快速查詢儲存格是否在合併範圍內
        /// 複雜度: 建立 O(M×C), 查詢 O(1), M=合併範圍數, C=每個範圍的儲存格數
        /// </summary>
        private class MergedCellIndex
        {
            // Key: "Row_Column", Value: 合併範圍地址 (如 "A1:B2")
            private readonly Dictionary<string, string> _cellToMergeMap = new();

            public MergedCellIndex(ExcelWorksheet worksheet)
            {
                if (worksheet.MergedCells == null || !worksheet.MergedCells.Any())
                    return;

                foreach (var mergeRange in worksheet.MergedCells)
                {
                    var range = worksheet.Cells[mergeRange];

                    for (int row = range.Start.Row; row <= range.End.Row; row++)
                    {
                        for (int col = range.Start.Column; col <= range.End.Column; col++)
                        {
                            var key = $"{row}_{col}";
                            _cellToMergeMap[key] = mergeRange;
                        }
                    }
                }
            }

            /// <summary>
            /// 取得指定儲存格所屬的合併範圍 - O(1) 複雜度
            /// </summary>
            public string? GetMergeRange(int row, int col)
            {
                _cellToMergeMap.TryGetValue($"{row}_{col}", out var range);
                return range;
            }

            /// <summary>
            /// 檢查指定儲存格是否在合併範圍內 - O(1) 複雜度
            /// </summary>
            public bool IsMergedCell(int row, int col)
            {
                return _cellToMergeMap.ContainsKey($"{row}_{col}");
            }

            /// <summary>
            /// 取得總合併範圍數量
            /// </summary>
            public int MergeCount => _cellToMergeMap.Values.Distinct().Count();
        }

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }

        static ExcelController()
        {
            // 設定EPPlus授權（非商業用途）- EPPlus 8.x 新 API
            ExcelPackage.License.SetNonCommercialPersonal("dek");//个人
        }

        #region 日誌輔助方法 - 統一日誌開關控制

        /// <summary>
        /// 條件式詳細日誌 (每個儲存格級別的日誌)
        /// 預設關閉以提升效能,開發時可啟用 ENABLE_VERBOSE_LOGGING
        /// </summary>
        private void LogVerbose(string message)
        {
            if (ENABLE_VERBOSE_LOGGING)
            {
                _logger.LogInformation(message);
            }
        }

        /// <summary>
        /// 條件式調試日誌 (函數調用、狀態檢查等)
        /// 預設關閉以提升效能,調試時可啟用 ENABLE_DEBUG_LOGGING
        /// </summary>
        private void LogDebugConditional(string message)
        {
            if (ENABLE_DEBUG_LOGGING)
            {
                _logger.LogDebug(message);
            }
        }

        /// <summary>
        /// 效能關鍵點日誌 (索引建立、批次處理完成等)
        /// 預設啟用,用於監控系統效能
        /// </summary>
        private void LogPerformance(string message)
        {
            if (ENABLE_PERFORMANCE_LOGGING)
            {
                _logger.LogInformation(message);
            }
        }

        #endregion

        #region DRY 原則 - 共用邏輯提取方法

        /// <summary>
        /// 設定儲存格的合併資訊
        /// </summary>
        private void SetCellMergedInfo(ExcelCellInfo cellInfo, int fromRow, int fromCol, int toRow, int toCol)
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

        /// <summary>
        /// 合併浮動物件的文字內容到儲存格文字中
        /// </summary>
        private void MergeFloatingObjectText(ExcelCellInfo cellInfo, string? floatingObjectText, string cellAddress)
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

            _logger.LogInformation($"✅ 已將浮動物件文字合併到儲存格 {cellAddress}: '{floatingObjectText}'");
        }

        /// <summary>
        /// 在工作表的繪圖集合中查找指定名稱的圖片
        /// </summary>
        private OfficeOpenXml.Drawing.ExcelPicture? FindPictureInDrawings(ExcelWorksheet worksheet, string imageName)
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
        /// </summary>
        private void ProcessImageCrossCells(ExcelCellInfo cellInfo, ExcelRange cell, ExcelWorksheet worksheet)
        {
            if (cellInfo.Images == null || !cellInfo.Images.Any())
                return;
            if(cell.Address.Contains("H2"))
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
                        _logger.LogInformation($"⚠️  儲存格 {cell.Address} 已合併 ({mergedRange})，圖片 '{image.Name}' 範圍: {GetColumnName(fromCol)}{fromRow}:{GetColumnName(toCol)}{toRow}");

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
                                _logger.LogInformation($"✅ 圖片 '{image.Name}' 完全在已存在的合併範圍內");
                            }
                        }
                    }
                    else if (toRow > fromRow || toCol > fromCol)
                    {
                        // 原始邏輯：儲存格未合併時，根據圖片範圍自動設定合併
                        int rowSpan = toRow - fromRow + 1;
                        int colSpan = toCol - fromCol + 1;

                        _logger.LogInformation($"圖片 '{image.Name}' 跨越 {rowSpan} 行 x {colSpan} 欄，自動設定合併儲存格");

                        SetCellMergedInfo(cellInfo, fromRow, fromCol, toRow, toCol);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// 處理浮動物件跨儲存格邏輯 (包含文字合併)
        /// ⭐ 修復: 考慮已存在的合併儲存格範圍
        /// </summary>
        private void ProcessFloatingObjectCrossCells(ExcelCellInfo cellInfo, ExcelRange cell)
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
                    _logger.LogInformation($"⚠️  儲存格 {cell.Address} 已合併 ({mergedRange})，浮動物件 '{floatingObj.Name}' 範圍: {GetColumnName(fromCol)}{fromRow}:{GetColumnName(toCol)}{toRow}");

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
                            _logger.LogInformation($"✅ 浮動物件 '{floatingObj.Name}' 完全在已存在的合併範圍內，合併文字內容");
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

                    _logger.LogInformation($"浮動物件 '{floatingObj.Name}' (類型: {floatingObj.ObjectType}) 跨越 {rowSpan} 行 x {colSpan} 欄，自動設定合併儲存格");

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

        /// <summary>
        /// 創建預設字體資訊（避免顏色解析錯誤）
        /// </summary>
        private FontInfo CreateDefaultFontInfo()
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
        /// </summary>
        private AlignmentInfo CreateDefaultAlignmentInfo()
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
        /// </summary>
        private BorderInfo CreateDefaultBorderInfo()
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
        /// </summary>
        private FillInfo CreateDefaultFillInfo()
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
        /// 智能檢測儲存格的主要內容類型
        /// </summary>
        private enum CellContentType
        {
            Empty,          // 空儲存格
            TextOnly,       // 純文字內容
            ImageOnly,      // 純圖片內容
            Mixed           // 混合內容
        }

        /// <summary>
        /// 檢測儲存格的主要內容類型 (使用索引優化版 + EPPlus 8.x In-Cell Picture API)
        /// </summary>
        private CellContentType DetectCellContentType(ExcelRange cell, WorksheetImageIndex? imageIndex)
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
        /// 檢測儲存格的主要內容類型 (舊版本 - 相容性保留 + EPPlus 8.x In-Cell Picture API)
        /// </summary>
        private CellContentType DetectCellContentType(ExcelRange cell, ExcelWorksheet worksheet)
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
                        if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
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

        /// <summary>
        /// 創建儲存格資訊 (使用索引優化版 + 快取)
        /// </summary>
        private ExcelCellInfo CreateCellInfo(
            ExcelRange cell,
            ExcelWorksheet worksheet,
            WorksheetImageIndex imageIndex,
            ColorCache? colorCache = null,
            MergedCellIndex? mergedCellIndex = null)
        {
            if (cell == null || worksheet == null)
                throw new ArgumentNullException("Cell or worksheet cannot be null");

            var cellInfo = new ExcelCellInfo();

            try
            {
                // 🔍 調試用中斷點區域 - A30:H39 範圍 (In-Cell 圖片測試)
                var row = cell.Start.Row;
                var col = cell.Start.Column;
                if (row >= 30 && row <= 39 && col >= 1 && col <= 8) // A30:H39
                {
                    // ⚠️ 在此設定中斷點以調試 In-Cell 圖片
                    var debugAddress = cell.Address;
                    var debugValue = cell.Value;
                    var debugValueType = cell.Value?.GetType().FullName;
                    _logger.LogWarning($"🔍 調試點: {debugAddress}, Value Type: {debugValueType}, Value: {debugValue}");
                    // 👈 在這一行設定中斷點 (F9)
                }

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
                cellInfo.Value = GetSafeValue(cell.Value);
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
                        Color = GetColorFromExcelColor(cell.Style.Font.Color, colorCache),
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
                                Color = cell.Style.Border?.Top?.Color != null ? GetColorFromExcelColor(cell.Style.Border.Top.Color, colorCache) : null
                            },
                            Bottom = new BorderStyle
                            {
                                Style = cell.Style.Border?.Bottom?.Style.ToString() ?? "None",
                                Color = cell.Style.Border?.Bottom?.Color != null ? GetColorFromExcelColor(cell.Style.Border.Bottom.Color, colorCache) : null
                            },
                            Left = new BorderStyle
                            {
                                Style = cell.Style.Border?.Left?.Style.ToString() ?? "None",
                                Color = cell.Style.Border?.Left?.Color != null ? GetColorFromExcelColor(cell.Style.Border.Left.Color, colorCache) : null
                            },
                            Right = new BorderStyle
                            {
                                Style = cell.Style.Border?.Right?.Style.ToString() ?? "None",
                                Color = cell.Style.Border?.Right?.Color != null ? GetColorFromExcelColor(cell.Style.Border.Right.Color, colorCache) : null
                            },
                            Diagonal = new BorderStyle
                            {
                                Style = cell.Style.Border?.Diagonal?.Style.ToString() ?? "None",
                                Color = cell.Style.Border?.Diagonal?.Color != null ? GetColorFromExcelColor(cell.Style.Border.Diagonal.Color, colorCache) : null
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
                        BackgroundColor = GetBackgroundColor(cell),
                        PatternColor = GetColorFromExcelColor(cell.Style.Fill.PatternColor, colorCache),
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
                        mergedRange = FindMergedRange(worksheet, cell.Start.Row, cell.Start.Column);
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
                            cellInfo.Border = GetMergedCellBorder(worksheet, mergedRange, cell);
                        }
                        else
                        {
                            cellInfo.Dimensions.RowSpan = 1;
                            cellInfo.Dimensions.ColSpan = 1;
                        }
                    }
                }

                // Rich Text 處理 (與原版相同,省略)
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
                        mergedRange = FindMergedRange(worksheet, cell.Start.Row, cell.Start.Column);
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
                    cellInfo.Images = ENABLE_CELL_IMAGES_CHECK ? GetCellImages(rangeToCheck, imageIndex, worksheet) : null;
                }

                // 圖片跨儲存格處理 - 使用 DRY 共用方法
                ProcessImageCrossCells(cellInfo, cell, worksheet);

                // 浮動物件 - ⭐ 修復: 只在合併儲存格的主要儲存格中查找浮動物件
                if (cell.Merge && cellInfo.Dimensions?.IsMainMergedCell != true)
                {
                    // 如果是合併儲存格但不是主要儲存格，則不查找浮動物件
                    cellInfo.FloatingObjects = null;
                    _logger.LogDebug($"儲存格 {cell.Address} 是合併儲存格的次要儲存格，跳過浮動物件檢查");
                }
                else
                {
                    cellInfo.FloatingObjects = ENABLE_FLOATING_OBJECTS_CHECK ? GetCellFloatingObjects(worksheet, rangeToCheck) : null;
                }

                // 浮動物件跨儲存格處理 - 使用 DRY 共用方法
                ProcessFloatingObjectCrossCells(cellInfo, cell);

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
        /// 創建儲存格資訊 (舊版本 - 相容性保留,內部使用索引優化版本)
        /// </summary>
        private ExcelCellInfo CreateCellInfo(ExcelRange cell, ExcelWorksheet worksheet)
        {
            // 為了保持向後相容,創建臨時索引並調用優化版本
            var imageIndex = new WorksheetImageIndex(worksheet);
            return CreateCellInfo(cell, worksheet, imageIndex, null, null);
        }


        /// <summary>
        /// 獲取指定儲存格範圍內的所有圖片 (使用索引優化版 + EPPlus 8.x In-Cell Picture API)
        /// </summary>
        private List<ImageInfo>? GetCellImages(ExcelRange cell, WorksheetImageIndex imageIndex, ExcelWorksheet worksheet)
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
                            _logger.LogInformation($"✅ 儲存格 {cell.Address} 包含 In-Cell 圖片 (EPPlus 8.x API)");

                            var cellPicture = cell.Picture.Get();
                            if (cellPicture != null)
                            {
                                var imageBytes = cellPicture.GetImageBytes();
                                var imageType = GetImageTypeFromFileName(cellPicture.FileName);

                                // 🔍 計算儲存格/合併範圍的像素尺寸 (In-Cell 圖片會填滿整個儲存格)
                                // 1. 計算單一儲存格的基準高度
                                var (cellWidthPixels, singleCellHeightPixels) = GetCellPixelDimensions(worksheet, cell.Start.Row, cell.Start.Column);

                                // 2. 計算合併範圍的總高度
                                int rowSpan = cell.End.Row - cell.Start.Row + 1; // 合併的行數
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
                                    OriginalWidth = (int)Math.Round(cellWidthPixels),
                                    OriginalHeight = (int)Math.Round(totalHeightPixels),
                                    ExcelWidthCm = 0,
                                    ExcelHeightCm = 0,
                                    ScaleFactor = 1.0,
                                    IsScaled = false,
                                    ScaleMethod = $"In-Cell 圖片 (自動填滿 {rowSpan} 行合併儲存格)"
                                };

                                images.Add(imageInfo);
                                _logger.LogInformation($"成功讀取 In-Cell 圖片: {imageInfo.Name}, 大小: {imageInfo.FileSize} bytes, 尺寸: {cellWidthPixels:F0}×{totalHeightPixels:F0}px");
                                return images.Any() ? images : null;
                            }
                        }
                    }
                    catch (Exception inCellEx)
                    {
                        _logger.LogWarning($"讀取 In-Cell 圖片失敗 (儲存格 {cell.Address}): {inCellEx.Message}");
                    }



                // 使用索引快速查詢浮動圖片 (Drawing Pictures) - O(1) 複雜度
                var pictures = imageIndex.GetImagesAtCell(cell.Start.Row, cell.Start.Column);

                if (pictures == null)
                {
                    _logger.LogDebug($"儲存格 {cell.Address} 沒有圖片");
                    return null;
                }

                _logger.LogInformation($"儲存格 {cell.Address} 找到 {pictures.Count} 張圖片 (來自索引)");

                // 處理找到的圖片
                foreach (var picture in pictures)
                {
                    try
                    {
                        // 安全獲取圖片位置
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

                        _logger.LogInformation($"處理圖片: '{picture.Name ?? "未命名"}' 位置: Row {fromRow}-{toRow}, Col {fromCol}-{toCol}");

                        // 獲取圖片原始尺寸
                        var (actualWidth, actualHeight) = GetActualImageDimensions(picture);

                        // 計算 Excel 顯示尺寸
                        int excelDisplayWidth = actualWidth;
                        int excelDisplayHeight = actualHeight;
                        double excelWidthCm = 0;
                        double excelHeightCm = 0;
                        double scalePercentage = 100.0;

                        try
                        {
                            // 從 From/To 計算 Excel 顯示尺寸
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

                                _logger.LogDebug($"📐 Excel 顯示尺寸 - 像素: {excelDisplayWidth}×{excelDisplayHeight}px, 厘米: {excelWidthCm:F2}×{excelHeightCm:F2}cm, 縮放: {scalePercentage:F1}%");
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
                        _logger.LogInformation($"成功解析圖片: {imageInfo.Name}, 大小: {imageInfo.FileSize} bytes");
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

        /// <summary>
        /// 獲取指定儲存格範圍內的所有圖片 (舊版本 - 相容性保留)
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

                // 初始化計數器（只在第一次請求時）
                if (_requestStartTime == DateTime.MinValue)
                {
                    _requestStartTime = DateTime.Now;
                    ResetWorksheetDrawingObjectCounts();
                }

                // 安全檢查：如果已經檢查太多物件，直接跳過這個儲存格
                // if (_globalDrawingObjectCount > MAX_DRAWING_OBJECTS_TO_CHECK)
                // {
                //     _logger.LogDebug($"儲存格 {cell.Address} 跳過圖片檢查 - 已達到檢查限制");
                //     return null;
                // }

                // 1. 檢查所有工作表中的圖片 (採用寬鬆匹配策略)
                if (worksheet.Drawings != null && worksheet.Drawings.Any())
                {
                    _logger.LogDebug($"工作表 '{worksheet.Name}' 包含 {worksheet.Drawings.Count} 個繪圖物件 (已檢查: {GetWorksheetDrawingObjectCount(worksheet.Name)})");

                    foreach (var drawing in worksheet.Drawings)
                    {
                        // 安全檢查：防止處理過多物件
                        // if (++_globalDrawingObjectCount > MAX_DRAWING_OBJECTS_TO_CHECK)
                        // {
                        //     _logger.LogWarning($"已檢查 {MAX_DRAWING_OBJECTS_TO_CHECK} 個繪圖物件，停止進一步檢查以避免效能問題");
                        //     return images.Any() ? images : null;
                        // }

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

                                // 只在圖片的起始儲存格（From位置）添加圖片
                                // 避免同一張圖片被重複添加到多個儲存格，造成資料量過大
                                bool shouldInclude = (fromRow >= cellStartRow && fromRow <= cellEndRow &&
                                                     fromCol >= cellStartCol && fromCol <= cellEndCol);

                                // 記錄詳細的檢查結果
                                _logger.LogDebug($"圖片 '{picture.Name ?? "未命名"}' 位置檢查: " +
                                               $"From({fromRow},{fromCol}) 是否在儲存格 [{cellStartRow},{cellEndRow}] x [{cellStartCol},{cellEndCol}] 內? " +
                                               $"結果: {shouldInclude}");

                                if (shouldInclude)
                                {
                                    try
                                    {
                                        // 獲取圖片原始尺寸
                                        var (actualWidth, actualHeight) = GetActualImageDimensions(picture);

                                        // 使用 ExcelDrawingSize 獲取 Excel 中的顯示尺寸
                                        int excelDisplayWidth = actualWidth;
                                        int excelDisplayHeight = actualHeight;
                                        double excelWidthCm = 0;
                                        double excelHeightCm = 0;
                                        double scalePercentage = 100.0;

                                        try
                                        {
                                            // 從 From/To 計算 Excel 顯示尺寸
                                            if (picture.From != null && picture.To != null)
                                            {
                                                const double emuPerPixel = 9525.0; // 914400 EMU / 96 DPI
                                                const double emuPerInch = 914400.0;
                                                const double emuPerCm = emuPerInch / 2.54;

                                                //  正確計算: 需要加上中間儲存格的尺寸
                                                long totalWidthEmu = 0;
                                                long totalHeightEmu = 0;

                                                // 計算總寬度
                                                for (int col = picture.From.Column; col <= picture.To.Column; col++)
                                                {
                                                    var column = worksheet.Column(col + 1); // EPPlus column index is 1-based
                                                    var colWidth = column.Width > 0 ? column.Width : worksheet.DefaultColWidth;
                                                    // Excel 欄寬單位轉 EMU: 欄寬 * 字符寬度(7px) * 9525 EMU/px
                                                    long colWidthEmu = (long)(colWidth * 7.0 * emuPerPixel);

                                                    if (col == picture.From.Column && col == picture.To.Column)
                                                    {
                                                        // 同一欄: To.ColumnOff - From.ColumnOff
                                                        totalWidthEmu = picture.To.ColumnOff - picture.From.ColumnOff;
                                                    }
                                                    else if (col == picture.From.Column)
                                                    {
                                                        // 起始欄: 儲存格總寬 - From.ColumnOff
                                                        totalWidthEmu += colWidthEmu - picture.From.ColumnOff;
                                                    }
                                                    else if (col == picture.To.Column)
                                                    {
                                                        // 結束欄: To.ColumnOff
                                                        totalWidthEmu += picture.To.ColumnOff;
                                                    }
                                                    else
                                                    {
                                                        // 中間欄: 完整寬度
                                                        totalWidthEmu += colWidthEmu;
                                                    }
                                                }

                                                // 計算總高度
                                                for (int row = picture.From.Row; row <= picture.To.Row; row++)
                                                {
                                                    var rowObj = worksheet.Row(row + 1); // EPPlus row index is 1-based
                                                    var rowHeight = rowObj.Height > 0 ? rowObj.Height : worksheet.DefaultRowHeight;
                                                    // 行高單位是點數(points): 1 point = 12700 EMU
                                                    long rowHeightEmu = (long)(rowHeight * 12700);

                                                    if (row == picture.From.Row && row == picture.To.Row)
                                                    {
                                                        // 同一行: To.RowOff - From.RowOff
                                                        totalHeightEmu = picture.To.RowOff - picture.From.RowOff;
                                                    }
                                                    else if (row == picture.From.Row)
                                                    {
                                                        // 起始行: 儲存格總高 - From.RowOff
                                                        totalHeightEmu += rowHeightEmu - picture.From.RowOff;
                                                    }
                                                    else if (row == picture.To.Row)
                                                    {
                                                        // 結束行: To.RowOff
                                                        totalHeightEmu += picture.To.RowOff;
                                                    }
                                                    else
                                                    {
                                                        // 中間行: 完整高度
                                                        totalHeightEmu += rowHeightEmu;
                                                    }
                                                }

                                                // 轉換為像素和公分
                                                excelDisplayWidth = (int)(totalWidthEmu / emuPerPixel);
                                                excelDisplayHeight = (int)(totalHeightEmu / emuPerPixel);
                                                excelWidthCm = totalWidthEmu / emuPerCm;
                                                excelHeightCm = totalHeightEmu / emuPerCm;

                                                // 計算縮放比例
                                                if (actualWidth > 0 && actualHeight > 0)
                                                {
                                                    double scaleX = (double)excelDisplayWidth / actualWidth * 100.0;
                                                    double scaleY = (double)excelDisplayHeight / actualHeight * 100.0;
                                                    scalePercentage = (scaleX + scaleY) / 2.0;
                                                }

                                                _logger.LogDebug($"📐 Excel 顯示尺寸 - 像素: {excelDisplayWidth}×{excelDisplayHeight}px, 厘米: {excelWidthCm:F2}×{excelHeightCm:F2}cm, 縮放: {scalePercentage:F1}%");
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
                                            Width = excelDisplayWidth, // 使用 Excel 顯示寬度
                                            Height = excelDisplayHeight, // 使用 Excel 顯示高度
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

                                            // 原始尺寸和 Excel 縮放資訊
                                            OriginalWidth = actualWidth,
                                            OriginalHeight = actualHeight,
                                            ExcelWidthCm = excelWidthCm,
                                            ExcelHeightCm = excelHeightCm,
                                            ScaleFactor = scalePercentage / 100.0,
                                            IsScaled = Math.Abs(scalePercentage - 100.0) > 1.0,
                                            ScaleMethod = $"Excel 縮放 {scalePercentage:F1}% (顯示: {excelWidthCm:F2}×{excelHeightCm:F2}cm)"
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



                return images.Any() ? images : null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"讀取儲存格 {cell.Address} 的圖片時發生錯誤: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 獲取指定儲存格範圍內的所有浮動物件（文字框、形狀等）
        /// ⭐ 修復: 解決合併儲存格與浮動圖片跨足範圍不一致的問題
        /// </summary>
        private List<FloatingObjectInfo>? GetCellFloatingObjects(ExcelWorksheet worksheet, ExcelRange cell)
        { if(cell.Address.Contains("B19")){
            var a = 0;
        }
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
                            if (drawing is OfficeOpenXml.Drawing.ExcelPicture)
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

                            _logger.LogInformation($"發現浮動物件: '{drawing.Name ?? "未命名"}' 類型: {drawing.GetType().Name} 位置: Row {fromRow}-{toRow}, Col {fromCol}-{toCol}");

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

                            // 記錄詳細的檢查結果
                            _logger.LogDebug($"浮動物件 '{drawing.Name ?? "未命名"}' 位置檢查: " +
                                           $"浮動物件範圍: Row {fromRow}-{toRow}, Col {fromCol}-{toCol} | " +
                                           $"儲存格範圍: Row {cellStartRow}-{cellEndRow}, Col {cellStartCol}-{cellEndCol} | " +
                                           $"有交集: {hasOverlap} | 是錨點儲存格: {isAnchorCell} | 結果: {shouldInclude}");

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
                                    _logger.LogInformation($"✅ 成功解析浮動物件: {floatingObjectInfo.Name}, 類型: {floatingObjectInfo.ObjectType}");
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

        /// <summary>
        /// 獲取繪圖物件類型
        /// </summary>
        private string GetDrawingObjectType(OfficeOpenXml.Drawing.ExcelDrawing drawing)
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
        private string? ExtractTextFromDrawing(OfficeOpenXml.Drawing.ExcelDrawing drawing)
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

                // 如果是 TextBox，嘗試特殊處理
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
        private string? ExtractStyleFromDrawing(OfficeOpenXml.Drawing.ExcelDrawing drawing)
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
        private string? ExtractHyperlinkFromDrawing(OfficeOpenXml.Drawing.ExcelDrawing drawing)
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
                        //return (calculatedWidth, calculatedHeight);
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
                ".webp" => "WEBP",
                ".ico" => "ICO",
                _ => "Unknown"
            };
        }

        /// <summary>
        /// 從檔案名稱獲取圖片格式類型 (別名方法)
        /// </summary>
        private string GetImageTypeFromFileName(string? fileName)
        {
            return GetImageTypeFromName(fileName);
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

                    // EMF 檔頭: 檢查 EMF 格式 (會自動轉換為 PNG)
                    if (IsEmfFormat(bytes))
                    {
                        _logger.LogInformation($"圖片 {picture.Name} 是 EMF 格式，將自動轉換為 PNG 格式");
                        return "PNG"; // 因為會自動轉換，所以返回 PNG 類型
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
        /// 將圖片轉換為 Base64 字串 (支援 EMF 格式自動轉換為 PNG)
        /// </summary>
        private string ConvertImageToBase64(OfficeOpenXml.Drawing.ExcelPicture picture)
        {
            try
            {
                if (picture.Image?.ImageBytes != null && picture.Image.ImageBytes.Length > 0)
                {
                    var imageBytes = picture.Image.ImageBytes;
                    
                    // 檢查是否為 EMF 格式 (Enhanced Metafile)
                    if (IsEmfFormat(imageBytes))
                    {
                        _logger.LogInformation($"🔄 檢測到 EMF 格式圖片: {picture.Name}，正在轉換為 PNG 格式...");
                        
                        // 嘗試轉換 EMF 到 PNG
                        var pngBytes = ConvertEmfToPng(imageBytes);
                        
                        if (pngBytes != null && pngBytes.Length > 0)
                        {
                            _logger.LogInformation($"✅ EMF 轉 PNG 成功: {picture.Name} ({imageBytes.Length} -> {pngBytes.Length} bytes)");
                            return Convert.ToBase64String(pngBytes);
                        }
                        else
                        {
                            _logger.LogWarning($"⚠️ EMF 轉 PNG 失敗: {picture.Name}，使用錯誤提示圖片");
                            var errorPngBytes = CreateEmfErrorPng();
                            return Convert.ToBase64String(errorPngBytes);
                        }
                    }
                    
                    // 非 EMF 格式，直接返回原始資料
                    return Convert.ToBase64String(imageBytes);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"轉換圖片 {picture.Name} 為 Base64 時發生錯誤: {ex.Message}");
                
                // 發生錯誤時，返回錯誤提示圖片
                try
                {
                    var errorPngBytes = CreateEmfErrorPng();
                    return Convert.ToBase64String(errorPngBytes);
                }
                catch
                {
                    return string.Empty;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// 檢查圖片是否為 EMF 格式
        /// </summary>
        private bool IsEmfFormat(byte[] imageBytes)
        {
            if (imageBytes.Length < 44) return false;
            
            // EMF 文件的特徵：在偏移量 40 處有 " EMF" 標識
            return imageBytes[40] == 0x20 && 
                   imageBytes[41] == 0x45 && 
                   imageBytes[42] == 0x4D && 
                   imageBytes[43] == 0x46;
        }

        /// <summary>
        /// 獲取圖片格式信息
        /// </summary>
        private string GetImageFormat(byte[] imageBytes)
        {
            if (imageBytes.Length < 8) return "unknown";
            
            // PNG 格式檢查
            if (imageBytes[0] == 0x89 && imageBytes[1] == 0x50 && imageBytes[2] == 0x4E && imageBytes[3] == 0x47)
                return "png";
            
            // JPEG 格式檢查
            if (imageBytes[0] == 0xFF && imageBytes[1] == 0xD8 && imageBytes[2] == 0xFF)
                return "jpeg";
            
            // GIF 格式檢查
            if (imageBytes[0] == 0x47 && imageBytes[1] == 0x49 && imageBytes[2] == 0x46)
                return "gif";
            
            // EMF 格式檢查
            if (IsEmfFormat(imageBytes))
                return "emf";
            
            return "unknown";
        }

        /// <summary>
        /// 將EMF格式轉換為PNG格式 (跨平台支援)
        /// </summary>
        private byte[]? ConvertEmfToPng(byte[] emfBytes, int width = 800, int height = 600)
        {
            try
            {
                _logger.LogInformation($"開始轉換 EMF 到 PNG，原始大小: {emfBytes.Length} bytes，目標尺寸: {width}x{height}px");

                // 檢查平台支援
                var isWindows = System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows);
                
                // 方法1: Windows 平台使用 System.Drawing 進行實際轉換
                if (isWindows)
                {
                    try
                    {
                        _logger.LogInformation("嘗試使用 System.Drawing 進行 EMF 轉換...");
                        
                        using var emfStream = new MemoryStream(emfBytes);
                        using var emfImage = Image.FromStream(emfStream);
                        
                        // 獲取EMF的實際尺寸
                        var emfWidth = emfImage.Width;
                        var emfHeight = emfImage.Height;
                        
                        _logger.LogInformation($"EMF 原始尺寸: {emfWidth}x{emfHeight}px");
                        
                        // 如果沒有指定目標尺寸，使用EMF的原始尺寸
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
                        
                        // 繪製EMF到位圖 - 保持縱橫比
                        var targetRect = new Rectangle(0, 0, targetWidth, targetHeight);
                        graphics.DrawImage(emfImage, targetRect);
                        
                        // 轉換為PNG
                        using var pngStream = new MemoryStream();
                        pngBitmap.Save(pngStream, ImageFormat.Png);
                        var pngBytes = pngStream.ToArray();
                        
                        _logger.LogInformation($"✅ System.Drawing EMF轉換成功: {emfBytes.Length} -> {pngBytes.Length} bytes, 尺寸: {targetWidth}x{targetHeight}px");
                        return pngBytes;
                    }
                    catch (Exception systemDrawingEx)
                    {
                        _logger.LogError(systemDrawingEx, $"System.Drawing EMF轉換失敗: {systemDrawingEx.Message}");
                        _logger.LogWarning("回退到 SkiaSharp 提示圖片方案");
                    }
                }
                else
                {
                    _logger.LogInformation("非 Windows 平台，EMF格式無法直接轉換，使用提示圖片");
                }

                // 方法2: 跨平台使用 SkiaSharp 創建提示圖片
                _logger.LogInformation("使用 SkiaSharp 創建 EMF 格式提示圖片");
                return CreateEmfPlaceholderPng(width, height, $"EMF 檔案 ({emfBytes.Length} bytes)");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"EMF轉PNG轉換過程中發生未預期的錯誤: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 創建EMF格式的PNG提示圖片
        /// </summary>
        private byte[] CreateEmfPlaceholderPng(int width = 400, int height = 200, string? additionalInfo = null)
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
                    return Array.Empty<byte>();
                }
            }
        }

        /// <summary>
        /// 創建EMF轉換失敗的錯誤提示PNG圖片 (向後相容)
        /// </summary>
        private byte[] CreateEmfErrorPng(int width = 400, int height = 200)
        {
            return CreateEmfPlaceholderPng(width, height, "轉換失敗，請檢查檔案格式");
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
                // var ooxmlImage = TryDirectOoxmlImageSearch(workbook, imageId);
                // if (ooxmlImage != null)
                // {
                //     return ooxmlImage;
                // }

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

            // 使用條件式詳細日誌 (可透過 ENABLE_VERBOSE_LOGGING 開關控制)
            LogVerbose($"Cell {cell.Address} - PatternType: {fill.PatternType}, " +
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
        /// 從 cell.Value 提取安全的基本類型值，避免 EPPlus 內部物件造成 JSON 序列化循環引用
        /// </summary>
        private object? GetSafeValue(object? value)
        {
            if (value == null)
                return null;

            try
            {
                // 獲取值的類型
                var valueType = value.GetType();

                // 如果是基本類型（string, int, double, bool, DateTime 等），直接返回
                if (valueType.IsPrimitive || value is string || value is DateTime || value is decimal)
                {
                    return value;
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

        /// <summary>
        /// 從 EPPlus ExcelColor 物件提取顏色值 (帶快取支援)
        /// </summary>
        private string? GetColorFromExcelColor(OfficeOpenXml.Style.ExcelColor excelColor, ColorCache? cache = null)
        {
            if (excelColor == null)
                return null;

            // 嘗試從快取讀取
            if (cache != null)
            {
                var cacheKey = cache.GetCacheKey(excelColor);
                if (cache.TryGetCachedColor(cacheKey, out var cachedColor))
                {
                    return cachedColor;
                }
            }

            string? result = null;
            try
            {
                // 1. 優先使用 RGB 值 (靜默處理錯誤)
                string? rgbValue = null;
                try
                {
                    rgbValue = excelColor.Rgb;
                }
                catch
                {
                    // 靜默處理 RGB 存取錯誤
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
                        result = colorValue.ToUpperInvariant();
                    }
                    // 處理3位短格式（例如：F00 -> FF0000）
                    else if (colorValue.Length == 3)
                    {
                        result = $"{colorValue[0]}{colorValue[0]}{colorValue[1]}{colorValue[1]}{colorValue[2]}{colorValue[2]}";
                    }
                }

                // 2. 嘗試使用索引顏色 (加強錯誤處理)
                if (result == null)
                {
                    try
                    {
                        if (excelColor.Indexed >= 0)
                        {
                            result = GetIndexedColor(excelColor.Indexed);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogDebug($"無法存取 Indexed 值: {ex.Message}");
                    }
                }

                // 3. 嘗試使用主題顏色 (加強錯誤處理)
                if (result == null)
                {
                    try
                    {
                        if (excelColor.Theme != null)
                        {
                            var themeValue = (int)excelColor.Theme;
                            var tintValue = (double)excelColor.Tint;
                            result = GetThemeColor(themeValue, tintValue);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogDebug($"無法存取 Theme 值: {ex.Message}");
                    }
                }

                // 4. 嘗試自動顏色 (加強錯誤處理)
                if (result == null)
                {
                    try
                    {
                        if (excelColor.Auto == true)
                        {
                            result = "000000"; // 預設黑色
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogDebug($"無法存取 Auto 值: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "解析顏色時發生錯誤");
            }

            // 存入快取
            if (cache != null)
            {
                var cacheKey = cache.GetCacheKey(excelColor);
                cache.CacheColor(cacheKey, result);
            }

            return result;
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
            _logger.LogInformation($"開始處理檔案上傳: {file?.FileName ?? "null"}, 大小: {file?.Length ?? 0} bytes");

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

                // 擴展範圍以包含所有圖片
                if (worksheet.Drawings != null && worksheet.Drawings.Any())
                {
                    foreach (var drawing in worksheet.Drawings)
                    {
                        if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
                        {
                            var picToRow = picture.To.Row + 1;
                            var picToCol = picture.To.Column + 1;

                            if (picToRow > rowCount) rowCount = picToRow;
                            if (picToCol > colCount) colCount = picToCol;

                            _logger.LogDebug($"圖片 '{picture.Name}' 擴展範圍到: Row {picToRow}, Col {picToCol}");
                        }
                    }
                    _logger.LogInformation($"包含圖片後的範圍: {rowCount} 行 x {colCount} 欄");
                }

                excelData.TotalRows = rowCount;
                excelData.TotalColumns = colCount;

                // 🚀 Phase 1 優化: 建立圖片位置索引 (一次性遍歷所有 Drawings)
                var imageIndexStopwatch = System.Diagnostics.Stopwatch.StartNew();
                var imageIndex = new WorksheetImageIndex(worksheet);
                imageIndexStopwatch.Stop();

                // 🚀 Phase 3.1 優化: 建立快取索引 (樣式、顏色、合併儲存格)
                var cacheStopwatch = System.Diagnostics.Stopwatch.StartNew();
                var styleCache = new StyleCache();
                var colorCache = new ColorCache();
                var mergedCellIndex = new MergedCellIndex(worksheet);
                cacheStopwatch.Stop();

                _logger.LogInformation($"⚡ 索引建立完成 - 圖片: {imageIndex.TotalImageCount} 張 ({imageIndexStopwatch.ElapsedMilliseconds}ms), " +
                    $"合併儲存格: {mergedCellIndex.MergeCount} 個 ({cacheStopwatch.ElapsedMilliseconds}ms)");

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

                // 只提供欄位標頭（A、B、C...），避免與資料行重複
                excelData.Headers = new[] { columnHeaders.ToArray() };

                // 讀取資料行，保留原始格式（包含Rich Text） - 使用索引 + 快取優化 + 並行處理
                var processingStopwatch = System.Diagnostics.Stopwatch.StartNew();

                // 🚀 Phase 3.2.3: 使用刪去法處理合併儲存格 - 建立待排除儲存格集合
                var excludedCells = new HashSet<string>(); // 儲存格位址 (如 "B2", "C2", ...)

                var rows = new List<object[]>();
                for (int row = 1; row <= rowCount; row++)
                {
                    var rowData = new List<object>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        var cellAddress = cell.Address; // 如 "A2", "B2", ...

                        // 檢查是否在待排除集合中
                        if (excludedCells.Contains(cellAddress))
                        {
                            excludedCells.Remove(cellAddress); // 刪去法:處理後移除
                            continue; // 跳過此儲存格,不加入 rowData
                        }
                        // 調試用程式碼已移除
                        var cellInfo = CreateCellInfo(cell, worksheet, imageIndex, colorCache, mergedCellIndex);

                        // 如果遇到主合併儲存格,建立待排除集合
                        if (cellInfo.Dimensions?.MergedRangeAddress != null)
                        {
                            var mergedRange = worksheet.Cells[cellInfo.Dimensions.MergedRangeAddress];

                            // 建立該合併範圍內所有儲存格的位址(除了主儲存格)
                            for (int r = mergedRange.Start.Row; r <= mergedRange.End.Row; r++)
                            {
                                for (int c = mergedRange.Start.Column; c <= mergedRange.End.Column; c++)
                                {
                                    var addr = new ExcelCellAddress(r, c).Address;
                                    if (addr != cellAddress) // 排除主儲存格本身
                                    {
                                        excludedCells.Add(addr);
                                    }
                                }
                            }
                        }

                        rowData.Add(cellInfo);
                    }
                    rows.Add(rowData.ToArray());
                }

                processingStopwatch.Stop();

                excelData.Rows = rows.ToArray();

                LogPerformance($"✅ 成功讀取 Excel 檔案: {file.FileName}, 行數: {rowCount}, 欄數: {colCount}, 處理耗時: {processingStopwatch.ElapsedMilliseconds}ms, 平均每行: {processingStopwatch.ElapsedMilliseconds / (double)rowCount:F2}ms");

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

        [HttpGet("test-smart-detection")]
        public ActionResult<object> TestSmartDetection()
        {
            try
            {
                _logger.LogInformation("開始測試智慧內容檢測功能");

                // 使用現有的 Excel 檔案進行測試
                var testFilePath = Path.Combine("d:", "VUE_EPPLUS", "有圖片的excel.xlsx");

                if (!System.IO.File.Exists(testFilePath))
                {
                    return BadRequest($"測試檔案不存在: {testFilePath}");
                }

                using var package = new ExcelPackage(new FileInfo(testFilePath));
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension == null)
                {
                    return BadRequest("Excel 檔案為空");
                }

                // 測試 A1 儲存格
                var cellA1 = worksheet.Cells["A1"];
                var contentType = DetectCellContentType(cellA1, worksheet);

                _logger.LogInformation($"A1 儲存格內容類型檢測結果: {contentType}");

                var cellInfo = CreateCellInfo(cellA1, worksheet);

                return Ok(new
                {
                    Message = "智慧內容檢測測試完成",
                    CellAddress = "A1",
                    DetectedContentType = contentType.ToString(),
                    CellValue = cellA1.Value,
                    CellText = cellA1.Text,
                    HasImages = cellInfo.Images?.Count > 0,
                    ImageCount = cellInfo.Images?.Count ?? 0
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "測試智慧內容檢測時發生錯誤");
                return StatusCode(500, $"測試失敗: {ex.Message}");
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

                        // 邊框 - 使用 GetColorFromExcelColor 避免循環引用
                        Border = new
                        {
                            Top = new { Style = cell.Style.Border.Top.Style.ToString(), Color = GetColorFromExcelColor(cell.Style.Border.Top.Color) },
                            Bottom = new { Style = cell.Style.Border.Bottom.Style.ToString(), Color = GetColorFromExcelColor(cell.Style.Border.Bottom.Color) },
                            Left = new { Style = cell.Style.Border.Left.Style.ToString(), Color = GetColorFromExcelColor(cell.Style.Border.Left.Color) },
                            Right = new { Style = cell.Style.Border.Right.Style.ToString(), Color = GetColorFromExcelColor(cell.Style.Border.Right.Color) },
                            Diagonal = new { Style = cell.Style.Border.Diagonal.Style.ToString(), Color = GetColorFromExcelColor(cell.Style.Border.Diagonal.Color) },
                            DiagonalUp = cell.Style.Border.DiagonalUp,
                            DiagonalDown = cell.Style.Border.DiagonalDown
                        },

                        // 填充/背景 - 使用 GetColorFromExcelColor 避免循環引用
                        Fill = new
                        {
                            PatternType = cell.Style.Fill.PatternType.ToString(),
                            BackgroundColor = GetColorFromExcelColor(cell.Style.Fill.BackgroundColor),
                            PatternColor = GetColorFromExcelColor(cell.Style.Fill.PatternColor),
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