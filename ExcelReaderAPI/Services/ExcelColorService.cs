using ExcelReaderAPI.Models.Caches;
using ExcelReaderAPI.Services.Interfaces;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReaderAPI.Services
{
    /// <summary>
    /// Excel 顏色處理服務 - 顏色轉換、主題顏色、Tint 效果
    /// Phase 2.4: 從 ExcelController 搬移 5 個方法,保持邏輯完全不變
    /// </summary>
    public class ExcelColorService : IExcelColorService
    {
        private readonly ILogger<ExcelColorService> _logger;

        public ExcelColorService(ILogger<ExcelColorService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// 取得背景顏色
        /// 從 ExcelController.GetBackgroundColor 搬移 (24行)
        /// </summary>
        public string? GetBackgroundColor(ExcelFill fill)
        {
            // 檢查填充類型，只有 Solid 或 Pattern 類型才有背景色
            if (fill.PatternType == ExcelFillStyle.Solid)
            {
                // Solid 填充：使用背景色
                return GetColorFromExcelColor(fill.BackgroundColor);
            }
            else if (fill.PatternType != ExcelFillStyle.None)
            {
                // Pattern 填充：優先使用 BackgroundColor，其次使用 PatternColor
                return GetColorFromExcelColor(fill.BackgroundColor) ??
                       GetColorFromExcelColor(fill.PatternColor);
            }

            return null;
        }

        /// <summary>
        /// 從 ExcelColor 取得顏色字串
        /// 從 ExcelController.GetColorFromExcelColor 搬移 (113行)
        /// </summary>
        public string? GetColorFromExcelColor(ExcelColor? excelColor)
        {
            if (excelColor == null)
                return null;

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

            return result;
        }

        /// <summary>
        /// 取得索引顏色
        /// 從 ExcelController.GetIndexedColor 搬移 (87行)
        /// </summary>
        public string? GetIndexedColor(int? index)
        {
            if (index == null || index < 0)
                return null;

            var colorIndex = index.Value;

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
        /// 取得主題顏色
        /// 從 ExcelController.GetThemeColor 搬移 (35行)
        /// </summary>
        public string? GetThemeColor(int? theme, double? tint)
        {
            if (theme == null)
                return null;

            var themeIndex = theme.Value;
            var tintValue = tint ?? 0.0;

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
            if (Math.Abs(tintValue) > 0.001)
            {
                return ApplyTint(baseColor, tintValue);
            }

            return baseColor;
        }

        /// <summary>
        /// 套用 Tint 效果
        /// 從 ExcelController.ApplyTint 搬移 (37行)
        /// </summary>
        public string ApplyTint(string hexColor, double tint)
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
    }
}
