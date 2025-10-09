using System.Collections.Concurrent;
using System.Drawing;
using OfficeOpenXml.Style;

namespace ExcelReaderAPI.Models.Caches
{
    /// <summary>
    /// 樣式快取 - 減少重複物件建立以降低 GC 壓力
    /// 使用 ConcurrentDictionary 確保執行緒安全
    /// </summary>
    public class StyleCache
    {
        private readonly ConcurrentDictionary<string, object> _fontCache = new();
        private readonly ConcurrentDictionary<string, object> _borderCache = new();
        private readonly ConcurrentDictionary<string, object> _fillCache = new();
        private readonly ConcurrentDictionary<string, object> _alignmentCache = new();

        public object GetOrAddFont(ExcelFont font, Func<object> factory)
        {
            if (font == null) return factory();
            
            string key = $"{font.Name}_{font.Size}_{font.Bold}_{font.Italic}_{font.UnderLine}_{font.Strike}";
            return _fontCache.GetOrAdd(key, _ => factory());
        }

        public object GetOrAddBorder(ExcelBorderItem border, string position, Func<object> factory)
        {
            if (border == null || border.Style == ExcelBorderStyle.None)
                return factory();
            
            string key = $"{position}_{border.Style}_{border.Color?.Rgb}";
            return _borderCache.GetOrAdd(key, _ => factory());
        }

        public object GetOrAddFill(ExcelFill fill, Func<object> factory)
        {
            if (fill == null) return factory();
            
            string key = fill.PatternType switch
            {
                ExcelFillStyle.Solid => $"Solid_{fill.BackgroundColor?.Rgb}",
                ExcelFillStyle.None => "None",
                _ => $"{fill.PatternType}_{fill.BackgroundColor?.Rgb}_{fill.PatternColor?.Rgb}"
            };
            
            return _fillCache.GetOrAdd(key, _ => factory());
        }

        public object GetOrAddAlignment(ExcelHorizontalAlignment horizontal, ExcelVerticalAlignment vertical, 
            bool wrapText, int textRotation, Func<object> factory)
        {
            string key = $"{horizontal}_{vertical}_{wrapText}_{textRotation}";
            return _alignmentCache.GetOrAdd(key, _ => factory());
        }

        public void Clear()
        {
            _fontCache.Clear();
            _borderCache.Clear();
            _fillCache.Clear();
            _alignmentCache.Clear();
        }

        public int GetCacheSize()
        {
            return _fontCache.Count + _borderCache.Count + _fillCache.Count + _alignmentCache.Count;
        }

        public Dictionary<string, int> GetCacheStats()
        {
            return new Dictionary<string, int>
            {
                { "Font", _fontCache.Count },
                { "Border", _borderCache.Count },
                { "Fill", _fillCache.Count },
                { "Alignment", _alignmentCache.Count }
            };
        }
    }
}
