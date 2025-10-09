using OfficeOpenXml.Style;
using System.Collections.Concurrent;

namespace ExcelReaderAPI.Models.Caches
{
    /// <summary>
    /// 顏色轉換快取 - 避免重複轉換相同顏色 (執行緒安全)
    /// Phase 3.2: 使用 ConcurrentDictionary 支援並行處理
    /// </summary>
    public class ColorCache
    {
        private readonly ConcurrentDictionary<string, string?> _cache = new();

        public string GetCacheKey(ExcelColor color)
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
}
