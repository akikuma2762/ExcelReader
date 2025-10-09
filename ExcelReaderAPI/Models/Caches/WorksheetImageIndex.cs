using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace ExcelReaderAPI.Models.Caches
{
    /// <summary>
    /// 工作表圖片位置索引 - 用於效能優化
    /// 一次性建立索引,避免每個儲存格都遍歷所有 Drawings
    /// 複雜度: 建立 O(D), 查詢 O(1), D = Drawings 數量
    /// </summary>
    public class WorksheetImageIndex
    {
        // Key: "Row_Column" (例: "5_3" 代表 Row=5, Col=3)
        // Value: 該儲存格起始位置的所有圖片
        private readonly Dictionary<string, List<ExcelPicture>> _cellImageMap;

        public WorksheetImageIndex(ExcelWorksheet worksheet)
        {
            _cellImageMap = new Dictionary<string, List<ExcelPicture>>();

            if (worksheet.Drawings == null || !worksheet.Drawings.Any())
                return;

            // 一次性遍歷所有繪圖物件建立索引
            foreach (var drawing in worksheet.Drawings)
            {
                if (drawing is ExcelPicture picture && picture.From != null)
                {
                    int fromRow = picture.From.Row + 1; // EPPlus 使用 0-based, 轉為 1-based
                    int fromCol = picture.From.Column + 1;
                    string key = $"{fromRow}_{fromCol}";

                    if (!_cellImageMap.ContainsKey(key))
                        _cellImageMap[key] = new List<ExcelPicture>();

                    _cellImageMap[key].Add(picture);
                }
            }
        }

        /// <summary>
        /// 取得指定儲存格位置的所有圖片
        /// 複雜度: O(1)
        /// </summary>
        public List<ExcelPicture> GetImagesAtCell(int row, int column)
        {
            string key = $"{row}_{column}";
            return _cellImageMap.TryGetValue(key, out var images) ? images : new List<ExcelPicture>();
        }

        /// <summary>
        /// 檢查指定儲存格是否有圖片
        /// 複雜度: O(1)
        /// </summary>
        public bool HasImagesAtCell(int row, int column)
        {
            string key = $"{row}_{column}";
            return _cellImageMap.ContainsKey(key);
        }

        /// <summary>
        /// 取得所有有圖片的儲存格數量
        /// </summary>
        public int GetCellWithImageCount()
        {
            return _cellImageMap.Count;
        }

        /// <summary>
        /// 取得總圖片數量
        /// </summary>
        public int GetTotalImageCount()
        {
            return _cellImageMap.Values.Sum(list => list.Count);
        }
    }
}
