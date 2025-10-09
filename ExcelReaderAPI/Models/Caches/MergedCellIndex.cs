using OfficeOpenXml;

namespace ExcelReaderAPI.Models.Caches
{
    /// <summary>
    /// 合併儲存格索引 - 快速查詢儲存格是否在合併範圍內
    /// 複雜度: 建立 O(M×C), 查詢 O(1), M=合併範圍數, C=每個範圍的儲存格數
    /// </summary>
    public class MergedCellIndex
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
}
