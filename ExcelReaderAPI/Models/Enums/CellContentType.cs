namespace ExcelReaderAPI.Models.Enums
{
    /// <summary>
    /// 智能檢測儲存格的主要內容類型
    /// </summary>
    public enum CellContentType
    {
        Empty,          // 空儲存格
        TextOnly,       // 純文字內容
        ImageOnly,      // 純圖片內容
        Mixed           // 混合內容
    }
}
