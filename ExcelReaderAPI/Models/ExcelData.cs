namespace ExcelReaderAPI.Models
{
    public class RichTextPart
    {
        public string Text { get; set; } = string.Empty;
        public bool? FontBold { get; set; }
        public bool? FontItalic { get; set; }
        public bool? FontUnderline { get; set; }
        public float? FontSize { get; set; }
        public string? FontName { get; set; }
        public string? FontColor { get; set; }
    }

    public class ExcelCellInfo
    {
        public object? Value { get; set; }
        public string DisplayText { get; set; } = string.Empty;
        public string FormatCode { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public bool? FontBold { get; set; }
        public float? FontSize { get; set; }
        public string? FontName { get; set; }
        public string? BackgroundColor { get; set; }
        public string? FontColor { get; set; }
        public string? TextAlign { get; set; }
        public double? ColumnWidth { get; set; }
        public List<RichTextPart>? RichText { get; set; }
        public bool IsRichText { get; set; }
        public int? RowSpan { get; set; }
        public int? ColSpan { get; set; }
        public bool IsMerged { get; set; }
        public bool IsMainMergedCell { get; set; }
    }

    public class ExcelData
    {
        public object[][] Headers { get; set; } = Array.Empty<object[]>();
        public object[][] Rows { get; set; } = Array.Empty<object[]>();
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string WorksheetName { get; set; } = string.Empty;
        public List<string> AvailableWorksheets { get; set; } = new List<string>();
    }

    public class UploadResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public ExcelData? Data { get; set; }
    }
}