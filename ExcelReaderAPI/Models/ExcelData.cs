namespace ExcelReaderAPI.Models
{
    /// <summary>
    /// Rich Text 部分的詳細資訊
    /// </summary>
    public class RichTextPart
    {
        public string Text { get; set; } = string.Empty;
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public bool? UnderLine { get; set; }
        public bool? Strike { get; set; }
        public float? Size { get; set; }
        public string? FontName { get; set; }
        public string? Color { get; set; }
        public string? VerticalAlign { get; set; }
    }

    /// <summary>
    /// 儲存格位置資訊
    /// </summary>
    public class CellPosition
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Address { get; set; } = string.Empty;
    }

    /// <summary>
    /// 字體樣式詳細資訊
    /// </summary>
    public class FontInfo
    {
        public string? Name { get; set; }
        public float? Size { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public string? UnderLine { get; set; }
        public bool? Strike { get; set; }
        public string? Color { get; set; }
        public string? ColorTheme { get; set; }
        public double? ColorTint { get; set; }
        public int? Charset { get; set; }
        public string? Scheme { get; set; }
        public int? Family { get; set; }
    }

    /// <summary>
    /// 對齊方式詳細資訊
    /// </summary>
    public class AlignmentInfo
    {
        public string? Horizontal { get; set; }
        public string? Vertical { get; set; }
        public bool? WrapText { get; set; }
        public int? Indent { get; set; }
        public string? ReadingOrder { get; set; }
        public int? TextRotation { get; set; }
        public bool? ShrinkToFit { get; set; }
    }

    /// <summary>
    /// 邊框樣式資訊
    /// </summary>
    public class BorderStyle
    {
        public string? Style { get; set; }
        public string? Color { get; set; }
    }

    /// <summary>
    /// 邊框詳細資訊
    /// </summary>
    public class BorderInfo
    {
        public BorderStyle? Top { get; set; }
        public BorderStyle? Bottom { get; set; }
        public BorderStyle? Left { get; set; }
        public BorderStyle? Right { get; set; }
        public BorderStyle? Diagonal { get; set; }
        public bool? DiagonalUp { get; set; }
        public bool? DiagonalDown { get; set; }
    }

    /// <summary>
    /// 填充/背景詳細資訊
    /// </summary>
    public class FillInfo
    {
        public string? PatternType { get; set; }
        public string? BackgroundColor { get; set; }
        public string? PatternColor { get; set; }
        public string? BackgroundColorTheme { get; set; }
        public double? BackgroundColorTint { get; set; }
    }

    /// <summary>
    /// 尺寸和合併詳細資訊
    /// </summary>
    public class DimensionInfo
    {
        public double? ColumnWidth { get; set; }
        public double? RowHeight { get; set; }
        public bool? IsMerged { get; set; }
        public string? MergedRangeAddress { get; set; }
        public bool? IsMainMergedCell { get; set; }
        public int? RowSpan { get; set; }
        public int? ColSpan { get; set; }
    }

    /// <summary>
    /// 註解資訊
    /// </summary>
    public class CommentInfo
    {
        public string? Text { get; set; }
        public string? Author { get; set; }
        public bool? AutoFit { get; set; }
        public bool? Visible { get; set; }
    }

    /// <summary>
    /// 超連結資訊
    /// </summary>
    public class HyperlinkInfo
    {
        public string? AbsoluteUri { get; set; }
        public string? OriginalString { get; set; }
        public bool? IsAbsoluteUri { get; set; }
    }

    /// <summary>
    /// 儲存格中繼資料
    /// </summary>
    public class CellMetadata
    {
        public bool? HasFormula { get; set; }
        public bool? IsRichText { get; set; }
        public int? StyleId { get; set; }
        public string? StyleName { get; set; }
        public int? Rows { get; set; }
        public int? Columns { get; set; }
        public CellPosition? Start { get; set; }
        public CellPosition? End { get; set; }
    }

    /// <summary>
    /// 完整的Excel儲存格資訊（基於EPPlus所有屬性）
    /// </summary>
    public class ExcelCellInfo
    {
        // 基本位置和值
        public CellPosition Position { get; set; } = new();
        
        // 基本值和顯示
        public object? Value { get; set; }
        public string Text { get; set; } = string.Empty;
        public string? Formula { get; set; }
        public string? FormulaR1C1 { get; set; }
        
        // 資料類型
        public string? ValueType { get; set; }
        public string DataType { get; set; } = string.Empty;
        
        // 格式化
        public string? NumberFormat { get; set; }
        public int? NumberFormatId { get; set; }
        
        // 字體樣式
        public FontInfo Font { get; set; } = new();
        
        // 對齊方式
        public AlignmentInfo Alignment { get; set; } = new();
        
        // 邊框
        public BorderInfo Border { get; set; } = new();
        
        // 填充/背景
        public FillInfo Fill { get; set; } = new();
        
        // 尺寸和合併
        public DimensionInfo Dimensions { get; set; } = new();
        
        // Rich Text
        public List<RichTextPart>? RichText { get; set; }
        
        // 註解
        public CommentInfo? Comment { get; set; }
        
        // 超連結
        public HyperlinkInfo? Hyperlink { get; set; }
        
        // 中繼資料
        public CellMetadata Metadata { get; set; } = new();

        // 舊屬性保持兼容性（標記為過時但仍保留）
        [Obsolete("請使用 Text 屬性")]
        public string DisplayText 
        { 
            get => Text; 
            set => Text = value; 
        }
        
        [Obsolete("請使用 NumberFormat 屬性")]
        public string FormatCode 
        { 
            get => NumberFormat ?? string.Empty; 
            set => NumberFormat = value; 
        }
        
        [Obsolete("請使用 Font.Bold 屬性")]
        public bool? FontBold 
        { 
            get => Font.Bold; 
            set => Font.Bold = value; 
        }
        
        [Obsolete("請使用 Font.Size 屬性")]
        public float? FontSize 
        { 
            get => Font.Size; 
            set => Font.Size = value; 
        }
        
        [Obsolete("請使用 Font.Name 屬性")]
        public string? FontName 
        { 
            get => Font.Name; 
            set => Font.Name = value; 
        }
        
        [Obsolete("請使用 Fill.BackgroundColor 屬性")]
        public string? BackgroundColor 
        { 
            get => Fill.BackgroundColor; 
            set => Fill.BackgroundColor = value; 
        }
        
        [Obsolete("請使用 Font.Color 屬性")]
        public string? FontColor 
        { 
            get => Font.Color; 
            set => Font.Color = value; 
        }
        
        [Obsolete("請使用 Alignment.Horizontal 屬性")]
        public string? TextAlign 
        { 
            get => Alignment.Horizontal; 
            set => Alignment.Horizontal = value; 
        }
        
        [Obsolete("請使用 Dimensions.ColumnWidth 屬性")]
        public double? ColumnWidth 
        { 
            get => Dimensions.ColumnWidth; 
            set => Dimensions.ColumnWidth = value; 
        }
        
        [Obsolete("請使用 Metadata.IsRichText 屬性")]
        public bool IsRichText 
        { 
            get => Metadata.IsRichText ?? false; 
            set => Metadata.IsRichText = value; 
        }
        
        [Obsolete("請使用 Dimensions.RowSpan 屬性")]
        public int? RowSpan 
        { 
            get => Dimensions.RowSpan; 
            set => Dimensions.RowSpan = value; 
        }
        
        [Obsolete("請使用 Dimensions.ColSpan 屬性")]
        public int? ColSpan 
        { 
            get => Dimensions.ColSpan; 
            set => Dimensions.ColSpan = value; 
        }
        
        [Obsolete("請使用 Dimensions.IsMerged 屬性")]
        public bool IsMerged 
        { 
            get => Dimensions.IsMerged ?? false; 
            set => Dimensions.IsMerged = value; 
        }
        
        [Obsolete("請使用 Dimensions.IsMainMergedCell 屬性")]
        public bool IsMainMergedCell 
        { 
            get => Dimensions.IsMainMergedCell ?? false; 
            set => Dimensions.IsMainMergedCell = value; 
        }
    }

    /// <summary>
    /// 工作表資訊
    /// </summary>
    public class WorksheetInfo
    {
        public string Name { get; set; } = string.Empty;
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }
        public double DefaultColWidth { get; set; }
        public double DefaultRowHeight { get; set; }
    }

    /// <summary>
    /// Excel 檔案資料
    /// </summary>
    public class ExcelData
    {
        public object[][] Headers { get; set; } = Array.Empty<object[]>();
        public object[][] Rows { get; set; } = Array.Empty<object[]>();
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string WorksheetName { get; set; } = string.Empty;
        public List<string> AvailableWorksheets { get; set; } = new List<string>();
        public WorksheetInfo? WorksheetInfo { get; set; }
    }

    /// <summary>
    /// 上傳回應
    /// </summary>
    public class UploadResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public ExcelData? Data { get; set; }
    }

    /// <summary>
    /// Debug模式的完整Excel資料
    /// </summary>
    public class DebugExcelData
    {
        public string FileName { get; set; } = string.Empty;
        public WorksheetInfo? WorksheetInfo { get; set; }
        public object[,]? SampleCells { get; set; }
        public List<object>? AllWorksheets { get; set; }
    }
}