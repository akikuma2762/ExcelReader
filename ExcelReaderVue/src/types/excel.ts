/**
 * Excel 資料類型定義 - 對應後端 EPPlus 完整屬性
 */

/**
 * Rich Text 部分的詳細資訊
 */
export interface RichTextPart {
  text: string
  bold?: boolean
  italic?: boolean
  underLine?: boolean
  strike?: boolean
  size?: number
  fontName?: string
  color?: string
  verticalAlign?: string
}

/**
 * 圖片資訊
 */
export interface ImageInfo {
  name: string
  description?: string
  imageType: string // 如：PNG, JPEG, GIF 等
  width: number
  height: number
  left: number
  top: number
  base64Data: string // Base64 編碼的圖片數據
  fileName?: string
  fileSize: number
  anchorCell?: CellPosition // 錨點儲存格
  hyperlinkAddress?: string // 如果圖片有超連結
}

/**
 * 浮動物件資訊（包含文字框、形狀等）
 */
export interface FloatingObjectInfo {
  name: string
  description?: string
  objectType: string // TextBox, Shape, Drawing 等
  width: number
  height: number
  left: number
  top: number
  text?: string // 文字內容
  anchorCell?: CellPosition // 錨點儲存格
  style?: string // 樣式資訊（顏色、字型等）
  hyperlinkAddress?: string // 如果有超連結
  fromCell?: CellPosition // 起始位置
  toCell?: CellPosition // 結束位置
  isFloating: boolean // 是否為浮動物件
}

/**
 * 儲存格位置資訊
 */
export interface CellPosition {
  row: number
  column: number
  address: string
}

/**
 * 字體樣式詳細資訊
 */
export interface FontInfo {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underLine?: string
  strike?: boolean
  color?: string
  colorTheme?: string
  colorTint?: number
  charset?: number
  scheme?: string
  family?: number
}

/**
 * 對齊方式詳細資訊
 */
export interface AlignmentInfo {
  horizontal?: string
  vertical?: string
  wrapText?: boolean
  indent?: number
  readingOrder?: string
  textRotation?: number
  shrinkToFit?: boolean
}

/**
 * 邊框樣式資訊
 */
export interface BorderStyle {
  style?: string
  color?: string
}

/**
 * 邊框詳細資訊
 */
export interface BorderInfo {
  top?: BorderStyle
  bottom?: BorderStyle
  left?: BorderStyle
  right?: BorderStyle
  diagonal?: BorderStyle
  diagonalUp?: boolean
  diagonalDown?: boolean
}

/**
 * 填充/背景詳細資訊
 */
export interface FillInfo {
  patternType?: string
  backgroundColor?: string
  patternColor?: string
  backgroundColorTheme?: string
  backgroundColorTint?: number
}

/**
 * 尺寸和合併詳細資訊
 */
export interface DimensionInfo {
  columnWidth?: number
  rowHeight?: number
  isMerged?: boolean
  mergedRangeAddress?: string
  isMainMergedCell?: boolean
  rowSpan?: number
  colSpan?: number
}

/**
 * 註解資訊
 */
export interface CommentInfo {
  text?: string
  author?: string
  autoFit?: boolean
  visible?: boolean
}

/**
 * 超連結資訊
 */
export interface HyperlinkInfo {
  absoluteUri?: string
  originalString?: string
  isAbsoluteUri?: boolean
}

/**
 * 儲存格中繼資料
 */
export interface CellMetadata {
  hasFormula?: boolean
  isRichText?: boolean
  styleId?: number
  styleName?: string
  rows?: number
  columns?: number
  start?: CellPosition
  end?: CellPosition
}

/**
 * 完整的Excel儲存格資訊（基於EPPlus所有屬性）
 */
export interface ExcelCellInfo {
  // 基本位置和值
  position: CellPosition

  // 基本值和顯示
  value?: string | number | boolean | Date | null
  text: string
  formula?: string
  formulaR1C1?: string

  // 資料類型
  valueType?: string
  dataType: string

  // 格式化
  numberFormat?: string
  numberFormatId?: number

  // 字體樣式
  font: FontInfo

  // 對齊方式
  alignment: AlignmentInfo

  // 邊框
  border: BorderInfo

  // 填充/背景
  fill: FillInfo

  // 尺寸和合併
  dimensions: DimensionInfo

  // Rich Text
  richText?: RichTextPart[]

  // 註解
  comment?: CommentInfo

  // 超連結
  hyperlink?: HyperlinkInfo

  // 圖片
  images?: ImageInfo[]

  // 浮動物件（文字框、形狀等）
  floatingObjects?: FloatingObjectInfo[]

  // 中繼資料
  metadata: CellMetadata

  // 舊屬性（向後兼容，已棄用）
  // displayText 屬性已移除，請使用 text 屬性
  /** @deprecated 請使用 numberFormat 屬性 */
  formatCode?: string
  /** @deprecated 請使用 font.bold 屬性 */
  fontBold?: boolean
  /** @deprecated 請使用 font.size 屬性 */
  fontSize?: number
  /** @deprecated 請使用 font.name 屬性 */
  fontName?: string
  /** @deprecated 請使用 fill.backgroundColor 屬性 */
  backgroundColor?: string
  /** @deprecated 請使用 font.color 屬性 */
  fontColor?: string
  /** @deprecated 請使用 alignment.horizontal 屬性 */
  textAlign?: string
  /** @deprecated 請使用 dimensions.columnWidth 屬性 */
  columnWidth?: number
  /** @deprecated 請使用 metadata.isRichText 屬性 */
  isRichText?: boolean
  /** @deprecated 請使用 dimensions.rowSpan 屬性 */
  rowSpan?: number
  /** @deprecated 請使用 dimensions.colSpan 屬性 */
  colSpan?: number
  /** @deprecated 請使用 dimensions.isMerged 屬性 */
  isMerged?: boolean
  /** @deprecated 請使用 dimensions.isMainMergedCell 屬性 */
  isMainMergedCell?: boolean
}

/**
 * 工作表資訊
 */
export interface WorksheetInfo {
  name: string
  totalRows: number
  totalColumns: number
  defaultColWidth: number
  defaultRowHeight: number
}

/**
 * Excel 檔案資料
 */
export interface ExcelData {
  headers: ExcelCellInfo[][]
  rows: ExcelCellInfo[][]
  totalRows: number
  totalColumns: number
  fileName: string
  worksheetName: string
  availableWorksheets: string[]
  worksheetInfo?: WorksheetInfo
}

/**
 * 上傳回應
 */
export interface UploadResponse {
  success: boolean
  message: string
  data?: ExcelData
}

/**
 * Debug模式的工作表資訊
 */
export interface DebugWorksheetInfo {
  name: string
  index: number
  state: string
}

/**
 * Debug模式的完整Excel資料
 */
export interface DebugExcelData {
  fileName: string
  worksheetInfo?: WorksheetInfo
  sampleCells?: Record<string, unknown>
  allWorksheets?: DebugWorksheetInfo[]
}

/**
 * 簡化的儲存格資訊（向後兼容）
 */
export interface SimpleCellInfo {
  value: string | number | boolean | Date | null
  text: string
  formatCode: string
  dataType: string
  fontBold?: boolean
  fontSize?: number
  fontName?: string
  backgroundColor?: string
  fontColor?: string
  textAlign?: string
  columnWidth?: number
  richText?: RichTextPart[]
  isRichText?: boolean
  rowSpan?: number
  colSpan?: number
  isMerged?: boolean
  isMainMergedCell?: boolean
}
