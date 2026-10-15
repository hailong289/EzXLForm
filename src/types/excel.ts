// Excel Editor - TypeScript Type Definitions

export interface CellFormat {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  fontColor?: string
  bgColor?: string
  alignment?: 'left' | 'center' | 'right'
  wrapText?: boolean
  numFmt?: string
}

export type CellValue = string | number | boolean | Date | null | undefined

export interface CellData {
  value: CellValue
  formula?: string
  format?: CellFormat
}

export interface ColumnHeader {
  key: string          // unique key (used as data key)
  label: string        // display label
  originalIndex: number // original column index in Excel
  width?: number
  dataType?: 'text' | 'number' | 'date' | 'boolean'
  /** Cách hiện ngày, ví dụ dd/MM/yyyy. Chỉ dùng khi dataType là date */
  dateFormat?: string
  /** Mã hiện số kiểu Excel, ví dụ #,##0.00 hoặc 0%. Chỉ dùng khi dataType là number */
  numFmt?: string
  visible?: boolean
}

export type RowData = Record<string, CellValue>

export interface ColumnSumStat {
  sum: number
  count: number
}

export interface SheetData {
  id: string
  name: string
  headers: ColumnHeader[]
  rows: RowData[]
  headerRowIndex: number  // which row (0-based) in sourceGrid is the header row
  /** File lớn: dòng nằm trong IndexedDB, `rows` chỉ để trống */
  paged?: boolean
  rowCount?: number
  columnSums?: Record<string, ColumnSumStat>
  /** Mọi dòng gốc của sheet, kể cả dòng tiêu đề và dòng phía trên nó */
  sourceGrid?: CellValue[][]
  rawData?: CellData[][]  // raw cell data with formatting
  frozenRows?: number
  frozenCols?: number
}

export interface WorkbookMeta {
  fileName: string
  fileSize: number
  lastModified: Date
  sheetCount: number
}

export interface WorkbookData {
  meta: WorkbookMeta
  sheets: SheetData[]
  activeSheetId: string
}

// History for Undo/Redo
export type HistoryActionType =
  | 'ADD_ROW'
  | 'DELETE_ROWS'
  | 'EDIT_CELL'
  | 'EDIT_HEADER'
  | 'ADD_COLUMN'
  | 'DELETE_COLUMN'
  | 'MOVE_COLUMN'
  | 'BATCH_EDIT'
  | 'FIND_REPLACE'
  | 'ADD_SHEET'
  | 'DELETE_SHEET'
  | 'RENAME_SHEET'
  | 'MOVE_ROW'

export interface HistoryEntry {
  id: string
  timestamp: number
  type: HistoryActionType
  description: string
  sheetId: string
  before: Partial<SheetData>
  after: Partial<SheetData>
  /** Dòng IndexedDB bị sửa, khi sheet đang chia trang */
  rowIndex?: number
  /** Vị trí đích khi kéo thả sắp xếp hàng */
  toIndex?: number
}

// Theme & Settings
export interface ThemeConfig {
  name: string
  primaryColor?: string   // HSL format: "120 60% 40%"
  secondaryColor?: string
  accentColor?: string
}

export interface AppSettings {
  theme: ThemeConfig
  viewMode: 'editor' | 'preview'
  splitRatio: number
  language: 'vi' | 'en'
}

// Search & Filter
export interface SearchFilter {
  query: string
  columns: string[]   // empty = search all columns
  caseSensitive: boolean
  exactMatch: boolean
}

export type FilterJoin = 'and' | 'or'

export type FilterOperator =
  | 'contains'
  | 'not_contains'
  | 'equals'
  | 'not_equals'
  | 'starts'
  | 'ends'
  | 'like'
  | 'not_like'
  | 'empty'
  | 'not_empty'
  | 'gt'
  | 'gte'
  | 'lt'
  | 'lte'

export interface AdvancedFilterRule {
  id: string
  columnKey: string
  operator: FilterOperator
  value: string
}

export interface FindReplaceOptions {
  findText: string
  replaceText: string
  columns: string[]   // empty = all columns
  caseSensitive: boolean
  exactMatch: boolean
  useRegex: boolean
  matchEmpty?: boolean
}

export interface FindReplaceResult {
  matchCount: number
  replacedCount: number
  affectedRows: number[]
  affectedRowCount?: number
  matchedRowCount?: number
}

export interface ReplaceScanStats {
  matchedRows: number
  changedRows: number
}

// Batch Operations
export type BatchOperationType =
  | 'SET_VALUE'
  | 'ADD_PREFIX'
  | 'ADD_SUFFIX'
  | 'UPPERCASE'
  | 'LOWERCASE'
  | 'CAPITALIZE'
  | 'TRIM'
  | 'CLEAR'

export interface BatchOperation {
  type: BatchOperationType
  columnKey: string
  value?: string
  targetRowIndexes: number[]
}

// Import Options
export interface ImportOptions {
  headerRowIndex: number  // 0-based, which row to use as header
  skipRows: number        // rows to skip from top
  sheetName?: string      // if null, use all sheets
}

// Column Context Menu
export interface ColumnAction {
  type: 'rename' | 'add-left' | 'add-right' | 'delete' | 'move-left' | 'move-right' | 'set-type' | 'hide'
  columnKey: string
  payload?: unknown
}

// Row Context Menu
export interface RowAction {
  type: 'edit' | 'delete' | 'duplicate' | 'insert-above' | 'insert-below'
  rowIndex: number
}

// Toast Notification
export type ToastType = 'success' | 'error' | 'warning' | 'info'

export interface Toast {
  id: string
  type: ToastType
  message: string
  duration?: number
}

// Excel Column Letter Utilities
export function columnIndexToLetter(index: number): string {
  let letter = ''
  let n = index + 1
  while (n > 0) {
    const remainder = (n - 1) % 26
    letter = String.fromCharCode(65 + remainder) + letter
    n = Math.floor((n - 1) / 26)
  }
  return letter
}

export function letterToColumnIndex(letter: string): number {
  let result = 0
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + letter.charCodeAt(i) - 64
  }
  return result - 1
}
