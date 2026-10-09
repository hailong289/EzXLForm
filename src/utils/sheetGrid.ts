import type { CellValue, ColumnHeader, RowData, SheetData } from '@/types/excel'

export function sheetRowCount(sheet: SheetData): number {
  return sheet.paged ? sheet.rowCount ?? 0 : sheet.rows.length
}

/** Ghép lại lưới gốc từ header + dòng hiện tại, để lần chọn header sau không mất dòng tiêu đề cũ */
export function synthesizeGrid(sheet: SheetData): CellValue[][] {
  const headerLine = sheet.headers.map((h) => h.label)
  const body = sheet.rows.map((row) => sheet.headers.map((h) => row[h.key] ?? null))
  return [headerLine, ...body]
}

export function ensureSourceGrid(sheet: SheetData): CellValue[][] {
  if (sheet.sourceGrid && sheet.sourceGrid.length > 0) return sheet.sourceGrid
  return synthesizeGrid(sheet)
}

function gridRows(sheet: SheetData): CellValue[][] {
  const grid = ensureSourceGrid(sheet)
  sheet.sourceGrid = grid
  return grid
}

function bodyLineIndex(sheet: SheetData, rowIndex: number): number {
  return Math.max(0, sheet.headerRowIndex) + 1 + rowIndex
}

/** Ghi một dòng dữ liệu vào lưới gốc, để xem trước và lần chọn header sau thấy giá trị vừa sửa */
export function writeBodyRow(sheet: SheetData, rowIndex: number, data: RowData): void {
  if (sheet.paged) return
  const grid = gridRows(sheet)
  const lineIndex = bodyLineIndex(sheet, rowIndex)
  while (grid.length <= lineIndex) grid.push([])
  const line = grid[lineIndex]
  sheet.headers.forEach((header, columnIndex) => {
    while (line.length <= columnIndex) line.push(null)
    line[columnIndex] = data[header.key] ?? null
  })
}

export function insertBodyRow(sheet: SheetData, rowIndex: number, data: RowData): void {
  if (sheet.paged) return
  const grid = gridRows(sheet)
  const line = sheet.headers.map((header) => data[header.key] ?? null)
  const at = Math.min(bodyLineIndex(sheet, rowIndex), grid.length)
  grid.splice(at, 0, line)
}

/** Đồng bộ toàn bộ dòng dữ liệu hiện tại vào lưới gốc sau hoàn tác */
export function syncBodyRows(sheet: SheetData): void {
  if (sheet.paged) return
  const grid = gridRows(sheet)
  const base = Math.max(0, sheet.headerRowIndex) + 1
  const previous = Math.max(0, grid.length - base)
  if (previous > sheet.rows.length) grid.splice(base + sheet.rows.length, previous - sheet.rows.length)
  sheet.rows.forEach((row, index) => writeBodyRow(sheet, index, row))
}

export function removeBodyRows(sheet: SheetData, rowIndexes: number[]): void {
  if (sheet.paged) return
  const grid = gridRows(sheet)
  const sorted = [...rowIndexes].sort((a, b) => b - a)
  for (const rowIndex of sorted) {
    const lineIndex = bodyLineIndex(sheet, rowIndex)
    if (lineIndex >= 0 && lineIndex < grid.length) grid.splice(lineIndex, 1)
  }
}

/** Ghi nhãn header vào đúng cột trên dòng tiêu đề của lưới gốc */
export function writeHeaderLabel(sheet: SheetData, columnIndex: number, label: string): void {
  const grid = gridRows(sheet)
  const headerIndex = Math.min(Math.max(0, sheet.headerRowIndex), Math.max(grid.length - 1, 0))
  const row = grid[headerIndex] ? [...grid[headerIndex]] : []
  while (row.length <= columnIndex) row.push(null)
  row[columnIndex] = label
  grid[headerIndex] = row
}

/** Chèn một cột vào lưới gốc, nhãn nằm trên dòng header */
export function insertGridColumn(sheet: SheetData, columnIndex: number, label: string): void {
  const grid = gridRows(sheet)
  const headerIndex = Math.min(Math.max(0, sheet.headerRowIndex), Math.max(grid.length - 1, 0))
  if (grid.length === 0) grid.push([])
  for (let rowIndex = 0; rowIndex < grid.length; rowIndex++) {
    const row = grid[rowIndex] ? [...grid[rowIndex]] : []
    row.splice(columnIndex, 0, rowIndex === headerIndex ? label : null)
    grid[rowIndex] = row
  }
}

/** Đưa một cột của lưới gốc sang vị trí mới, cùng thứ tự với header */
export function moveGridColumn(sheet: SheetData, fromIndex: number, toIndex: number): void {
  if (fromIndex === toIndex) return
  const grid = gridRows(sheet)
  for (const row of grid) {
    const line = [...row]
    while (line.length <= Math.max(fromIndex, toIndex)) line.push(null)
    const [cell] = line.splice(fromIndex, 1)
    line.splice(toIndex, 0, cell ?? null)
    row.splice(0, row.length, ...line)
  }
}
export function removeGridColumn(sheet: SheetData, columnIndex: number): void {
  if (!sheet.sourceGrid) return
  for (const row of sheet.sourceGrid) {
    if (columnIndex >= 0 && columnIndex < row.length) row.splice(columnIndex, 1)
  }
}

/** Chỉ đưa các dòng đầu vào menu, vì header thường nằm gần đầu sheet */
export function headerChoiceIndexes(gridLength: number, currentIndex: number, limit = 40): number[] {
  const indexes: number[] = []
  const max = Math.min(gridLength, limit)
  for (let index = 0; index < max; index++) indexes.push(index)
  if (currentIndex >= limit && currentIndex < gridLength) indexes.push(currentIndex)
  return indexes
}

/** Đọc ô Excel (rich text, hyperlink, formula) thành chữ, không để [object Object] */
export function cellToDisplayString(value: unknown): string {
  if (value == null) return ''
  if (typeof value === 'string') return value === '[object Object]' ? '' : value
  if (typeof value === 'number' || typeof value === 'boolean') return String(value)
  if (value instanceof Date) return value.toISOString().split('T')[0]
  if (Array.isArray(value)) {
    return value.map((item) => cellToDisplayString(item)).filter(Boolean).join('')
  }
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>
    if (Array.isArray(record.richText)) {
      return record.richText.map((part) => cellToDisplayString((part as { text?: unknown }).text)).join('')
    }
    if ('text' in record && record.text != null) return cellToDisplayString(record.text)
    if ('result' in record) return cellToDisplayString(record.result)
    if ('error' in record && record.error != null) return String(record.error)
    return ''
  }
  const text = String(value)
  return text === '[object Object]' ? '' : text
}

export function previewGridLine(grid: CellValue[][], rowIndex: number): string {
  const line = grid[rowIndex] ?? []
  const parts = line
    .map((cell) => cellToDisplayString(cell).trim())
    .filter((text) => text !== '')
    .slice(0, 3)
  const text = parts.join(' · ')
  if (!text) return '(trống)'
  return text.length > 42 ? `${text.slice(0, 42)}…` : text
}

/** Tách một dòng trong lưới thành header, các dòng phía dưới thành dữ liệu */
export function materializeSheet(
  grid: CellValue[][],
  headerRowIndex: number,
  dataStartRow?: number,
): { headers: ColumnHeader[]; rows: RowData[]; headerRowIndex: number } {
  const safeHeader = grid.length === 0 ? 0 : Math.min(Math.max(0, headerRowIndex), grid.length - 1)
  const width = grid.reduce((max, row) => Math.max(max, row.length), 1)
  const headerLine = grid[safeHeader] ?? []
  const headers: ColumnHeader[] = []

  for (let col = 0; col < width; col++) {
    const labelText = cellToDisplayString(headerLine[col]).trim()
    const label = labelText || `Cột ${col + 1}`
    headers.push({
      key: `col_${col}`,
      label,
      originalIndex: col,
      dataType: 'text',
    })
  }

  const start = dataStartRow ?? safeHeader + 1
  const rows: RowData[] = []
  for (let r = Math.max(start, safeHeader + 1); r < grid.length; r++) {
    const line = grid[r] ?? []
    const row: RowData = {}
    let hasValue = false
    headers.forEach((header, index) => {
      const value = normalizeCell(line[index])
      row[header.key] = value
      if (value != null && value !== '') hasValue = true
    })
    if (hasValue) rows.push(row)
  }

  detectColumnDataTypes(headers, rows)
  return { headers, rows, headerRowIndex: safeHeader }
}

function normalizeCell(value: CellValue | undefined): CellValue {
  if (value == null || value === '') return null
  if (typeof value === 'string') return value === '[object Object]' ? null : value
  if (typeof value === 'number' || typeof value === 'boolean') return value
  if (value instanceof Date) return value.toISOString().split('T')[0]
  const text = cellToDisplayString(value).trim()
  return text || null
}

function detectColumnDataTypes(headers: ColumnHeader[], rows: RowData[]) {
  for (const header of headers) {
    let numberCount = 0
    let total = 0
    for (const row of rows) {
      const value = row[header.key]
      if (value !== null && value !== undefined && value !== '') {
        total++
        if (!Number.isNaN(Number(value)) && String(value).trim() !== '') numberCount++
      }
    }
    header.dataType = total > 0 && numberCount === total ? 'number' : 'text'
  }
}
