import type { CellValue, WorkbookData } from '@/types/excel'
import { materializeSheet } from '@/utils/sheetGrid'
import { formatDateValue } from '@/utils/dateFormat'
import { LIVE_MAX_COLS, LIVE_MAX_ROWS } from '@/cloud/types'

export function columnLetter(index: number): string {
  let current = index + 1
  let letters = ''
  while (current > 0) {
    const remainder = (current - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    current = Math.floor((current - 1) / 26)
  }
  return letters
}

function cellFromApi(value: unknown): CellValue {
  if (value == null || value === '') return null
  if (typeof value === 'number' || typeof value === 'boolean') return value
  if (typeof value === 'string') return value
  return String(value)
}

function trimMatrix(rows: unknown[][]): CellValue[][] {
  let width = 0
  const grid = rows.map((row) => {
    const line = Array.isArray(row) ? row : []
    let last = -1
    for (let index = line.length - 1; index >= 0; index--) {
      const value = line[index]
      if (value != null && value !== '') {
        last = index
        break
      }
    }
    width = Math.max(width, last + 1)
    return line
  })
  if (width > LIVE_MAX_COLS) {
    throw new Error(`Sheet rộng hơn ${LIVE_MAX_COLS} cột. Hãy mở bằng chế độ file.`)
  }
  const dataRows = Math.max(0, grid.length - 1)
  if (dataRows > LIVE_MAX_ROWS) {
    throw new Error(`Sheet dài hơn ${LIVE_MAX_ROWS.toLocaleString('vi-VN')} dòng. Hãy mở bằng chế độ file.`)
  }
  return grid.map((line) => line.slice(0, width).map(cellFromApi))
}

export function workbookFromSheetValues(
  fileName: string,
  sheets: { name: string; values: unknown[][] }[],
): WorkbookData {
  if (sheets.length === 0) throw new Error('File không chứa sheet dữ liệu nào.')
  const built = sheets.map((sheet, index) => {
    const grid = trimMatrix(sheet.values.length ? sheet.values : [[]])
    const materialized = materializeSheet(grid, 0)
    return {
      id: `cloud_sheet_${index}_${Date.now()}`,
      name: sheet.name || `Sheet${index + 1}`,
      headers: materialized.headers,
      rows: materialized.rows,
      headerRowIndex: materialized.headerRowIndex,
      sourceGrid: grid,
    }
  })
  return {
    meta: {
      fileName,
      fileSize: 0,
      lastModified: new Date(),
      sheetCount: built.length,
    },
    sheets: built,
    activeSheetId: built[0].id,
  }
}

export function sheetToMatrix(sheet: WorkbookData['sheets'][number]): (string | number | boolean)[][] {
  const header = sheet.headers.map((item) => item.label)
  const rows = sheet.rows.map((row) =>
    sheet.headers.map((item) => {
      const value = row[item.key]
      if (value == null || value === '') return ''
      if (typeof value === 'number' || typeof value === 'boolean') return value
      if (value instanceof Date) return value.toISOString().slice(0, 10)
      if (item.dataType === 'date') return formatDateValue(value, item.dateFormat)
      return String(value)
    }),
  )
  const matrix = [header, ...rows]
  if (sheet.headers.length > LIVE_MAX_COLS || sheet.rows.length > LIVE_MAX_ROWS) {
    throw new Error(
      `Chế độ trực tiếp chỉ ghi tối đa ${LIVE_MAX_ROWS.toLocaleString('vi-VN')} dòng và ${LIVE_MAX_COLS} cột. Hãy dùng chế độ file.`,
    )
  }
  return matrix
}

export type LiveCell = string | number | boolean

export interface LiveRowWrite {
  sheetName: string
  excelRow: number
  /** Cột bắt đầu, 0 = A */
  startCol: number
  values: LiveCell[]
}

export interface LiveRowClear {
  sheetName: string
  fromRow: number
  toRow: number
  cols: number
}

export type LiveDelta =
  | { kind: 'full' }
  | { kind: 'rows'; writes: LiveRowWrite[]; clears: LiveRowClear[] }

export function snapshotSheetMatrices(workbook: WorkbookData): Map<string, LiveCell[][]> {
  return new Map(workbook.sheets.map((sheet) => [sheet.name, sheetToMatrix(sheet)]))
}

function sameRow(left: LiveCell[], right: LiveCell[]): boolean {
  if (left.length !== right.length) return false
  return left.every((cell, index) => cell === right[index])
}

/** So với bản đã sync. Sửa một dòng chỉ trả về đúng dòng đó. */
export function diffLiveWorkbook(current: WorkbookData, previous: Map<string, LiveCell[][]> | null): LiveDelta {
  if (!previous) return { kind: 'full' }
  const next = snapshotSheetMatrices(current)
  if (next.size !== previous.size) return { kind: 'full' }
  const writes: LiveRowWrite[] = []
  const clears: LiveRowClear[] = []
  for (const [name, matrix] of next) {
    const before = previous.get(name)
    if (!before) return { kind: 'full' }
    const width = matrix[0]?.length ?? 0
    const beforeWidth = before[0]?.length ?? 0
    if (width !== beforeWidth) return { kind: 'full' }
    const shared = Math.min(matrix.length, before.length)
    for (let index = 0; index < shared; index += 1) {
      const nextRow = matrix[index]
      const prevRow = before[index]
      if (sameRow(nextRow, prevRow)) continue
      let first = -1
      let last = -1
      const width = Math.max(nextRow.length, prevRow.length)
      for (let column = 0; column < width; column += 1) {
        if ((nextRow[column] ?? '') === (prevRow[column] ?? '')) continue
        if (first < 0) first = column
        last = column
      }
      if (first < 0) continue
      writes.push({
        sheetName: name,
        excelRow: index + 1,
        startCol: first,
        values: nextRow.slice(first, last + 1).map((cell) => cell ?? ''),
      })
    }
    for (let index = before.length; index < matrix.length; index += 1) {
      writes.push({ sheetName: name, excelRow: index + 1, startCol: 0, values: matrix[index] })
    }
    if (matrix.length < before.length && width > 0) {
      clears.push({
        sheetName: name,
        fromRow: matrix.length + 1,
        toRow: before.length,
        cols: width,
      })
    }
  }
  return { kind: 'rows', writes, clears }
}

export function assertLiveWorkbook(workbook: WorkbookData): void {
  if (workbook.sheets.some((sheet) => sheet.paged)) {
    throw new Error('File lớn không ghi từng ô lên cloud. Hãy lưu bằng chế độ file .xlsx.')
  }
  for (const sheet of workbook.sheets) sheetToMatrix(sheet)
}

export function quotedSheet(title: string): string {
  return `'${title.replace(/'/g, "''")}'`
}
