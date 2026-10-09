// stores/workbookStore.ts
// Pinia store quan ly toan bo Workbook/Sheet data
// Bao gom: CRUD rows/cols, find & replace, sheet management

import { defineStore } from 'pinia'
import type {
  WorkbookData,
  SheetData,
  RowData,
  CellValue,
  ColumnHeader,
  FindReplaceOptions,
  FindReplaceResult,
  ReplaceScanStats,
} from '@/types/excel'
import {
  cellToDisplayString,
  ensureSourceGrid,
  insertGridColumn,
  materializeSheet,
  moveGridColumn,
  removeGridColumn,
  insertBodyRow,
  moveBodyRow,
  removeBodyRows,
  writeBodyRow,
  writeHeaderLabel,
} from '@/utils/sheetGrid'
import { clearAllSheetOrders, deleteRowKey, getRow, hasSheetOrder, loadSheetMeta, moveLogicalRow, persistSheetOrder, physicalOf, placeInsertedRow, putRow, readLogicalRange, removeLogicalRow, restoreSheetOrders, saveSheetMeta, sheetOrderLength, writeRows } from '@/composables/rowPageStore'
import { buildFindRegex, isBlankCell } from '@/utils/rowFilter'

// State Interface

interface WorkbookState {
  workbook: WorkbookData | null
  isLoading: boolean
  /** true khi co thay doi chua duoc luu */
  isDirty: boolean
  /** Tăng mỗi lần sửa dòng để bảng và preview vẽ lại ngay */
  dataRevision: number
  /** Dòng vừa thêm, chưa bấm lưu. Khóa theo sheet id. */
  freshRows: Record<string, number[]>
}

// Helper

/** Tim sheet theo id, nem loi neu khong ton tai */
function getSheetOrThrow(workbook: WorkbookData, sheetId: string): SheetData {
  const sheet = workbook.sheets.find((s) => s.id === sheetId)
  if (!sheet) throw new Error(`Sheet "${sheetId}" not found`)
  return sheet
}

function numericOf(value: CellValue): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value === 'string' && value.trim() !== '') {
    const numeric = Number(value.replace(/,/g, ''))
    if (!Number.isNaN(numeric)) return numeric
  }
  return null
}

function applySumDelta(sheet: SheetData, oldRow: RowData | null, newRow: RowData): void {
  if (!sheet.columnSums) sheet.columnSums = {}
  for (const header of sheet.headers) {
    if (header.dataType === 'date' || header.dataType === 'boolean') continue
    const stat = sheet.columnSums[header.key] ?? { sum: 0, count: 0 }
    const previous = numericOf(oldRow?.[header.key] ?? null)
    const next = numericOf(newRow[header.key] ?? null)
    if (previous != null) {
      stat.sum -= previous
      stat.count = Math.max(0, stat.count - 1)
    }
    if (next != null) {
      stat.sum += next
      stat.count += 1
    }
    sheet.columnSums[header.key] = stat
  }
}

async function rememberPagedSheet(sheet: SheetData): Promise<void> {
  await saveSheetMeta(sheet.id, {
    rowCount: sheet.rowCount ?? 0,
    columnSums: sheet.columnSums ?? {},
  })
}
function buildSearchRegex(options: FindReplaceOptions): RegExp | null {
  return buildFindRegex(options)
}

const REPLACE_CHUNK = 8000

function replaceCellText(cell: string, regex: RegExp, replaceText: string): { text: string; count: number } {
  let count = 0
  regex.lastIndex = 0
  const text = cell.replace(regex, () => {
    count += 1
    return replaceText
  })
  return { text, count }
}

function noteFreshMove(map: Record<string, number[]>, sheetId: string, from: number, to: number): Record<string, number[]> {
  const list = map[sheetId]
  if (!list?.length || from === to) return map
  const next = list.map((index) => {
    if (index === from) return to
    if (from < to && index > from && index <= to) return index - 1
    if (to < from && index >= to && index < from) return index + 1
    return index
  })
  return { ...map, [sheetId]: next }
}

function noteFreshDelete(map: Record<string, number[]>, sheetId: string, indexes: number[]): Record<string, number[]> {
  let list = map[sheetId] ?? []
  for (const index of [...indexes].sort((a, b) => b - a)) {
    list = list.filter((item) => item !== index).map((item) => (item > index ? item - 1 : item))
  }
  return { ...map, [sheetId]: list }
}

function yieldTurn(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0))
}

let rowMutationQueue: Promise<void> = Promise.resolve()

/** Ghi IndexedDB lần lượt để lần gõ sau không bị lần gõ trước đè lên. */
function enqueueRowMutation(task: () => Promise<void>): Promise<void> {
  const run = rowMutationQueue.then(task, task)
  rowMutationQueue = run.then(
    () => undefined,
    () => undefined,
  )
  return run
}

// Store Definition

export const useWorkbookStore = defineStore('workbook', {
  state: (): WorkbookState => ({
    workbook: null,
    isLoading: false,
    isDirty: false,
    dataRevision: 0,
    freshRows: {},
  }),

  getters: {
    /** Sheet dang duoc active (theo activeSheetId cua workbook) */
    activeSheet(state): SheetData | null {
      if (!state.workbook) return null
      return (
        state.workbook.sheets.find(
          (s) => s.id === state.workbook!.activeSheetId,
        ) ?? null
      )
    },

    /** Danh sach tat ca sheets */
    sheetList(state): SheetData[] {
      return state.workbook?.sheets ?? []
    },

    /** Index cua sheet dang active trong mang sheets */
    isFreshRow: (state) => (sheetId: string, index: number) =>
      (state.freshRows[sheetId] ?? []).includes(index),

    activeSheetIndex(state): number {
      if (!state.workbook) return -1
      return state.workbook.sheets.findIndex(
        (s) => s.id === state.workbook!.activeSheetId,
      )
    },
  },

  actions: {
    // Workbook-level

    /** Gan workbook moi vao store */
    async setWorkbook(wb: WorkbookData): Promise<void> {
      for (const sheet of wb.sheets) {
        if (sheet.paged) continue
        const broken = sheet.headers.some((header) => String(header.label).includes('[object Object]'))
        if (broken && sheet.sourceGrid?.length) {
          const built = materializeSheet(sheet.sourceGrid, sheet.headerRowIndex)
          sheet.headers = built.headers
          sheet.rows = built.rows
          sheet.headerRowIndex = built.headerRowIndex
        } else {
          sheet.headers.forEach((header, index) => {
            const label = cellToDisplayString(header.label).trim()
            if (!label || String(header.label).includes('[object Object]')) {
              header.label = `Cột ${index + 1}`
            }
          })
        }
      }
      clearAllSheetOrders()
      const pagedIds = wb.sheets.filter((sheet) => sheet.paged).map((sheet) => sheet.id)
      await restoreSheetOrders(pagedIds)
      for (const sheet of wb.sheets) {
        if (!sheet.paged) continue
        const meta = await loadSheetMeta(sheet.id)
        const stored = Math.max(meta?.rowCount ?? 0, sheetOrderLength(sheet.id))
        if (stored > (sheet.rowCount ?? 0)) sheet.rowCount = stored
      }
      this.workbook = wb
      this.isDirty = false
    },

    /**
     * Chọn dòng header riêng cho một sheet.
     * Lưới gốc được giữ lại nên mỗi sheet có thể trỏ tới một dòng khác nhau.
     */
    setHeaderRow(sheetId: string, headerRowIndex: number): void {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      if (sheet.paged) return
      const grid = ensureSourceGrid(sheet)
      sheet.sourceGrid = grid
      const built = materializeSheet(grid, headerRowIndex)
      sheet.headers = built.headers
      sheet.rows = built.rows
      sheet.headerRowIndex = built.headerRowIndex
      this.isDirty = true
      this.dataRevision += 1
    },

    /** Doi active sheet theo id */
    setActiveSheet(sheetId: string): void {
      if (!this.workbook) return
      const exists = this.workbook.sheets.some((s) => s.id === sheetId)
      if (!exists) return
      this.workbook.activeSheetId = sheetId
    },

    /** Reset ve trang thai ban dau */
    resetWorkbook(): void {
      clearAllSheetOrders()
      this.workbook = null
      this.isLoading = false
      this.isDirty = false
      this.freshRows = {}
    },

    commitFreshRow(sheetId: string, index: number): void {
      this.freshRows = {
        ...this.freshRows,
        [sheetId]: (this.freshRows[sheetId] ?? []).filter((item) => item !== index),
      }
    },

    // Row Operations

    /**
     * Them mot dong vao sheet.
     * @param position Neu khong truyen, them vao cuoi
     */
    async addRow(sheetId: string, rowData: RowData, position?: number): Promise<void> {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      if (sheet.paged) {
        return enqueueRowMutation(async () => {
        const count = sheet.rowCount ?? 0
        const at = position === undefined || position >= count ? count : Math.max(0, position)
        const physical = placeInsertedRow(sheet.id, at, count)
        await putRow(sheet.id, physical, { ...rowData })
        sheet.rowCount = count + 1
        applySumDelta(sheet, null, rowData)
        await rememberPagedSheet(sheet)
        await persistSheetOrder(sheet.id)
        this.isDirty = true
        this.dataRevision += 1
        })
      }

      const at = position === undefined || position >= sheet.rows.length
        ? sheet.rows.length
        : Math.max(0, position)
      const next = { ...rowData }
      if (at >= sheet.rows.length) sheet.rows.push(next)
      else sheet.rows.splice(at, 0, next)
      insertBodyRow(sheet, at, next)
      this.isDirty = true
      this.dataRevision += 1
    },

    /** Đổi vị trí một dòng. File lớn chỉ đổi thứ tự hiển thị, không ghi lại các dòng khác. */
    async moveRow(sheetId: string, from: number, to: number): Promise<void> {
      if (!this.workbook || from === to) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      const count = sheet.paged ? sheet.rowCount ?? 0 : sheet.rows.length
      if (from < 0 || to < 0 || from >= count || to >= count) return
      if (sheet.paged) {
        return enqueueRowMutation(async () => {
          if (!moveLogicalRow(sheet.id, from, to, count)) return
          await persistSheetOrder(sheet.id)
          this.freshRows = noteFreshMove(this.freshRows, sheet.id, from, to)
          this.isDirty = true
          this.dataRevision += 1
        })
      }
      const [row] = sheet.rows.splice(from, 1)
      sheet.rows.splice(to, 0, row)
      moveBodyRow(sheet, from, to)
      this.freshRows = noteFreshMove(this.freshRows, sheet.id, from, to)
      this.isDirty = true
      this.dataRevision += 1
    },

    /**
     * Xoa nhieu dong theo indexes.
     * Tu dong sort descending de tranh index bi lech khi xoa.
     */
    async deleteRows(sheetId: string, rowIndexes: number[]): Promise<void> {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      if (sheet.paged) {
        return enqueueRowMutation(async () => {
          const sorted = [...rowIndexes].sort((a, b) => b - a)
          if (hasSheetOrder(sheet.id)) {
            for (const idx of sorted) {
              const physical = removeLogicalRow(sheet.id, idx)
              if (physical === null) continue
              const removed = await getRow(sheet.id, physical)
              await deleteRowKey(sheet.id, physical)
              sheet.rowCount = Math.max(0, (sheet.rowCount ?? 1) - 1)
              applySumDelta(sheet, removed, {})
            }
            await rememberPagedSheet(sheet)
            await persistSheetOrder(sheet.id)
            this.isDirty = true
            this.dataRevision += 1
            return
          }
          for (const idx of sorted) {
            const last = (sheet.rowCount ?? 0) - 1
            if (idx < 0 || idx > last) continue
            const removed = await getRow(sheet.id, idx)
            if (idx !== last) {
              const tail = (await getRow(sheet.id, last)) ?? {}
              await putRow(sheet.id, idx, tail)
            }
            await deleteRowKey(sheet.id, last)
            sheet.rowCount = last
            applySumDelta(sheet, removed, {})
          }
          await rememberPagedSheet(sheet)
          this.isDirty = true
          this.dataRevision += 1
        })
      }

      const sorted = [...rowIndexes].sort((a, b) => b - a)
      const removed: number[] = []
      for (const idx of sorted) {
        if (idx >= 0 && idx < sheet.rows.length) {
          sheet.rows.splice(idx, 1)
          removed.push(idx)
        }
      }
      removeBodyRows(sheet, removed)
      this.freshRows = noteFreshDelete(this.freshRows, sheetId, removed)
      this.isDirty = true
      this.dataRevision += 1
    },

    // Cell Operations

    /** Cap nhat gia tri mot cell */
    async updateCell(
      sheetId: string,
      rowIndex: number,
      columnKey: string,
      value: CellValue,
    ): Promise<void> {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      if (sheet.paged) {
        return enqueueRowMutation(async () => {
          if (rowIndex < 0 || rowIndex >= (sheet.rowCount ?? 0)) return
          const physical = physicalOf(sheet.id, rowIndex)
          const old = (await getRow(sheet.id, physical)) ?? {}
          const next = { ...old, [columnKey]: value }
          await putRow(sheet.id, physical, next)
          applySumDelta(sheet, old, next)
          await rememberPagedSheet(sheet)
          this.isDirty = true
          this.dataRevision += 1
        })
      }
      if (rowIndex < 0 || rowIndex >= sheet.rows.length) return

      const next = { ...sheet.rows[rowIndex], [columnKey]: value }
      sheet.rows.splice(rowIndex, 1, next)
      writeBodyRow(sheet, rowIndex, next)
      this.isDirty = true
      this.dataRevision += 1
    },

    /** Thay cả dòng — splice để bảng và preview thấy ngay khi gõ form */
    async replaceRow(sheetId: string, rowIndex: number, data: RowData): Promise<void> {
      return enqueueRowMutation(async () => {
        if (!this.workbook) return
        const sheet = getSheetOrThrow(this.workbook, sheetId)
        if (sheet.paged) {
          if (rowIndex < 0 || rowIndex >= (sheet.rowCount ?? 0)) return
          const physical = physicalOf(sheet.id, rowIndex)
          const old = (await getRow(sheet.id, physical)) ?? {}
          const next = { ...data }
          await putRow(sheet.id, physical, next)
          applySumDelta(sheet, old, next)
          await rememberPagedSheet(sheet)
          this.isDirty = true
          this.dataRevision += 1
          return
        }
        if (rowIndex < 0 || rowIndex >= sheet.rows.length) return
        const next = { ...data }
        sheet.rows.splice(rowIndex, 1, next)
        writeBodyRow(sheet, rowIndex, next)
        this.isDirty = true
        this.dataRevision += 1
      })
    },

    // Column (Header) Operations

    /** Doi label (tieu de hien thi) cua mot column */
    updateHeader(sheetId: string, columnKey: string, newLabel: string): void {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      const header = sheet.headers.find((h) => h.key === columnKey)
      if (!header) return

      header.label = newLabel
      const columnIndex = sheet.headers.findIndex((item) => item.key === columnKey)
      if (columnIndex >= 0) writeHeaderLabel(sheet, columnIndex, newLabel)
      this.isDirty = true
      this.dataRevision += 1
    },

    /**
     * Them column moi.
     * @param afterColumnKey Them sau column nay; null = them vao dau
     */
    addColumn(
      sheetId: string,
      afterColumnKey: string | null,
      header: ColumnHeader,
    ): void {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)

      let insertAt: number
      if (afterColumnKey === null) {
        insertAt = 0
      } else {
        const idx = sheet.headers.findIndex((h) => h.key === afterColumnKey)
        insertAt = idx === -1 ? sheet.headers.length : idx + 1
      }

      sheet.headers.splice(insertAt, 0, { ...header })
      insertGridColumn(sheet, insertAt, header.label)

      // Them key vao moi row voi gia tri null
      for (const row of sheet.rows) {
        if (!(header.key in row)) {
          row[header.key] = null
        }
      }
      this.isDirty = true
      this.dataRevision += 1
    },

    /** Xoa mot column khoi headers va tat ca rows */
    deleteColumn(sheetId: string, columnKey: string): void {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)

      const columnIndex = sheet.headers.findIndex((h) => h.key === columnKey)
      sheet.headers = sheet.headers.filter((h) => h.key !== columnKey)
      if (columnIndex >= 0) removeGridColumn(sheet, columnIndex)

      for (const row of sheet.rows) {
        delete row[columnKey]
      }
      this.isDirty = true
      this.dataRevision += 1
    },

    /** Đưa cột fromKey vào đúng vị trí toIndex (trước khi gỡ cột nguồn) */
    reorderColumn(sheetId: string, fromKey: string, toIndex: number): void {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      const fromIdx = sheet.headers.findIndex((h) => h.key === fromKey)
      if (fromIdx < 0 || toIndex < 0 || toIndex >= sheet.headers.length || fromIdx === toIndex) return

      const [header] = sheet.headers.splice(fromIdx, 1)
      sheet.headers.splice(toIndex, 0, header)
      moveGridColumn(sheet, fromIdx, toIndex)
      this.isDirty = true
      this.dataRevision += 1
    },

    /** Swap vi tri cua 2 columns theo key */
    moveColumn(sheetId: string, fromKey: string, toKey: string): void {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      const toIdx = sheet.headers.findIndex((h) => h.key === toKey)
      if (toIdx < 0) return
      this.reorderColumn(sheetId, fromKey, toIdx)
    },

    // Sheet Management

    /** Them mot sheet moi vao workbook */
    addSheet(sheet: SheetData): void {
      if (!this.workbook) return
      this.workbook.sheets.push({ ...sheet })
      this.workbook.meta.sheetCount = this.workbook.sheets.length
      this.isDirty = true
    },

    /** Xoa sheet theo id; neu dang active thi chuyen sang sheet ke can */
    deleteSheet(sheetId: string): void {
      if (!this.workbook) return
      const idx = this.workbook.sheets.findIndex((s) => s.id === sheetId)
      if (idx === -1) return

      const wasActive = this.workbook.activeSheetId === sheetId
      this.workbook.sheets.splice(idx, 1)
      this.workbook.meta.sheetCount = this.workbook.sheets.length

      if (wasActive && this.workbook.sheets.length > 0) {
        const nextIdx = Math.min(idx, this.workbook.sheets.length - 1)
        this.workbook.activeSheetId = this.workbook.sheets[nextIdx].id
      }
      this.isDirty = true
    },

    /** Doi ten sheet */
    renameSheet(sheetId: string, newName: string): void {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      sheet.name = newName
      this.isDirty = true
    },

    /** Doi thu tu sheet (drag & drop) */
    reorderSheet(fromIndex: number, toIndex: number): void {
      if (!this.workbook) return
      const sheets = this.workbook.sheets
      if (
        fromIndex < 0 ||
        fromIndex >= sheets.length ||
        toIndex < 0 ||
        toIndex >= sheets.length ||
        fromIndex === toIndex
      ) {
        return
      }

      const [removed] = sheets.splice(fromIndex, 1)
      sheets.splice(toIndex, 0, removed)
      this.isDirty = true
    },

    // Batch & Advanced Operations

    /**
     * Cap nhat hang loat cells trong cung mot column.
     * @param valueFactory Ham nhan gia tri cu, tra ve gia tri moi
     */
    batchUpdateRows(
      sheetId: string,
      rowIndexes: number[],
      columnKey: string,
      valueFactory: (oldVal: CellValue) => CellValue,
    ): void {
      if (!this.workbook) return
      const sheet = getSheetOrThrow(this.workbook, sheetId)

      for (const idx of rowIndexes) {
        if (idx >= 0 && idx < sheet.rows.length) {
          const oldVal = sheet.rows[idx][columnKey] ?? null
          sheet.rows[idx] = {
            ...sheet.rows[idx],
            [columnKey]: valueFactory(oldVal),
          }
        }
      }
      this.isDirty = true
    },

    /**
     * Tim va thay the trong sheet.
     * Tra ve so match, so da replace va cac row index bi anh huong.
     */
    async findAndReplace(
      sheetId: string,
      options: FindReplaceOptions,
      onProgress?: (done: number, total: number, stats?: ReplaceScanStats) => void,
      control?: { checkpoint: () => Promise<boolean> },
    ): Promise<FindReplaceResult> {
      const result: FindReplaceResult = {
        matchCount: 0,
        replacedCount: 0,
        affectedRows: [],
      }

      if (!this.workbook || (!options.findText && !options.matchEmpty)) return result

      const sheet = getSheetOrThrow(this.workbook, sheetId)
      const regex = options.matchEmpty ? null : buildSearchRegex(options)
      if (!options.matchEmpty && !regex) return result

      // Xac dinh cac column can tim kiem
      const targetKeys =
        options.columns.length > 0
          ? options.columns
          : sheet.headers.map((h) => h.key)

      const affectedSet = new Set<number>()
      let matchedRows = 0
      const totalRows = Math.max(sheet.rows.length, 1)
      const report = (scanned: number) => onProgress?.(scanned, totalRows, { matchedRows, changedRows: affectedSet.size })
      report(0)

      for (let rowIdx = 0; rowIdx < sheet.rows.length; rowIdx++) {
        if (rowIdx % REPLACE_CHUNK === 0 && control && !(await control.checkpoint())) break
        const row = sheet.rows[rowIdx]
        let rowMatched = false
        let rowAffected = false

        for (const key of targetKeys) {
          if (!options.matchEmpty && !(key in row)) continue

          const cellStr = row[key] != null ? String(row[key]) : ''
          if (options.matchEmpty) {
            if (!isBlankCell(row[key])) continue
            rowMatched = true
            result.matchCount += 1
            if (options.replaceText !== cellStr) {
              row[key] = options.replaceText
              result.replacedCount += 1
              rowAffected = true
            }
            continue
          }
          if (!cellStr || !regex) continue

          const replaced = replaceCellText(cellStr, regex, options.replaceText)
          if (replaced.count === 0) continue
          rowMatched = true
          result.matchCount += replaced.count
          if (replaced.text !== cellStr) {
            row[key] = replaced.text
            result.replacedCount += replaced.count
            rowAffected = true
          }
        }

        if (rowMatched) matchedRows += 1
        if (rowAffected) affectedSet.add(rowIdx)
        if ((rowIdx + 1) % REPLACE_CHUNK === 0 || rowIdx + 1 === sheet.rows.length) {
          report(rowIdx + 1)
          await yieldTurn()
          if (control && !(await control.checkpoint())) break
        }
      }

      result.affectedRows = Array.from(affectedSet).sort((a, b) => a - b)
      result.matchedRowCount = matchedRows
      result.affectedRowCount = affectedSet.size
      if (result.replacedCount > 0) {
        this.isDirty = true
        this.dataRevision += 1
      }

      return result
    },

    async findAndReplacePaged(
      sheetId: string,
      options: FindReplaceOptions,
      onProgress?: (done: number, total: number, stats?: ReplaceScanStats) => void,
      control?: { checkpoint: () => Promise<boolean> },
    ): Promise<FindReplaceResult> {
      const result: FindReplaceResult = { matchCount: 0, replacedCount: 0, affectedRows: [] }
      if (!this.workbook || (!options.findText && !options.matchEmpty)) return result
      const sheet = getSheetOrThrow(this.workbook, sheetId)
      if (!sheet.paged) return this.findAndReplace(sheetId, options, onProgress, control)
      const regex = options.matchEmpty ? null : buildSearchRegex(options)
      if (!options.matchEmpty && !regex) return result
      const targetKeys = options.columns.length > 0 ? options.columns : sheet.headers.map((header) => header.key)
      const rowCount = sheet.rowCount ?? 0
      const totalRows = Math.max(rowCount, 1)
      let matchedRows = 0
      let changedRows = 0
      const report = (scanned: number) => onProgress?.(scanned, totalRows, { matchedRows, changedRows })
      report(0)
      const dirty: { index: number; row: RowData }[] = []
      let stopped = false
      for (let start = 0; start < rowCount && !stopped; start += REPLACE_CHUNK) {
        if (control && !(await control.checkpoint())) break
        const count = Math.min(REPLACE_CHUNK, rowCount - start)
        const rows = await readLogicalRange(sheet.id, start, count)
        for (let offset = 0; offset < rows.length; offset++) {
          const row = rows[offset]
          let matched = false
          let changed = false
          const next: RowData = { ...row }
          for (const key of targetKeys) {
            const cellStr = next[key] != null ? String(next[key]) : ''
            if (options.matchEmpty) {
              if (!isBlankCell(next[key])) continue
              matched = true
              result.matchCount += 1
              if (options.replaceText !== cellStr) {
                next[key] = options.replaceText
                result.replacedCount += 1
                changed = true
              }
              continue
            }
            if (!cellStr || !regex) continue
            const replaced = replaceCellText(cellStr, regex, options.replaceText)
            if (replaced.count === 0) continue
            matched = true
            result.matchCount += replaced.count
            if (replaced.text !== cellStr) {
              next[key] = replaced.text
              result.replacedCount += replaced.count
              changed = true
            }
          }
          if (matched) matchedRows += 1
          if (changed) {
            changedRows += 1
            applySumDelta(sheet, row, next)
            dirty.push({ index: physicalOf(sheet.id, start + offset), row: next })
          }
        }
        if (dirty.length > 0) {
          await writeRows(sheet.id, dirty)
          dirty.length = 0
        }
        report(start + rows.length)
        await yieldTurn()
        if (control && !(await control.checkpoint())) stopped = true
      }
      result.matchedRowCount = matchedRows
      result.affectedRowCount = changedRows
      if (result.replacedCount > 0) {
        await rememberPagedSheet(sheet)
        this.isDirty = true
        this.dataRevision += 1
      }
      return result
    },
  },
})

export type WorkbookStore = ReturnType<typeof useWorkbookStore>
