import { ref } from 'vue'
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'
import type {
  WorkbookData,
  SheetData,
  ColumnHeader,
  RowData,
  CellValue,
  ImportOptions,
} from '@/types/excel'
import { cellToDisplayString, materializeSheet } from '@/utils/sheetGrid'
import { formatDateValue } from '@/utils/dateFormat'
import {
  LARGE_FILE_BYTES,
  exportLargeCsv,
  exportLargeWorkbook,
  importLargeWorkbook,
  isImportStopped,
  isUserAbort,
  type LargeImportWatch,
} from '@/composables/importLargeWorkbook'

export function useExcelParser() {
  const isLoading = ref<boolean>(false)
  const error = ref<string | null>(null)
  const importProgress = ref('')

  /**
   * Đọc file Excel (.xlsx, .xls, .csv)
   */
  async function readExcelFile(
    file: File,
    options?: Partial<ImportOptions>,
    watch?: LargeImportWatch & { onProgress: (message: string) => void },
  ): Promise<WorkbookData> {
    error.value = null
    const extension = file.name.split('.').pop()?.toLowerCase()
    const large = (extension === 'xlsx' || extension === 'csv') && file.size >= LARGE_FILE_BYTES
    if (large && watch) {
      try {
        return await importLargeWorkbook(
          file,
          options,
          (message) => {
            watch.onProgress(message)
          },
          watch,
        )
      } catch (err: unknown) {
        if (isImportStopped(err)) throw err
        const msg = err instanceof Error ? err.message : 'Lỗi không xác định khi đọc file Excel'
        error.value = msg
        throw new Error(msg)
      }
    }

    isLoading.value = true

    try {
      if ((extension === 'xlsx' || extension === 'csv') && file.size >= LARGE_FILE_BYTES) {
        importProgress.value = 'Đang mở file lớn theo từng trang...'
        const workbookData = await importLargeWorkbook(file, options, (message) => {
          importProgress.value = message
        })
        return workbookData
      }
      if (extension === 'xls' && file.size >= LARGE_FILE_BYTES) {
        throw new Error('File .xls quá lớn. Hãy lưu lại thành .xlsx rồi mở.')
      }

      const arrayBuffer = await file.arrayBuffer()
      const wb = new ExcelJS.Workbook()
      await wb.xlsx.load(arrayBuffer)

      const headerRowIndex = options?.headerRowIndex ?? 0 // 0-based
      const skipRows = options?.skipRows ?? 0

      const sheets: SheetData[] = []

      wb.eachSheet((worksheet, sheetId) => {
        const rowCount = worksheet.rowCount
        const colCount = valueColumnCount(worksheet)
        const sourceGrid: CellValue[][] = []

        for (let r = 1; r <= rowCount; r++) {
          const row = worksheet.getRow(r)
          const line: CellValue[] = []
          for (let c = 1; c <= colCount; c++) {
            line.push(readCellValue(row.getCell(c)))
          }
          sourceGrid.push(line)
        }

        if (sourceGrid.length === 0) {
          sourceGrid.push([])
        }

        const safeHeader = Math.min(headerRowIndex, sourceGrid.length - 1)
        const built = materializeSheet(sourceGrid, safeHeader, safeHeader + 1 + skipRows)

        sheets.push({
          id: `sheet_${sheetId}_${Date.now()}`,
          name: worksheet.name,
          headers: built.headers,
          rows: built.rows,
          headerRowIndex: built.headerRowIndex,
          sourceGrid,
        })
      })

      if (sheets.length === 0) {
        throw new Error('File không chứa sheet dữ liệu nào.')
      }

      const workbookData: WorkbookData = {
        meta: {
          fileName: file.name,
          fileSize: file.size,
          lastModified: new Date(file.lastModified),
          sheetCount: sheets.length,
        },
        sheets,
        activeSheetId: sheets[0].id,
      }

      return workbookData
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Lỗi không xác định khi đọc file Excel'
      error.value = msg
      throw new Error(msg)
    } finally {
      isLoading.value = false
      importProgress.value = ''
    }
  }

  /**
   * Cột cuối cùng có giá trị thật.
   * Bỏ qua ô chỉ có định dạng: Excel có thể ghi cả dòng tới cột XFC (16.383 cột).
   */
  function valueColumnCount(worksheet: ExcelJS.Worksheet): number {
    let last = 0
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (colNumber > last && cellHasValue(cell)) last = colNumber
      })
    })
    return Math.max(last, 1)
  }

  function cellHasValue(cell: ExcelJS.Cell): boolean {
    const value = cell.value
    if (value == null || value === '') return false
    if (typeof value === 'object' && 'result' in value) {
      const result = value.result
      return result != null && result !== ''
    }
    return true
  }

  function readCellValue(cell: ExcelJS.Cell): CellValue {
    const value = cell.value
    if (value == null) return null
    if (typeof value === 'number' || typeof value === 'boolean') return value
    if (typeof value === 'string') return value === '[object Object]' ? null : value
    if (value instanceof Date) return value.toISOString().split('T')[0]

    if (typeof value === 'object' && 'result' in value) {
      const result = value.result
      if (typeof result === 'number' || typeof result === 'boolean') return result
      if (typeof result === 'string') return result === '[object Object]' ? null : result
      if (result instanceof Date) return result.toISOString().split('T')[0]
    }

    const fromValue = cellToDisplayString(value).trim()
    if (fromValue) return fromValue

    const shown = typeof cell.text === 'string' ? cell.text.trim() : ''
    if (shown && shown !== '[object Object]') return shown
    return null
  }

  /**
   * Xuất toàn bộ workbook ra file .xlsx
   */
  async function exportToExcel(workbook: WorkbookData, customFileName?: string): Promise<void> {
    isLoading.value = true
    error.value = null

    try {
      if (workbook.sheets.some((sheet) => sheet.paged)) {
        importProgress.value = 'Đang ghi file lớn ra đĩa...'
        const finalFileName = customFileName || workbook.meta.fileName || 'exported.xlsx'
        await exportLargeWorkbook(workbook, finalFileName, (message) => {
          importProgress.value = message
        })
        return
      }

      const wb = new ExcelJS.Workbook()
      wb.creator = 'Excel Editor'
      wb.created = new Date()

      for (const sheet of workbook.sheets) {
        const ws = wb.addWorksheet(sheet.name)

        // Cấu hình columns
        ws.columns = sheet.headers.map((h) => ({
          header: h.label,
          key: h.key,
          width: Math.max(12, Math.min(40, h.label.length + 5)),
        }))

        // Style header row
        const headerRow = ws.getRow(1)
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } }
        headerRow.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF217346' }, // Màu xanh Excel
        }
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' }

        // Thêm data rows
        for (const row of sheet.rows) {
          const rowValues: Record<string, CellValue> = {}
          sheet.headers.forEach((h) => {
            const v = row[h.key]
            if (h.dataType === 'number' && v !== null && v !== undefined && v !== '') {
              rowValues[h.key] = Number(v)
            } else if (h.dataType === 'date') {
              rowValues[h.key] = formatDateValue(v, h.dateFormat)
            } else {
              rowValues[h.key] = v ?? ''
            }
          })
          ws.addRow(rowValues)
        }

        // Tự động căn lề
        ws.eachRow((r, rowNum) => {
          if (rowNum > 1) {
            r.alignment = { vertical: 'middle' }
          }
        })
      }

      const buffer = await wb.xlsx.writeBuffer()
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      })
      const finalFileName = customFileName || workbook.meta.fileName || 'exported.xlsx'
      saveAs(blob, finalFileName.endsWith('.xlsx') ? finalFileName : `${finalFileName}.xlsx`)
    } catch (err: unknown) {
      if (isUserAbort(err)) throw err
      const msg = err instanceof Error ? err.message : 'Lỗi khi xuất file Excel'
      error.value = msg
      throw new Error(msg)
    } finally {
      isLoading.value = false
      importProgress.value = ''
    }
  }

  /**
   * Xuất sheet hiện tại ra file .csv
   */
  async function exportToCsv(sheetData: SheetData, customFileName?: string): Promise<void> {
    try {
      if (sheetData.paged) {
        isLoading.value = true
        importProgress.value = 'Đang ghi CSV ra đĩa...'
        await exportLargeCsv(sheetData, customFileName, (message) => {
          importProgress.value = message
        })
        return
      }
      const headerLine = sheetData.headers.map((h) => `"${h.label.replace(/"/g, '""')}"`).join(',')
      const rowLines = sheetData.rows.map((row) =>
        sheetData.headers
          .map((h) => {
            const val = row[h.key] ?? ''
            return `"${String(val).replace(/"/g, '""')}"`
          })
          .join(','),
      )

      // Thêm BOM \uFEFF để Excel mở UTF-8 (Tiếng Việt) không bị lỗi font
      const csvContent = '\uFEFF' + [headerLine, ...rowLines].join('\r\n')
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
      const fileName = customFileName || `${sheetData.name || 'sheet'}.csv`
      saveAs(blob, fileName.endsWith('.csv') ? fileName : `${fileName}.csv`)
    } catch (err: unknown) {
      if (isUserAbort(err)) throw err
      const msg = err instanceof Error ? err.message : 'Lỗi khi xuất CSV'
      error.value = msg
      throw new Error(msg)
    } finally {
      if (sheetData.paged) {
        isLoading.value = false
        importProgress.value = ''
      }
    }
  }

  return {
    isLoading,
    importProgress,
    error,
    readExcelFile,
    exportToExcel,
    exportToCsv,
  }
}
