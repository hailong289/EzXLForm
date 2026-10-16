import ExcelJS from 'exceljs'
import type { CellValue, WorkbookData } from '@/types/excel'
import { formatDateValue } from '@/utils/dateFormat'
import { buildPagedXlsxBlob } from '@/composables/importLargeWorkbook'
import { FORM_SHEET } from '@/cloud/formShare'
import { XLSX_MIME } from '@/cloud/types'

export async function buildXlsxBlob(
  workbook: WorkbookData,
  onProgress?: (message: string) => void,
  formJson?: string | null,
): Promise<Blob> {
  if (workbook.sheets.some((sheet) => sheet.paged)) {
    return buildPagedXlsxBlob(workbook, onProgress ?? (() => {}))
  }

  const excel = new ExcelJS.Workbook()
  excel.creator = 'Excel Editor'
  excel.created = new Date()

  for (const sheet of workbook.sheets) {
    const worksheet = excel.addWorksheet(sheet.name)
    worksheet.columns = sheet.headers.map((header) => ({
      header: header.label,
      key: header.key,
      width: Math.max(12, Math.min(40, header.label.length + 5)),
    }))

    const headerRow = worksheet.getRow(1)
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF217346' },
    }
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' }

    for (const row of sheet.rows) {
      const rowValues: Record<string, CellValue> = {}
      sheet.headers.forEach((header) => {
        const value = row[header.key]
        if (header.dataType === 'number' && value !== null && value !== undefined && value !== '') {
          rowValues[header.key] = Number(value)
        } else if (header.dataType === 'date') {
          rowValues[header.key] = formatDateValue(value, header.dateFormat)
        } else {
          rowValues[header.key] = value ?? ''
        }
      })
      worksheet.addRow(rowValues)
    }

    worksheet.eachRow((gridRow, rowNumber) => {
      if (rowNumber > 1) gridRow.alignment = { vertical: 'middle' }
    })
  }

  if (formJson) {
    const hidden = excel.addWorksheet(FORM_SHEET)
    hidden.getCell('A1').value = formJson
    hidden.state = 'hidden'
  }

  const buffer = await excel.xlsx.writeBuffer()
  return new Blob([buffer], { type: XLSX_MIME })
}
