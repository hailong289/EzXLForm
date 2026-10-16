import { microsoftAccessToken } from '@/cloud/microsoftAuth'
import { microsoftRevision } from '@/cloud/microsoftFiles'
import { apiError, authedJson } from '@/cloud/http'
import { FORM_SHEET } from '@/cloud/formShare'
import { LIVE_MAX_COLS, type CloudFileRef } from '@/cloud/types'
import { columnLetter, sheetToMatrix, workbookFromSheetValues, type LiveRowClear, type LiveRowWrite } from '@/cloud/values'
import type { WorkbookData } from '@/types/excel'

interface Worksheet {
  id?: string
  name?: string
}

async function graph<T>(path: string, init?: RequestInit): Promise<T> {
  const token = await microsoftAccessToken()
  return authedJson<T>(`https://graph.microsoft.com/v1.0/${path}`, token, init)
}

function rangePath(file: CloudFileRef, sheetId: string, address: string): string {
  if (!file.graphPath) throw new Error('File OneDrive thiếu đường dẫn')
  return `${file.graphPath}/workbook/worksheets/${encodeURIComponent(sheetId)}/range(address='${address}')`
}

export async function pullMicrosoftWorkbook(file: CloudFileRef): Promise<{ workbook: WorkbookData; revision: string }> {
  if (!file.graphPath) throw new Error('File OneDrive thiếu đường dẫn')
  const listed = await graph<{ value?: Worksheet[] }>(`${file.graphPath}/workbook/worksheets`)
  const sheets = listed.value ?? []
  const values: { name: string; values: unknown[][] }[] = []
  for (const sheet of sheets) {
    if (!sheet.id || sheet.name === FORM_SHEET) continue
    const path = `${file.graphPath}/workbook/worksheets/${encodeURIComponent(sheet.id)}/usedRange?$select=values`
    const token = await microsoftAccessToken()
    const response = await fetch(`https://graph.microsoft.com/v1.0/${path}`, {
      headers: { Authorization: `Bearer ${token}` },
    })
    if (response.status === 404) {
      values.push({ name: sheet.name || 'Sheet', values: [] })
      continue
    }
    if (!response.ok) throw await apiError(response)
    const data = (await response.json()) as { values?: unknown[][] }
    values.push({ name: sheet.name || 'Sheet', values: data.values ?? [] })
  }
  return {
    workbook: workbookFromSheetValues(file.name, values),
    revision: await microsoftRevision(file),
  }
}

async function ensureWorksheet(file: CloudFileRef, title: string, existing: Map<string, string>): Promise<string> {
  const found = existing.get(title)
  if (found) return found
  if (!file.graphPath) throw new Error('File OneDrive thiếu đường dẫn')
  const created = await graph<Worksheet>(`${file.graphPath}/workbook/worksheets`, {
    method: 'POST',
    body: JSON.stringify({ name: title }),
  })
  if (!created.id) throw new Error(`Không tạo được sheet ${title}`)
  existing.set(title, created.id)
  return created.id
}

export async function pushMicrosoftWorkbook(file: CloudFileRef, workbook: WorkbookData): Promise<string> {
  if (!file.graphPath) throw new Error('File OneDrive thiếu đường dẫn')
  const listed = await graph<{ value?: Worksheet[] }>(`${file.graphPath}/workbook/worksheets`)
  const existing = new Map<string, string>()
  for (const sheet of listed.value ?? []) {
    if (sheet.name && sheet.id) existing.set(sheet.name, sheet.id)
  }
  for (const sheet of workbook.sheets) {
    const sheetId = await ensureWorksheet(file, sheet.name, existing)
    const matrix = sheetToMatrix(sheet)
    const cols = Math.max(matrix[0]?.length ?? 1, 1)
    const address = `A1:${columnLetter(cols - 1)}${Math.max(matrix.length, 1)}`
    await graph(rangePath(file, sheetId, address), {
      method: 'PATCH',
      body: JSON.stringify({ values: matrix }),
    })
    const token = await microsoftAccessToken()
    const usedResponse = await fetch(
      `https://graph.microsoft.com/v1.0/${file.graphPath}/workbook/worksheets/${encodeURIComponent(sheetId)}/usedRange?$select=rowCount`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (usedResponse.ok) {
      const used = (await usedResponse.json()) as { rowCount?: number }
      const usedRows = used.rowCount ?? 0
      if (usedRows > matrix.length) {
        const clearAddress = `A${matrix.length + 1}:${columnLetter(LIVE_MAX_COLS - 1)}${usedRows}`
        await graph(`${rangePath(file, sheetId, clearAddress)}/clear`, {
          method: 'POST',
          body: JSON.stringify({ applyTo: 'Contents' }),
        })
      }
    }
  }
  return microsoftRevision(file)
}

export async function pushMicrosoftRows(
  file: CloudFileRef,
  writes: LiveRowWrite[],
  clears: LiveRowClear[],
): Promise<string> {
  if (!file.graphPath) throw new Error('File OneDrive thiếu đường dẫn')
  const listed = await graph<{ value?: Worksheet[] }>(`${file.graphPath}/workbook/worksheets`)
  const existing = new Map<string, string>()
  for (const sheet of listed.value ?? []) {
    if (sheet.name && sheet.id) existing.set(sheet.name, sheet.id)
  }
  const names = new Set([...writes.map((row) => row.sheetName), ...clears.map((item) => item.sheetName)])
  for (const name of names) {
    if (!existing.has(name)) throw new Error(`Sheet ${name} chưa có trên cloud`)
  }
  for (const row of writes) {
    const sheetId = existing.get(row.sheetName)
    if (!sheetId) continue
    const start = columnLetter(row.startCol)
    const end = columnLetter(row.startCol + Math.max(row.values.length, 1) - 1)
    const address = `${start}${row.excelRow}:${end}${row.excelRow}`
    await graph(rangePath(file, sheetId, address), {
      method: 'PATCH',
      body: JSON.stringify({ values: [row.values] }),
    })
  }
  for (const item of clears) {
    const sheetId = existing.get(item.sheetName)
    if (!sheetId) continue
    const address = `A${item.fromRow}:${columnLetter(item.cols - 1)}${item.toRow}`
    await graph(`${rangePath(file, sheetId, address)}/clear`, {
      method: 'POST',
      body: JSON.stringify({ applyTo: 'Contents' }),
    })
  }
  return microsoftRevision(file)
}
