import ExcelJS from 'exceljs'
import { listCloudGrants } from '@/cloud/access'
import { googleAccessToken, googleAccountLabel } from '@/cloud/googleAuth'
import { downloadGoogleFile, googleFileById, uploadGoogleXlsx } from '@/cloud/googleFiles'
import { authedJson } from '@/cloud/http'
import { downloadMicrosoftFile, uploadMicrosoftXlsx } from '@/cloud/microsoftFiles'
import { microsoftAccessToken, microsoftAccountLabel } from '@/cloud/microsoftAuth'
import { FORM_SHEET, canEditCloud, parseFormBook, type SharedForm } from '@/cloud/formShare'
import { GOOGLE_SHEET_MIME, XLSX_MIME, type CloudFileRef, type CloudProviderId } from '@/cloud/types'
import { columnLetter, quotedSheet } from '@/cloud/values'
import type { CellValue, RowData } from '@/types/excel'

function sheetApi(fileId: string): string {
  return `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(fileId)}`
}

function cellText(value: unknown): string {
  if (value == null) return ''
  if (typeof value === 'object' && value && 'text' in value) return String((value as { text?: string }).text ?? '')
  return String(value)
}

export async function signedInEmail(provider: CloudProviderId): Promise<string | null> {
  if (provider === 'microsoft') {
    const label = await microsoftAccountLabel()
    return label?.includes('@') ? label : null
  }
  const label = await googleAccountLabel()
  if (label?.includes('@')) return label
  const token = await googleAccessToken(false)
  const data = await authedJson<{ email?: string }>('https://www.googleapis.com/oauth2/v2/userinfo', token)
  return data.email ?? null
}

export async function assertCloudEditor(provider: CloudProviderId, file: CloudFileRef): Promise<void> {
  const email = await signedInEmail(provider)
  const grants = await listCloudGrants({ provider, mode: 'file', file, revision: '' })
  if (!canEditCloud(grants, email)) throw new Error('Cần quyền sửa file này')
}

export function rowLine(form: SharedForm, row: RowData): CellValue[] {
  const width = Math.max(...form.columns.map((column) => column.index), 0) + 1
  const line: CellValue[] = Array.from({ length: width }, () => '')
  for (const column of form.columns) {
    const value = row[column.key]
    line[column.index] = value == null ? '' : value
  }
  return line
}

async function googleSheetTabs(fileId: string, token: string) {
  const meta = await authedJson<{ sheets?: { properties?: { title?: string; sheetId?: number; gridProperties?: { rowCount?: number } } }[] }>(
    `${sheetApi(fileId)}?fields=sheets.properties(title,sheetId,gridProperties.rowCount)`,
    token,
  )
  return meta.sheets ?? []
}

async function readGoogleSheetForm(file: CloudFileRef): Promise<SharedForm | null> {
  const token = await googleAccessToken()
  try {
    const data = await authedJson<{ values?: unknown[][] }>(
      `${sheetApi(file.id)}/values/${encodeURIComponent(`${quotedSheet(FORM_SHEET)}!A1`)}`,
      token,
    )
    return parseFormBook(cellText(data.values?.[0]?.[0]))
  } catch {
    return null
  }
}

async function writeGoogleSheetForm(file: CloudFileRef, form: SharedForm): Promise<void> {
  const token = await googleAccessToken()
  const tabs = await googleSheetTabs(file.id, token)
  let tab = tabs.find((item) => item.properties?.title === FORM_SHEET)
  if (!tab) {
    const created = await authedJson<{ replies?: { addSheet?: { properties?: { sheetId?: number; title?: string } } }[] }>(
      `${sheetApi(file.id)}:batchUpdate`,
      token,
      { method: 'POST', body: JSON.stringify({ requests: [{ addSheet: { properties: { title: FORM_SHEET, hidden: true } } }] }) },
    )
    tab = created.replies?.[0]?.addSheet
  }
  await authedJson(
    `${sheetApi(file.id)}/values/${encodeURIComponent(`${quotedSheet(FORM_SHEET)}!A1`)}?valueInputOption=RAW`,
    token,
    { method: 'PUT', body: JSON.stringify({ values: [[JSON.stringify(form)]] }) },
  )
  const sheetId = tab?.properties?.sheetId
  if (sheetId == null) return
  await authedJson(`${sheetApi(file.id)}:batchUpdate`, token, {
    method: 'POST',
    body: JSON.stringify({
      requests: [{ updateSheetProperties: { properties: { sheetId, hidden: true }, fields: 'hidden' } }],
    }),
  }).catch(() => undefined)
}

async function appendGoogleSheet(file: CloudFileRef, form: SharedForm, row: RowData): Promise<void> {
  const token = await googleAccessToken()
  const tabs = await googleSheetTabs(file.id, token)
  const tab = tabs.find((item) => item.properties?.title === form.sheet)
  const sheetId = tab?.properties?.sheetId
  const gridRows = tab?.properties?.gridProperties?.rowCount ?? 0
  const postAppend = () => authedJson(
    `${sheetApi(file.id)}/values/${encodeURIComponent(`${quotedSheet(form.sheet)}!A1`)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    token,
    { method: 'POST', body: JSON.stringify({ values: [rowLine(form, row)] }) },
  )
  try {
    await postAppend()
  } catch (error) {
    const message = error instanceof Error ? error.message : ''
    if (sheetId == null || !/grid|exceeds/i.test(message)) throw error
    await authedJson(`${sheetApi(file.id)}:batchUpdate`, token, {
      method: 'POST',
      body: JSON.stringify({
        requests: [{
          updateSheetProperties: {
            properties: { sheetId, gridProperties: { rowCount: Math.max(gridRows, 1) + 1 } },
            fields: 'gridProperties.rowCount',
          },
        }],
      }),
    })
    await postAppend()
  }
}

async function peersGoogleSheet(file: CloudFileRef, form: SharedForm, index: number): Promise<string[]> {
  const token = await googleAccessToken()
  const letter = columnLetter(index)
  const start = form.headerRow + 2
  const data = await authedJson<{ values?: unknown[][] }>(
    `${sheetApi(file.id)}/values/${encodeURIComponent(`${quotedSheet(form.sheet)}!${letter}${start}:${letter}`)}`,
    token,
  )
  return (data.values ?? []).map((line) => cellText(line[0])).filter(Boolean)
}

function graphPath(file: CloudFileRef): string {
  if (!file.graphPath) throw new Error('File OneDrive thiếu đường dẫn')
  return file.graphPath
}

async function graph<T>(path: string, token: string, init?: RequestInit): Promise<T> {
  return authedJson<T>(`https://graph.microsoft.com/v1.0/${path}`, token, init)
}

async function microsoftSheets(file: CloudFileRef, token: string) {
  const listed = await graph<{ value?: { id?: string; name?: string }[] }>(`${graphPath(file)}/workbook/worksheets`, token)
  return listed.value ?? []
}

async function readMicrosoftForm(file: CloudFileRef): Promise<SharedForm | null> {
  const token = await microsoftAccessToken()
  const sheets = await microsoftSheets(file, token)
  const formSheet = sheets.find((sheet) => sheet.name === FORM_SHEET)
  if (!formSheet?.id) return null
  const data = await graph<{ values?: unknown[][] }>(
    `${graphPath(file)}/workbook/worksheets/${encodeURIComponent(formSheet.id)}/range(address='A1')`,
    token,
  )
  return parseFormBook(cellText(data.values?.[0]?.[0]))
}

async function writeMicrosoftForm(file: CloudFileRef, form: SharedForm): Promise<void> {
  const token = await microsoftAccessToken()
  const sheets = await microsoftSheets(file, token)
  let formSheet = sheets.find((sheet) => sheet.name === FORM_SHEET)
  if (!formSheet) {
    formSheet = await graph<{ id?: string }>(`${graphPath(file)}/workbook/worksheets`, token, {
      method: 'POST',
      body: JSON.stringify({ name: FORM_SHEET }),
    })
  }
  if (!formSheet.id) throw new Error('Không tạo được sheet form')
  await graph(
    `${graphPath(file)}/workbook/worksheets/${encodeURIComponent(formSheet.id)}/range(address='A1')`,
    token,
    { method: 'PATCH', body: JSON.stringify({ values: [[JSON.stringify(form)]] }) },
  )
  await graph(`${graphPath(file)}/workbook/worksheets/${encodeURIComponent(formSheet.id)}`, token, {
    method: 'PATCH',
    body: JSON.stringify({ visibility: 'Hidden' }),
  }).catch(() => undefined)
}

async function appendMicrosoft(file: CloudFileRef, form: SharedForm, row: RowData): Promise<void> {
  const token = await microsoftAccessToken()
  const sheets = await microsoftSheets(file, token)
  const dataSheet = sheets.find((sheet) => sheet.name === form.sheet)
  if (!dataSheet?.id) throw new Error('Không thấy sheet dữ liệu')
  const usedResponse = await fetch(
    `https://graph.microsoft.com/v1.0/${graphPath(file)}/workbook/worksheets/${encodeURIComponent(dataSheet.id)}/usedRange?$select=rowCount`,
    { headers: { Authorization: `Bearer ${token}` } },
  )
  const usedRows = usedResponse.ok ? ((await usedResponse.json()) as { rowCount?: number }).rowCount ?? 0 : 0
  const next = Math.max(usedRows, form.headerRow + 1) + 1
  const width = rowLine(form, row).length
  await graph(
    `${graphPath(file)}/workbook/worksheets/${encodeURIComponent(dataSheet.id)}/range(address='A${next}:${columnLetter(width - 1)}${next}')`,
    token,
    { method: 'PATCH', body: JSON.stringify({ values: [rowLine(form, row)] }) },
  )
}

async function peersMicrosoft(file: CloudFileRef, form: SharedForm, index: number): Promise<string[]> {
  const token = await microsoftAccessToken()
  const sheets = await microsoftSheets(file, token)
  const dataSheet = sheets.find((sheet) => sheet.name === form.sheet)
  if (!dataSheet?.id) return []
  const letter = columnLetter(index)
  const data = await graph<{ values?: unknown[][] }>(
    `${graphPath(file)}/workbook/worksheets/${encodeURIComponent(dataSheet.id)}/range(address='${letter}${form.headerRow + 2}:${letter}5000')`,
    token,
  )
  return (data.values ?? []).map((line) => cellText(line[0])).filter(Boolean)
}

async function loadXlsx(file: CloudFileRef, provider: CloudProviderId): Promise<ExcelJS.Workbook> {
  const downloaded = provider === 'google' ? await downloadGoogleFile(file) : await downloadMicrosoftFile(file)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await downloaded.arrayBuffer())
  return workbook
}

async function saveXlsx(file: CloudFileRef, provider: CloudProviderId, workbook: ExcelJS.Workbook): Promise<void> {
  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: XLSX_MIME })
  if (provider === 'google') await uploadGoogleXlsx(file.id, blob)
  else await uploadMicrosoftXlsx(file, blob)
}

function readXlsxForm(workbook: ExcelJS.Workbook): SharedForm | null {
  const sheet = workbook.getWorksheet(FORM_SHEET)
  if (!sheet) return null
  return parseFormBook(cellText(sheet.getCell('A1').value))
}

async function writeXlsxForm(file: CloudFileRef, provider: CloudProviderId, form: SharedForm): Promise<void> {
  const workbook = await loadXlsx(file, provider)
  let sheet = workbook.getWorksheet(FORM_SHEET)
  if (!sheet) sheet = workbook.addWorksheet(FORM_SHEET)
  sheet.getCell('A1').value = JSON.stringify(form)
  sheet.state = 'hidden'
  await saveXlsx(file, provider, workbook)
}

async function appendXlsx(file: CloudFileRef, provider: CloudProviderId, form: SharedForm, row: RowData): Promise<void> {
  const workbook = await loadXlsx(file, provider)
  const sheet = workbook.getWorksheet(form.sheet)
  if (!sheet) throw new Error('Không thấy sheet dữ liệu')
  sheet.addRow(rowLine(form, row).map((value) => (value == null || value === '' ? null : value)))
  await saveXlsx(file, provider, workbook)
}

function peersXlsx(workbook: ExcelJS.Workbook, form: SharedForm, index: number): string[] {
  const sheet = workbook.getWorksheet(form.sheet)
  if (!sheet) return []
  const values: string[] = []
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber <= form.headerRow + 1) return
    const text = cellText(row.getCell(index + 1).value)
    if (text) values.push(text)
  })
  return values
}

function sheetBacked(provider: CloudProviderId, file: CloudFileRef): boolean {
  return provider === 'google' && file.mime === GOOGLE_SHEET_MIME
}

export async function readSharedForm(provider: CloudProviderId, file: CloudFileRef): Promise<SharedForm | null> {
  if (sheetBacked(provider, file)) return readGoogleSheetForm(file)
  if (provider === 'microsoft') return readMicrosoftForm(file)
  return readXlsxForm(await loadXlsx(file, provider))
}

export async function writeSharedForm(provider: CloudProviderId, file: CloudFileRef, form: SharedForm): Promise<void> {
  if (sheetBacked(provider, file)) return writeGoogleSheetForm(file, form)
  if (provider === 'microsoft') return writeMicrosoftForm(file, form)
  return writeXlsxForm(file, provider, form)
}

export async function appendSharedRow(provider: CloudProviderId, file: CloudFileRef, form: SharedForm, row: RowData): Promise<void> {
  if (sheetBacked(provider, file)) return appendGoogleSheet(file, form, row)
  if (provider === 'microsoft') return appendMicrosoft(file, form, row)
  return appendXlsx(file, provider, form, row)
}

export async function uniquePeers(provider: CloudProviderId, file: CloudFileRef, form: SharedForm, index: number): Promise<string[]> {
  if (sheetBacked(provider, file)) return peersGoogleSheet(file, form, index)
  if (provider === 'microsoft') return peersMicrosoft(file, form, index)
  // shortcut: file .xlsx tải cả workbook cho mỗi cột duy nhất, gộp một lần khi form có nhiều cột unique
  return peersXlsx(await loadXlsx(file, provider), form, index)
}

export async function sharedFormJson(provider: CloudProviderId, file: CloudFileRef): Promise<string | null> {
  if (sheetBacked(provider, file)) return null
  try {
    const form = await readSharedForm(provider, file)
    return form ? JSON.stringify(form) : null
  } catch {
    return null
  }
}

export async function fillFile(provider: CloudProviderId, fileId: string): Promise<CloudFileRef> {
  if (provider === 'google') return googleFileById(fileId)
  return {
    id: fileId,
    name: 'Excel',
    mime: XLSX_MIME,
    graphPath: `me/drive/items/${fileId}`,
  }
}
