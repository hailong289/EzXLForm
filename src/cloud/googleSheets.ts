import { googleAccessToken } from '@/cloud/googleAuth'
import { googleRevision } from '@/cloud/googleFiles'
import { authedJson } from '@/cloud/http'
import { LIVE_MAX_COLS, LIVE_MAX_ROWS, type CloudFileRef } from '@/cloud/types'
import {
  columnLetter,
  quotedSheet,
  sheetToMatrix,
  workbookFromSheetValues,
  type LiveRowClear,
  type LiveRowWrite,
} from '@/cloud/values'
import type { WorkbookData } from '@/types/excel'

interface SheetTab {
  properties?: {
    title?: string
    sheetId?: number
    gridProperties?: { rowCount?: number; columnCount?: number }
  }
}

interface SheetProps {
  sheets?: SheetTab[]
}

function sheetUrl(fileId: string): string {
  return `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(fileId)}`
}

async function loadTabs(fileId: string, token: string): Promise<SheetTab[]> {
  const meta = await authedJson<SheetProps>(
    `${sheetUrl(fileId)}?fields=sheets.properties(title,sheetId,gridProperties(rowCount,columnCount))`,
    token,
  )
  return meta.sheets ?? []
}

async function ensureSheet(fileId: string, title: string, tabs: SheetTab[], token: string): Promise<SheetTab> {
  const found = tabs.find((tab) => tab.properties?.title === title)
  if (found) return found
  const created = await authedJson<{ replies?: { addSheet?: SheetTab }[] }>(`${sheetUrl(fileId)}:batchUpdate`, token, {
    method: 'POST',
    body: JSON.stringify({
      requests: [{ addSheet: { properties: { title } } }],
    }),
  })
  const tab = created.replies?.[0]?.addSheet
  if (!tab?.properties?.sheetId) throw new Error(`Không tạo được sheet ${title}`)
  tabs.push(tab)
  return tab
}

async function resizeGrid(
  fileId: string,
  sheetId: number,
  rowCount: number,
  columnCount: number,
  token: string,
) {
  await authedJson(`${sheetUrl(fileId)}:batchUpdate`, token, {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          updateSheetProperties: {
            properties: {
              sheetId,
              gridProperties: { rowCount, columnCount },
            },
            fields: 'gridProperties.rowCount,gridProperties.columnCount',
          },
        },
      ],
    }),
  })
}

export async function pullGoogleSheet(file: CloudFileRef): Promise<{ workbook: WorkbookData; revision: string }> {
  const token = await googleAccessToken()
  const tabs = await loadTabs(file.id, token)
  const titles = tabs.map((tab) => tab.properties?.title).filter((title): title is string => Boolean(title))
  const values = []
  for (const title of titles) {
    const range = encodeURIComponent(`${quotedSheet(title)}!A1:${columnLetter(LIVE_MAX_COLS - 1)}${LIVE_MAX_ROWS + 1}`)
    const data = await authedJson<{ values?: unknown[][] }>(
      `${sheetUrl(file.id)}/values/${range}?valueRenderOption=FORMATTED_VALUE`,
      token,
    )
    values.push({ name: title, values: data.values ?? [] })
  }
  const revision = await googleRevision(file.id)
  return {
    workbook: workbookFromSheetValues(file.name, values),
    revision,
  }
}

export async function pushGoogleSheet(file: CloudFileRef, workbook: WorkbookData): Promise<string> {
  const token = await googleAccessToken()
  const tabs = await loadTabs(file.id, token)
  for (const sheet of workbook.sheets) {
    const tab = await ensureSheet(file.id, sheet.name, tabs, token)
    const sheetId = tab.properties?.sheetId
    if (sheetId == null) throw new Error(`Không thấy mã sheet ${sheet.name}`)
    const matrix = sheetToMatrix(sheet)
    const rows = Math.max(matrix.length, 1)
    const cols = Math.max(matrix[0]?.length ?? 1, 1)
    const gridRows = tab.properties?.gridProperties?.rowCount ?? 0
    const gridCols = tab.properties?.gridProperties?.columnCount ?? 0
    if (gridRows < rows || gridCols < cols) {
      await resizeGrid(file.id, sheetId, Math.max(gridRows, rows), Math.max(gridCols, cols), token)
    }
    const writeRange = encodeURIComponent(
      `${quotedSheet(sheet.name)}!A1:${columnLetter(cols - 1)}${rows}`,
    )
    await authedJson(`${sheetUrl(file.id)}/values/${writeRange}?valueInputOption=USER_ENTERED`, token, {
      method: 'PUT',
      body: JSON.stringify({ values: matrix }),
    })
    if (gridRows > rows) {
      const clearRows = encodeURIComponent(
        `${quotedSheet(sheet.name)}!A${rows + 1}:${columnLetter(Math.max(gridCols, cols) - 1)}${gridRows}`,
      )
      await authedJson(`${sheetUrl(file.id)}/values/${clearRows}:clear`, token, {
        method: 'POST',
        body: JSON.stringify({}),
      })
    }
    if (gridCols > cols) {
      const clearCols = encodeURIComponent(
        `${quotedSheet(sheet.name)}!${columnLetter(cols)}1:${columnLetter(gridCols - 1)}${Math.max(gridRows, rows)}`,
      )
      await authedJson(`${sheetUrl(file.id)}/values/${clearCols}:clear`, token, {
        method: 'POST',
        body: JSON.stringify({}),
      })
    }
  }
  return googleRevision(file.id)
}

export async function pushGoogleRows(
  file: CloudFileRef,
  writes: LiveRowWrite[],
  clears: LiveRowClear[],
): Promise<string> {
  const token = await googleAccessToken()
  if (writes.length > 0) {
    await authedJson(`${sheetUrl(file.id)}/values:batchUpdate`, token, {
      method: 'POST',
      body: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: writes.map((row) => {
          const start = columnLetter(row.startCol)
          const end = columnLetter(row.startCol + Math.max(row.values.length, 1) - 1)
          const span = start === end ? `${start}${row.excelRow}` : `${start}${row.excelRow}:${end}${row.excelRow}`
          return {
            range: `${quotedSheet(row.sheetName)}!${span}`,
            values: [row.values],
          }
        }),
      }),
    })
  }
  if (clears.length > 0) {
    await authedJson(`${sheetUrl(file.id)}/values:batchClear`, token, {
      method: 'POST',
      body: JSON.stringify({
        ranges: clears.map(
          (item) =>
            `${quotedSheet(item.sheetName)}!A${item.fromRow}:${columnLetter(item.cols - 1)}${item.toRow}`,
        ),
      }),
    })
  }
  return googleRevision(file.id)
}
