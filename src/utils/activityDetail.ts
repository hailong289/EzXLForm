import type { CellValue, ColumnHeader, HistoryEntry, RowData } from '@/types/excel'

const MAX_LINES = 6
const MAX_CHARS = 48

function clip(value: CellValue): string {
  if (value === null || value === undefined || value === '') return '∅'
  const text = String(value)
  return text.length > MAX_CHARS ? `${text.slice(0, MAX_CHARS)}…` : text
}

function labelOf(headers: ColumnHeader[], key: string): string {
  return headers.find((header) => header.key === key)?.label || key
}

function changedFields(before: RowData | undefined, after: RowData | undefined, headers: ColumnHeader[]): string[] {
  const keys = new Set([...Object.keys(before ?? {}), ...Object.keys(after ?? {})])
  const lines: string[] = []
  for (const key of keys) {
    const left = before?.[key]
    const right = after?.[key]
    if (left === right) continue
    if (String(left ?? '') === String(right ?? '')) continue
    lines.push(`${labelOf(headers, key)}: ${clip(left)} → ${clip(right)}`)
  }
  return lines
}

function filledFields(row: RowData | undefined, headers: ColumnHeader[]): string[] {
  if (!row) return []
  const lines: string[] = []
  for (const [key, value] of Object.entries(row)) {
    if (value === null || value === undefined || value === '' || value === false) continue
    lines.push(`${labelOf(headers, key)}: ${clip(value)}`)
  }
  return lines
}

function cap(lines: string[]): string[] {
  if (lines.length <= MAX_LINES) return lines
  return [...lines.slice(0, MAX_LINES), `+${lines.length - MAX_LINES}`]
}

export function activityLines(
  entry: Pick<HistoryEntry, 'type' | 'before' | 'after'>,
  headers: ColumnHeader[] = [],
): string[] {
  if (entry.type === 'ADD_COLUMN' || entry.type === 'DELETE_COLUMN') {
    const column = entry.after.headers?.[0] ?? entry.before.headers?.[0]
    return column ? [`${column.label}${column.dataType ? ` · ${column.dataType}` : ''}`] : []
  }

  const beforeRows = entry.before.rows ?? []
  const afterRows = entry.after.rows ?? []

  if (beforeRows.length === 0 && afterRows.length === 1) return cap(filledFields(afterRows[0], headers))
  if (beforeRows.length === 1 && afterRows.length === 0) return cap(filledFields(beforeRows[0], headers))
  if (beforeRows.length === 1 && afterRows.length === 1) return cap(changedFields(beforeRows[0], afterRows[0], headers))
  if (beforeRows.length > 1 && beforeRows.length === afterRows.length) {
    const lines: string[] = []
    for (let index = 0; index < beforeRows.length; index += 1) {
      const fields = changedFields(beforeRows[index], afterRows[index], headers)
      for (const field of fields) lines.push(`#${index + 1} ${field}`)
    }
    return cap(lines)
  }
  return []
}
