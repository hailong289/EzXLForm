import type { CellValue, ColumnHeader, RowData } from '@/types/excel'
import { DEFAULT_DATE_FORMAT, parseDateToIso } from '@/utils/dateFormat'
import { peerValues, ruleError } from '@/utils/columnRules'

export interface CoerceResult {
  value: CellValue
  error: string | null
}

/**
 * Chuyển chuỗi người dùng nhập thành CellValue theo kiểu cột.
 * - number: rỗng → null; NaN → null + lỗi (chặn lưu giá trị không hợp lệ).
 * - boolean: 'true' / '1' → true, còn lại → false.
 * - date: rỗng → null; parse được → ISO; không parse được → giữ chuỗi gốc + lỗi.
 * - text/khác: giữ nguyên chuỗi.
 */
export function coerceCellValue(
  raw: string,
  dataType?: ColumnHeader['dataType'],
  dateFormat?: string,
): CoerceResult {
  if (dataType === 'number') {
    if (raw.trim() === '') return { value: null, error: null }
    const n = Number(raw)
    if (Number.isNaN(n)) return { value: null, error: 'Không phải số hợp lệ' }
    return { value: n, error: null }
  }
  if (dataType === 'boolean') {
    return { value: raw === 'true' || raw === '1', error: null }
  }
  if (dataType === 'date') {
    if (raw.trim() === '') return { value: null, error: null }
    const iso = parseDateToIso(raw, dateFormat ?? DEFAULT_DATE_FORMAT)
    if (iso) return { value: iso, error: null }
    return { value: raw, error: `Ngày chưa đúng định dạng ${dateFormat ?? DEFAULT_DATE_FORMAT}` }
  }
  return { value: raw, error: null }
}

/** Parse toàn bộ form data thành RowData, kèm danh sách lỗi theo key. */
export function coerceRowData(
  formData: Record<string, string>,
  headers: ColumnHeader[],
  peers?: { rows: RowData[]; skipIndex?: number | null },
): { row: RowData; errors: Record<string, string> } {
  const row: RowData = {}
  const errors: Record<string, string> = {}
  for (const h of headers) {
    const { value, error } = coerceCellValue(formData[h.key] ?? '', h.dataType, h.dateFormat)
    row[h.key] = value
    const broken = error ?? ruleError(h, value, peerValues(peers?.rows ?? [], h.key, peers?.skipIndex))
    if (broken) errors[h.key] = broken
  }
  return { row, errors }
}
