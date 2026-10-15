import { DATE_FORMATS, DEFAULT_DATE_FORMAT } from './dateFormat.ts'

export type FormatCategory = 'general' | 'number' | 'date' | 'percentage' | 'text' | 'boolean' | 'custom'

export interface ColumnFormatInput {
  dataType?: 'text' | 'number' | 'date' | 'boolean'
  dateFormat?: string
  numFmt?: string
}

export interface ColumnFormat {
  dataType: 'text' | 'number' | 'date' | 'boolean'
  dateFormat?: string
  numFmt?: string
}

export function categoryOf(header: ColumnFormatInput): FormatCategory {
  if (header.dataType === 'boolean') return 'boolean'
  if (header.dataType === 'date') return 'date'
  if (header.dataType === 'number' && header.numFmt?.includes('%')) return 'percentage'
  if (header.dataType === 'number') return 'number'
  return 'general'
}

export function decimalCount(numFmt?: string): number {
  const fraction = (numFmt ?? '').split('.')[1] ?? ''
  return fraction.replace('%', '').replace(/[^0]/g, '').length
}

export function numberCode(decimals: number, thousands: boolean): string {
  const places = Math.min(10, Math.max(0, decimals))
  const fraction = places > 0 ? `.${'0'.repeat(places)}` : ''
  return `${thousands ? '#,##0' : '0'}${fraction}`
}

export function percentCode(decimals: number): string {
  return `${numberCode(decimals, false)}%`
}

export function formatNumber(value: unknown, numFmt?: string): string {
  const numeric = typeof value === 'number' ? value : Number(value)
  if (Number.isNaN(numeric)) return value == null ? '' : String(value)
  if (!numFmt) return numeric.toLocaleString('vi-VN')
  const percent = numFmt.includes('%')
  const shown = percent ? numeric * 100 : numeric
  const text = shown.toLocaleString('vi-VN', {
    minimumFractionDigits: decimalCount(numFmt),
    maximumFractionDigits: decimalCount(numFmt),
    useGrouping: numFmt.includes(','),
  })
  return percent ? `${text}%` : text
}

export function formatFromDraft(draft: {
  category: FormatCategory
  decimals: number
  thousands: boolean
  dateFormat: string
  custom: string
}): ColumnFormat {
  if (draft.category === 'boolean') return { dataType: 'boolean' }
  if (draft.category === 'date') return { dataType: 'date', dateFormat: draft.dateFormat || DEFAULT_DATE_FORMAT }
  if (draft.category === 'number') return { dataType: 'number', numFmt: numberCode(draft.decimals, draft.thousands) }
  if (draft.category === 'percentage') return { dataType: 'number', numFmt: percentCode(draft.decimals) }
  if (draft.category === 'text') return { dataType: 'text' }
  if (draft.category === 'custom') {
    const code = draft.custom.trim()
    if (/yyyy|yy|dd/.test(code)) return { dataType: 'date', dateFormat: code || DEFAULT_DATE_FORMAT }
    if (/[0#]/.test(code)) return { dataType: 'number', numFmt: code }
    return { dataType: 'text' }
  }
  return { dataType: 'text' }
}

export const FORMAT_CATEGORIES: FormatCategory[] = [
  'general',
  'number',
  'date',
  'percentage',
  'text',
  'boolean',
  'custom',
]

export { DATE_FORMATS, DEFAULT_DATE_FORMAT }
