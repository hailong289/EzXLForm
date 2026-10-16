import { DATE_FORMATS, DEFAULT_DATE_FORMAT } from './dateFormat.ts'

export type FormatCategory =
  | 'general'
  | 'number'
  | 'currency'
  | 'accounting'
  | 'date'
  | 'time'
  | 'percentage'
  | 'fraction'
  | 'scientific'
  | 'text'
  | 'boolean'
  | 'custom'

export type NegativeStyle = 'minus' | 'parens'

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

export interface FormatDraft {
  category: FormatCategory
  decimals: number
  thousands: boolean
  dateFormat: string
  custom: string
  symbol: string
  negative: NegativeStyle
  fraction: string
  timeFormat: string
}

export const CURRENCY_SYMBOLS = ['₫', '$', '€', '¥', '£'] as const
export const FRACTION_CODES = ['# ?/?', '# ??/??', '# ???/???', '?/?', '??/??'] as const
export const TIME_FORMATS = ['hh:mm', 'hh:mm:ss'] as const

/** Danh sách Kiểu của Excel. Bỏ mã tên tháng và AM/PM vì chưa hiện được. */
export const CUSTOM_FORMATS = [
  'General',
  '0',
  '0.00',
  '#,##0',
  '#,##0.00',
  '#,##0_);(#,##0)',
  '#,##0_);[Red](#,##0)',
  '#,##0.00_);(#,##0.00)',
  '#,##0.00_);[Red](#,##0.00)',
  '$#,##0_);($#,##0)',
  '$#,##0_);[Red]($#,##0)',
  '$#,##0.00_);($#,##0.00)',
  '$#,##0.00_);[Red]($#,##0.00)',
  '"₫"#,##0_);("₫"#,##0)',
  '"₫"#,##0.00_);("₫"#,##0.00)',
  '0%',
  '0.00%',
  '0.00E+00',
  '##0.0E+0',
  '# ?/?',
  '# ??/??',
  'dd/MM/yyyy',
  'dd-MM-yyyy',
  'yyyy-MM-dd',
  'MM/dd/yyyy',
  'dd/MM/yy',
  'hh:mm',
  'hh:mm:ss',
  '@',
] as const

export function categoryOf(header: ColumnFormatInput): FormatCategory {
  if (header.dataType === 'boolean') return 'boolean'
  if (header.dataType === 'date') return 'date'
  const fmt = header.numFmt ?? ''
  if (header.dataType !== 'number' || !fmt) return 'general'
  if (fmt.includes('%')) return 'percentage'
  if (/[0#][0#.,]*[eE][+-]/.test(fmt)) return 'scientific'
  if (fmt.includes('?/')) return 'fraction'
  if (/h/i.test(fmt) && /m/i.test(fmt)) return 'time'
  const sections = splitSections(fmt)
  const symbol = /[₫$€¥£]/.test(fmt)
  if (symbol && sections[2] && !/[0#]/.test(sections[2])) return 'accounting'
  if (symbol) return 'currency'
  return 'number'
}

export function decimalCount(numFmt?: string): number {
  const fraction = (splitSections(numFmt ?? '')[0] ?? '').split('.')[1] ?? ''
  return (fraction.match(/^[0#?]*/)?.[0] ?? '').replace(/[^0]/g, '').length
}

export function readFormatOptions(numFmt?: string) {
  const first = splitSections(numFmt ?? '')[0] ?? ''
  const second = splitSections(numFmt ?? '')[1] ?? ''
  const fraction = FRACTION_CODES.find((code) => first.replace(/\s/g, '') === code.replace(/\s/g, ''))
  const time = TIME_FORMATS.find((code) => first === code)
  return {
    decimals: decimalCount(numFmt),
    thousands: /#,/.test(first),
    negative: (second.includes('(') ? 'parens' : 'minus') as NegativeStyle,
    symbol: first.match(/[₫$€¥£]/)?.[0] ?? '₫',
    fraction: fraction ?? '# ?/?',
    timeFormat: time ?? 'hh:mm',
  }
}

export function numberCode(decimals: number, thousands: boolean): string {
  const places = Math.min(10, Math.max(0, decimals))
  const fraction = places > 0 ? `.${'0'.repeat(places)}` : ''
  return `${thousands ? '#,##0' : '0'}${fraction}`
}

export function percentCode(decimals: number): string {
  return `${numberCode(decimals, false)}%`
}

export function currencyCode(decimals: number, symbol: string, negative: NegativeStyle): string {
  const base = quoteSymbol(symbol) + numberCode(decimals, true)
  return negative === 'parens' ? `${base};(${base})` : base
}

export function accountingCode(decimals: number, symbol: string): string {
  const base = quoteSymbol(symbol) + numberCode(decimals, true)
  return `${base};-${base};${quoteSymbol(symbol)}-`
}

export function scientificCode(decimals: number): string {
  const places = Math.min(10, Math.max(0, decimals))
  return places > 0 ? `0.${'0'.repeat(places)}E+00` : '0E+00'
}

export function formatNumber(value: unknown, numFmt?: string): string {
  const numeric = typeof value === 'number' ? value : Number(value)
  if (Number.isNaN(numeric)) return value == null ? '' : String(value)
  if (!numFmt || numFmt === 'General') return numeric.toLocaleString('vi-VN')
  const sections = splitSections(numFmt)
  const section = numeric < 0 ? (sections[1] || sections[0]) : numeric === 0 && sections[2] ? sections[2] : sections[0]
  return renderSection(Math.abs(numeric), section, numeric < 0 && !sections[1])
}

export function formatFromDraft(draft: FormatDraft): ColumnFormat {
  if (draft.category === 'boolean') return { dataType: 'boolean' }
  if (draft.category === 'date') return { dataType: 'date', dateFormat: draft.dateFormat || DEFAULT_DATE_FORMAT }
  if (draft.category === 'time') return { dataType: 'number', numFmt: draft.timeFormat || 'hh:mm' }
  if (draft.category === 'number') {
    const base = numberCode(draft.decimals, draft.thousands)
    return { dataType: 'number', numFmt: draft.negative === 'parens' ? `${base};(${base})` : base }
  }
  if (draft.category === 'currency') return { dataType: 'number', numFmt: currencyCode(draft.decimals, draft.symbol, draft.negative) }
  if (draft.category === 'accounting') return { dataType: 'number', numFmt: accountingCode(draft.decimals, draft.symbol) }
  if (draft.category === 'percentage') return { dataType: 'number', numFmt: percentCode(draft.decimals) }
  if (draft.category === 'fraction') return { dataType: 'number', numFmt: draft.fraction || '# ?/?' }
  if (draft.category === 'scientific') return { dataType: 'number', numFmt: scientificCode(draft.decimals) }
  if (draft.category === 'text') return { dataType: 'text' }
  if (draft.category === 'custom') {
    const code = draft.custom.trim()
    if (/yyyy|yy|dd/.test(code)) return { dataType: 'date', dateFormat: code || DEFAULT_DATE_FORMAT }
    if (/h/i.test(code) && /m/i.test(code)) return { dataType: 'number', numFmt: code }
    if (/[0#?]/.test(code)) return { dataType: 'number', numFmt: code }
    return { dataType: 'text' }
  }
  return { dataType: 'text' }
}

export const FORMAT_CATEGORIES: FormatCategory[] = [
  'general',
  'number',
  'currency',
  'accounting',
  'date',
  'time',
  'percentage',
  'fraction',
  'scientific',
  'text',
  'boolean',
  'custom',
]

function quoteSymbol(symbol: string): string {
  return symbol ? `"${symbol}"` : ''
}

function splitSections(code: string): string[] {
  const parts: string[] = []
  let current = ''
  let quote = false
  for (const char of code) {
    if (char === '"') quote = !quote
    if (char === ';' && !quote) {
      parts.push(current)
      current = ''
      continue
    }
    current += char
  }
  parts.push(current)
  return parts
}

function renderSection(abs: number, raw: string, forceMinus: boolean): string {
  const section = raw.replace(/\[[^\]]*\]/g, '')
  if (/[0#][0#.,]*[eE][+-]0+/.test(section)) return renderScientific(abs, section, forceMinus)
  if (/\?+\/\?+/.test(section)) return renderFraction(abs, section, forceMinus)
  // shortcut: giờ là phần của một ngày (0,5 = 12:00), không đọc giờ từ cột ngày
  if (/h/i.test(section) && /m/i.test(section) && !/[yd]/i.test(section)) {
    return `${forceMinus ? '-' : ''}${renderTime(abs, section)}`
  }

  const pieces: string[] = []
  let mask = ''
  const flush = () => {
    if (!mask) return
    pieces.push(`\0${mask}`)
    mask = ''
  }
  for (let i = 0; i < section.length; i++) {
    const char = section[i]
    if (char === '"') {
      flush()
      const end = section.indexOf('"', i + 1)
      pieces.push(end === -1 ? section.slice(i + 1) : section.slice(i + 1, end))
      i = end === -1 ? section.length : end
      continue
    }
    if ((char === '\\' || char === '_' || char === '*') && i + 1 < section.length) {
      flush()
      if (char === '\\') pieces.push(section[i + 1])
      i += 1
      continue
    }
    if ('#0.,%'.includes(char)) {
      mask += char
      continue
    }
    flush()
    pieces.push(char)
  }
  flush()
  const maskIndex = pieces.findIndex((piece) => piece.startsWith('\0'))
  if (maskIndex === -1) return pieces.join('')
  pieces[maskIndex] = renderMask(abs, pieces[maskIndex].slice(1))
  const text = pieces.join('')
  return forceMinus ? `-${text}` : text
}

function renderMask(abs: number, mask: string): string {
  let scaled = abs
  if (mask.includes('%')) scaled *= 100
  const body = mask.replace(/%/g, '')
  const fraction = (body.split('.')[1] ?? '').replace(/[^0#]/g, '')
  const min = fraction.replace(/[^0]/g, '').length
  const text = scaled.toLocaleString('vi-VN', {
    minimumFractionDigits: min,
    maximumFractionDigits: Math.max(min, fraction.length),
    useGrouping: body.includes(','),
  })
  return mask.includes('%') ? `${text}%` : text
}

function renderScientific(abs: number, section: string, forceMinus: boolean): string {
  const match = section.match(/[0#]+(?:\.([0#]+))?[eE]([+-])(0+)/)
  const frac = match?.[1] ?? ''
  const min = frac.replace(/[^0]/g, '').length
  let exp = 0
  let mantissa = abs
  if (abs !== 0) {
    exp = Math.floor(Math.log10(abs))
    mantissa = abs / 10 ** exp
  }
  const body = mantissa.toLocaleString('vi-VN', {
    minimumFractionDigits: min,
    maximumFractionDigits: Math.max(min, frac.length),
    useGrouping: false,
  })
  const expSign = exp < 0 ? '-' : match?.[2] === '+' ? '+' : ''
  const expText = String(Math.abs(exp)).padStart(match?.[3].length ?? 2, '0')
  return `${forceMinus ? '-' : ''}${body}E${expSign}${expText}`
}

function renderFraction(abs: number, section: string, forceMinus: boolean): string {
  const match = section.match(/(\?+)\/(\?+)/)
  const maxDen = 10 ** (match?.[2].length ?? 1) - 1
  const withWhole = section.includes('#')
  const whole = withWhole ? Math.trunc(abs) : 0
  const [numerator, denominator] = bestFraction(withWhole ? abs - whole : abs, maxDen)
  const sign = forceMinus ? '-' : ''
  if (numerator === 0) return `${sign}${whole}`
  const fraction = `${numerator}/${denominator}`
  if (!withWhole || whole === 0) return `${sign}${fraction}`
  return `${sign}${whole} ${fraction}`
}

function renderTime(abs: number, pattern: string): string {
  const total = Math.round(abs * 86400)
  const hours = Math.floor(total / 3600) % 24
  const minutes = Math.floor((total % 3600) / 60)
  const seconds = total % 60
  const hh = pattern.toLowerCase().includes('hh') ? String(hours).padStart(2, '0') : String(hours)
  const clock = `${hh}:${String(minutes).padStart(2, '0')}`
  return pattern.toLowerCase().includes('ss') ? `${clock}:${String(seconds).padStart(2, '0')}` : clock
}

function bestFraction(value: number, maxDen: number): [number, number] {
  if (value <= 1e-12) return [0, 1]
  let bestN = 0
  let bestD = 1
  let bestErr = value
  for (let denominator = 1; denominator <= maxDen; denominator++) {
    const numerator = Math.round(value * denominator)
    const err = Math.abs(value - numerator / denominator)
    if (err < bestErr - 1e-12) {
      bestN = numerator
      bestD = denominator
      bestErr = err
      if (err < 1e-10) break
    }
  }
  const divisor = gcd(bestN, bestD)
  return [bestN / divisor, bestD / divisor]
}

function gcd(a: number, b: number): number {
  let x = Math.abs(a)
  let y = Math.abs(b)
  while (y) {
    const next = x % y
    x = y
    y = next
  }
  return x || 1
}

export { DATE_FORMATS, DEFAULT_DATE_FORMAT }
