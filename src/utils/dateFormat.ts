export const DATE_FORMATS = ['dd/MM/yyyy', 'dd-MM-yyyy', 'yyyy-MM-dd', 'MM/dd/yyyy', 'dd/MM/yy'] as const

export const DEFAULT_DATE_FORMAT = 'dd/MM/yyyy'

type DateParts = { year: number; month: number; day: number }

function pad(value: number, length = 2): string {
  return String(value).padStart(length, '0')
}

function validParts(year: number, month: number, day: number): DateParts | null {
  if (!year || month < 1 || month > 12 || day < 1 || day > 31) return null
  const date = new Date(year, month - 1, day)
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null
  return { year, month, day }
}

function toIso(parts: DateParts): string {
  return `${pad(parts.year, 4)}-${pad(parts.month)}-${pad(parts.day)}`
}

function applyPattern(parts: DateParts, pattern: string): string {
  return pattern
    .replace('yyyy', pad(parts.year, 4))
    .replace('yy', pad(parts.year % 100))
    .replace('MM', pad(parts.month))
    .replace('dd', pad(parts.day))
}

function parseWithPattern(text: string, pattern: string): string | null {
  const tokens = pattern.match(/yyyy|yy|MM|dd/g)
  const numbers = text.trim().match(/\d+/g)
  if (!tokens || !numbers || tokens.length !== numbers.length) return null
  let year = 0
  let month = 0
  let day = 0
  tokens.forEach((token, index) => {
    const value = Number(numbers[index])
    if (token === 'yyyy') year = value
    else if (token === 'yy') year = value < 70 ? 2000 + value : 1900 + value
    else if (token === 'MM') month = value
    else day = value
  })
  const parts = validParts(year, month, day)
  return parts ? toIso(parts) : null
}

/** Đọc ngày từ ISO, Date, hoặc chuỗi theo format cột. Trả về yyyy-MM-dd. */
export function parseDateToIso(value: unknown, pattern = DEFAULT_DATE_FORMAT): string | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return toIso({ year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() })
  }
  const text = String(value ?? '').trim()
  if (!text) return null
  const patterns = [pattern, 'yyyy-MM-dd', ...DATE_FORMATS.filter((item) => item !== pattern)]
  for (const item of patterns) {
    const iso = parseWithPattern(text, item)
    if (iso) return iso
  }
  return null
}

/** Hiện ngày theo format. Chuỗi chưa phải ngày thì giữ nguyên để đang gõ không bị nhảy. */
export function formatDateValue(value: unknown, pattern = DEFAULT_DATE_FORMAT): string {
  if (value == null || value === '') return ''
  const iso = parseDateToIso(value, pattern)
  if (!iso) return String(value)
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso)
  if (!match) return String(value)
  return applyPattern(
    { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) },
    pattern,
  )
}
