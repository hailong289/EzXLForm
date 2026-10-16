import type { CellValue, ColumnHeader, RowData } from '@/types/excel'

type Token = number | string

export function ruleError(header: ColumnHeader, value: CellValue, peers: CellValue[] = []): string | null {
  if (header.formula) return null
  const empty = value == null || value === ''
  if (header.required && empty) return 'Bắt buộc'
  if (empty) return null
  const text = String(value)
  if (header.pattern) {
    try {
      if (!new RegExp(header.pattern).test(text)) return 'Không đúng mẫu'
    } catch {
      return 'Mẫu regex không hợp lệ'
    }
  }
  if (header.options?.length && !header.options.includes(text)) return 'Không nằm trong danh sách'
  if (header.unique && peers.some((peer) => peer != null && peer !== '' && String(peer) === text)) {
    return 'Đã có giá trị này'
  }
  return null
}

export function peerValues(rows: RowData[], key: string, skipIndex?: number | null): CellValue[] {
  return rows.filter((_, index) => index !== skipIndex).map((row) => row[key])
}

export function evalFormula(
  source: string,
  lookup: (name: string) => number,
): { value: number | null; error: string | null } {
  try {
    const tokens = tokenize(source)
    let index = 0
    const peek = () => tokens[index]
    const eat = () => tokens[index++]
    const parseExpr = (): number => {
      let left = parseTerm()
      while (peek() === '+' || peek() === '-') {
        const op = eat()
        const right = parseTerm()
        left = op === '+' ? left + right : left - right
      }
      return left
    }
    const parseTerm = (): number => {
      let left = parseFactor()
      while (peek() === '*' || peek() === '/') {
        const op = eat()
        const right = parseFactor()
        if (op === '/' && right === 0) throw new Error('chia cho 0')
        left = op === '*' ? left * right : left / right
      }
      return left
    }
    const parseFactor = (): number => {
      const token = eat()
      if (token === '(') {
        const value = parseExpr()
        if (eat() !== ')') throw new Error('thiếu ngoặc')
        return value
      }
      if (token === '-') return -parseFactor()
      if (typeof token === 'number') return token
      if (typeof token === 'string' && !'+-*/()'.includes(token)) return lookup(token)
      throw new Error('công thức không đọc được')
    }
    const value = parseExpr()
    if (index < tokens.length) throw new Error('công thức không đọc được')
    if (!Number.isFinite(value)) throw new Error('công thức không đọc được')
    return { value, error: null }
  } catch (error) {
    return { value: null, error: error instanceof Error ? error.message : 'công thức không đọc được' }
  }
}

export function applyFormulas(row: RowData, headers: ColumnHeader[]): RowData {
  const next = { ...row }
  const formulas = headers.filter((header) => header.formula)
  for (let pass = 0; pass < formulas.length; pass++) {
    for (const header of formulas) {
      const result = evalFormula(header.formula ?? '', (name) => {
        const raw = next[name]
        const numeric = typeof raw === 'number' ? raw : Number(raw)
        if (raw == null || raw === '' || Number.isNaN(numeric)) throw new Error(`thiếu ${name}`)
        return numeric
      })
      if (result.value != null) next[header.key] = result.value
    }
  }
  return next
}

export function formulaError(header: ColumnHeader, row: RowData): string | null {
  if (!header.formula) return null
  return evalFormula(header.formula, (name) => {
    const raw = row[name]
    const numeric = typeof raw === 'number' ? raw : Number(raw)
    if (raw == null || raw === '' || Number.isNaN(numeric)) throw new Error(`thiếu ${name}`)
    return numeric
  }).error
}

function tokenize(source: string): Token[] {
  const tokens: Token[] = []
  const pattern = /\s+|([()+\-*/])|(\d+(?:\.\d+)?)|([A-Za-z_][\w]*)/gy
  let cursor = 0
  while (cursor < source.length) {
    pattern.lastIndex = cursor
    const match = pattern.exec(source)
    if (!match || match.index !== cursor) throw new Error('công thức không đọc được')
    cursor = pattern.lastIndex
    if (match[1]) tokens.push(match[1])
    else if (match[2]) tokens.push(Number(match[2]))
    else if (match[3]) tokens.push(match[3])
  }
  return tokens
}
