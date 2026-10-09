import type { CellValue } from '@/types/excel'

export function decodeXml(text: string): string {
  return text
    .replace(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, code: string) => String.fromCodePoint(parseInt(code, 16)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

export function columnFromRef(ref: string): number {
  let value = 0
  for (let index = 0; index < ref.length; index++) {
    const code = ref.charCodeAt(index)
    if (code < 65 || code > 90) break
    value = value * 26 + (code - 64)
  }
  return value - 1
}

const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47])

export function dateStyleIndexes(stylesXml: string): Set<number> {
  const custom = new Map<number, string>()
  for (const match of stylesXml.matchAll(/<numFmt\b([^>]*)\/?>/g)) {
    const id = Number(match[1].match(/numFmtId="(\d+)"/)?.[1] ?? -1)
    const code = match[1].match(/formatCode="([^"]*)"/)?.[1] ?? ''
    if (id >= 0) custom.set(id, decodeXml(code))
  }

  const block = stylesXml.split(/<cellXfs\b[^>]*>/)[1]?.split('</cellXfs>')[0] ?? ''
  const dates = new Set<number>()
  let index = 0
  for (const match of block.matchAll(/<xf\b([^>]*)\/?>/g)) {
    const formatId = Number(match[1].match(/numFmtId="(\d+)"/)?.[1] ?? 0)
    const code = custom.get(formatId) ?? ''
    const looksDate = /[dy]/i.test(code) && !/^0/.test(code)
    if (BUILTIN_DATE_FORMATS.has(formatId) || looksDate) dates.add(index)
    index += 1
  }
  return dates
}

export function excelSerialToIso(serial: number): string {
  const utc = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000
  const date = new Date(utc)
  const month = String(date.getUTCMonth() + 1).padStart(2, '0')
  const day = String(date.getUTCDate()).padStart(2, '0')
  return `${date.getUTCFullYear()}-${month}-${day}`
}

export interface ParsedSheetRow {
  excelIndex: number
  cells: CellValue[]
}

/** Đọc XML worksheet theo từng mẩu, không giữ cả sheet trong bộ nhớ */
export class WorksheetXmlParser {
  private buffer = ''
  private inCell = false
  private capture: 'v' | 't' | null = null
  private cellType = ''
  private cellStyle = ''
  private cellRef = ''
  private text = ''
  private valueText = ''
  private inline = ''
  private cells: CellValue[] = []
  private rowExcelIndex = -1
  private nextRow = 0

  constructor(
    private sharedStrings: string[],
    private dateStyles: Set<number>,
  ) {}

  push(chunk: string, emit: (row: ParsedSheetRow) => void): void {
    this.buffer += chunk
    this.drain(emit, false)
  }

  finish(emit: (row: ParsedSheetRow) => void): void {
    this.drain(emit, true)
  }

  private drain(emit: (row: ParsedSheetRow) => void, end: boolean): void {
    while (this.buffer.length > 0) {
      if (this.capture) {
        const close = this.capture === 'v' ? '</v>' : '</t>'
        const found = this.buffer.indexOf(close)
        if (found < 0) {
          const keep = end ? 0 : close.length - 1
          if (this.buffer.length <= keep) return
          this.text += this.buffer.slice(0, this.buffer.length - keep)
          this.buffer = keep ? this.buffer.slice(-keep) : ''
          return
        }
        this.text += this.buffer.slice(0, found)
        this.buffer = this.buffer.slice(found + close.length)
        if (this.capture === 't') this.inline += decodeXml(this.text)
        else this.valueText = decodeXml(this.text)
        this.capture = null
        this.text = ''
        continue
      }

      const start = this.buffer.indexOf('<')
      if (start < 0) {
        this.buffer = ''
        return
      }
      if (start > 0) this.buffer = this.buffer.slice(start)
      const endTag = this.buffer.indexOf('>')
      if (endTag < 0) return
      const tag = this.buffer.slice(1, endTag)
      this.buffer = this.buffer.slice(endTag + 1)
      this.onTag(tag, emit)
    }
  }

  private onTag(tag: string, emit: (row: ParsedSheetRow) => void): void {
    const name = tag.replace(/^\//, '').split(/\s/)[0].replace(/\/$/, '')
    const closing = tag.startsWith('/')
    if (!closing && name === 'row') {
      this.cells = []
      const rowNumber = tag.match(/\br="(\d+)"/)?.[1]
      this.rowExcelIndex = rowNumber ? Number(rowNumber) - 1 : this.nextRow
      this.nextRow = this.rowExcelIndex + 1
      return
    }
    if (closing && name === 'row') {
      emit({ excelIndex: this.rowExcelIndex, cells: this.cells })
      this.cells = []
      return
    }
    if (!closing && name === 'c') {
      this.inCell = true
      this.cellType = tag.match(/\bt="([^"]+)"/)?.[1] ?? ''
      this.cellStyle = tag.match(/\bs="(\d+)"/)?.[1] ?? ''
      this.cellRef = tag.match(/\br="([^"]+)"/)?.[1] ?? ''
      this.inline = ''
      this.text = ''
      this.valueText = ''
      return
    }
    if (closing && name === 'c') {
      this.commitCell()
      this.inCell = false
      return
    }
    if (!closing && (name === 'v' || name === 't') && this.inCell) {
      this.capture = name
      this.text = ''
    }
  }

  private commitCell(): void {
    const column = this.cellRef ? columnFromRef(this.cellRef) : this.cells.length
    if (column < 0) return
    while (this.cells.length < column) this.cells.push(null)
    this.cells[column] = this.cellValue()
  }

  private cellValue(): CellValue {
    const body = this.valueText.trim()
    if (this.cellType === 's') {
      const text = this.sharedStrings[Number(body)] ?? ''
      return text === '' ? null : text
    }
    if (this.cellType === 'inlineStr') {
      const text = this.inline.trim()
      return text === '' ? null : text
    }
    if (this.cellType === 'b') return body === '1'
    if (body === '') return this.inline.trim() || null
    const numeric = Number(body)
    if (!Number.isNaN(numeric) && body !== '') {
      const style = Number(this.cellStyle)
      if (this.dateStyles.has(style)) return excelSerialToIso(numeric)
      return numeric
    }
    return body
  }
}

export class SharedStringParser {
  readonly strings: string[] = []
  private buffer = ''
  private capture = false
  private parts: string[] = []
  private text = ''

  push(chunk: string): void {
    this.buffer += chunk
    this.drain(false)
  }

  finish(): void {
    this.drain(true)
  }

  private drain(end: boolean): void {
    while (this.buffer.length > 0) {
      if (this.capture) {
        const found = this.buffer.indexOf('</t>')
        if (found < 0) {
          const keep = end ? 0 : 3
          if (this.buffer.length <= keep) return
          this.text += this.buffer.slice(0, this.buffer.length - keep)
          this.buffer = keep ? this.buffer.slice(-keep) : ''
          return
        }
        this.text += this.buffer.slice(0, found)
        this.parts.push(decodeXml(this.text))
        this.text = ''
        this.capture = false
        this.buffer = this.buffer.slice(found + 4)
        continue
      }
      const start = this.buffer.indexOf('<')
      if (start < 0) {
        this.buffer = ''
        return
      }
      if (start > 0) this.buffer = this.buffer.slice(start)
      const endTag = this.buffer.indexOf('>')
      if (endTag < 0) return
      const tag = this.buffer.slice(1, endTag)
      this.buffer = this.buffer.slice(endTag + 1)
      const name = tag.replace(/^\//, '').split(/\s/)[0]
      if (!tag.startsWith('/') && name === 'si') this.parts = []
      if (tag.startsWith('/') && name === 'si') this.strings.push(this.parts.join(''))
      if (!tag.startsWith('/') && name === 't') {
        this.capture = true
        this.text = ''
      }
    }
  }
}
