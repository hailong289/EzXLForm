import type {
  CellValue,
  ColumnHeader,
  ColumnSumStat,
  ImportOptions,
  RowData,
  SheetData,
  WorkbookData,
} from '@/types/excel'
import { discardSheet, saveSheetMeta, scanLogicalRows, writeRowBatch } from '@/composables/rowPageStore'
import { DEFAULT_DATE_FORMAT, formatDateValue } from '@/utils/dateFormat'
import { entryStream, listZipEntries, readEntryText, type ZipEntry } from '@/utils/zipRead'
import { StoredZipWriter } from '@/utils/zipWrite'
import {
  SharedStringParser,
  WorksheetXmlParser,
  dateStyleIndexes,
  decodeXml,
} from '@/utils/xlsxSheetXml'

export const LARGE_FILE_BYTES = 4 * 1024 * 1024
const WRITE_BATCH = 2000

export class ImportStopped extends Error {
  constructor() {
    super('Đã dừng nhập file')
    this.name = 'ImportStopped'
  }
}

export function isImportStopped(error: unknown): boolean {
  return error instanceof ImportStopped
}

export interface LargeImportWatch {
  checkpoint: () => Promise<boolean>
  onRatio: (ratio: number) => void
}

async function gate(watch?: LargeImportWatch): Promise<void> {
  if (!watch) return
  const keepGoing = await watch.checkpoint()
  if (!keepGoing) throw new ImportStopped()
}

type SaveSink = {
  write: (chunk: Uint8Array) => Promise<void>
  close: () => Promise<void>
}

export function isUserAbort(error: unknown): boolean {
  return error instanceof DOMException && error.name === 'AbortError'
}

async function openSaveStream(
  filename: string,
  description: string,
  mime: string,
  extension: string,
): Promise<SaveSink> {
  const picker = (
    window as Window & {
      showSaveFilePicker?: (options: {
        suggestedName?: string
        types?: { description?: string; accept: Record<string, string[]> }[]
      }) => Promise<{ createWritable: () => Promise<SaveSink> }>
    }
  ).showSaveFilePicker
  if (!picker) {
    throw new Error('File lớn cần Chrome hoặc Edge để ghi thẳng ra đĩa, không giữ cả file trong bộ nhớ.')
  }
  const handle = await picker({
    suggestedName: filename,
    types: [{ description, accept: { [mime]: [extension] } }],
  })
  return handle.createWritable()
}

function zipPath(target: string): string {
  const clean = decodeXml(target).replace(/\\/g, '/').replace(/^\//, '')
  if (clean.startsWith('xl/')) return clean
  return `xl/${clean.replace(/^\.\//, '')}`
}

function findEntry(entries: ZipEntry[], path: string): ZipEntry | undefined {
  const want = path.replace(/\\/g, '/').toLowerCase()
  return entries.find((entry) => entry.name.replace(/\\/g, '/').toLowerCase() === want)
}

async function readTextChunks(
  file: File,
  entry: ZipEntry,
  onChunk: (chunk: string) => void,
  onBytes?: (byteLength: number) => void,
  watch?: LargeImportWatch,
): Promise<void> {
  const reader = entryStream(file, entry, onBytes).getReader()
  const decoder = new TextDecoder()
  try {
    while (true) {
      await gate(watch)
      const { done, value } = await reader.read()
      if (done) {
        const tail = decoder.decode()
        if (tail) onChunk(tail)
        return
      }
      const text = decoder.decode(value, { stream: true })
      if (text) onChunk(text)
    }
  } finally {
    await reader.cancel().catch(() => undefined)
  }
}

function workbookSheetList(workbookXml: string, relsXml: string): { name: string; path: string }[] {
  const rels = new Map<string, string>()
  for (const match of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const id = match[1].match(/\bId="([^"]+)"/)?.[1]
    const target = match[1].match(/\bTarget="([^"]+)"/)?.[1]
    const type = match[1].match(/\bType="([^"]+)"/)?.[1] ?? ''
    if (!id || !target || !type.includes('worksheet')) continue
    rels.set(id, zipPath(target))
  }
  const sheets: { name: string; path: string }[] = []
  for (const match of workbookXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const name = decodeXml(match[1].match(/\bname="([^"]*)"/)?.[1] ?? 'Sheet')
    const rid = match[1].match(/\br:id="([^"]+)"/)?.[1]
    const path = rid ? rels.get(rid) : undefined
    if (path) sheets.push({ name, path })
  }
  return sheets
}

interface Tally {
  number: number
  date: number
  bool: number
  total: number
}

function buildHeaders(headerCells: CellValue[], width: number, tallies: Tally[]): ColumnHeader[] {
  const headers: ColumnHeader[] = []
  for (let column = 0; column < width; column++) {
    const raw = headerCells[column]
    const labelText = raw == null ? '' : String(raw).trim()
    const tally = tallies[column] ?? { number: 0, date: 0, bool: 0, total: 0 }
    let dataType: ColumnHeader['dataType'] = 'text'
    if (tally.total > 0 && tally.date > 0 && tally.date >= tally.number && tally.date * 2 >= tally.total) {
      dataType = 'date'
    } else if (tally.total > 0 && tally.bool === tally.total) {
      dataType = 'boolean'
    } else if (tally.total > 0 && tally.number === tally.total) {
      dataType = 'number'
    }
    headers.push({
      key: `col_${column}`,
      label: labelText || `Cột ${column + 1}`,
      originalIndex: column,
      dataType,
      dateFormat: dataType === 'date' ? DEFAULT_DATE_FORMAT : undefined,
    })
  }
  return headers
}

class RowSink {
  readonly tallies: Tally[] = []
  readonly sums: Record<string, ColumnSumStat> = {}
  readonly preview: CellValue[][] = []
  headerCells: CellValue[] = []
  maxCols = 1
  rowCount = 0
  private written = 0
  private batch: RowData[] = []
  private chain: Promise<void> = Promise.resolve()

  constructor(
    private sheetId: string,
    private sheetName: string,
    private headerRowIndex: number,
    private dataStart: number,
    private onProgress: (message: string) => void,
    private watch?: LargeImportWatch,
  ) {}

  accept(excelIndex: number, cells: CellValue[]): void {
    if (cells.length > this.maxCols) this.maxCols = cells.length
    if (excelIndex < 40) this.preview[excelIndex] = cells
    if (excelIndex === this.headerRowIndex) this.headerCells = cells
    if (excelIndex < this.dataStart) return

    const row: RowData = {}
    let hasValue = false
    for (let column = 0; column < cells.length; column++) {
      const value = cells[column]
      if (value == null || value === '') continue
      row[`col_${column}`] = value
      hasValue = true
      const tally = this.tallies[column] ?? { number: 0, date: 0, bool: 0, total: 0 }
      if (tally.total < 400) {
        tally.total += 1
        if (typeof value === 'number') tally.number += 1
        else if (typeof value === 'boolean') tally.bool += 1
        else if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) tally.date += 1
        this.tallies[column] = tally
      }
      if (typeof value === 'number') {
        const key = `col_${column}`
        const stat = this.sums[key] ?? { sum: 0, count: 0 }
        stat.sum += value
        stat.count += 1
        this.sums[key] = stat
      }
    }
    if (!hasValue) return
    this.batch.push(row)
    this.rowCount += 1
    if (this.batch.length >= WRITE_BATCH) {
      const pending = this.batch
      this.batch = []
      const start = this.written
      this.written += pending.length
      this.chain = this.chain.then(async () => {
        await writeRowBatch(this.sheetId, start, pending)
        this.onProgress(`Đang đọc ${this.sheetName}: ${this.written.toLocaleString('vi-VN')} dòng`)
        await gate(this.watch)
      })
    }
  }

  drain(): Promise<void> {
    return this.chain
  }

  async finish(): Promise<SheetData> {
    if (this.batch.length > 0) {
      const pending = this.batch
      this.batch = []
      const start = this.written
      this.written += pending.length
      this.chain = this.chain.then(() => writeRowBatch(this.sheetId, start, pending))
    }
    await this.chain
    const width = Math.max(this.maxCols, this.headerCells.length, 1)
    const headers = buildHeaders(this.headerCells, width, this.tallies)
    const grid: CellValue[][] = []
    const previewLength = Math.min(this.preview.length, 40)
    for (let index = 0; index < previewLength; index++) grid.push(this.preview[index] ?? [])
    let headerRowIndex = this.headerRowIndex
    if (headerRowIndex >= grid.length) {
      grid.unshift(this.headerCells)
      headerRowIndex = 0
    }
    const sheet: SheetData = {
      id: this.sheetId,
      name: this.sheetName,
      headers,
      rows: [],
      paged: true,
      rowCount: this.rowCount,
      columnSums: this.sums,
      headerRowIndex,
      sourceGrid: grid,
    }
    await saveSheetMeta(this.sheetId, { rowCount: this.rowCount, columnSums: this.sums })
    this.onProgress(`Đã nạp ${this.sheetName}: ${this.rowCount.toLocaleString('vi-VN')} dòng`)
    return sheet
  }
}

export async function importLargeWorkbook(
  file: File,
  options: Partial<ImportOptions> | undefined,
  onProgress: (message: string) => void,
  watch?: LargeImportWatch,
): Promise<WorkbookData> {
  const extension = file.name.split('.').pop()?.toLowerCase()
  if (extension === 'csv') return importLargeCsv(file, options, onProgress, watch)
  return importLargeXlsx(file, options, onProgress, watch)
}

async function importLargeXlsx(
  file: File,
  options: Partial<ImportOptions> | undefined,
  onProgress: (message: string) => void,
  watch?: LargeImportWatch,
): Promise<WorkbookData> {
  let bytesRead = 0
  const note = (message: string) => {
    onProgress(message)
    watch?.onRatio(file.size > 0 ? bytesRead / file.size : 0)
  }
  const countBytes = (byteLength: number) => {
    bytesRead += byteLength
    watch?.onRatio(file.size > 0 ? bytesRead / file.size : 0)
  }

  note('Đang đọc cấu trúc file...')
  const entries = await listZipEntries(file)
  const workbookEntry = findEntry(entries, 'xl/workbook.xml')
  const relsEntry = findEntry(entries, 'xl/_rels/workbook.xml.rels')
  if (!workbookEntry || !relsEntry) throw new Error('File Excel thiếu workbook.')

  const [workbookXml, relsXml] = await Promise.all([
    readEntryText(file, workbookEntry),
    readEntryText(file, relsEntry),
  ])
  countBytes(workbookEntry.compressedSize + relsEntry.compressedSize)
  const listed = workbookSheetList(workbookXml, relsXml)
  if (listed.length === 0) throw new Error('File không chứa sheet dữ liệu nào.')

  const stylesEntry = findEntry(entries, 'xl/styles.xml')
  const dateStyles = stylesEntry ? dateStyleIndexes(await readEntryText(file, stylesEntry)) : new Set<number>()
  if (stylesEntry) countBytes(stylesEntry.compressedSize)

  note('Đang đọc bảng chuỗi dùng chung...')
  const shared = new SharedStringParser()
  const sharedEntry = findEntry(entries, 'xl/sharedStrings.xml')
  if (sharedEntry) {
    await readTextChunks(file, sharedEntry, (chunk) => shared.push(chunk), countBytes, watch)
    shared.finish()
  }

  const headerRowIndex = options?.headerRowIndex ?? 0
  const dataStart = headerRowIndex + 1 + (options?.skipRows ?? 0)
  const createdAt = Date.now()
  const sheets: SheetData[] = []
  const createdIds: string[] = []

  try {
  for (let index = 0; index < listed.length; index++) {
    const item = listed[index]
    const entry = findEntry(entries, item.path)
    if (!entry) continue
    await gate(watch)
    const sheetId = `sheet_${index}_${createdAt}`
    createdIds.push(sheetId)
    const sink = new RowSink(sheetId, item.name, headerRowIndex, dataStart, note, watch)
    const parser = new WorksheetXmlParser(shared.strings, dateStyles)
    const reader = entryStream(file, entry, countBytes).getReader()
    const decoder = new TextDecoder()
    try {
      while (true) {
        await gate(watch)
        const { done, value } = await reader.read()
        if (done) {
          const tail = decoder.decode()
          if (tail) parser.push(tail, (row) => sink.accept(row.excelIndex, row.cells))
          break
        }
        const text = decoder.decode(value, { stream: true })
        if (text) parser.push(text, (row) => sink.accept(row.excelIndex, row.cells))
        await sink.drain()
      }
      parser.finish((row) => sink.accept(row.excelIndex, row.cells))
      sheets.push(await sink.finish())
    } finally {
      await reader.cancel().catch(() => undefined)
    }
  }
  } catch (error) {
    await Promise.all(createdIds.map((id) => discardSheet(id)))
    throw error
  }

  if (sheets.length === 0) throw new Error('File không chứa sheet dữ liệu nào.')
  return {
    meta: {
      fileName: file.name,
      fileSize: file.size,
      lastModified: new Date(file.lastModified),
      sheetCount: sheets.length,
    },
    sheets,
    activeSheetId: sheets[0].id,
  }
}

class CsvParser {
  private row: string[] = []
  private cell = ''
  private quoted = false

  push(chunk: string, emit: (cells: string[]) => void): void {
    for (let index = 0; index < chunk.length; index++) {
      const char = chunk[index]
      if (this.quoted) {
        if (char === '"') {
          if (chunk[index + 1] === '"') {
            this.cell += '"'
            index += 1
          } else {
            this.quoted = false
          }
          continue
        }
        this.cell += char
        continue
      }
      if (char === '"') {
        this.quoted = true
        continue
      }
      if (char === ',') {
        this.row.push(this.cell)
        this.cell = ''
        continue
      }
      if (char === '\n') {
        this.row.push(this.cell)
        this.cell = ''
        emit(this.row)
        this.row = []
        continue
      }
      if (char === '\r') continue
      this.cell += char
    }
  }

  finish(emit: (cells: string[]) => void): void {
    if (this.cell.length > 0 || this.row.length > 0) {
      this.row.push(this.cell)
      emit(this.row)
      this.row = []
      this.cell = ''
    }
  }
}

function csvCell(value: string): CellValue {
  const trimmed = value.trim()
  if (trimmed === '') return null
  if (trimmed === 'true' || trimmed === 'TRUE') return true
  if (trimmed === 'false' || trimmed === 'FALSE') return false
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) return Number(trimmed)
  return value
}

async function importLargeCsv(
  file: File,
  options: Partial<ImportOptions> | undefined,
  onProgress: (message: string) => void,
  watch?: LargeImportWatch,
): Promise<WorkbookData> {
  const headerRowIndex = options?.headerRowIndex ?? 0
  const dataStart = headerRowIndex + 1 + (options?.skipRows ?? 0)
  const sheetId = `sheet_0_${Date.now()}`
  const sheetName = file.name.replace(/\.csv$/i, '') || 'CSV'
  let bytesRead = 0
  const note = (message: string) => {
    onProgress(message)
    watch?.onRatio(file.size > 0 ? bytesRead / file.size : 0)
  }
  const sink = new RowSink(sheetId, sheetName, headerRowIndex, dataStart, note, watch)
  const parser = new CsvParser()
  let excelIndex = 0
  const reader = file.stream().getReader()
  const decoder = new TextDecoder()
  const take = (cells: string[]) => {
    sink.accept(excelIndex, cells.map(csvCell))
    excelIndex += 1
  }
  try {
    while (true) {
      await gate(watch)
      const { done, value } = await reader.read()
      if (value) {
        bytesRead += value.byteLength
        watch?.onRatio(file.size > 0 ? bytesRead / file.size : 0)
      }
      if (done) {
        const tail = decoder.decode()
        if (tail) parser.push(tail, take)
        break
      }
      const text = decoder.decode(value, { stream: true })
      if (text) parser.push(text, take)
    }
    parser.finish(take)
    const sheet = await sink.finish()
    return {
      meta: {
        fileName: file.name,
        fileSize: file.size,
        lastModified: new Date(file.lastModified),
        sheetCount: 1,
      },
      sheets: [sheet],
      activeSheetId: sheet.id,
    }
  } catch (error) {
    await discardSheet(sheetId)
    throw error
  } finally {
    await reader.cancel().catch(() => undefined)
  }
}

function xmlEscape(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function columnLetter(index: number): string {
  let current = index + 1
  let letters = ''
  while (current > 0) {
    const remainder = (current - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    current = Math.floor((current - 1) / 26)
  }
  return letters
}

function cellXml(column: number, excelRow: number, value: CellValue, header: ColumnHeader): string {
  if (value == null || value === '') return ''
  const ref = `${columnLetter(column)}${excelRow}`
  if (typeof value === 'number') return `<c r="${ref}"><v>${value}</v></c>`
  if (typeof value === 'boolean') return `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`
  const shown = header.dataType === 'date' ? formatDateValue(value, header.dateFormat) : String(value)
  if (!shown) return ''
  return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(shown)}</t></is></c>`
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>`

async function writePagedXlsxToSink(
  workbook: WorkbookData,
  sink: SaveSink,
  onProgress: (message: string) => void,
): Promise<void> {
  const zip = new StoredZipWriter((chunk) => sink.write(chunk))
  const overrides = workbook.sheets
    .map(
      (_, index) =>
        `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join('')
  await zip.addText(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${overrides}</Types>`,
  )
  await zip.addText(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  const sheetTags = workbook.sheets
    .map((sheet, index) => `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join('')
  await zip.addText(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheetTags}</sheets></workbook>`,
  )
  const relTags = workbook.sheets
    .map(
      (_, index) =>
        `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
    )
    .join('')
  await zip.addText(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relTags}<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  )
  await zip.addText('xl/styles.xml', STYLES_XML)

  for (let index = 0; index < workbook.sheets.length; index++) {
    const sheet = workbook.sheets[index]
    onProgress(`Đang ghi ${sheet.name}...`)
    await zip.streamText(`xl/worksheets/sheet${index + 1}.xml`, async (push) => {
      await push(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>`,
      )
      const headerXml = sheet.headers
        .map((header, column) => cellXml(column, 1, header.label, header))
        .join('')
      await push(`<row r="1">${headerXml}</row>`)
      let line = ''
      let pendingRows = 0
      await scanLogicalRows(sheet.id, sheet.rowCount ?? 0, async (rowIndex, row) => {
        const excelRow = rowIndex + 2
        const cells = sheet.headers.map((header, column) => cellXml(column, excelRow, row[header.key] ?? null, header)).join('')
        line += `<row r="${excelRow}">${cells}</row>`
        pendingRows += 1
        if (pendingRows >= 200) {
          await push(line)
          line = ''
          pendingRows = 0
        }
        if (rowIndex > 0 && rowIndex % 20000 === 0) {
          onProgress(`Đang ghi ${sheet.name}: ${rowIndex.toLocaleString('vi-VN')} dòng`)
        }
        return true
      })
      if (line) await push(line)
      await push(`</sheetData></worksheet>`)
    })
  }

  await zip.close()
  await sink.close()
}

async function writePagedXlsx(workbook: WorkbookData, filename: string, onProgress: (message: string) => void): Promise<void> {
  const sink = await openSaveStream(
    filename,
    'Excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xlsx',
  )
  await writePagedXlsxToSink(workbook, sink, onProgress)
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

/** Gom file lớn thành blob để tải lên cloud. File rất lớn có thể vượt bộ nhớ trình duyệt. */
export async function buildPagedXlsxBlob(
  workbook: WorkbookData,
  onProgress: (message: string) => void,
): Promise<Blob> {
  const chunks: BlobPart[] = []
  const sink: SaveSink = {
    write: async (chunk) => {
      chunks.push(chunk)
    },
    close: async () => {},
  }
  await writePagedXlsxToSink(workbook, sink, onProgress)
  return new Blob(chunks, { type: XLSX_MIME })
}

async function writePagedCsv(sheet: SheetData, filename: string, onProgress: (message: string) => void): Promise<void> {
  const sink = await openSaveStream(filename, 'CSV', 'text/csv', '.csv')
  const encoder = new TextEncoder()
  const header = sheet.headers.map((item) => `"${item.label.replace(/"/g, '""')}"`).join(',')
  await sink.write(encoder.encode(`\uFEFF${header}\r\n`))
  let block = ''
  let pending = 0
  await scanLogicalRows(sheet.id, sheet.rowCount ?? 0, async (rowIndex, row) => {
    const line = sheet.headers
      .map((item) => {
        const value = item.dataType === 'date' ? formatDateValue(row[item.key], item.dateFormat) : row[item.key] ?? ''
        return `"${String(value).replace(/"/g, '""')}"`
      })
      .join(',')
    block += `${line}\r\n`
    pending += 1
    if (pending >= 200) {
      await sink.write(encoder.encode(block))
      block = ''
      pending = 0
    }
    if (rowIndex > 0 && rowIndex % 20000 === 0) {
      onProgress(`Đang ghi CSV: ${rowIndex.toLocaleString('vi-VN')} dòng`)
    }
    return true
  })
  if (block) await sink.write(encoder.encode(block))
  await sink.close()
}

export async function exportLargeWorkbook(
  workbook: WorkbookData,
  customFileName: string | undefined,
  onProgress: (message: string) => void,
): Promise<void> {
  const base = customFileName || workbook.meta.fileName || 'exported.xlsx'
  const filename = base.endsWith('.xlsx') ? base : `${base}.xlsx`
  await writePagedXlsx(workbook, filename, onProgress)
}

export async function exportLargeCsv(
  sheet: SheetData,
  customFileName: string | undefined,
  onProgress: (message: string) => void,
): Promise<void> {
  const base = customFileName || `${sheet.name || 'sheet'}.csv`
  const filename = base.endsWith('.csv') ? base : `${base}.csv`
  await writePagedCsv(sheet, filename, onProgress)
}
