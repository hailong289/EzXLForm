import type { ColumnSumStat, RowData } from '@/types/excel'

const DB_NAME = 'excel_editor_row_pages'
const DB_VERSION = 1

export interface PagedSheetMeta {
  rowCount: number
  columnSums: Record<string, ColumnSumStat>
}

let databasePromise: Promise<IDBDatabase> | null = null

function openDatabase(): Promise<IDBDatabase> {
  if (!databasePromise) {
    databasePromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION)
      request.onupgradeneeded = () => {
        const db = request.result
        if (!db.objectStoreNames.contains('rows')) db.createObjectStore('rows')
        if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta')
      }
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => {
        databasePromise = null
        reject(request.error)
      }
    })
  }
  return databasePromise
}

/** IndexedDB từ chối Vue proxy. Chỉ ghi bản dữ liệu thuần. */
function storable<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}

function transactionDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
    tx.onabort = () => reject(tx.error)
  })
}

export async function writeRowBatch(sheetId: string, start: number, rows: RowData[]): Promise<void> {
  await writeRows(sheetId, rows.map((row, index) => ({ index: start + index, row })))
}

export async function writeRows(sheetId: string, entries: { index: number; row: RowData }[]): Promise<void> {
  if (entries.length === 0) return
  const db = await openDatabase()
  const tx = db.transaction('rows', 'readwrite')
  const store = tx.objectStore('rows')
  for (const entry of entries) store.put(storable(entry.row), [sheetId, entry.index])
  await transactionDone(tx)
}

export async function putRow(sheetId: string, index: number, row: RowData): Promise<void> {
  await writeRowBatch(sheetId, index, [row])
}

export async function getRow(sheetId: string, index: number): Promise<RowData | null> {
  const rows = await readRange(sheetId, index, 1)
  return rows[0] ?? null
}

export async function deleteRowKey(sheetId: string, index: number): Promise<void> {
  const db = await openDatabase()
  const tx = db.transaction('rows', 'readwrite')
  tx.objectStore('rows').delete([sheetId, index])
  await transactionDone(tx)
}

export async function readRange(sheetId: string, start: number, count: number): Promise<RowData[]> {
  if (count <= 0) return []
  const rows: RowData[] = Array.from({ length: count }, () => ({}))
  const db = await openDatabase()
  const tx = db.transaction('rows', 'readonly')
  const request = tx.objectStore('rows').openCursor(IDBKeyRange.bound([sheetId, start], [sheetId, start + count - 1]))
  request.onsuccess = () => {
    const cursor = request.result
    if (!cursor) return
    const index = (cursor.key as [string, number])[1]
    rows[index - start] = cursor.value as RowData
    cursor.continue()
  }
  await transactionDone(tx)
  return rows
}

export async function scanRows(
  sheetId: string,
  rowCount: number,
  visit: (index: number, row: RowData) => boolean | Promise<boolean>,
): Promise<void> {
  const batchSize = 2000
  for (let start = 0; start < rowCount; start += batchSize) {
    const count = Math.min(batchSize, rowCount - start)
    const rows = await readRange(sheetId, start, count)
    for (let offset = 0; offset < rows.length; offset++) {
      const keep = await visit(start + offset, rows[offset])
      if (keep === false) return
    }
    await new Promise((resolve) => setTimeout(resolve, 0))
  }
}

export async function discardSheet(sheetId: string): Promise<void> {
  const db = await openDatabase()
  const tx = db.transaction(['rows', 'meta'], 'readwrite')
  tx.objectStore('rows').delete(IDBKeyRange.bound([sheetId, 0], [sheetId, Number.MAX_SAFE_INTEGER]))
  tx.objectStore('meta').delete(sheetId)
  await transactionDone(tx)
}

export async function saveSheetMeta(sheetId: string, meta: PagedSheetMeta): Promise<void> {
  const db = await openDatabase()
  const tx = db.transaction('meta', 'readwrite')
  tx.objectStore('meta').put(storable(meta), sheetId)
  await transactionDone(tx)
}

export async function loadSheetMeta(sheetId: string): Promise<PagedSheetMeta | null> {
  const db = await openDatabase()
  const tx = db.transaction('meta', 'readonly')
  const request = tx.objectStore('meta').get(sheetId)
  const value = await new Promise<PagedSheetMeta | undefined>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result as PagedSheetMeta | undefined)
    request.onerror = () => reject(request.error)
  })
  await transactionDone(tx)
  return value ?? null
}
