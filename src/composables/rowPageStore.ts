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
  fromIndex = 0,
): Promise<void> {
  const batchSize = 2000
  for (let start = Math.max(0, fromIndex); start < rowCount; start += batchSize) {
    const count = Math.min(batchSize, rowCount - start)
    const rows = await readRange(sheetId, start, count)
    for (let offset = 0; offset < rows.length; offset++) {
      const keep = await visit(start + offset, rows[offset])
      if (keep === false) return
    }
    await new Promise((resolve) => setTimeout(resolve, 0))
  }
}

interface SheetOrder {
  order: number[] | null
  nextPhysical: number
}

const sheetOrders = new Map<string, SheetOrder>()

function orderKey(sheetId: string): string {
  return `order:${sheetId}`
}

export function clearSheetOrder(sheetId: string): void {
  sheetOrders.delete(sheetId)
}

export function clearAllSheetOrders(): void {
  sheetOrders.clear()
}

export function hasSheetOrder(sheetId: string): boolean {
  return Boolean(sheetOrders.get(sheetId)?.order)
}

export function physicalOf(sheetId: string, logical: number): number {
  const state = sheetOrders.get(sheetId)
  if (!state?.order) return logical
  return state.order[logical] ?? logical
}

/** Chèn một khóa vật lý vào thứ tự hiển thị. Không dịch các dòng đã lưu. */
export function placeInsertedRow(sheetId: string, at: number, count: number): number {
  const state = sheetOrders.get(sheetId) ?? { order: null, nextPhysical: count }
  const physical = state.order ? state.nextPhysical : count
  state.nextPhysical = physical + 1
  if (at >= count && !state.order) {
    sheetOrders.set(sheetId, state)
    return physical
  }
  if (!state.order) {
    const order = new Array<number>(count + 1)
    for (let index = 0; index < at; index++) order[index] = index
    order[at] = physical
    for (let index = at; index < count; index++) order[index + 1] = index
    state.order = order
  } else if (at >= state.order.length) {
    state.order.push(physical)
  } else {
    state.order.splice(at, 0, physical)
  }
  sheetOrders.set(sheetId, state)
  return physical
}

export function moveLogicalRow(sheetId: string, from: number, to: number, count: number): boolean {
  if (from === to || from < 0 || to < 0 || from >= count || to >= count) return false
  const state = sheetOrders.get(sheetId) ?? { order: null, nextPhysical: count }
  if (!state.order) {
    state.order = Array.from({ length: count }, (_, index) => index)
  }
  const [moved] = state.order.splice(from, 1)
  state.order.splice(to, 0, moved)
  sheetOrders.set(sheetId, state)
  return true
}

export function removeLogicalRow(sheetId: string, logical: number): number | null {
  const state = sheetOrders.get(sheetId)
  if (!state?.order) return null
  if (logical < 0 || logical >= state.order.length) return null
  const physical = state.order[logical]
  state.order.splice(logical, 1)
  if (state.order.length === 0) state.order = null
  return physical
}

export async function readLogicalRange(sheetId: string, start: number, count: number): Promise<RowData[]> {
  const order = sheetOrders.get(sheetId)?.order
  if (!order || count <= 0) return readRange(sheetId, start, count)
  const rows: RowData[] = []
  let offset = 0
  while (offset < count) {
    const physical = order[start + offset] ?? start + offset
    let run = 1
    while (offset + run < count && order[start + offset + run] === physical + run) run += 1
    const chunk = await readRange(sheetId, physical, run)
    for (let index = 0; index < run; index++) rows.push(chunk[index] ?? {})
    offset += run
  }
  return rows
}

export async function scanLogicalRows(
  sheetId: string,
  rowCount: number,
  visit: (index: number, row: RowData) => boolean | Promise<boolean>,
  fromIndex = 0,
): Promise<void> {
  if (!sheetOrders.get(sheetId)?.order) {
    await scanRows(sheetId, rowCount, visit, fromIndex)
    return
  }
  const batchSize = 2000
  for (let start = Math.max(0, fromIndex); start < rowCount; start += batchSize) {
    const count = Math.min(batchSize, rowCount - start)
    const rows = await readLogicalRange(sheetId, start, count)
    for (let offset = 0; offset < rows.length; offset++) {
      const keep = await visit(start + offset, rows[offset])
      if (keep === false) return
    }
    await new Promise((resolve) => setTimeout(resolve, 0))
  }
}

export function sheetOrderLength(sheetId: string): number {
  return sheetOrders.get(sheetId)?.order?.length ?? 0
}

export async function persistSheetOrder(sheetId: string): Promise<void> {
  const state = sheetOrders.get(sheetId)
  if (!state?.order) return
  const packed = Uint32Array.from(state.order)
  const db = await openDatabase()
  const tx = db.transaction('meta', 'readwrite')
  tx.objectStore('meta').put(
    { nextPhysical: state.nextPhysical, order: packed },
    orderKey(sheetId),
  )
  await transactionDone(tx)
}

function unpackOrder(raw: unknown): number[] | null {
  if (raw instanceof Uint32Array) return Array.from(raw)
  if (raw instanceof ArrayBuffer) return Array.from(new Uint32Array(raw))
  if (Array.isArray(raw)) return raw.map((item) => Number(item))
  return null
}

export async function restoreSheetOrders(sheetIds: string[]): Promise<boolean> {
  let restored = false
  for (const sheetId of sheetIds) {
    const db = await openDatabase()
    const tx = db.transaction('meta', 'readonly')
    const request = tx.objectStore('meta').get(orderKey(sheetId))
    const value = await new Promise<{ order?: unknown; nextPhysical?: number } | undefined>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result as { order?: unknown; nextPhysical?: number } | undefined)
      request.onerror = () => reject(request.error)
    })
    await transactionDone(tx)
    const order = unpackOrder(value?.order)
    if (!order?.length) continue
    sheetOrders.set(sheetId, {
      order,
      nextPhysical: value?.nextPhysical ?? order.reduce((max, item) => Math.max(max, item), 0) + 1,
    })
    restored = true
  }
  return restored
}

export async function discardSheet(sheetId: string): Promise<void> {
  clearSheetOrder(sheetId)
  const db = await openDatabase()
  const tx = db.transaction(['rows', 'meta'], 'readwrite')
  tx.objectStore('rows').delete(IDBKeyRange.bound([sheetId, 0], [sheetId, Number.MAX_SAFE_INTEGER]))
  tx.objectStore('meta').delete(sheetId)
  tx.objectStore('meta').delete(orderKey(sheetId))
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
