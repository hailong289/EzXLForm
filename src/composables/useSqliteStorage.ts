import { ref } from 'vue'
import initSqlJs from 'sql.js'
import sqlWasmUrl from 'sql.js/dist/sql-wasm-browser.wasm?url'
import type { Database } from 'sql.js'
import { saveAs } from 'file-saver'
import type { WorkbookData, SheetData, ColumnHeader, RowData } from '@/types/excel'
import { sheetRowCount } from '@/utils/sheetGrid'
import {
  applyStoredRowMarks,
  collectRowMarksForSave,
  type StoredRowMark,
} from '@/composables/rowMarks'

const DB_STORE_NAME = 'excel_sqlite_store'
const DB_NAME = 'excel_editor_db'
const DB_KEY = 'current_session_sqlite'
const INDEX_KEY = 'sessions_index'
const ACTIVE_ID_STORAGE_KEY = 'excel_editor_active_session'

export interface FileSession {
  id: string
  fileName: string
  updatedAt: number
  sheetCount: number
  rowCount: number
}

const sessions = ref<FileSession[]>([])
const activeSessionId = ref<string | null>(
  typeof localStorage !== 'undefined' ? localStorage.getItem(ACTIVE_ID_STORAGE_KEY) : null,
)

function sessionKey(id: string) {
  return `session:${id}`
}

function setActiveSessionId(id: string | null) {
  activeSessionId.value = id
  if (typeof localStorage === 'undefined') return
  if (id) localStorage.setItem(ACTIVE_ID_STORAGE_KEY, id)
  else localStorage.removeItem(ACTIVE_ID_STORAGE_KEY)
}

// IndexedDB helpers to store SQLite binary Uint8Array
function openIndexedDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1)
    request.onupgradeneeded = () => {
      const db = request.result
      if (!db.objectStoreNames.contains(DB_STORE_NAME)) {
        db.createObjectStore(DB_STORE_NAME)
      }
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })
}

async function saveBinaryToIndexedDB(key: string, data: Uint8Array): Promise<void> {
  const db = await openIndexedDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE_NAME, 'readwrite')
    const store = tx.objectStore(DB_STORE_NAME)
    const req = store.put(data, key)
    req.onsuccess = () => resolve()
    req.onerror = () => reject(req.error)
  })
}

async function getBinaryFromIndexedDB(key: string): Promise<Uint8Array | null> {
  const db = await openIndexedDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE_NAME, 'readonly')
    const store = tx.objectStore(DB_STORE_NAME)
    const req = store.get(key)
    req.onsuccess = () => resolve((req.result as Uint8Array) || null)
    req.onerror = () => reject(req.error)
  })
}

async function saveJsonToIndexedDB(key: string, value: unknown): Promise<void> {
  const bytes = new TextEncoder().encode(JSON.stringify(value))
  await saveBinaryToIndexedDB(key, bytes)
}

async function loadJsonFromIndexedDB<T>(key: string): Promise<T | null> {
  const bytes = await getBinaryFromIndexedDB(key)
  if (!bytes) return null
  try {
    return JSON.parse(new TextDecoder().decode(bytes)) as T
  } catch {
    return null
  }
}

async function clearBinaryFromIndexedDB(key: string): Promise<void> {
  const db = await openIndexedDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE_NAME, 'readwrite')
    const store = tx.objectStore(DB_STORE_NAME)
    const req = store.delete(key)
    req.onsuccess = () => resolve()
    req.onerror = () => reject(req.error)
  })
}

// Module-level SQLite instance
let sqlInstance: any = null
let currentDb: Database | null = null

export function useSqliteStorage() {
  const isSaving = ref(false)
  const isLoading = ref(false)
  const lastSavedAt = ref<Date | null>(null)
  const isInitialized = ref(false)

  async function getSqlInstance() {
    if (!sqlInstance) {
      sqlInstance = await initSqlJs({
        locateFile: () => sqlWasmUrl,
      })
    }
    return sqlInstance
  }

  function initTables(db: Database) {
    db.run(`
      CREATE TABLE IF NOT EXISTS metadata (
        key TEXT PRIMARY KEY,
        value TEXT
      );

      CREATE TABLE IF NOT EXISTS workbook (
        id INTEGER PRIMARY KEY,
        file_name TEXT,
        file_size INTEGER,
        active_sheet_id TEXT,
        updated_at INTEGER
      );

      CREATE TABLE IF NOT EXISTS sheets (
        id TEXT PRIMARY KEY,
        name TEXT,
        sort_order INTEGER,
        header_row_index INTEGER,
        source_grid_json TEXT
      );

      CREATE TABLE IF NOT EXISTS columns (
        id TEXT PRIMARY KEY,
        sheet_id TEXT,
        col_key TEXT,
        label TEXT,
        data_type TEXT,
        sort_order INTEGER
      );

      CREATE TABLE IF NOT EXISTS rows (
        id TEXT PRIMARY KEY,
        sheet_id TEXT,
        row_index INTEGER,
        data_json TEXT
      );

      CREATE TABLE IF NOT EXISTS sheet_page_meta (
        sheet_id TEXT PRIMARY KEY,
        row_count INTEGER,
        column_sums_json TEXT
      );

      CREATE TABLE IF NOT EXISTS row_marks (
        sheet_id TEXT NOT NULL,
        sheet_name TEXT NOT NULL,
        row_index INTEGER NOT NULL,
        color TEXT NOT NULL,
        PRIMARY KEY (sheet_id, row_index)
      );
    `)
  }

  function readRowMarksFromDatabase(db: Database): StoredRowMark[] {
    try {
      const res = db.exec(
        `SELECT sheet_id, sheet_name, row_index, color FROM row_marks ORDER BY sheet_name ASC, row_index ASC`,
      )
      if (!res?.[0]?.values?.length) return []
      return res[0].values.map((row) => ({
        sheetId: String(row[0]),
        sheetName: String(row[1]),
        rowIndex: Number(row[2]),
        color: String(row[3]) as StoredRowMark['color'],
      }))
    } catch {
      return []
    }
  }

  function writeRowMarksToDatabase(db: Database, workbook: WorkbookData) {
    try {
      db.run(`DELETE FROM row_marks`)
      for (const mark of collectRowMarksForSave(workbook.sheets)) {
        db.run(
          `INSERT INTO row_marks (sheet_id, sheet_name, row_index, color) VALUES (?, ?, ?, ?)`,
          [mark.sheetId, mark.sheetName, mark.rowIndex, mark.color],
        )
      }
    } catch {
      // Bảng chưa có trên DB cũ — initTables đã tạo ở phiên lưu mới
    }
  }

  function workbookFromDatabase(db: Database): WorkbookData | null {
    const wbRes = db.exec(`SELECT file_name, file_size, active_sheet_id, updated_at FROM workbook LIMIT 1`)
    if (!wbRes || wbRes.length === 0 || wbRes[0].values.length === 0) {
      return null
    }

    const wbRow = wbRes[0].values[0]
    const fileName = String(wbRow[0] || 'Phien_Lam_Viec.xlsx')
    const fileSize = Number(wbRow[1] || 0)
    const activeSheetId = String(wbRow[2] || '')
    const updatedAt = new Date(Number(wbRow[3] || Date.now()))

    let sheetsRes: ReturnType<Database['exec']> | null = null
    try {
      sheetsRes = db.exec(
        `SELECT id, name, header_row_index, source_grid_json FROM sheets ORDER BY sort_order ASC`,
      )
    } catch {
      sheetsRes = db.exec(`SELECT id, name, header_row_index FROM sheets ORDER BY sort_order ASC`)
    }
    if (!sheetsRes || sheetsRes.length === 0 || sheetsRes[0].values.length === 0) {
      return null
    }

    const sheets: SheetData[] = []

    for (const sRow of sheetsRes[0].values) {
      const sheetId = String(sRow[0])
      const sheetName = String(sRow[1])
        const headerRowIndex = Number(sRow[2] || 0)
        let sourceGrid: SheetData['sourceGrid']
        if (sRow.length > 3 && sRow[3] != null && String(sRow[3]) !== '') {
          try {
            sourceGrid = JSON.parse(String(sRow[3]))
          } catch {
            sourceGrid = undefined
          }
        }

      const stmtCols = db.prepare(`SELECT col_key, label, data_type FROM columns WHERE sheet_id = ? ORDER BY sort_order ASC`)
      stmtCols.bind([sheetId])
      const headers: ColumnHeader[] = []
      let colIdx = 0
      while (stmtCols.step()) {
        const colRow = stmtCols.get()
        headers.push({
          key: String(colRow[0]),
          label: String(colRow[1]),
          dataType: (colRow[2] as ColumnHeader['dataType']) || 'text',
          originalIndex: colIdx++,
        })
      }
      stmtCols.free()

      const stmtRows = db.prepare(`SELECT data_json FROM rows WHERE sheet_id = ? ORDER BY row_index ASC`)
      stmtRows.bind([sheetId])
      const rows: RowData[] = []
      while (stmtRows.step()) {
        const rRow = stmtRows.get()
        try {
          rows.push(JSON.parse(String(rRow[0])))
        } catch {
          rows.push({})
        }
      }
      stmtRows.free()

      let paged = false
      let rowCount: number | undefined
      let columnSums: SheetData['columnSums']
      try {
        const metaStmt = db.prepare(
          `SELECT row_count, column_sums_json FROM sheet_page_meta WHERE sheet_id = ?`,
        )
        metaStmt.bind([sheetId])
        if (metaStmt.step()) {
          const metaRow = metaStmt.get()
          paged = true
          rowCount = Number(metaRow[0] || 0)
          if (metaRow[1]) columnSums = JSON.parse(String(metaRow[1]))
        }
        metaStmt.free()
      } catch {
        paged = false
      }

      sheets.push({
        id: sheetId,
        name: sheetName,
        headers,
        rows: paged ? [] : rows,
        paged,
        rowCount,
        columnSums,
        headerRowIndex,
        sourceGrid,
      })
    }

    if (sheets.length === 0) return null

    return {
      meta: {
        fileName,
        fileSize,
        lastModified: updatedAt,
        sheetCount: sheets.length,
      },
      sheets,
      activeSheetId: activeSheetId || sheets[0].id,
    }
  }

  async function refreshSessions(): Promise<void> {
    const list = await loadJsonFromIndexedDB<FileSession[]>(INDEX_KEY)
    sessions.value = Array.isArray(list)
      ? [...list].sort((a, b) => b.updatedAt - a.updatedAt)
      : []
  }

  /** Gắn file đang mở vào một phiên (cùng tên file thì dùng lại phiên đó) */
  function adoptSessionForFile(fileName: string): string {
    const existing = sessions.value.find((s) => s.fileName === fileName)
    const id = existing?.id ?? crypto.randomUUID()
    setActiveSessionId(id)
    return id
  }

  /**
   * Lưu toàn bộ Workbook vào SQLite database và lưu xuống IndexedDB
   */
  async function saveSessionToSqlite(workbook: WorkbookData, sessionId?: string): Promise<void> {
    isSaving.value = true
    const id = sessionId || activeSessionId.value || crypto.randomUUID()
    if (!sessionId) setActiveSessionId(id)
    try {
      const SQL = await getSqlInstance()
      const db = new SQL.Database()
      currentDb = db

      initTables(db)

      // 1. Lưu Workbook info
      db.run(
        `INSERT OR REPLACE INTO workbook (id, file_name, file_size, active_sheet_id, updated_at) VALUES (1, ?, ?, ?, ?)`,
        [
          workbook.meta.fileName,
          workbook.meta.fileSize,
          workbook.activeSheetId,
          Date.now(),
        ],
      )

      // 2. Lưu Sheets, Columns, Rows
      workbook.sheets.forEach((sheet, sheetIdx) => {
        db.run(
          `INSERT INTO sheets (id, name, sort_order, header_row_index, source_grid_json) VALUES (?, ?, ?, ?, ?)`,
          [
            sheet.id,
            sheet.name,
            sheetIdx,
            sheet.headerRowIndex,
            sheet.sourceGrid ? JSON.stringify(sheet.sourceGrid) : null,
          ],
        )

        sheet.headers.forEach((col, colIdx) => {
          db.run(
            `INSERT INTO columns (id, sheet_id, col_key, label, data_type, sort_order) VALUES (?, ?, ?, ?, ?, ?)`,
            [
              `${sheet.id}_${col.key}`,
              sheet.id,
              col.key,
              col.label,
              col.dataType || 'text',
              colIdx,
            ],
          )
        })

        if (sheet.paged) {
          db.run(
            `INSERT INTO sheet_page_meta (sheet_id, row_count, column_sums_json) VALUES (?, ?, ?)`,
            [sheet.id, sheet.rowCount ?? 0, JSON.stringify(sheet.columnSums ?? {})],
          )
        } else {
          sheet.rows.forEach((row, rowIdx) => {
            db.run(
              `INSERT INTO rows (id, sheet_id, row_index, data_json) VALUES (?, ?, ?, ?)`,
              [`${sheet.id}_row_${rowIdx}`, sheet.id, rowIdx, JSON.stringify(row)],
            )
          })
        }
      })

      writeRowMarksToDatabase(db, workbook)

      const binaryData = db.export()
      await saveBinaryToIndexedDB(sessionKey(id), binaryData)

      const rowCount = workbook.sheets.reduce((sum, sheet) => sum + sheetRowCount(sheet), 0)
      const meta: FileSession = {
        id,
        fileName: workbook.meta.fileName,
        updatedAt: Date.now(),
        sheetCount: workbook.sheets.length,
        rowCount,
      }
      const next = sessions.value.filter((s) => s.id !== id)
      next.unshift(meta)
      sessions.value = next.sort((a, b) => b.updatedAt - a.updatedAt)
      await saveJsonToIndexedDB(INDEX_KEY, sessions.value)

      lastSavedAt.value = new Date()
    } catch (err) {
      console.error('Lỗi khi lưu SQLite session:', err)
    } finally {
      isSaving.value = false
    }
  }

  async function loadSessionById(id: string): Promise<WorkbookData | null> {
    const binaryData = await getBinaryFromIndexedDB(sessionKey(id))
    if (!binaryData) return null

    const SQL = await getSqlInstance()
    const db = new SQL.Database(binaryData)
    currentDb = db
    const result = workbookFromDatabase(db)
    if (result) {
      applyStoredRowMarks(result.sheets, readRowMarksFromDatabase(db))
      lastSavedAt.value = result.meta.lastModified
    }
    return result
  }

  async function deleteFileSession(id: string): Promise<void> {
    await clearBinaryFromIndexedDB(sessionKey(id))
    sessions.value = sessions.value.filter((s) => s.id !== id)
    await saveJsonToIndexedDB(INDEX_KEY, sessions.value)
    if (activeSessionId.value === id) {
      setActiveSessionId(sessions.value[0]?.id ?? null)
    }
  }

  /**
   * Khôi phục phiên đang mở. Nếu chỉ còn dữ liệu cũ một-file thì chuyển thành một phiên.
   */
  async function loadSessionFromSqlite(): Promise<WorkbookData | null> {
    isLoading.value = true
    try {
      await refreshSessions()

      if (sessions.value.length === 0) {
        const legacy = await getBinaryFromIndexedDB(DB_KEY)
        if (legacy) {
          const SQL = await getSqlInstance()
          const db = new SQL.Database(legacy)
          currentDb = db
          const legacyWb = workbookFromDatabase(db)
          if (legacyWb) {
            const id = crypto.randomUUID()
            setActiveSessionId(id)
            await saveSessionToSqlite(legacyWb, id)
          }
        }
      }

      const preferred = activeSessionId.value
      const id =
        preferred && sessions.value.some((s) => s.id === preferred)
          ? preferred
          : sessions.value[0]?.id
      if (!id) return null

      setActiveSessionId(id)
      return await loadSessionById(id)
    } catch (err) {
      console.error('Lỗi khi nạp SQLite session:', err)
      return null
    } finally {
      isLoading.value = false
      isInitialized.value = true
    }
  }

  /**
   * Tải file SQLite .db về máy tính cho người dùng
   */
  async function downloadSqliteDatabase(fileName = 'excel_session.sqlite') {
    const id = activeSessionId.value
    if (!currentDb && id) {
      const binaryData = await getBinaryFromIndexedDB(sessionKey(id))
      if (binaryData) {
        const blob = new Blob([binaryData], { type: 'application/x-sqlite3' })
        saveAs(blob, fileName)
        return
      }
    }
    if (!currentDb) {
      const binaryData = await getBinaryFromIndexedDB(DB_KEY)
      if (binaryData) {
        const blob = new Blob([binaryData], { type: 'application/x-sqlite3' })
        saveAs(blob, fileName)
      }
      return
    }
    const binary = currentDb.export()
    const blob = new Blob([binary], { type: 'application/x-sqlite3' })
    saveAs(blob, fileName)
  }

  /**
   * Xóa phiên làm việc trong SQLite
   */
  async function clearSqliteSession(): Promise<void> {
    await clearBinaryFromIndexedDB(DB_KEY)
    if (currentDb) {
      currentDb.close()
      currentDb = null
    }
    lastSavedAt.value = null
  }

  return {
    isSaving,
    isLoading,
    lastSavedAt,
    isInitialized,
    sessions,
    activeSessionId,
    saveSessionToSqlite,
    loadSessionFromSqlite,
    loadSessionById,
    deleteFileSession,
    adoptSessionForFile,
    setActiveSessionId,
    downloadSqliteDatabase,
    clearSqliteSession,
  }
}
