import initSqlJs from 'sql.js'
import sqlWasmUrl from 'sql.js/dist/sql-wasm-browser.wasm?url'
import type { Database, SqlJsStatic } from 'sql.js'

const DB_NAME = 'excel_editor_db'
const DB_STORE_NAME = 'excel_sqlite_store'
const AUTH_DB_KEY = 'google_auth_sqlite'
const ROW_ID = 'google'

let sqlInstance: SqlJsStatic | null = null

function openIndexedDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1)
    request.onupgradeneeded = () => {
      const db = request.result
      if (!db.objectStoreNames.contains(DB_STORE_NAME)) db.createObjectStore(DB_STORE_NAME)
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })
}

async function readBinary(): Promise<Uint8Array | null> {
  const db = await openIndexedDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE_NAME, 'readonly')
    const req = tx.objectStore(DB_STORE_NAME).get(AUTH_DB_KEY)
    req.onsuccess = () => resolve((req.result as Uint8Array) || null)
    req.onerror = () => reject(req.error)
  })
}

async function writeBinary(data: Uint8Array | null): Promise<void> {
  const db = await openIndexedDB()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE_NAME, 'readwrite')
    const store = tx.objectStore(DB_STORE_NAME)
    const req = data ? store.put(data, AUTH_DB_KEY) : store.delete(AUTH_DB_KEY)
    req.onsuccess = () => resolve()
    req.onerror = () => reject(req.error)
  })
}

async function database(): Promise<Database> {
  if (!sqlInstance) {
    sqlInstance = await initSqlJs({ locateFile: () => sqlWasmUrl })
  }
  const binary = await readBinary()
  const db = binary ? new sqlInstance.Database(binary) : new sqlInstance.Database()
  db.run(`CREATE TABLE IF NOT EXISTS auth_token (id TEXT PRIMARY KEY, cipher TEXT NOT NULL)`)
  return db
}

export async function readGoogleTokenCipher(): Promise<string | null> {
  const db = await database()
  try {
    const statement = db.prepare(`SELECT cipher FROM auth_token WHERE id = ?`)
    statement.bind([ROW_ID])
    const value = statement.step() ? statement.getAsObject().cipher : null
    statement.free()
    return typeof value === 'string' && value ? value : null
  } finally {
    db.close()
  }
}

export async function writeGoogleTokenCipher(cipher: string | null): Promise<void> {
  const db = await database()
  try {
    if (!cipher) db.run(`DELETE FROM auth_token WHERE id = ?`, [ROW_ID])
    else db.run(`INSERT OR REPLACE INTO auth_token (id, cipher) VALUES (?, ?)`, [ROW_ID, cipher])
    const left = db.exec(`SELECT COUNT(*) FROM auth_token`)
    const count = Number(left[0]?.values[0]?.[0] ?? 0)
    await writeBinary(count > 0 ? db.export() : null)
  } finally {
    db.close()
  }
}
