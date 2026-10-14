import { googleClientId, googleClientSecret, googleTokenKey } from '@/cloud/config'
import { openToken, sealToken } from '@/cloud/tokenCipher'
import { readGoogleTokenCipher, writeGoogleTokenCipher } from '@/cloud/googleTokenStore'

const TOKEN_KEY = 'ezxl_google_token'
const TOKEN_URL = 'https://oauth2.googleapis.com/token'
const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive'
const SCOPES = ['openid', 'email', 'profile', DRIVE_SCOPE, 'https://www.googleapis.com/auth/spreadsheets'].join(' ')

interface StoredToken {
  accessToken: string
  expiresAt: number
  refreshToken?: string
  email?: string
  scope?: string
}

interface GoogleTokenPayload {
  access_token?: string
  expires_in?: number
  refresh_token?: string
  scope?: string
  error?: string
  error_description?: string
}

let tokenClient: GoogleTokenClient | null = null
let codeClient: GoogleCodeClient | null = null
let pending: {
  resolve: (token: StoredToken) => void
  reject: (error: Error) => void
} | null = null

let memory: StoredToken | null | undefined
let io: Promise<void> = Promise.resolve()

function enqueue<T>(job: () => Promise<T>): Promise<T> {
  const run = io.then(job, job)
  io = run.then(
    () => undefined,
    () => undefined,
  )
  return run
}

function scopeAllowsDrive(token: StoredToken): boolean {
  const granted = token.scope?.split(' ') ?? []
  return granted.includes(DRIVE_SCOPE)
}

function accessFresh(token: StoredToken | null | undefined): boolean {
  return !!token?.accessToken && token.expiresAt > Date.now() + 30_000 && scopeAllowsDrive(token)
}

function parseStored(value: unknown): StoredToken | null {
  if (!value || typeof value !== 'object') return null
  const token = value as StoredToken
  if (!token.accessToken && !token.refreshToken) return null
  if (token.scope && !scopeAllowsDrive(token)) return null
  return token
}

function readSession(): StoredToken | null {
  const raw = sessionStorage.getItem(TOKEN_KEY)
  if (!raw) return null
  try {
    return parseStored(JSON.parse(raw))
  } catch {
    return null
  }
}

async function persist(token: StoredToken | null): Promise<void> {
  memory = token
  const secret = googleTokenKey()
  if (!secret) {
    if (!token) sessionStorage.removeItem(TOKEN_KEY)
    else sessionStorage.setItem(TOKEN_KEY, JSON.stringify(token))
    return
  }
  sessionStorage.removeItem(TOKEN_KEY)
  if (!token) {
    await writeGoogleTokenCipher(null)
    return
  }
  await writeGoogleTokenCipher(await sealToken(secret, JSON.stringify(token)))
}

async function loadRecord(): Promise<StoredToken | null> {
  if (memory !== undefined) return memory
  const secret = googleTokenKey()
  if (!secret) {
    memory = readSession()
    return memory
  }
  const cipher = await readGoogleTokenCipher()
  if (cipher) {
    try {
      const parsed = parseStored(JSON.parse(await openToken(secret, cipher)))
      if (parsed && (accessFresh(parsed) || parsed.refreshToken)) {
        memory = parsed
        sessionStorage.removeItem(TOKEN_KEY)
        return parsed
      }
      memory = null
      await writeGoogleTokenCipher(null)
      return null
    } catch {
      memory = null
      await writeGoogleTokenCipher(null)
      return null
    }
  }
  const legacy = readSession()
  memory = legacy
  if (legacy) await persist(legacy)
  return legacy
}

function storedFromPayload(data: GoogleTokenPayload, previous?: StoredToken | null): StoredToken {
  return {
    accessToken: data.access_token ?? previous?.accessToken ?? '',
    expiresAt: Date.now() + (data.expires_in ?? 3600) * 1000,
    refreshToken: data.refresh_token || previous?.refreshToken,
    email: previous?.email,
    scope: data.scope || previous?.scope,
  }
}

async function postToken(body: URLSearchParams): Promise<GoogleTokenPayload> {
  const response = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  })
  const data = (await response.json()) as GoogleTokenPayload
  if (!response.ok || !data.access_token) {
    throw new Error(data.error_description || data.error || 'Google từ chối phiên đăng nhập')
  }
  return data
}

async function refreshAccess(record: StoredToken): Promise<StoredToken> {
  const data = await postToken(
    new URLSearchParams({
      client_id: googleClientId(),
      client_secret: googleClientSecret(),
      refresh_token: record.refreshToken ?? '',
      grant_type: 'refresh_token',
    }),
  )
  const next = storedFromPayload(data, record)
  if (!scopeAllowsDrive(next)) {
    throw new Error('Tài khoản chưa được cấp quyền Google Drive.')
  }
  await persist(next)
  return next
}

async function ensureFreshAccess(): Promise<StoredToken | null> {
  const record = await loadRecord()
  if (accessFresh(record)) return record
  if (record?.refreshToken && googleClientSecret()) {
    try {
      return await refreshAccess(record)
    } catch {
      await persist(null)
      return null
    }
  }
  if (record && !record.refreshToken) await persist(null)
  return null
}

function loadScript(src: string): Promise<void> {
  const existing = document.querySelector<HTMLScriptElement>(`script[src="${src}"]`)
  if (existing?.dataset.loaded === '1') return Promise.resolve()
  return new Promise((resolve, reject) => {
    const script = existing ?? document.createElement('script')
    script.src = src
    script.async = true
    script.onload = () => {
      script.dataset.loaded = '1'
      resolve()
    }
    script.onerror = () => reject(new Error('Không tải được Google Identity Services'))
    if (!existing) document.head.appendChild(script)
  })
}

async function oauthApi() {
  await loadScript('https://accounts.google.com/gsi/client')
  const oauth = window.google?.accounts.oauth2
  if (!oauth) throw new Error('Google Identity Services chưa sẵn sàng')
  return oauth
}

function rejectPending(error: Error) {
  const waiter = pending
  pending = null
  waiter?.reject(error)
}

function rememberEmail(accessToken: string) {
  void fetchEmail(accessToken).then((email) => {
    if (!email) return
    return enqueue(async () => {
      if (memory?.accessToken !== accessToken) return
      await persist({ ...memory, email })
    })
  })
}

async function acceptPayload(data: GoogleTokenPayload, previous?: StoredToken | null): Promise<StoredToken> {
  const granted = data.scope?.split(' ').filter(Boolean) ?? (await grantedScopes(data.access_token ?? ''))
  const stored = await enqueue(async () => {
    const next = storedFromPayload({ ...data, scope: granted.join(' ') }, previous ?? memory)
    if (!scopeAllowsDrive(next)) {
      throw new Error(
        'Tài khoản chưa được cấp quyền Google Drive. Trên OAuth consent screen hãy thêm scope https://www.googleapis.com/auth/drive, bật Google Drive API, rồi đăng nhập lại.',
      )
    }
    await persist(next)
    return next
  })
  rememberEmail(stored.accessToken)
  return stored
}

async function codeClientReady(): Promise<GoogleCodeClient> {
  if (codeClient) return codeClient
  const oauth = await oauthApi()
  codeClient = oauth.initCodeClient({
    client_id: googleClientId(),
    scope: SCOPES,
    ux_mode: 'popup',
    error_callback: (error) => {
      rejectPending(new Error(error.type === 'popup_closed' ? 'Đã đóng cửa sổ đăng nhập Google' : 'Không mở được cửa sổ đăng nhập Google'))
    },
    callback: (response) => {
      void (async () => {
        const waiter = pending
        pending = null
        if (!waiter) return
        if (response.error || !response.code) {
          waiter.reject(new Error(response.error_description || response.error || 'Google từ chối đăng nhập'))
          return
        }
        try {
          const data = await postToken(
            new URLSearchParams({
              code: response.code,
              client_id: googleClientId(),
              client_secret: googleClientSecret(),
              redirect_uri: 'postmessage',
              grant_type: 'authorization_code',
            }),
          )
          waiter.resolve(await acceptPayload({ ...data, scope: data.scope || response.scope }))
        } catch (error) {
          waiter.reject(error instanceof Error ? error : new Error('Google từ chối đăng nhập'))
        }
      })()
    },
  })
  return codeClient
}

async function tokenClientReady(): Promise<GoogleTokenClient> {
  if (tokenClient) return tokenClient
  const oauth = await oauthApi()
  tokenClient = oauth.initTokenClient({
    client_id: googleClientId(),
    scope: SCOPES,
    error_callback: (error) => {
      rejectPending(new Error(error.type === 'popup_closed' ? 'Đã đóng cửa sổ đăng nhập Google' : 'Không mở được cửa sổ đăng nhập Google'))
    },
    callback: (response) => {
      void (async () => {
        const waiter = pending
        pending = null
        if (!waiter) return
        if (response.error || !response.access_token) {
          waiter.reject(new Error(response.error_description || response.error || 'Google từ chối đăng nhập'))
          return
        }
        try {
          waiter.resolve(
            await acceptPayload({
              access_token: response.access_token,
              expires_in: response.expires_in,
              scope: response.scope,
            }),
          )
        } catch (error) {
          waiter.reject(error instanceof Error ? error : new Error('Google từ chối đăng nhập'))
        }
      })()
    },
  })
  return tokenClient
}

async function grantedScopes(token: string): Promise<string[]> {
  try {
    const response = await fetch(
      `https://www.googleapis.com/oauth2/v3/tokeninfo?access_token=${encodeURIComponent(token)}`,
    )
    if (!response.ok) return []
    const data = (await response.json()) as { scope?: string }
    return data.scope?.split(' ').filter(Boolean) ?? []
  } catch {
    return []
  }
}

async function fetchEmail(token: string): Promise<string | undefined> {
  try {
    const response = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: `Bearer ${token}` },
    })
    if (!response.ok) return undefined
    const data = (await response.json()) as { email?: string }
    return data.email
  } catch {
    return undefined
  }
}

function requestInteractive(): Promise<string> {
  return new Promise((resolve, reject) => {
    pending = {
      resolve: (token) => resolve(token.accessToken),
      reject,
    }
    const open = googleClientSecret()
      ? codeClientReady().then((client) => client.requestCode())
      : tokenClientReady().then((client) => client.requestAccessToken({ prompt: 'consent' }))
    void open
      .catch((error: unknown) => {
        pending = null
        reject(error instanceof Error ? error : new Error('Không mở được cửa sổ đăng nhập Google'))
      })
  })
}

export async function googleAccountLabel(): Promise<string | null> {
  const stored = await enqueue(() => loadRecord())
  if (!stored || (!accessFresh(stored) && !stored.refreshToken)) return null
  return stored.email ?? 'Google'
}

export async function googleAccessToken(interactive = true): Promise<string> {
  const stored = await enqueue(() => ensureFreshAccess())
  if (stored) return stored.accessToken
  if (!interactive) throw new Error('Chưa đăng nhập Google')
  return requestInteractive()
}

export async function signOutGoogle(): Promise<void> {
  const token = await enqueue(async () => {
    if (memory === undefined) await loadRecord()
    const current = memory?.refreshToken || memory?.accessToken || ''
    await persist(null)
    return current
  })
  if (token && window.google?.accounts.oauth2) {
    window.google.accounts.oauth2.revoke(token, () => {})
  }
}
