import { googleClientId } from '@/cloud/config'

const TOKEN_KEY = 'ezxl_google_token'
const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive'
const SCOPES = ['openid', 'email', 'profile', DRIVE_SCOPE, 'https://www.googleapis.com/auth/spreadsheets'].join(' ')

interface StoredToken {
  accessToken: string
  expiresAt: number
  email?: string
  scope?: string
}

let client: GoogleTokenClient | null = null
let pending: {
  resolve: (token: StoredToken) => void
  reject: (error: Error) => void
} | null = null

function readStored(): StoredToken | null {
  const raw = sessionStorage.getItem(TOKEN_KEY)
  if (!raw) return null
  try {
    const parsed = JSON.parse(raw) as StoredToken
    if (!parsed.accessToken || parsed.expiresAt <= Date.now() + 30_000) return null
    const granted = parsed.scope?.split(' ') ?? []
    if (!granted.includes(DRIVE_SCOPE)) return null
    return parsed
  } catch {
    return null
  }
}

function writeStored(token: StoredToken | null) {
  if (!token) sessionStorage.removeItem(TOKEN_KEY)
  else sessionStorage.setItem(TOKEN_KEY, JSON.stringify(token))
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

async function tokenClient(): Promise<GoogleTokenClient> {
  if (client) return client
  await loadScript('https://accounts.google.com/gsi/client')
  const oauth = window.google?.accounts.oauth2
  if (!oauth) throw new Error('Google Identity Services chưa sẵn sàng')
  client = oauth.initTokenClient({
    client_id: googleClientId(),
    scope: SCOPES,
    callback: (response) => {
      void (async () => {
      const waiter = pending
      pending = null
      if (!waiter) return
      if (response.error || !response.access_token) {
        waiter.reject(new Error(response.error_description || response.error || 'Google từ chối đăng nhập'))
        return
      }
      const granted = response.scope?.split(' ').filter(Boolean) ?? (await grantedScopes(response.access_token))
      if (!granted.includes(DRIVE_SCOPE)) {
        waiter.reject(
          new Error(
            'Tài khoản chưa được cấp quyền Google Drive. Trên OAuth consent screen hãy thêm scope https://www.googleapis.com/auth/drive, bật Google Drive API, rồi đăng nhập lại.',
          ),
        )
        return
      }
      const stored: StoredToken = {
        accessToken: response.access_token,
        expiresAt: Date.now() + (response.expires_in ?? 3600) * 1000,
        scope: granted.join(' '),
      }
      writeStored(stored)
      void fetchEmail(stored.accessToken).then((email) => {
        if (!email) return
        const current = readStored()
        if (current?.accessToken === stored.accessToken) writeStored({ ...current, email })
      })
      waiter.resolve(stored)
      })()
    },
  })
  return client
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

export function googleAccountLabel(): string | null {
  return readStored()?.email ?? (readStored() ? 'Google' : null)
}

export async function googleAccessToken(interactive = true): Promise<string> {
  const stored = readStored()
  if (stored) return stored.accessToken
  if (!interactive) throw new Error('Chưa đăng nhập Google')
  const oauth = await tokenClient()
  return new Promise((resolve, reject) => {
    pending = {
      resolve: (token) => resolve(token.accessToken),
      reject,
    }
    oauth.requestAccessToken({ prompt: 'consent' })
  })
}

export function signOutGoogle(): void {
  const raw = sessionStorage.getItem(TOKEN_KEY)
  let accessToken = ''
  try {
    accessToken = raw ? (JSON.parse(raw) as StoredToken).accessToken : ''
  } catch {
    accessToken = ''
  }
  writeStored(null)
  if (accessToken && window.google?.accounts.oauth2) {
    window.google.accounts.oauth2.revoke(accessToken, () => {})
  }
}
