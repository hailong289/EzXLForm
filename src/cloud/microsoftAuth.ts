import { PublicClientApplication, type AccountInfo } from '@azure/msal-browser'
import { microsoftClientId } from '@/cloud/config'

const SCOPES = ['User.Read', 'Files.ReadWrite', 'offline_access']

let app: PublicClientApplication | null = null
let ready: Promise<PublicClientApplication> | null = null

function createApp(): PublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: microsoftClientId(),
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: 'sessionStorage',
    },
  })
}

export async function microsoftApp(): Promise<PublicClientApplication> {
  if (!ready) {
    app = createApp()
    ready = app.initialize().then(async () => {
      await app?.handleRedirectPromise()
      return app as PublicClientApplication
    })
  }
  return ready
}

function accountLabel(account: AccountInfo): string {
  return account.username || account.name || 'Microsoft'
}

export async function microsoftAccountLabel(): Promise<string | null> {
  if (!microsoftClientId()) return null
  const pca = await microsoftApp()
  const account = pca.getAllAccounts()[0]
  return account ? accountLabel(account) : null
}

export async function microsoftAccessToken(): Promise<string> {
  const pca = await microsoftApp()
  let account = pca.getAllAccounts()[0]
  if (!account) {
    const login = await pca.loginPopup({ scopes: SCOPES })
    account = login.account
    return login.accessToken
  }
  try {
    const silent = await pca.acquireTokenSilent({ scopes: SCOPES, account })
    return silent.accessToken
  } catch {
    const popup = await pca.acquireTokenPopup({ scopes: SCOPES, account })
    return popup.accessToken
  }
}

export async function signOutMicrosoft(): Promise<void> {
  const pca = await microsoftApp()
  const account = pca.getAllAccounts()[0]
  if (!account) return
  await pca.logoutPopup({ account })
}
