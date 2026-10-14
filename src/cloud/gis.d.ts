interface GoogleTokenResponse {
  access_token?: string
  expires_in?: number
  scope?: string
  error?: string
  error_description?: string
}

interface GoogleTokenClient {
  requestAccessToken: (override?: { prompt?: string }) => void
}

interface GoogleCodeResponse {
  code?: string
  scope?: string
  error?: string
  error_description?: string
}

interface GoogleCodeClient {
  requestCode: () => void
}

interface GooglePickerDoc {
  id: string
  name?: string
  mimeType?: string
  url?: string
}

interface GooglePickerData {
  action: string
  docs?: GooglePickerDoc[]
}

interface GooglePickerBuilder {
  addView: (view: unknown) => GooglePickerBuilder
  setOAuthToken: (token: string) => GooglePickerBuilder
  setDeveloperKey: (key: string) => GooglePickerBuilder
  setAppId: (appId: string) => GooglePickerBuilder
  setOrigin: (origin: string) => GooglePickerBuilder
  setCallback: (callback: (data: GooglePickerData) => void) => GooglePickerBuilder
  setTitle: (title: string) => GooglePickerBuilder
  build: () => { setVisible: (visible: boolean) => void }
}

interface GoogleDocsView {
  setMimeTypes: (mime: string) => GoogleDocsView
}

interface Window {
  gapi?: {
    load: (name: string, callback: () => void) => void
  }
  google?: {
    accounts: {
      oauth2: {
        initTokenClient: (config: {
          client_id: string
          scope: string
          callback: (response: GoogleTokenResponse) => void
          error_callback?: (error: { type: string }) => void
        }) => GoogleTokenClient
        initCodeClient: (config: {
          client_id: string
          scope: string
          ux_mode?: 'popup' | 'redirect'
          callback: (response: GoogleCodeResponse) => void
          error_callback?: (error: { type: string }) => void
        }) => GoogleCodeClient
        revoke: (token: string, done: () => void) => void
      }
    }
    picker?: {
      Action: { PICKED: string; CANCEL: string }
      ViewId: { SPREADSHEETS: string; DOCS: string }
      DocsView: new (viewId: string) => GoogleDocsView
      PickerBuilder: new () => GooglePickerBuilder
    }
  }
}
