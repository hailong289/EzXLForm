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
        }) => GoogleTokenClient
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
