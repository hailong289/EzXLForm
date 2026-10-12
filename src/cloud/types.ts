export type CloudProviderId = 'google' | 'microsoft'
export type CloudOpenMode = 'file' | 'live'

export const GOOGLE_SHEET_MIME = 'application/vnd.google-apps.spreadsheet'
export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

export const LIVE_MAX_ROWS = 10_000
export const LIVE_MAX_COLS = 50

export interface CloudFileRef {
  id: string
  name: string
  mime: string
  webUrl?: string
  /** Đường dẫn Graph, ví dụ me/drive/items/{id} */
  graphPath?: string
}

export interface CloudBinding {
  provider: CloudProviderId
  mode: CloudOpenMode
  file: CloudFileRef
  revision: string
}

export interface CloudAccount {
  provider: CloudProviderId
  label: string
}
