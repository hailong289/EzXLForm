import { googleAccessToken } from '@/cloud/googleAuth'
import { apiError, authedJson } from '@/cloud/http'
import { GOOGLE_SHEET_MIME, XLSX_MIME, type CloudFileRef } from '@/cloud/types'

interface DriveFile {
  id?: string
  name?: string
  mimeType?: string
  webViewLink?: string
}

export async function listGoogleFiles(liveOnly: boolean): Promise<CloudFileRef[]> {
  const token = await googleAccessToken()
  const mimeQuery = liveOnly
    ? `mimeType='${GOOGLE_SHEET_MIME}'`
    : `(mimeType='${GOOGLE_SHEET_MIME}' or mimeType='${XLSX_MIME}')`
  const query = encodeURIComponent(`${mimeQuery} and trashed=false`)
  const data = await authedJson<{ files?: DriveFile[] }>(
    `https://www.googleapis.com/drive/v3/files?q=${query}&orderBy=modifiedTime%20desc&pageSize=40&fields=files(id,name,mimeType,webViewLink)`,
    token,
  )
  return (data.files ?? [])
    .filter((file): file is DriveFile & { id: string } => Boolean(file.id))
    .map((file) => ({
      id: file.id,
      name: file.name || 'Sheet',
      mime: file.mimeType || GOOGLE_SHEET_MIME,
      webUrl: file.webViewLink,
    }))
}

export async function googleFileById(fileId: string): Promise<CloudFileRef> {
  const token = await googleAccessToken()
  const data = await authedJson<DriveFile>(
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?supportsAllDrives=true&fields=id,name,mimeType,webViewLink`,
    token,
  )
  if (!data.id) throw new Error('Không đọc được file từ link Google.')
  const mime = data.mimeType || ''
  if (mime !== GOOGLE_SHEET_MIME && mime !== XLSX_MIME) {
    throw new Error('Link này không phải Google Sheet hoặc file .xlsx.')
  }
  return {
    id: data.id,
    name: data.name || 'Sheet',
    mime,
    webUrl: data.webViewLink,
  }
}

export async function googleRevision(fileId: string): Promise<string> {
  const token = await googleAccessToken()
  const data = await authedJson<{ version?: string | number }>(
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?fields=version`,
    token,
  )
  return String(data.version ?? '')
}

export async function downloadGoogleFile(file: CloudFileRef): Promise<File> {
  const token = await googleAccessToken()
  const url =
    file.mime === GOOGLE_SHEET_MIME
      ? `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}/export?mimeType=${encodeURIComponent(XLSX_MIME)}`
      : `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}?alt=media`
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } })
  if (!response.ok) throw await apiError(response)
  const buffer = await response.arrayBuffer()
  const name = file.name.toLowerCase().endsWith('.xlsx') ? file.name : `${file.name}.xlsx`
  return new File([buffer], name, { type: XLSX_MIME, lastModified: Date.now() })
}

export async function uploadGoogleXlsx(fileId: string, blob: Blob): Promise<string> {
  const token = await googleAccessToken()
  const response = await fetch(
    `https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(fileId)}?uploadType=media&fields=version`,
    {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': XLSX_MIME,
      },
      body: blob,
    },
  )
  if (!response.ok) throw await apiError(response)
  const data = (await response.json()) as { version?: string | number }
  return String(data.version ?? '')
}

export async function createGoogleXlsx(name: string, blob: Blob): Promise<CloudFileRef> {
  const token = await googleAccessToken()
  const fileName = name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`
  const body = new FormData()
  body.append('metadata', new Blob([JSON.stringify({ name: fileName, mimeType: XLSX_MIME })], { type: 'application/json' }))
  body.append('file', blob, fileName)
  const response = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,webViewLink',
    { method: 'POST', headers: { Authorization: `Bearer ${token}` }, body },
  )
  if (!response.ok) throw await apiError(response)
  const data = (await response.json()) as DriveFile
  if (!data.id) throw new Error('Không tạo được file trên Google Drive')
  return { id: data.id, name: data.name || fileName, mime: data.mimeType || XLSX_MIME, webUrl: data.webViewLink }
}
