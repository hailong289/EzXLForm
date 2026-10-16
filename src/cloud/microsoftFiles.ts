import { microsoftAccessToken } from '@/cloud/microsoftAuth'
import { apiError, authedJson } from '@/cloud/http'
import { XLSX_MIME, type CloudFileRef } from '@/cloud/types'

interface DriveItem {
  id?: string
  name?: string
  webUrl?: string
  file?: { mimeType?: string }
  remoteItem?: {
    id?: string
    parentReference?: { driveId?: string }
  }
}

function graphPath(item: DriveItem): string | null {
  const remote = item.remoteItem
  if (remote?.id && remote.parentReference?.driveId) {
    return `drives/${remote.parentReference.driveId}/items/${remote.id}`
  }
  if (!item.id) return null
  return `me/drive/items/${item.id}`
}

function toRef(item: DriveItem): CloudFileRef | null {
  const path = graphPath(item)
  const name = item.name || ''
  const mime = item.file?.mimeType || ''
  const excel = mime === XLSX_MIME || name.toLowerCase().endsWith('.xlsx')
  if (!path || !excel || name.toLowerCase().endsWith('.xls')) return null
  return {
    id: item.remoteItem?.id || item.id || path,
    name,
    mime: XLSX_MIME,
    webUrl: item.webUrl,
    graphPath: path,
  }
}

async function graph(path: string, init?: RequestInit): Promise<Response> {
  const token = await microsoftAccessToken()
  const response = await fetch(`https://graph.microsoft.com/v1.0/${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      ...init?.headers,
    },
  })
  if (!response.ok) throw await apiError(response)
  return response
}

export async function listRecentExcel(): Promise<CloudFileRef[]> {
  const token = await microsoftAccessToken()
  const data = await authedJson<{ value?: DriveItem[] }>(
    'https://graph.microsoft.com/v1.0/me/drive/recent',
    token,
  )
  const files: CloudFileRef[] = []
  const seen = new Set<string>()
  for (const item of data.value ?? []) {
    const ref = toRef(item)
    if (!ref || seen.has(ref.graphPath || ref.id)) continue
    seen.add(ref.graphPath || ref.id)
    files.push(ref)
  }
  return files
}

function encodeShareId(url: string): string {
  const bytes = new TextEncoder().encode(url.trim())
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  const base64 = btoa(binary).replace(/=+$/g, '').replace(/\//g, '_').replace(/\+/g, '-')
  return `u!${base64}`
}

export async function microsoftFileByShareUrl(url: string): Promise<CloudFileRef> {
  const token = await microsoftAccessToken()
  const shareId = encodeURIComponent(encodeShareId(url))
  const data = await authedJson<DriveItem & { parentReference?: { driveId?: string } }>(
    `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem?$select=id,name,webUrl,file,parentReference,remoteItem`,
    token,
  )
  const driveId = data.parentReference?.driveId || data.remoteItem?.parentReference?.driveId
  const itemId = data.remoteItem?.id || data.id
  if (!itemId) throw new Error('Không đọc được file từ link Excel.')
  const graphPath = driveId ? `drives/${driveId}/items/${itemId}` : `me/drive/items/${itemId}`
  const name = data.name || 'Excel'
  const mime = data.file?.mimeType || ''
  const excel = mime === XLSX_MIME || name.toLowerCase().endsWith('.xlsx')
  if (!excel) throw new Error('Link này không phải file .xlsx.')
  return {
    id: itemId,
    name,
    mime: XLSX_MIME,
    webUrl: data.webUrl,
    graphPath,
  }
}

export async function microsoftRevision(file: CloudFileRef): Promise<string> {
  const path = file.graphPath
  if (!path) throw new Error('File OneDrive thiếu đường dẫn')
  const token = await microsoftAccessToken()
  const data = await authedJson<{ eTag?: string }>(
    `https://graph.microsoft.com/v1.0/${path}?select=eTag`,
    token,
  )
  return data.eTag || ''
}

export async function downloadMicrosoftFile(file: CloudFileRef): Promise<File> {
  if (!file.graphPath) throw new Error('File OneDrive thiếu đường dẫn')
  const response = await graph(`${file.graphPath}/content`)
  const buffer = await response.arrayBuffer()
  const name = file.name.toLowerCase().endsWith('.xlsx') ? file.name : `${file.name}.xlsx`
  return new File([buffer], name, { type: XLSX_MIME, lastModified: Date.now() })
}

export async function uploadMicrosoftXlsx(file: CloudFileRef, blob: Blob): Promise<string> {
  if (!file.graphPath) throw new Error('File OneDrive thiếu đường dẫn')
  const response = await graph(`${file.graphPath}/content`, {
    method: 'PUT',
    headers: { 'Content-Type': XLSX_MIME },
    body: blob,
  })
  const data = (await response.json()) as { eTag?: string }
  return data.eTag || ''
}

export async function createMicrosoftXlsx(name: string, blob: Blob): Promise<CloudFileRef> {
  const token = await microsoftAccessToken()
  const fileName = name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fileName)}:/content?@microsoft.graph.conflictBehavior=rename`,
    {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': XLSX_MIME },
      body: blob,
    },
  )
  if (!response.ok) throw await apiError(response)
  const item = (await response.json()) as DriveItem
  const file = toRef(item)
  if (!file) throw new Error('Không tạo được file trên OneDrive')
  return file
}
