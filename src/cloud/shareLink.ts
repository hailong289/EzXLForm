import { googleFileById } from '@/cloud/googleFiles'
import { microsoftFileByShareUrl } from '@/cloud/microsoftFiles'
import type { CloudFileRef, CloudProviderId } from '@/cloud/types'

export function shareProvider(raw: string): CloudProviderId | null {
  let url: URL
  try {
    url = new URL(raw.trim())
  } catch {
    return null
  }
  const host = url.hostname.toLowerCase().replace(/^www\./, '')
  if (host === 'docs.google.com' || host === 'drive.google.com') return 'google'
  if (
    host === '1drv.ms' ||
    host.endsWith('onedrive.live.com') ||
    host.endsWith('sharepoint.com') ||
    host.endsWith('office.com') ||
    host.endsWith('office365.com') ||
    host === 'excel.cloud.microsoft'
  ) {
    return 'microsoft'
  }
  return null
}

function googleFileId(raw: string): string | null {
  const url = new URL(raw.trim())
  const parts = url.pathname.split('/').filter(Boolean)
  const marker = parts.indexOf('d')
  if (marker >= 0 && parts[marker + 1]) return parts[marker + 1]
  return url.searchParams.get('id')
}

export async function fileFromShareLink(raw: string): Promise<{ provider: CloudProviderId; file: CloudFileRef }> {
  const provider = shareProvider(raw)
  if (!provider) throw new Error('Dán link chia sẻ Google Sheet hoặc Excel (OneDrive, SharePoint).')
  if (provider === 'google') {
    const id = googleFileId(raw)
    if (!id) throw new Error('Link Google thiếu mã file.')
    return { provider, file: await googleFileById(id) }
  }
  return { provider, file: await microsoftFileByShareUrl(raw.trim()) }
}
