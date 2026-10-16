import { grantsFromGoogle, grantsFromMicrosoft, type CloudGrant, type CloudLinkLevel, type GooglePermission, type GraphPermission } from '@/cloud/accessMap'
import { googleAccessToken } from '@/cloud/googleAuth'
import { apiError, authedJson } from '@/cloud/http'
import { microsoftAccessToken } from '@/cloud/microsoftAuth'
import type { CloudBinding } from '@/cloud/types'

export type { CloudAccessRole, CloudGrant, CloudLinkLevel } from '@/cloud/accessMap'

async function googlePermissions(fileId: string, token: string) {
  const data = await authedJson<{ permissions?: GooglePermission[] }>(
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}/permissions?supportsAllDrives=true&fields=permissions(id,type,role,emailAddress,displayName)`,
    token,
  )
  return grantsFromGoogle(data.permissions ?? [])
}

function graphPath(binding: CloudBinding): string {
  if (!binding.file.graphPath) throw new Error('File OneDrive thiếu đường dẫn')
  return binding.file.graphPath
}

export async function listCloudGrants(binding: CloudBinding): Promise<CloudGrant[]> {
  if (binding.provider === 'google') return googlePermissions(binding.file.id, await googleAccessToken())
  const token = await microsoftAccessToken()
  const data = await authedJson<{ value?: GraphPermission[] }>(
    `https://graph.microsoft.com/v1.0/${graphPath(binding)}/permissions`,
    token,
  )
  return grantsFromMicrosoft(data.value ?? [])
}

export async function grantCloudAccess(binding: CloudBinding, email: string, role: 'writer' | 'reader'): Promise<void> {
  const address = email.trim()
  if (!address) throw new Error('Thiếu email')
  if (binding.provider === 'google') {
    const token = await googleAccessToken()
    await authedJson(
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(binding.file.id)}/permissions?supportsAllDrives=true&sendNotificationEmail=true`,
      token,
      {
        method: 'POST',
        body: JSON.stringify({ role, type: 'user', emailAddress: address }),
      },
    )
    return
  }
  const token = await microsoftAccessToken()
  await authedJson(`https://graph.microsoft.com/v1.0/${graphPath(binding)}/invite`, token, {
    method: 'POST',
    body: JSON.stringify({
      recipients: [{ email: address }],
      requireSignIn: true,
      sendInvitation: true,
      roles: [role === 'writer' ? 'write' : 'read'],
    }),
  })
}

export async function setCloudGrantRole(binding: CloudBinding, grantId: string, role: 'writer' | 'reader'): Promise<void> {
  if (binding.provider === 'google') {
    const token = await googleAccessToken()
    await authedJson(
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(binding.file.id)}/permissions/${encodeURIComponent(grantId)}?supportsAllDrives=true`,
      token,
      { method: 'PATCH', body: JSON.stringify({ role }) },
    )
    return
  }
  const token = await microsoftAccessToken()
  await authedJson(
    `https://graph.microsoft.com/v1.0/${graphPath(binding)}/permissions/${encodeURIComponent(grantId)}`,
    token,
    { method: 'PATCH', body: JSON.stringify({ roles: [role === 'writer' ? 'write' : 'read'] }) },
  )
}

export async function setCloudLinkAccess(binding: CloudBinding, level: CloudLinkLevel): Promise<void> {
  const grants = await listCloudGrants(binding)
  const links = grants.filter((grant) => grant.link)
  if (level === 'none') {
    for (const grant of links) await removeCloudGrant(binding, grant.id)
    return
  }
  if (binding.provider === 'google') {
    const token = await googleAccessToken()
    const url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(binding.file.id)}/permissions`
    const link = links[0]
    if (link) {
      await authedJson(`${url}/${encodeURIComponent(link.id)}?supportsAllDrives=true`, token, {
        method: 'PATCH',
        body: JSON.stringify({ role: level }),
      })
      return
    }
    await authedJson(`${url}?supportsAllDrives=true&sendNotificationEmail=false`, token, {
      method: 'POST',
      body: JSON.stringify({ role: level, type: 'anyone' }),
    })
    return
  }
  const token = await microsoftAccessToken()
  const keep = level === 'writer' ? 'writer' : 'reader'
  for (const grant of links) {
    if (grant.role !== keep) await removeCloudGrant(binding, grant.id)
  }
  if (links.some((grant) => grant.role === keep)) return
  await authedJson(`https://graph.microsoft.com/v1.0/${graphPath(binding)}/createLink`, token, {
    method: 'POST',
    body: JSON.stringify({ type: level === 'writer' ? 'edit' : 'view', scope: 'anonymous' }),
  })
}

export async function removeCloudGrant(binding: CloudBinding, grantId: string): Promise<void> {
  const token = binding.provider === 'google' ? await googleAccessToken() : await microsoftAccessToken()
  const url = binding.provider === 'google'
    ? `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(binding.file.id)}/permissions/${encodeURIComponent(grantId)}?supportsAllDrives=true`
    : `https://graph.microsoft.com/v1.0/${graphPath(binding)}/permissions/${encodeURIComponent(grantId)}`
  const response = await fetch(url, { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } })
  if (!response.ok) throw await apiError(response)
}
