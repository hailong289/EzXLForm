export type CloudAccessRole = 'owner' | 'writer' | 'reader'

export interface CloudGrant {
  id: string
  label: string
  role: CloudAccessRole
  editable: boolean
  link: boolean
}

export type CloudLinkLevel = 'none' | 'reader' | 'writer'

export function linkAccess(grants: CloudGrant[]): CloudLinkLevel {
  const links = grants.filter((grant) => grant.link)
  if (links.some((grant) => grant.role === 'writer' || grant.role === 'owner')) return 'writer'
  if (links.length) return 'reader'
  return 'none'
}

export interface GooglePermission {
  id?: string
  type?: string
  role?: string
  emailAddress?: string
  displayName?: string
}

export interface GraphPermission {
  id?: string
  roles?: string[]
  grantedTo?: { user?: { email?: string; displayName?: string } }
  grantedToV2?: { user?: { email?: string; displayName?: string } }
  invitation?: { email?: string }
  link?: { scope?: string; type?: string }
}

function googleRole(role: string | undefined): CloudAccessRole {
  if (role === 'owner') return 'owner'
  if (role === 'writer') return 'writer'
  return 'reader'
}

export function grantsFromGoogle(items: GooglePermission[]): CloudGrant[] {
  return items.flatMap((item) => {
    if (!item.id) return []
    const role = googleRole(item.role)
    return [{
      id: item.id,
      label: item.emailAddress || item.displayName || (item.type === 'anyone' ? 'Bất kỳ ai có link' : 'Google'),
      role,
      editable: item.type === 'user' && role !== 'owner',
      link: item.type === 'anyone',
    }]
  })
}

export function grantsFromMicrosoft(items: GraphPermission[]): CloudGrant[] {
  return items.flatMap((item) => {
    if (!item.id) return []
    const publicLink = item.link?.scope === 'anonymous'
    const roles = item.roles ?? []
    const role: CloudAccessRole = publicLink && item.link?.type === 'edit'
      ? 'writer'
      : publicLink && item.link?.type === 'view'
        ? 'reader'
        : roles.includes('owner') ? 'owner' : roles.includes('write') ? 'writer' : 'reader'
    const user = item.grantedToV2?.user || item.grantedTo?.user
    return [{
      id: item.id,
      label: user?.email || item.invitation?.email || user?.displayName || (publicLink ? 'Bất kỳ ai có link' : 'OneDrive'),
      role,
      editable: !publicLink && role !== 'owner' && Boolean(user?.email || item.invitation?.email),
      link: publicLink,
    }]
  })
}
