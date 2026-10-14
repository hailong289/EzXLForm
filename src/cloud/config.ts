export function googleClientId(): string {
  return import.meta.env.VITE_GOOGLE_CLIENT_ID?.trim() ?? ''
}

export function googleApiKey(): string {
  return import.meta.env.VITE_GOOGLE_API_KEY?.trim() ?? ''
}

/** Project number của Cloud. Picker bắt buộc số này, nếu không sẽ báo developer key không hợp lệ. */
export function googleAppId(): string {
  const configured = import.meta.env.VITE_GOOGLE_APP_ID?.trim()
  if (configured) return configured
  const projectNumber = googleClientId().split('-')[0]
  return /^\d+$/.test(projectNumber) ? projectNumber : ''
}

export function googleTokenKey(): string {
  return import.meta.env.VITE_GOOGLE_TOKEN_KEY?.trim() ?? ''
}

export function googleClientSecret(): string {
  return import.meta.env.VITE_GOOGLE_CLIENT_SECRET?.trim() ?? ''
}

export function microsoftClientId(): string {
  return import.meta.env.VITE_MICROSOFT_CLIENT_ID?.trim() ?? ''
}

export function googleConfigured(): boolean {
  return Boolean(googleClientId())
}

export function microsoftConfigured(): boolean {
  return Boolean(microsoftClientId())
}
