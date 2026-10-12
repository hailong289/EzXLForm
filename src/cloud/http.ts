export async function apiError(response: Response): Promise<Error> {
  let message = response.statusText || `HTTP ${response.status}`
  let reason = ''
  try {
    const data = (await response.json()) as {
      error?: { message?: string; status?: string; errors?: { reason?: string; message?: string }[] } | string
      error_description?: string
    }
    if (typeof data.error === 'string' && data.error) message = data.error
    else if (typeof data.error === 'object' && data.error?.message) {
      message = data.error.message
      reason = data.error.errors?.[0]?.reason || data.error.status || ''
    } else if (data.error_description) message = data.error_description
  } catch {
    // Giữ status text khi body không phải JSON
  }
  if (response.status === 403) message = explainGoogleDenied(message, reason)
  return new Error(message)
}

function explainGoogleDenied(message: string, reason: string): string {
  if (/has not been used in project|is disabled|SERVICE_DISABLED|accessNotConfigured/i.test(`${message} ${reason}`)) {
    return 'Chưa bật API trên Google Cloud. Mở đúng project của OAuth client, bật Google Drive API và Google Sheets API, đợi vài phút rồi thử lại.'
  }
  if (/insufficientFilePermissions|does not have sufficient permissions/i.test(`${message} ${reason}`)) {
    return 'Tài khoản Google xem được file này nhưng không được sửa. Hãy mở quyền Editor cho chính tài khoản đang đăng nhập, rồi bấm Đồng bộ.'
  }
  if (/caller does not have permission|insufficient|PERMISSION_DENIED|forbidden/i.test(`${message} ${reason}`)) {
    return 'Google từ chối quyền ghi. File phải do chính tài khoản này sửa được (quyền Editor), và OAuth consent screen phải có scope https://www.googleapis.com/auth/drive cùng https://www.googleapis.com/auth/spreadsheets. Sau đó đăng xuất rồi đăng nhập lại.'
  }
  return message
}

export async function authedJson<T>(url: string, token: string, init?: RequestInit): Promise<T> {
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(init?.body ? { 'Content-Type': 'application/json' } : {}),
      ...init?.headers,
    },
  })
  if (!response.ok) throw await apiError(response)
  return (await response.json()) as T
}
