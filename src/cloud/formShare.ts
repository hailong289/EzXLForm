import { linkAccess, type CloudGrant } from '@/cloud/accessMap'
import type { CloudProviderId } from '@/cloud/types'
import type { ColumnHeader } from '@/types/excel'

export const FORM_SHEET = '_ezxl_form'

export interface SharedFormColumn {
  key: string
  label: string
  index: number
  dataType?: ColumnHeader['dataType']
  dateFormat?: string
  required?: boolean
  unique?: boolean
  pattern?: string
  options?: string[]
  formula?: string
}

export interface SharedForm {
  sheet: string
  headerRow: number
  columns: SharedFormColumn[]
}

export function formSnapshot(
  sheetName: string,
  headerRow: number,
  headers: ColumnHeader[],
  visibleKeys: string[],
): SharedForm {
  const allowed = new Set(visibleKeys)
  return {
    sheet: sheetName,
    headerRow,
    columns: headers
      .filter((header) => header.visible !== false && allowed.has(header.key))
      .map((header) => ({
        key: header.key,
        label: header.label,
        index: header.originalIndex,
        dataType: header.dataType,
        dateFormat: header.dateFormat,
        required: header.required || undefined,
        unique: header.unique || undefined,
        pattern: header.pattern,
        options: header.options?.length ? header.options : undefined,
        formula: header.formula,
      })),
  }
}

export function parseFormBook(raw: unknown): SharedForm | null {
  if (typeof raw !== 'string' || !raw) return null
  try {
    const value = JSON.parse(raw) as Partial<SharedForm>
    if (!value.sheet || !Array.isArray(value.columns) || value.columns.length === 0) return null
    const columns = value.columns.flatMap((column) => {
      if (!column || typeof column.key !== 'string' || typeof column.label !== 'string' || typeof column.index !== 'number') return []
      return [column]
    })
    if (columns.length !== value.columns.length) return null
    return { sheet: value.sheet, headerRow: value.headerRow ?? 0, columns }
  } catch {
    return null
  }
}

export function canEditCloud(grants: CloudGrant[], email: string | null): boolean {
  if (linkAccess(grants) === 'writer') return true
  const me = email?.trim().toLocaleLowerCase()
  if (!me) return false
  return grants.some((grant) => !grant.link && grant.label.trim().toLocaleLowerCase() === me && (grant.role === 'owner' || grant.role === 'writer'))
}

export function fillHash(provider: CloudProviderId, fileId: string): string {
  return `#fill/${provider}/${encodeURIComponent(fileId)}`
}

export function parseFillHash(hash: string): { provider: CloudProviderId; fileId: string } | null {
  const match = /^#fill\/(google|microsoft)\/([^/]+)$/.exec(hash)
  if (!match) return null
  return { provider: match[1] as CloudProviderId, fileId: decodeURIComponent(match[2]) }
}
