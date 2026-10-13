export interface FormPreset {
  id: string
  sheetId: string
  name: string
  visibleKeys: string[]
}

export const EMPTY_PRESET_MESSAGE = 'Loại form này không còn cột. Hãy sửa lại trong Loại form.'

export type PresetDraftError = 'empty-name' | 'duplicate-name' | 'no-columns'

export function presetStorageKey(sessionId: string, sheetId: string): string {
  return `excel_editor_form_preset:${sessionId}:${sheetId}`
}

export function readSelectedPresetId(
  storage: { getItem(key: string): string | null },
  sessionId: string,
  sheetId: string,
): string {
  return storage.getItem(presetStorageKey(sessionId, sheetId)) ?? ''
}

export function writeSelectedPresetId(
  storage: { setItem(key: string, value: string): void; removeItem(key: string): void },
  sessionId: string,
  sheetId: string,
  presetId: string,
): void {
  const key = presetStorageKey(sessionId, sheetId)
  if (!presetId) storage.removeItem(key)
  else storage.setItem(key, presetId)
}

function nameKey(name: string): string {
  return name.trim().toLocaleLowerCase('vi')
}

export function validatePresetDraft(
  presets: FormPreset[],
  name: string,
  visibleKeys: string[],
  exceptId?: string,
): PresetDraftError | null {
  if (!nameKey(name)) return 'empty-name'
  if (visibleKeys.length === 0) return 'no-columns'
  const taken = presets.some((preset) => preset.id !== exceptId && nameKey(preset.name) === nameKey(name))
  if (taken) return 'duplicate-name'
  return null
}

export function headersForPreset<T extends { key: string; visible?: boolean }>(
  headers: T[],
  visibleKeys: string[] | null,
): T[] {
  const shown = headers.filter((header) => header.visible !== false)
  if (visibleKeys == null) return shown
  const allowed = new Set(visibleKeys)
  return shown.filter((header) => allowed.has(header.key))
}

export function prunePreset(
  preset: FormPreset,
  headers: { key: string; visible?: boolean }[],
): FormPreset {
  const allowed = new Set(headers.filter((header) => header.visible !== false).map((header) => header.key))
  return { ...preset, visibleKeys: preset.visibleKeys.filter((key) => allowed.has(key)) }
}

export function presetsForLiveSheets(presets: FormPreset[], sheetIds: string[]): FormPreset[] {
  const live = new Set(sheetIds)
  return presets.filter((preset) => live.has(preset.sheetId))
}

export function resolveSelectedPresetId(storedId: string, presets: FormPreset[]): string {
  if (!storedId) return ''
  return presets.some((preset) => preset.id === storedId) ? storedId : ''
}
