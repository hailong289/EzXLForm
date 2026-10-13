import { computed, ref } from 'vue'
import type { ColumnHeader } from '@/types/excel'
import { useSqliteStorage } from '@/composables/useSqliteStorage'
import {
  headersForPreset,
  prunePreset,
  readSelectedPresetId,
  resolveSelectedPresetId,
  validatePresetDraft,
  writeSelectedPresetId,
  type FormPreset,
  type PresetDraftError,
} from '@/utils/formPresets'

const allPresets = ref<FormPreset[]>([])
const sheetId = ref('')
const headers = ref<ColumnHeader[]>([])
const selectedPresetId = ref('')

export function useFormPresets() {
  const sqlite = useSqliteStorage()

  const presets = computed(() =>
    allPresets.value
      .filter((preset) => preset.sheetId === sheetId.value)
      .map((preset) => prunePreset(preset, headers.value)),
  )

  const activeKeys = computed(() => {
    const selected = presets.value.find((preset) => preset.id === selectedPresetId.value)
    return selected ? selected.visibleKeys : null
  })

  const shownHeaders = computed(() => headersForPreset(headers.value, activeKeys.value))
  const presetEmpty = computed(() => selectedPresetId.value !== '' && shownHeaders.value.length === 0)

  async function persist(next: FormPreset[]): Promise<void> {
    const sessionId = sqlite.activeSessionId.value
    if (!sessionId) return
    allPresets.value = next
    await sqlite.saveFormPresets(sessionId, next)
  }

  async function loadForSheet(nextSheetId: string, nextHeaders: ColumnHeader[]): Promise<void> {
    sheetId.value = nextSheetId
    headers.value = nextHeaders
    const sessionId = sqlite.activeSessionId.value
    if (!sessionId) {
      allPresets.value = []
      selectedPresetId.value = ''
      return
    }
    allPresets.value = await sqlite.listFormPresets(sessionId)
    const stored = readSelectedPresetId(localStorage, sessionId, nextSheetId)
    selectedPresetId.value = resolveSelectedPresetId(
      stored,
      allPresets.value.filter((preset) => preset.sheetId === nextSheetId),
    )
  }

  function selectPreset(presetId: string): void {
    const sessionId = sqlite.activeSessionId.value
    selectedPresetId.value = resolveSelectedPresetId(
      presetId,
      presets.value,
    )
    if (sessionId && sheetId.value) {
      writeSelectedPresetId(localStorage, sessionId, sheetId.value, selectedPresetId.value)
    }
  }

  async function createPreset(name: string, visibleKeys: string[]): Promise<PresetDraftError | null> {
    const error = validatePresetDraft(presets.value, name, visibleKeys)
    if (error || !sheetId.value) return error ?? 'empty-name'
    const preset: FormPreset = {
      id: crypto.randomUUID(),
      sheetId: sheetId.value,
      name: name.trim(),
      visibleKeys,
    }
    await persist([...allPresets.value, preset])
    selectPreset(preset.id)
    return null
  }

  async function updatePreset(id: string, name: string, visibleKeys: string[]): Promise<PresetDraftError | null> {
    const error = validatePresetDraft(presets.value, name, visibleKeys, id)
    if (error) return error
    await persist(allPresets.value.map((preset) => (
      preset.id === id ? { ...preset, name: name.trim(), visibleKeys } : preset
    )))
    return null
  }

  async function deletePreset(id: string): Promise<void> {
    await persist(allPresets.value.filter((preset) => preset.id !== id))
    if (selectedPresetId.value === id) selectPreset('')
  }

  async function deletePresetsForSheet(targetSheetId: string): Promise<void> {
    const sessionId = sqlite.activeSessionId.value
    if (!sessionId) return
    if (allPresets.value.length === 0) allPresets.value = await sqlite.listFormPresets(sessionId)
    await persist(allPresets.value.filter((preset) => preset.sheetId !== targetSheetId))
    writeSelectedPresetId(localStorage, sessionId, targetSheetId, '')
    if (sheetId.value === targetSheetId) selectedPresetId.value = ''
  }

  return {
    presets,
    selectedPresetId,
    shownHeaders,
    presetEmpty,
    loadForSheet,
    selectPreset,
    createPreset,
    updatePreset,
    deletePreset,
    deletePresetsForSheet,
  }
}
