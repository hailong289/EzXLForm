import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import {
  EMPTY_PRESET_MESSAGE,
  headersForPreset,
  presetStorageKey,
  presetsForLiveSheets,
  prunePreset,
  readSelectedPresetId,
  resolveSelectedPresetId,
  validatePresetDraft,
  writeSelectedPresetId,
  type FormPreset,
} from './formPresets.ts'

const goods: FormPreset = {
  id: 'p1',
  sheetId: 'sheet-a',
  name: 'Nhập hàng hóa',
  visibleKeys: ['sku', 'qty', 'gone'],
}

describe('form presets', () => {
  it('builds the localStorage key from the spec', () => {
    assert.equal(presetStorageKey('sess', 'sheet-a'), 'excel_editor_form_preset:sess:sheet-a')
  })

  it('treats a missing or blank stored id as all columns', () => {
    const storage = { getItem: () => null }
    assert.equal(readSelectedPresetId(storage, 'sess', 'sheet-a'), '')
  })

  it('removes the key when the selection is all columns', () => {
    const bag = new Map<string, string>()
    const storage = {
      setItem: (key: string, value: string) => bag.set(key, value),
      removeItem: (key: string) => bag.delete(key),
    }
    writeSelectedPresetId(storage, 'sess', 'sheet-a', 'p1')
    assert.equal(bag.get('excel_editor_form_preset:sess:sheet-a'), 'p1')
    writeSelectedPresetId(storage, 'sess', 'sheet-a', '')
    assert.equal(bag.has('excel_editor_form_preset:sess:sheet-a'), false)
  })

  it('rejects an empty name, a duplicate name, and a preset with no columns', () => {
    assert.equal(validatePresetDraft([], '   ', ['sku']), 'empty-name')
    assert.equal(validatePresetDraft([goods], ' nhập Hàng Hóa ', ['sku']), 'duplicate-name')
    assert.equal(validatePresetDraft([goods], 'Nhập hàng hóa', ['sku'], 'p1'), null)
    assert.equal(validatePresetDraft([], 'Kho', []), 'no-columns')
  })

  it('shows selected columns in table order and skips columns hidden on the table', () => {
    const headers = [
      { key: 'name', visible: true },
      { key: 'sku', visible: true },
      { key: 'secret', visible: false },
      { key: 'qty', visible: true },
    ]
    assert.deepEqual(
      headersForPreset(headers, ['qty', 'secret', 'sku']).map((header) => header.key),
      ['sku', 'qty'],
    )
    assert.deepEqual(
      headersForPreset(headers, null).map((header) => header.key),
      ['name', 'sku', 'qty'],
    )
  })

  it('drops deleted or table-hidden keys and does not add new columns', () => {
    const pruned = prunePreset(goods, [
      { key: 'sku', visible: true },
      { key: 'qty', visible: true },
      { key: 'note', visible: true },
      { key: 'gone', visible: false },
    ])
    assert.deepEqual(pruned.visibleKeys, ['sku', 'qty'])
  })

  it('drops presets whose sheet was deleted', () => {
    assert.deepEqual(
      presetsForLiveSheets([goods, { ...goods, id: 'p2', sheetId: 'sheet-b' }], ['sheet-a']).map((preset) => preset.id),
      ['p1'],
    )
  })

  it('falls back to all columns when the stored id is gone', () => {
    assert.equal(resolveSelectedPresetId('missing', [goods]), '')
    assert.equal(resolveSelectedPresetId('p1', [goods]), 'p1')
  })

  it('uses the exact empty-preset sentence', () => {
    assert.equal(EMPTY_PRESET_MESSAGE, 'Loại form này không còn cột. Hãy sửa lại trong Loại form.')
  })
})
