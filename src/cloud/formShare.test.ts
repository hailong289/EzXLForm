import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import type { CloudGrant } from './accessMap.ts'
import { canEditCloud, formSnapshot, parseFillHash, parseFormBook } from './formShare.ts'
import type { ColumnHeader } from '../types/excel.ts'

function grant(patch: Partial<CloudGrant>): CloudGrant {
  return { id: '1', label: 'a@b.c', role: 'reader', editable: true, link: false, ...patch }
}

describe('shared form', () => {
  it('keeps only the preset columns', () => {
    const headers = [
      { key: 'name', label: 'Tên', originalIndex: 0, required: true },
      { key: 'note', label: 'Ghi chú', originalIndex: 2 },
    ] as ColumnHeader[]
    const form = formSnapshot('Khách', 0, headers, ['name'])
    assert.equal(form.columns.length, 1)
    assert.equal(form.columns[0]?.index, 0)
    assert.equal(parseFormBook(JSON.stringify(form))?.sheet, 'Khách')
  })

  it('allows the editor and a full link', () => {
    assert.equal(canEditCloud([grant({ role: 'owner', label: 'Me@X.com' })], 'me@x.com'), true)
    assert.equal(canEditCloud([grant({ role: 'reader' })], 'a@b.c'), false)
    assert.equal(canEditCloud([grant({ link: true, role: 'writer', label: 'Bất kỳ ai có link' })], null), true)
  })

  it('reads the fill hash', () => {
    assert.deepEqual(parseFillHash('#fill/google/abc%2F1'), { provider: 'google', fileId: 'abc/1' })
    assert.equal(parseFillHash('#/editor'), null)
  })
})