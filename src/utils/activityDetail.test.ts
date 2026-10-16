import assert from 'node:assert/strict'
import test from 'node:test'
import { activityLines } from './activityDetail.ts'
import type { ColumnHeader } from '../types/excel.ts'

const headers: ColumnHeader[] = [
  { key: 'name', label: 'Tên', originalIndex: 0 },
  { key: 'qty', label: 'Số lượng', originalIndex: 1 },
]

test('edit lists each changed column', () => {
  const lines = activityLines(
    {
      type: 'EDIT_CELL',
      before: { rows: [{ name: 'An', qty: 1 }] },
      after: { rows: [{ name: 'Bình', qty: 1 }] },
    },
    headers,
  )
  assert.deepEqual(lines, ['Tên: An → Bình'])
})

test('new row lists filled cells', () => {
  const lines = activityLines(
    {
      type: 'ADD_ROW',
      before: {},
      after: { rows: [{ name: 'An', qty: null }] },
    },
    headers,
  )
  assert.deepEqual(lines, ['Tên: An'])
})
