import assert from 'node:assert/strict'
import test from 'node:test'
import { supportPayload } from './supportRequest.ts'

test('attaches a local file under the discord limit', () => {
  const result = supportPayload({
    title: 'Lỗi lưu',
    description: 'Không ghi được dòng',
    email: 'a@b.c',
    fileName: 'bang.sqlite',
    fileBytes: 1200,
  })
  assert.equal(result.attach, true)
  assert.match(result.payload, /bang\.sqlite/)
})

test('skips a file over 25 MB', () => {
  const result = supportPayload({
    title: 'Lỗi',
    description: 'Nặng',
    email: 'a@b.c',
    fileName: 'bang.sqlite',
    fileBytes: 26 * 1024 * 1024,
  })
  assert.equal(result.attach, false)
  assert.match(result.payload, /longdh2\.dev@gmail\.com/)
})
