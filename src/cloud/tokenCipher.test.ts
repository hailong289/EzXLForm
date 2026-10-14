import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { openToken, sealToken } from './tokenCipher.ts'

describe('token cipher', () => {
  it('round-trips a token and rejects the wrong key', async () => {
    const secret = 'local-test-key'
    const sealed = await sealToken(secret, '{"accessToken":"abc"}')
    assert.equal(sealed.includes('abc'), false)
    assert.equal(await openToken(secret, sealed), '{"accessToken":"abc"}')
    await assert.rejects(() => openToken('other-key', sealed))
  })
})
