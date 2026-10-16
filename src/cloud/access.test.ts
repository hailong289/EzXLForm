import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { grantsFromGoogle, grantsFromMicrosoft, linkAccess } from './accessMap.ts'

describe('cloud access', () => {
  it('maps google people and leaves the owner locked', () => {
    const grants = grantsFromGoogle([
      { id: '1', type: 'user', role: 'owner', emailAddress: 'me@gmail.com' },
      { id: '2', type: 'user', role: 'reader', emailAddress: 'a@gmail.com' },
      { id: '3', type: 'anyone', role: 'writer' },
    ])
    assert.equal(grants[0]?.editable, false)
    assert.equal(grants[1]?.role, 'reader')
    assert.equal(grants[1]?.editable, true)
    assert.equal(grants[2]?.label, 'Bất kỳ ai có link')
    assert.equal(grants[2]?.link, true)
    assert.equal(linkAccess(grants), 'writer')
  })

  it('maps onedrive invitees', () => {
    const grants = grantsFromMicrosoft([
      { id: 'p1', roles: ['write'], grantedToV2: { user: { email: 'b@contoso.com' } } },
      { id: 'p2', roles: ['owner'], grantedTo: { user: { email: 'me@contoso.com' } } },
    ])
    assert.equal(grants[0]?.role, 'writer')
    assert.equal(grants[0]?.editable, true)
    assert.equal(grants[1]?.editable, false)
  })

  it('treats an anonymous edit link as full', () => {
    const grants = grantsFromMicrosoft([
      { id: 'l1', roles: ['write'], link: { scope: 'anonymous', type: 'edit' } },
    ])
    assert.equal(grants[0]?.link, true)
    assert.equal(linkAccess(grants), 'writer')
  })
})
