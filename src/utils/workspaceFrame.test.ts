import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import {
  BATCH_BAR_BOTTOM_PX,
  allowInlineCellEdit,
  clampFormWidth,
  menuWidthPx,
  nextEditorVisibility,
  panesForTier,
  workspaceTier,
} from './workspaceFrame.ts'

describe('workspace frame', () => {
  it('splits phone, tablet, and desktop on the Tailwind breakpoints', () => {
    assert.equal(workspaceTier(767), 'phone')
    assert.equal(workspaceTier(768), 'tablet')
    assert.equal(workspaceTier(1023), 'tablet')
    assert.equal(workspaceTier(1024), 'desktop')
  })

  it('shows the table when a phone has both panes or neither', () => {
    assert.deepEqual(panesForTier('phone', { form: true, table: true }), { form: false, table: true })
    assert.deepEqual(panesForTier('phone', { form: false, table: false }), { form: false, table: true })
    assert.deepEqual(panesForTier('phone', { form: true, table: false }), { form: true, table: false })
  })

  it('keeps a visible pane when hiding the other on a phone', () => {
    assert.deepEqual(
      nextEditorVisibility('phone', { form: false, table: true }, 'hide-form'),
      { form: false, table: true },
    )
    assert.deepEqual(
      nextEditorVisibility('phone', { form: false, table: true }, 'select-row'),
      { form: true, table: false },
    )
  })

  it('hides the form on desktop and leaves the table', () => {
    assert.deepEqual(
      nextEditorVisibility('desktop', { form: true, table: true }, 'hide-form'),
      { form: false, table: true },
    )
  })

  it('clamps the form width', () => {
    assert.equal(clampFormWidth('tablet', 800, 100), 280)
    assert.equal(clampFormWidth('tablet', 800, 500), 336)
    assert.equal(clampFormWidth('desktop', 1000, 100), 340)
    assert.equal(clampFormWidth('desktop', 1000, 900), 500)
  })

  it('blocks inline cell edit on a phone', () => {
    assert.equal(allowInlineCellEdit('phone'), false)
    assert.equal(allowInlineCellEdit('tablet'), true)
    assert.equal(allowInlineCellEdit('desktop'), true)
  })

  it('keeps menus inside the viewport and at least 160px', () => {
    assert.equal(menuWidthPx(360, 320), 320)
    assert.equal(menuWidthPx(100, 320), 160)
    assert.equal(menuWidthPx(200, 320), 184)
  })

  it('places the batch bar above the sheet tabs and status bar', () => {
    assert.equal(BATCH_BAR_BOTTOM_PX, 60)
  })
})
