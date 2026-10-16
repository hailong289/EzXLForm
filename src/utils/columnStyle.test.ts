import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { columnFormatPayload, columnPaint, readColumnFormat, styleFromDraft } from './columnStyle.ts'

describe('column style', () => {
  it('keeps only the choices that change the cell', () => {
    assert.equal(styleFromDraft({
      align: '',
      vertical: '',
      wrap: false,
      indent: 0,
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      fontSize: 0,
      fontColor: '',
      fontName: '',
      border: '',
      borderColor: '#000000',
      fill: '',
      locked: false,
      hidden: false,
    }), undefined)
    const style = styleFromDraft({
      align: 'center',
      vertical: 'top',
      wrap: true,
      indent: 2,
      bold: true,
      italic: false,
      underline: true,
      strike: false,
      fontSize: 16,
      fontColor: '#ff0000',
      fontName: 'Arial',
      border: 'box',
      borderColor: '#000000',
      fill: '#fff2cc',
      locked: true,
      hidden: false,
    })
    assert.equal(style?.align, 'center')
    assert.equal(style?.locked, true)
    assert.equal(columnPaint(style).text.fontWeight, '700')
    assert.equal(columnPaint(style).box.backgroundColor, '#fff2cc')
    assert.match(columnPaint(style).box.boxShadow ?? '', /inset/)
  })

  it('round-trips column format json', () => {
    const json = columnFormatPayload({ dateFormat: 'dd/MM/yyyy', numFmt: '0.00', style: { bold: true } })
    assert.deepEqual(readColumnFormat(json), { dateFormat: 'dd/MM/yyyy', numFmt: '0.00', style: { bold: true } })
    assert.deepEqual(readColumnFormat(''), {})
  })
})