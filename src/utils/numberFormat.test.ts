import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { categoryOf, formatFromDraft, formatNumber, numberCode } from './numberFormat.ts'

describe('number format', () => {
  it('shows grouped decimals and percents', () => {
    assert.equal(formatNumber(1234.5, numberCode(2, true)), '1.234,50')
    assert.equal(formatNumber(1234.5, numberCode(0, false)), '1235')
    assert.equal(formatNumber(0.125, '0.0%'), '12,5%')
  })

  it('maps a column onto a format-cells category', () => {
    assert.equal(categoryOf({ dataType: 'date', dateFormat: 'dd/MM/yyyy' }), 'date')
    assert.equal(categoryOf({ dataType: 'number', numFmt: '0%' }), 'percentage')
    assert.equal(categoryOf({ dataType: 'text' }), 'general')
  })

  it('writes the category back onto the column', () => {
    assert.deepEqual(formatFromDraft({
      category: 'number',
      decimals: 2,
      thousands: true,
      dateFormat: 'dd/MM/yyyy',
      custom: '',
    }), { dataType: 'number', numFmt: '#,##0.00' })
    assert.deepEqual(formatFromDraft({
      category: 'custom',
      decimals: 0,
      thousands: false,
      dateFormat: '',
      custom: 'yyyy-MM-dd',
    }), { dataType: 'date', dateFormat: 'yyyy-MM-dd' })
  })
})
