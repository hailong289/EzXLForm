import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { CUSTOM_FORMATS, accountingCode, categoryOf, currencyCode, formatFromDraft, formatNumber, numberCode, scientificCode } from './numberFormat.ts'

describe('number format', () => {
  it('shows grouped decimals and percents', () => {
    assert.equal(formatNumber(1234.5, numberCode(2, true)), '1.234,50')
    assert.equal(formatNumber(1234.5, numberCode(0, false)), '1235')
    assert.equal(formatNumber(0.125, '0.0%'), '12,5%')
    assert.equal(formatNumber(-1234.5, currencyCode(2, '₫', 'parens')), '(₫1.234,50)')
    assert.equal(formatNumber(0, accountingCode(2, '₫')), '₫-')
    assert.equal(formatNumber(1234.5, scientificCode(2)), '1,23E+03')
    assert.equal(formatNumber(1.5, '# ?/?'), '1 1/2')
    assert.equal(formatNumber(0.5, 'hh:mm'), '12:00')
    assert.equal(formatNumber(-1234.5, '#,##0.00_);[Red](#,##0.00)'), '(1.234,50)')
    assert.ok(CUSTOM_FORMATS.includes('General') && CUSTOM_FORMATS.includes('@'))
  })

  it('maps a column onto a format-cells category', () => {
    assert.equal(categoryOf({ dataType: 'date', dateFormat: 'dd/MM/yyyy' }), 'date')
    assert.equal(categoryOf({ dataType: 'number', numFmt: '0%' }), 'percentage')
    assert.equal(categoryOf({ dataType: 'number', numFmt: currencyCode(0, '$', 'minus') }), 'currency')
    assert.equal(categoryOf({ dataType: 'number', numFmt: accountingCode(2, '₫') }), 'accounting')
    assert.equal(categoryOf({ dataType: 'text' }), 'general')
  })

  it('writes the category back onto the column', () => {
    assert.deepEqual(formatFromDraft({
      category: 'number',
      decimals: 2,
      thousands: true,
      dateFormat: 'dd/MM/yyyy',
      custom: '',
      symbol: '₫',
      negative: 'minus',
      fraction: '# ?/?',
      timeFormat: 'hh:mm',
    }), { dataType: 'number', numFmt: '#,##0.00' })
    assert.deepEqual(formatFromDraft({
      category: 'custom',
      decimals: 0,
      thousands: false,
      dateFormat: '',
      custom: 'yyyy-MM-dd',
      symbol: '',
      negative: 'minus',
      fraction: '',
      timeFormat: '',
    }), { dataType: 'date', dateFormat: 'yyyy-MM-dd' })
  })
})
