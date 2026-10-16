import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { applyFormulas, evalFormula, ruleError } from './columnRules.ts'
import { coerceRowData } from './cellValue.ts'
import type { ColumnHeader } from '@/types/excel'

function header(patch: Partial<ColumnHeader>): ColumnHeader {
  return { key: 'code', label: 'Mã', originalIndex: 0, dataType: 'text', ...patch }
}

describe('column rules', () => {
  it('blocks required, regex, and duplicates', () => {
    assert.equal(ruleError(header({ required: true }), ''), 'Bắt buộc')
    assert.equal(ruleError(header({ pattern: '^A\\d+$' }), 'B1'), 'Không đúng mẫu')
    assert.equal(ruleError(header({ pattern: '^A\\d+$' }), 'A1'), null)
    assert.equal(ruleError(header({ unique: true }), 'NV01', ['NV01']), 'Đã có giá trị này')
    const { errors } = coerceRowData(
      { code: '' },
      [header({ key: 'code', required: true })],
    )
    assert.equal(errors.code, 'Bắt buộc')
  })

  it('evaluates column arithmetic', () => {
    assert.equal(evalFormula('qty * price', (name) => (name === 'qty' ? 2 : 10)).value, 20)
    assert.equal(evalFormula('(qty + 1) * 2', () => 3).value, 8)
    const row = applyFormulas(
      { qty: 2, price: 10, total: null },
      [
        header({ key: 'qty', dataType: 'number' }),
        header({ key: 'price', dataType: 'number' }),
        header({ key: 'total', formula: 'qty * price' }),
      ],
    )
    assert.equal(row.total, 20)
  })
})