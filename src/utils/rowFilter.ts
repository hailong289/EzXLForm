import type { AdvancedFilterRule, CellValue, ColumnHeader, FilterOperator, FindReplaceOptions, RowData } from '@/types/excel'
import { parseDateToIso, formatDateValue } from '@/utils/dateFormat'

export interface FilterOperatorOption {
  id: FilterOperator
  label: string
}

const TEXT_OPERATORS: FilterOperatorOption[] = [
  { id: 'like', label: 'table.op.like' },
  { id: 'not_like', label: 'table.op.not.like' },
  { id: 'equals', label: 'table.op.equals' },
  { id: 'starts', label: 'table.op.starts' },
  { id: 'ends', label: 'table.op.ends' },
  { id: 'empty', label: 'table.op.empty' },
  { id: 'not_empty', label: 'table.op.filled' },
]

const NUMBER_OPERATORS: FilterOperatorOption[] = [
  { id: 'like', label: 'table.op.like' },
  { id: 'equals', label: 'table.op.eq' },
  { id: 'not_equals', label: 'table.op.ne' },
  { id: 'gt', label: 'table.op.gt' },
  { id: 'gte', label: 'table.op.gte' },
  { id: 'lt', label: 'table.op.lt' },
  { id: 'lte', label: 'table.op.lte' },
  { id: 'empty', label: 'table.op.empty' },
  { id: 'not_empty', label: 'table.op.filled' },
]

const DATE_OPERATORS: FilterOperatorOption[] = [
  { id: 'like', label: 'table.op.like' },
  { id: 'equals', label: 'table.op.on.date' },
  { id: 'gt', label: 'table.op.after.date' },
  { id: 'lt', label: 'table.op.before.date' },
  { id: 'empty', label: 'table.op.empty' },
  { id: 'not_empty', label: 'table.op.filled' },
]

const BOOLEAN_OPERATORS: FilterOperatorOption[] = [
  { id: 'equals', label: 'table.op.is' },
  { id: 'empty', label: 'table.op.empty' },
  { id: 'not_empty', label: 'table.op.filled' },
]

export function operatorsFor(header: ColumnHeader | undefined): FilterOperatorOption[] {
  if (header?.dataType === 'number') return NUMBER_OPERATORS
  if (header?.dataType === 'date') return DATE_OPERATORS
  if (header?.dataType === 'boolean') return BOOLEAN_OPERATORS
  return TEXT_OPERATORS
}

export function operatorNeedsValue(operator: FilterOperator): boolean {
  return operator !== 'empty' && operator !== 'not_empty'
}

export function activeFilterRules(rules: AdvancedFilterRule[]): AdvancedFilterRule[] {
  return rules.filter((rule) => rule.columnKey && (!operatorNeedsValue(rule.operator) || rule.value.trim() !== ''))
}

function asNumber(value: CellValue): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value === 'boolean' || value == null || value === '') return null
  const numeric = Number(String(value).replace(/,/g, '').trim())
  return Number.isNaN(numeric) ? null : numeric
}

function likeMatch(text: string, pattern: string): boolean {
  const needle = pattern.trim().toLowerCase()
  if (!needle) return false
  return text.toLowerCase().includes(needle)
}

function isEmpty(value: CellValue): boolean {
  return value == null || value === ''
}

function matchRule(row: RowData, rule: AdvancedFilterRule, header: ColumnHeader | undefined): boolean {
  const current = row[rule.columnKey] ?? null
  if (rule.operator === 'like' || rule.operator === 'not_like') {
    const texts = [String(current ?? '')]
    if (typeof current === 'number') texts.push(current.toLocaleString('vi-VN'))
    if (header?.dataType === 'date') texts.push(formatDateValue(current, header.dateFormat))
    const hit = texts.some((text) => likeMatch(text, rule.value))
    return rule.operator === 'like' ? hit : !hit
  }
  if (rule.operator === 'empty') return isEmpty(current)
  if (rule.operator === 'not_empty') return !isEmpty(current)

  if (header?.dataType === 'number') {
    const left = asNumber(current)
    const right = asNumber(rule.value)
    if (left == null || right == null) return false
    if (rule.operator === 'equals') return left === right
    if (rule.operator === 'not_equals') return left !== right
    if (rule.operator === 'gt') return left > right
    if (rule.operator === 'gte') return left >= right
    if (rule.operator === 'lt') return left < right
    if (rule.operator === 'lte') return left <= right
    return false
  }

  if (header?.dataType === 'date') {
    const left = parseDateToIso(current, header.dateFormat)
    const right = parseDateToIso(rule.value, header.dateFormat)
    if (!left || !right) return false
    if (rule.operator === 'equals') return left === right
    if (rule.operator === 'gt') return left > right
    if (rule.operator === 'lt') return left < right
    return false
  }

  if (header?.dataType === 'boolean') {
    const left = current === true || current === 'true' || current === 1 || current === '1'
    const right = rule.value === 'true'
    return rule.operator === 'equals' ? left === right : false
  }

  const left = String(current ?? '').toLowerCase()
  const right = rule.value.trim().toLowerCase()
  if (rule.operator === 'contains') return left.includes(right)
  if (rule.operator === 'not_contains') return !left.includes(right)
  if (rule.operator === 'equals') return left === right
  if (rule.operator === 'not_equals') return left !== right
  if (rule.operator === 'starts') return left.startsWith(right)
  if (rule.operator === 'ends') return left.endsWith(right)
  return false
}

export function isBlankCell(value: unknown): boolean {
  if (value == null) return true
  return typeof value === 'string' && value.trim() === ''
}

export function buildFindRegex(options: FindReplaceOptions): RegExp | null {
  if (options.matchEmpty || !options.findText) return null
  const flags = options.caseSensitive ? 'g' : 'gi'
  if (options.useRegex) {
    try {
      return new RegExp(options.findText, flags)
    } catch {
      return null
    }
  }
  const escaped = options.findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const pattern = options.exactMatch ? `^${escaped}$` : escaped
  return new RegExp(pattern, flags)
}

export function rowMatchesFind(row: RowData, headers: ColumnHeader[], options: FindReplaceOptions): boolean {
  const keys = options.columns.length > 0 ? options.columns : headers.map((header) => header.key)
  if (options.matchEmpty) return keys.some((key) => isBlankCell(row[key]))
  const regex = buildFindRegex(options)
  if (!regex) return false
  return keys.some((key) => {
    const value = row[key]
    if (value == null || value === '') return false
    regex.lastIndex = 0
    return regex.test(String(value))
  })
}
export function rowMatchesFilters(
  row: RowData,
  headers: ColumnHeader[],
  query: string,
  columnKey: string,
  rules: AdvancedFilterRule[],
  join: 'and' | 'or',
  findOptions?: FindReplaceOptions | null,
): boolean {
  if (findOptions?.findText) {
    if (!rowMatchesFind(row, headers, findOptions)) return false
  } else {
    const needle = query.trim().toLowerCase()
    if (needle) {
      const keys = columnKey ? [columnKey] : headers.map((header) => header.key)
      const hit = keys.some((key) => {
        const value = row[key]
        return value != null && String(value).toLowerCase().includes(needle)
      })
      if (!hit) return false
    }
  }

  const active = activeFilterRules(rules)
  if (active.length === 0) return true
  const checks = active.map((rule) => matchRule(row, rule, headers.find((header) => header.key === rule.columnKey)))
  return join === 'or' ? checks.some(Boolean) : checks.every(Boolean)
}
