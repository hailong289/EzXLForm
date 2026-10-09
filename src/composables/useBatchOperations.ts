import type {
  RowData,
  CellValue,
  BatchOperation,
  FindReplaceOptions,
  FindReplaceResult,
  SearchFilter,
} from '@/types/excel'

export function useBatchOperations() {
  /**
   * Áp dụng thao tác hàng loạt trên các dòng được chỉ định
   */
  function applyBatchOperation(rows: RowData[], operation: BatchOperation): RowData[] {
    const newRows = rows.map((r) => ({ ...r }))
    const targetSet = new Set(operation.targetRowIndexes)
    const { columnKey, type, value } = operation

    targetSet.forEach((idx) => {
      if (idx < 0 || idx >= newRows.length) return
      const currentVal = newRows[idx][columnKey]
      const strVal = currentVal != null ? String(currentVal) : ''

      let updatedVal: CellValue = currentVal

      switch (type) {
        case 'SET_VALUE':
          updatedVal = value ?? ''
          break
        case 'ADD_PREFIX':
          updatedVal = (value ?? '') + strVal
          break
        case 'ADD_SUFFIX':
          updatedVal = strVal + (value ?? '')
          break
        case 'UPPERCASE':
          updatedVal = strVal.toUpperCase()
          break
        case 'LOWERCASE':
          updatedVal = strVal.toLowerCase()
          break
        case 'CAPITALIZE':
          updatedVal = strVal
            .toLowerCase()
            .split(' ')
            .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
            .join(' ')
          break
        case 'TRIM':
          updatedVal = strVal.trim()
          break
        case 'CLEAR':
          updatedVal = null
          break
      }

      newRows[idx][columnKey] = updatedVal
    })

    return newRows
  }

  /**
   * Tìm kiếm các dòng khớp với FindReplaceOptions
   */
  function findMatches(rows: RowData[], options: FindReplaceOptions): number[] {
    if (!options.findText) return []

    let regex: RegExp
    try {
      if (options.useRegex) {
        regex = new RegExp(options.findText, options.caseSensitive ? 'g' : 'gi')
      } else {
        const escaped = options.findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
        const pattern = options.exactMatch ? `^${escaped}$` : escaped
        regex = new RegExp(pattern, options.caseSensitive ? 'g' : 'gi')
      }
    } catch {
      return []
    }

    const matchedIndexes: number[] = []

    rows.forEach((row, idx) => {
      const keys = options.columns.length > 0 ? options.columns : Object.keys(row)
      const matched = keys.some((k) => {
        const val = row[k]
        if (val == null) return false
        regex.lastIndex = 0
        return regex.test(String(val))
      })

      if (matched) {
        matchedIndexes.push(idx)
      }
    })

    return matchedIndexes
  }

  /**
   * Thực hiện Find & Replace trên tập rows
   */
  function performReplace(
    rows: RowData[],
    options: FindReplaceOptions,
  ): { rows: RowData[]; result: FindReplaceResult } {
    const result: FindReplaceResult = {
      matchCount: 0,
      replacedCount: 0,
      affectedRows: [],
    }

    if (!options.findText) return { rows, result }

    let regex: RegExp
    try {
      if (options.useRegex) {
        regex = new RegExp(options.findText, options.caseSensitive ? 'g' : 'gi')
      } else {
        const escaped = options.findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
        const pattern = options.exactMatch ? `^${escaped}$` : escaped
        regex = new RegExp(pattern, options.caseSensitive ? 'g' : 'gi')
      }
    } catch {
      return { rows, result }
    }

    const affectedSet = new Set<number>()
    const newRows = rows.map((row, idx) => {
      const updatedRow = { ...row }
      const keys = options.columns.length > 0 ? options.columns : Object.keys(row)
      let rowChanged = false

      keys.forEach((k) => {
        if (!(k in updatedRow)) return
        const val = updatedRow[k]
        if (val == null) return

        const strVal = String(val)
        regex.lastIndex = 0
        const matches = strVal.match(regex)
        if (!matches) return

        result.matchCount += matches.length
        regex.lastIndex = 0
        const replaced = strVal.replace(regex, options.replaceText)
        if (replaced !== strVal) {
          updatedRow[k] = replaced
          result.replacedCount += matches.length
          rowChanged = true
        }
      })

      if (rowChanged) {
        affectedSet.add(idx)
      }

      return updatedRow
    })

    result.affectedRows = Array.from(affectedSet).sort((a, b) => a - b)
    return { rows: newRows, result }
  }

  /**
   * Lọc dòng theo SearchFilter (dùng cho thanh tìm kiếm thông thường)
   */
  function filterRows(rows: RowData[], filter: SearchFilter): number[] {
    if (!filter.query.trim()) {
      return rows.map((_, i) => i)
    }

    const query = filter.caseSensitive ? filter.query : filter.query.toLowerCase()
    const matched: number[] = []

    rows.forEach((row, idx) => {
      const keys = filter.columns.length > 0 ? filter.columns : Object.keys(row)
      const found = keys.some((k) => {
        const val = row[k]
        if (val == null) return false
        const strVal = filter.caseSensitive ? String(val) : String(val).toLowerCase()
        return filter.exactMatch ? strVal === query : strVal.includes(query)
      })

      if (found) {
        matched.push(idx)
      }
    })

    return matched
  }

  return {
    applyBatchOperation,
    findMatches,
    performReplace,
    filterRows,
  }
}
