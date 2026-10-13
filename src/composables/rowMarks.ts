import { ref } from 'vue'

export const MARK_COLORS = [
  { id: 'amber', label: 'common.amber' },
  { id: 'emerald', label: 'common.green' },
  { id: 'sky', label: 'common.blue' },
  { id: 'rose', label: 'common.red' },
  { id: 'violet', label: 'common.purple' },
] as const

export type MarkColor = (typeof MARK_COLORS)[number]['id']

const DEFAULT_COLOR: MarkColor = 'amber'
const marksBySheet = new Map<string, Map<number, MarkColor>>()
const tick = ref(0)

/** Theo dõi thay đổi dấu để lưu phiên SQLite */
export const rowMarksRevision = tick

export interface StoredRowMark {
  sheetId: string
  sheetName: string
  rowIndex: number
  color: MarkColor
}

export function clearAllRowMarks() {
  marksBySheet.clear()
  touch()
}

export function collectRowMarksForSave(sheets: { id: string; name: string }[]): StoredRowMark[] {
  void tick.value
  const nameById = new Map(sheets.map((sheet) => [sheet.id, sheet.name]))
  const allowed = new Set(nameById.keys())
  const out: StoredRowMark[] = []
  for (const [sheetId, marks] of marksBySheet) {
    if (!allowed.has(sheetId)) continue
    const sheetName = nameById.get(sheetId) ?? sheetId
    for (const [rowIndex, color] of marks) {
      out.push({ sheetId, sheetName, rowIndex, color })
    }
  }
  return out
}

/** Ghép dấu đã lưu với sheet hiện tại (ưu tiên id, fallback theo tên sheet). */
export function applyStoredRowMarks(
  sheets: { id: string; name: string }[],
  stored: StoredRowMark[],
) {
  marksBySheet.clear()
  if (stored.length === 0 || sheets.length === 0) {
    touch()
    return
  }

  const bySheetId = new Map<string, Map<number, MarkColor>>()
  const bySheetName = new Map<string, Map<number, MarkColor>>()

  for (const row of stored) {
    if (!isMarkColor(row.color)) continue
    let byId = bySheetId.get(row.sheetId)
    if (!byId) {
      byId = new Map()
      bySheetId.set(row.sheetId, byId)
    }
    byId.set(row.rowIndex, row.color)

    let byName = bySheetName.get(row.sheetName)
    if (!byName) {
      byName = new Map()
      bySheetName.set(row.sheetName, byName)
    }
    byName.set(row.rowIndex, row.color)
  }

  for (const sheet of sheets) {
    const source = bySheetId.get(sheet.id) ?? bySheetName.get(sheet.name)
    if (!source || source.size === 0) continue
    marksBySheet.set(sheet.id, new Map(source))
  }
  touch()
}

function marksOf(sheetId: string): Map<number, MarkColor> {
  let marks = marksBySheet.get(sheetId)
  if (!marks) {
    marks = new Map()
    marksBySheet.set(sheetId, marks)
  }
  return marks
}

function touch() {
  tick.value += 1
}

function isMarkColor(value: string): value is MarkColor {
  return MARK_COLORS.some((item) => item.id === value)
}

function shifted(index: number, from: number, to: number): number {
  if (index === from) return to
  if (from < to && index > from && index <= to) return index - 1
  if (to < from && index >= to && index < from) return index + 1
  return index
}

export function useRowMarks() {
  function colorOf(sheetId: string, index: number): MarkColor | null {
    void tick.value
    return marksBySheet.get(sheetId)?.get(index) ?? null
  }

  function isMarked(sheetId: string, index: number): boolean {
    return colorOf(sheetId, index) !== null
  }

  function count(sheetId: string): number {
    void tick.value
    return marksBySheet.get(sheetId)?.size ?? 0
  }

  function list(sheetId: string): { index: number; color: MarkColor }[] {
    void tick.value
    const marks = marksBySheet.get(sheetId)
    if (!marks) return []
    return [...marks.entries()]
      .map(([index, color]) => ({ index, color }))
      .sort((left, right) => left.index - right.index)
  }

  function setMark(sheetId: string, index: number, color: string) {
    marksOf(sheetId).set(index, isMarkColor(color) ? color : DEFAULT_COLOR)
    touch()
  }

  function removeMark(sheetId: string, index: number) {
    const marks = marksBySheet.get(sheetId)
    if (!marks?.delete(index)) return
    touch()
  }

  function toggle(sheetId: string, index: number, color: string = DEFAULT_COLOR) {
    if (isMarked(sheetId, index)) removeMark(sheetId, index)
    else setMark(sheetId, index, color)
  }

  function nextAfter(sheetId: string, current: number): number | null {
    const items = list(sheetId)
    if (items.length === 0) return null
    return items.find((item) => item.index > current)?.index ?? items[0].index
  }

  function noteRowInserted(sheetId: string, at: number) {
    const marks = marksBySheet.get(sheetId)
    if (!marks || marks.size === 0) return
    const next = new Map<number, MarkColor>()
    for (const [index, color] of marks) next.set(index >= at ? index + 1 : index, color)
    marksBySheet.set(sheetId, next)
    touch()
  }

  /**
   * `swapTail`: file lớn không có thứ tự riêng. Dòng cuối được đưa vào chỗ vừa xóa,
   * nên dấu của dòng cuối đi theo nội dung đó.
   */
  function noteRowsDeleted(sheetId: string, indexes: number[], swapTail = false, rowCount = 0) {
    const marks = marksBySheet.get(sheetId)
    if (!marks || marks.size === 0 || indexes.length === 0) return
    if (swapTail) {
      let length = rowCount
      const removed = [...new Set(indexes)].sort((left, right) => right - left)
      for (const index of removed) {
        const last = length - 1
        if (index < 0 || index > last) continue
        const tailColor = index !== last ? marks.get(last) : undefined
        marks.delete(index)
        if (tailColor) {
          marks.delete(last)
          marks.set(index, tailColor)
        }
        length -= 1
      }
      touch()
      return
    }
    const removed = [...new Set(indexes)].sort((left, right) => left - right)
    const dropped = new Set(removed)
    const next = new Map<number, MarkColor>()
    for (const [index, color] of marks) {
      if (dropped.has(index)) continue
      const shift = removed.filter((item) => item < index).length
      next.set(index - shift, color)
    }
    marksBySheet.set(sheetId, next)
    touch()
  }

  function noteRowMoved(sheetId: string, from: number, to: number) {
    const marks = marksBySheet.get(sheetId)
    if (!marks || marks.size === 0 || from === to) return
    const next = new Map<number, MarkColor>()
    for (const [index, color] of marks) next.set(shifted(index, from, to), color)
    marksBySheet.set(sheetId, next)
    touch()
  }

  return {
    isMarked,
    colorOf,
    count,
    list,
    setMark,
    removeMark,
    toggle,
    nextAfter,
    noteRowInserted,
    noteRowsDeleted,
    noteRowMoved,
  }
}
