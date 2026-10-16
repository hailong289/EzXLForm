/**
 * useHistory.ts
 * Composable quản lý Undo/Redo history cho Excel Editor.
 * Singleton pattern: state được chia sẻ toàn bộ ứng dụng.
 */

import { computed, reactive } from 'vue'
import type { ComputedRef } from 'vue'
import { useWorkbookStore } from '@/stores/workbookStore'
import type { ColumnHeader, HistoryEntry } from '@/types/excel'
import { activityLines } from '@/utils/activityDetail'

interface ActivityItem {
  id: string
  at: number
  kind: 'do' | 'undo' | 'redo'
  text: string
  details: string[]
}

interface HistoryState {
  history: HistoryEntry[]
  currentIndex: number
  maxHistory: number
  activity: ActivityItem[]
}

const state = reactive<HistoryState>({
  history: [],
  currentIndex: -1,
  maxHistory: 100,
  activity: [],
})

function headersFor(sheetId: string): ColumnHeader[] {
  const workbook = useWorkbookStore().workbook
  return workbook?.sheets.find((sheet) => sheet.id === sheetId)?.headers ?? []
}

function noteActivity(kind: 'do' | 'undo' | 'redo', entry: HistoryEntry): void {
  state.activity.unshift({
    id: `act_${Date.now()}_${state.activity.length}`,
    at: Date.now(),
    kind,
    text: entry.description,
    details: activityLines(entry, headersFor(entry.sheetId)),
  })
  if (state.activity.length > 80) state.activity.pop()
}

function generateId(): string {
  return `hist_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`
}

export function useHistory() {
  const canUndo: ComputedRef<boolean> = computed(
    () => state.currentIndex >= 0,
  )

  const canRedo: ComputedRef<boolean> = computed(
    () => state.currentIndex < state.history.length - 1,
  )

  const historyLength: ComputedRef<number> = computed(
    () => state.history.length,
  )

  const currentHistoryIndex: ComputedRef<number> = computed(
    () => state.currentIndex,
  )

  const undoStackSize: ComputedRef<number> = computed(
    () => state.currentIndex + 1,
  )

  const redoStackSize: ComputedRef<number> = computed(
    () => Math.max(0, state.history.length - 1 - state.currentIndex),
  )

  const undoDescription: ComputedRef<string> = computed(() => {
    if (!canUndo.value) return ''
    return state.history[state.currentIndex]?.description ?? ''
  })

  const redoDescription: ComputedRef<string> = computed(() => {
    if (!canRedo.value) return ''
    return state.history[state.currentIndex + 1]?.description ?? ''
  })

  function push(entry: Omit<HistoryEntry, 'id' | 'timestamp'>): void {
    if (state.currentIndex < state.history.length - 1) {
      state.history.splice(state.currentIndex + 1)
    }

    const newEntry: HistoryEntry = {
      ...entry,
      id: generateId(),
      timestamp: Date.now(),
    }

    state.history.push(newEntry)

    if (state.history.length > state.maxHistory) {
      state.history.splice(0, state.history.length - state.maxHistory)
    }

    state.currentIndex = state.history.length - 1
    noteActivity('do', newEntry)
  }

  function undo(): HistoryEntry | null {
    if (!canUndo.value) return null
    const entry = state.history[state.currentIndex]
    state.currentIndex -= 1
    noteActivity('undo', entry)
    return entry
  }

  function redo(): HistoryEntry | null {
    if (!canRedo.value) return null
    state.currentIndex += 1
    const entry = state.history[state.currentIndex]
    noteActivity('redo', entry)
    return entry
  }

  function clear(): void {
    state.history.splice(0)
    state.currentIndex = -1
  }

  function getUndoDescription(): string {
    return undoDescription.value
  }

  function getRedoDescription(): string {
    return redoDescription.value
  }

  const activity = computed(() => state.activity)

  return {
    canUndo,
    canRedo,
    historyLength,
    currentHistoryIndex,
    undoStackSize,
    redoStackSize,
    undoDescription,
    redoDescription,
    activity,
    push,
    undo,
    redo,
    clear,
    getUndoDescription,
    getRedoDescription,
  }
}
