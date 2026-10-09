/**
 * useHistory.ts
 * Composable quản lý Undo/Redo history cho Excel Editor.
 * Singleton pattern: state được chia sẻ toàn bộ ứng dụng.
 */

import { computed, reactive } from 'vue'
import type { ComputedRef } from 'vue'
import type { HistoryEntry } from '@/types/excel'

interface HistoryState {
  history: HistoryEntry[]
  currentIndex: number
  maxHistory: number
}

const state = reactive<HistoryState>({
  history: [],
  currentIndex: -1,
  maxHistory: 100,
})

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
  }

  function undo(): HistoryEntry | null {
    if (!canUndo.value) return null
    const entry = state.history[state.currentIndex]
    state.currentIndex -= 1
    return entry
  }

  function redo(): HistoryEntry | null {
    if (!canRedo.value) return null
    state.currentIndex += 1
    const entry = state.history[state.currentIndex]
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

  return {
    canUndo,
    canRedo,
    historyLength,
    currentHistoryIndex,
    undoStackSize,
    redoStackSize,
    undoDescription,
    redoDescription,
    push,
    undo,
    redo,
    clear,
    getUndoDescription,
    getRedoDescription,
  }
}
