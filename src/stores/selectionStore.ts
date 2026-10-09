import { defineStore } from 'pinia'

interface SelectionState {
  selectedRowIndexes: number[]
  activeRowIndex: number | null
  activeColumnKey: string | null
  editingCell: { rowIndex: number; columnKey: string } | null
}

export const useSelectionStore = defineStore('selection', {
  state: (): SelectionState => ({
    selectedRowIndexes: [],
    activeRowIndex: null,
    activeColumnKey: null,
    editingCell: null,
  }),

  getters: {
    selectedCount: (state) => state.selectedRowIndexes.length,
    hasSelection: (state) => state.selectedRowIndexes.length > 0,
    isRowSelected: (state) => (rowIndex: number) => state.selectedRowIndexes.includes(rowIndex),
  },

  actions: {
    toggleRowSelection(rowIndex: number) {
      const idx = this.selectedRowIndexes.indexOf(rowIndex)
      if (idx === -1) {
        this.selectedRowIndexes.push(rowIndex)
      } else {
        this.selectedRowIndexes.splice(idx, 1)
      }
    },

    selectRow(rowIndex: number) {
      if (!this.selectedRowIndexes.includes(rowIndex)) {
        this.selectedRowIndexes.push(rowIndex)
      }
    },

    deselectRow(rowIndex: number) {
      const idx = this.selectedRowIndexes.indexOf(rowIndex)
      if (idx !== -1) {
        this.selectedRowIndexes.splice(idx, 1)
      }
    },

    selectAll(totalRows: number) {
      this.selectedRowIndexes = Array.from({ length: totalRows }, (_, i) => i)
    },

    clearSelection() {
      this.selectedRowIndexes = []
      this.activeRowIndex = null
      this.editingCell = null
    },

    invertSelection(totalRows: number) {
      const current = new Set(this.selectedRowIndexes)
      const inverted: number[] = []
      for (let i = 0; i < totalRows; i++) {
        if (!current.has(i)) {
          inverted.push(i)
        }
      }
      this.selectedRowIndexes = inverted
    },

    setEditingCell(rowIndex: number | null, columnKey: string | null) {
      if (rowIndex !== null && columnKey !== null) {
        this.editingCell = { rowIndex, columnKey }
      } else {
        this.editingCell = null
      }
    },

    setActiveRowIndex(index: number | null) {
      this.activeRowIndex = index
    },

    setActiveColumnKey(key: string | null) {
      this.activeColumnKey = key
    },
  },
})
