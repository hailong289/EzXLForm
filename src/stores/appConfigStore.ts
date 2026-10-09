import { defineStore } from 'pinia'
import type { Toast, AdvancedFilterRule, FilterJoin, FindReplaceOptions } from '@/types/excel'

export const PAGE_SIZE_OPTIONS = [20, 50, 100, 200, 500, 1000] as const

const PAGE_SIZE_KEY = 'excel_editor_rows_per_page'

function savedRowsPerPage(): number {
  try {
    const value = Number(localStorage.getItem(PAGE_SIZE_KEY))
    if ((PAGE_SIZE_OPTIONS as readonly number[]).includes(value)) return value
  } catch {
    /* trình duyệt chặn storage thì dùng mặc định */
  }
  return 50
}

interface AppConfigState {
  viewMode: 'editor' | 'preview'
  splitRatio: number
  rowsPerPage: number
  currentPage: number
  searchQuery: string
  filterColumn: string
  advancedRules: AdvancedFilterRule[]
  advancedJoin: FilterJoin
  pendingReplace: FindReplaceOptions | null
  toasts: Toast[]
}

export const useAppConfigStore = defineStore('appConfig', {
  state: (): AppConfigState => ({
    viewMode: 'editor',
    splitRatio: 55,
    rowsPerPage: savedRowsPerPage(),
    currentPage: 1,
    searchQuery: '',
    filterColumn: '',
    advancedRules: [],
    advancedJoin: 'and',
    pendingReplace: null,
    toasts: [],
  }),

  getters: {
    hasSearch: (state) => Boolean(state.searchQuery.trim()),
  },

  actions: {
    setViewMode(mode: 'editor' | 'preview') {
      this.viewMode = mode
    },

    setSplitRatio(ratio: number) {
      this.splitRatio = Math.max(20, Math.min(80, ratio))
    },

    setPage(page: number) {
      this.currentPage = Math.max(1, page)
    },

    setRowsPerPage(n: number) {
      const next = (PAGE_SIZE_OPTIONS as readonly number[]).includes(n) ? n : 50
      const firstRow = (this.currentPage - 1) * this.rowsPerPage
      this.rowsPerPage = next
      this.currentPage = Math.floor(firstRow / next) + 1
      try {
        localStorage.setItem(PAGE_SIZE_KEY, String(next))
      } catch {
        /* bỏ qua nếu không ghi được */
      }
    },

    setSearch(query: string, column: string = '') {
      this.searchQuery = query
      this.filterColumn = column
      this.pendingReplace = null
      this.currentPage = 1
    },

    applyFindFilter(options: FindReplaceOptions) {
      this.searchQuery = options.findText
      this.filterColumn = options.columns[0] ?? ''
      this.pendingReplace = {
        ...options,
        columns: [...options.columns],
      }
      this.currentPage = 1
    },

    clearSearch() {
      this.searchQuery = ''
      this.filterColumn = ''
      this.advancedRules = []
      this.advancedJoin = 'and'
      this.pendingReplace = null
    },

    setAdvancedFilter(rules: AdvancedFilterRule[], join: FilterJoin) {
      this.advancedRules = rules.map((rule) => ({ ...rule }))
      this.advancedJoin = join
      this.currentPage = 1
    },

    addToast(toast: Omit<Toast, 'id'>): string {
      const id = `toast_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`
      const newToast: Toast = { ...toast, id }
      this.toasts.push(newToast)

      const duration = toast.duration ?? 3000
      if (duration > 0) {
        setTimeout(() => {
          this.removeToast(id)
        }, duration)
      }

      return id
    },

    removeToast(id: string) {
      const idx = this.toasts.findIndex((t) => t.id === id)
      if (idx !== -1) {
        this.toasts.splice(idx, 1)
      }
    },

    clearAllToasts() {
      this.toasts = []
    },
  },
})
