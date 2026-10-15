import { defineStore } from 'pinia'
import type { Toast, AdvancedFilterRule, FilterJoin, FindReplaceOptions } from '@/types/excel'

export interface ListWindow {
  filtering: boolean
  from: number
  to: number
  total: number
}

interface AppConfigState {
  viewMode: 'editor' | 'preview'
  phoneScreen: 'table' | 'form'
  splitRatio: number
  searchQuery: string
  filterColumn: string
  advancedRules: AdvancedFilterRule[]
  advancedJoin: FilterJoin
  pendingReplace: FindReplaceOptions | null
  toasts: Toast[]
  listWindow: ListWindow | null
}

export const useAppConfigStore = defineStore('appConfig', {
  state: (): AppConfigState => ({
    viewMode: 'editor',
    phoneScreen: 'table',
    splitRatio: 55,
    searchQuery: '',
    filterColumn: '',
    advancedRules: [],
    advancedJoin: 'and',
    pendingReplace: null,
    toasts: [],
    listWindow: null,
  }),

  getters: {
    hasSearch: (state) => Boolean(state.searchQuery.trim()),
  },

  actions: {
    setViewMode(mode: 'editor' | 'preview') {
      this.viewMode = mode
    },

    setPhoneScreen(screen: 'table' | 'form') {
      this.phoneScreen = screen
      this.viewMode = 'editor'
    },

    setSplitRatio(ratio: number) {
      this.splitRatio = Math.max(20, Math.min(80, ratio))
    },

    setListWindow(window: ListWindow | null) {
      this.listWindow = window
    },

    setSearch(query: string, column: string = '') {
      this.searchQuery = query
      this.filterColumn = column
      this.pendingReplace = null
    },

    applyFindFilter(options: FindReplaceOptions) {
      this.searchQuery = options.findText
      this.filterColumn = options.columns[0] ?? ''
      this.pendingReplace = {
        ...options,
        columns: [...options.columns],
      }
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
