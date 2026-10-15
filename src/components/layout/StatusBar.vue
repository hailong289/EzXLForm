<script setup lang="ts">
/**
 * StatusBar.vue
 * Thanh trạng thái cuối trang, hiển thị thông tin tổng hợp:
 * số dòng, số cột, dòng đã chọn, tên file, trang hiện tại, undo stack.
 */
import { computed } from 'vue'
import { useWorkbookStore } from '@/stores/workbookStore'
import { useSelectionStore } from '@/stores/selectionStore'
import { useAppConfigStore } from '@/stores/appConfigStore'
import { useHistory } from '@/composables/useHistory'
import { sheetRowCount } from '@/utils/sheetGrid'
import { useLocale } from '@/composables/useLocale'

// ─── Stores & Composables ────────────────────────────────────────────────────
const workbookStore = useWorkbookStore()
const selectionStore = useSelectionStore()
const appConfigStore = useAppConfigStore()
const { undoStackSize, redoStackSize } = useHistory()
const { t, locale } = useLocale()

// ─── Computed ────────────────────────────────────────────────────────────────

/** Sheet đang active */
const activeSheet = computed(() => workbookStore.activeSheet)

/** Tổng số dòng dữ liệu (không tính header) */
const totalRows = computed(() => (activeSheet.value ? sheetRowCount(activeSheet.value) : 0))

/** Tổng số cột */
const totalColumns = computed(() => activeSheet.value?.headers.length ?? 0)

/** Số dòng đang được chọn (checkbox) */
const selectedRowCount = computed(() => selectionStore.selectedRowIndexes.length)

/** Tên file */
const fileName = computed(() => workbookStore.workbook?.meta.fileName ?? '')

const windowLabel = computed(() => {
  const window = appConfigStore.listWindow
  if (!window || window.total <= 0 || window.to < window.from) return ''
  const numberLocale = locale.value === 'en' ? 'en-US' : 'vi-VN'
  const from = window.from.toLocaleString(numberLocale)
  const to = window.to.toLocaleString(numberLocale)
  const total = window.total.toLocaleString(numberLocale)
  return window.filtering
    ? t('status.match.total', { from, to, total })
    : t('status.rows.total', { from, to, total })
})
</script>

<template>
  <div
    data-tour="status"
    class="flex h-6 select-none items-center justify-between overflow-x-auto border-t border-base-300 bg-base-200 px-4 text-xs tabular-nums opacity-70 max-md:h-7 max-md:px-3"
  >
    <!-- ── Left: Rows, Selected, Columns ─────────────────────────────────── -->
    <div class="flex shrink-0 items-center gap-3">

      <!-- Tổng số dòng -->
      <span v-if="activeSheet" class="flex items-center gap-1">
        <span class="opacity-60">Rows:</span>
        <span class="font-medium">{{ totalRows.toLocaleString() }}</span>
      </span>

      <!-- Số dòng đã chọn (chỉ hiện khi có chọn) -->
      <span
        v-if="selectedRowCount > 0"
        class="flex items-center gap-1 text-primary"
      >
        <span class="opacity-60">Selected:</span>
        <span class="font-medium">{{ selectedRowCount.toLocaleString() }}</span>
      </span>

      <!-- Tổng số cột -->
      <span v-if="activeSheet" class="flex items-center gap-1">
        <span class="opacity-60">Cols:</span>
        <span class="font-medium">{{ totalColumns }}</span>
      </span>

      <!-- Trạng thái khi chưa mở file -->
      <span v-if="!activeSheet" class="italic opacity-50">
        {{ t('status.no.file.open') }}
      </span>
    </div>

    <!-- ── Right: File name, Page info, Undo depth ───────────────────────── -->
    <div class="flex shrink-0 items-center gap-3">

      <!-- Undo stack depth (chỉ hiện khi có lịch sử) -->
      <span
        v-if="undoStackSize > 0"
        class="hidden items-center gap-1 md:flex"
        :title="t('status.can.undo.n.actions', { n: undoStackSize })"
      >
        <span class="opacity-60">Undo:</span>
        <span class="font-medium">{{ undoStackSize }}</span>
      </span>

      <!-- Divider -->
      <div v-if="fileName" class="w-px h-3 bg-base-content opacity-20" />

      <!-- Tên file -->
      <span
        v-if="fileName"
        class="max-w-[96px] truncate md:max-w-[180px]"
        :title="fileName"
      >
        {{ fileName }}
      </span>

      <!-- Divider -->
      <div v-if="activeSheet" class="w-px h-3 bg-base-content opacity-20" />

      <span v-if="windowLabel" class="font-medium">{{ windowLabel }}</span>
    </div>
  </div>
</template>
