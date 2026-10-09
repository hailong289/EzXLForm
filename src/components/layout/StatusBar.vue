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

// ─── Stores & Composables ────────────────────────────────────────────────────
const workbookStore = useWorkbookStore()
const selectionStore = useSelectionStore()
const appConfigStore = useAppConfigStore()
const { undoStackSize, redoStackSize } = useHistory()

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

/** Trang hiện tại */
const currentPage = computed(() => appConfigStore.currentPage ?? 1)

/** Tổng số trang */
const totalPages = computed(() => {
  if (!totalRows.value || !appConfigStore.rowsPerPage) return 1
  return Math.max(1, Math.ceil(totalRows.value / appConfigStore.rowsPerPage))
})
</script>

<template>
  <div
    class="flex items-center justify-between bg-base-200 border-t border-base-300 px-4 h-6 text-xs opacity-70 select-none"
  >
    <!-- ── Left: Rows, Selected, Columns ─────────────────────────────────── -->
    <div class="flex items-center gap-3">

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
        Chưa mở file nào
      </span>
    </div>

    <!-- ── Right: File name, Page info, Undo depth ───────────────────────── -->
    <div class="flex items-center gap-3">

      <!-- Undo stack depth (chỉ hiện khi có lịch sử) -->
      <span
        v-if="undoStackSize > 0"
        class="flex items-center gap-1"
        :title="`Có thể hoàn tác ${undoStackSize} thao tác`"
      >
        <span class="opacity-60">Undo:</span>
        <span class="font-medium">{{ undoStackSize }}</span>
      </span>

      <!-- Divider -->
      <div v-if="fileName" class="w-px h-3 bg-base-content opacity-20" />

      <!-- Tên file -->
      <span
        v-if="fileName"
        class="max-w-[180px] truncate"
        :title="fileName"
      >
        {{ fileName }}
      </span>

      <!-- Divider -->
      <div v-if="activeSheet" class="w-px h-3 bg-base-content opacity-20" />

      <!-- Thông tin phân trang -->
      <span v-if="activeSheet" class="flex items-center gap-1">
        <span class="opacity-60">Trang</span>
        <span class="font-medium">{{ currentPage }}</span>
        <span class="opacity-60">/</span>
        <span class="font-medium">{{ totalPages }}</span>
      </span>
    </div>
  </div>
</template>
