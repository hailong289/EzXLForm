<script setup lang="ts">
import { ref, computed, watch } from 'vue'
import { Eye, FileSpreadsheet } from 'lucide-vue-next'
import { useWorkbookStore } from '@/stores/workbookStore'
import PreviewSheetTabs from './PreviewSheetTabs.vue'
import { columnIndexToLetter } from '@/types/excel'
import type { CellValue, RowData } from '@/types/excel'
import { formatDateValue } from '@/utils/dateFormat'
import { readLogicalRange } from '@/composables/rowPageStore'
import { cellToDisplayString } from '@/utils/sheetGrid'

const workbookStore = useWorkbookStore()

// Sheet preview id (mặc định lấy activeSheetId)
const previewSheetId = ref<string>('')

watch(
  () => workbookStore.activeSheet?.id,
  (newId) => {
    if (newId) {
      previewSheetId.value = newId
    }
  },
  { immediate: true },
)

const currentPreviewSheet = computed(() => {
  if (!workbookStore.workbook) return null
  return (
    workbookStore.workbook.sheets.find((s) => s.id === previewSheetId.value) ??
    workbookStore.activeSheet
  )
})

const pagedPreviewRows = ref<RowData[]>([])

watch(
  () =>
    `${currentPreviewSheet.value?.id ?? ''}:${currentPreviewSheet.value?.paged ? 1 : 0}:${workbookStore.dataRevision}`,
  async () => {
    const sheet = currentPreviewSheet.value
    if (!sheet?.paged) {
      pagedPreviewRows.value = []
      return
    }
    pagedPreviewRows.value = await readLogicalRange(sheet.id, 0, Math.min(200, sheet.rowCount ?? 0))
  },
  { immediate: true },
)

const PREVIEW_PAGE_ROWS = 100
const previewPage = ref(1)

const fullGrid = computed(() => {
  void workbookStore.dataRevision
  const sheet = currentPreviewSheet.value
  if (!sheet) return [] as CellValue[][]
  const source = sheet.sourceGrid?.length ? sheet.sourceGrid : null
  const width = Math.max(
    source?.reduce((max, row) => Math.max(max, row.length), 0) ?? 0,
    sheet.headers.length,
    1,
  )
  const grid: CellValue[][] = (source ?? []).map((row) => {
    const line = row.slice()
    while (line.length < width) line.push(null)
    return line
  })
  const headerAt = Math.min(Math.max(0, sheet.headerRowIndex), Math.max(grid.length - 1, 0))
  while (grid.length <= headerAt) grid.push(Array.from({ length: width }, () => null))
  if (!grid[headerAt]) grid[headerAt] = []
  while (grid[headerAt].length < width) grid[headerAt].push(null)
  sheet.headers.forEach((header, column) => {
    grid[headerAt][column] = header.label
  })
  if (sheet.paged) {
    pagedPreviewRows.value.forEach((row, index) => {
      const lineIndex = headerAt + 1 + index
      while (grid.length <= lineIndex) grid.push(Array.from({ length: width }, () => null))
      sheet.headers.forEach((header, column) => {
        grid[lineIndex][column] = row[header.key] ?? null
      })
    })
    return grid
  }
  sheet.rows.forEach((row, index) => {
    const lineIndex = headerAt + 1 + index
    while (grid.length <= lineIndex) grid.push(Array.from({ length: width }, () => null))
    while (grid[lineIndex].length < width) grid[lineIndex].push(null)
    sheet.headers.forEach((header, column) => {
      grid[lineIndex][column] = row[header.key] ?? null
    })
  })
  return grid
})

const previewPageCount = computed(() => Math.max(1, Math.ceil(fullGrid.value.length / PREVIEW_PAGE_ROWS)))
const pagedGrid = computed(() => {
  const start = (previewPage.value - 1) * PREVIEW_PAGE_ROWS
  return fullGrid.value.slice(start, start + PREVIEW_PAGE_ROWS).map((row, offset) => ({
    index: start + offset,
    cells: row,
  }))
})
const columnCount = computed(() => fullGrid.value.reduce((max, row) => Math.max(max, row.length), 0))

watch(previewPageCount, (count) => {
  if (previewPage.value > count) previewPage.value = count
})
watch(
  () => currentPreviewSheet.value?.id,
  () => {
    previewPage.value = 1
  },
)

function onSelectSheet(sheetId: string) {
  previewSheetId.value = sheetId
}

function cellText(rowIndex: number, columnIndex: number, value: CellValue): string {
  const sheet = currentPreviewSheet.value
  const header = sheet && rowIndex > sheet.headerRowIndex ? sheet.headers[columnIndex] : undefined
  if (value === null || value === undefined || value === '') return ''
  if (header?.dataType === 'number') {
    const num = Number(value)
    return Number.isNaN(num) ? cellToDisplayString(value) : num.toLocaleString('vi-VN')
  }
  if (header?.dataType === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (header?.dataType === 'date') return formatDateValue(value, header.dateFormat)
  return cellToDisplayString(value)
}

function isHeaderRow(rowIndex: number): boolean {
  return rowIndex === (currentPreviewSheet.value?.headerRowIndex ?? 0)
}
</script>

<template>
  <div class="flex flex-col h-full bg-base-100 overflow-hidden select-none">
    <!-- ── Preview Toolbar ─────────────────────────────────────────────────── -->
    <div
      class="flex items-center justify-between p-2 bg-base-200/70 border-b border-base-300 text-xs"
    >
      <div class="flex items-center gap-2">
        <Eye class="w-3.5 h-3.5 text-primary" />
        <span class="font-bold">Xem trước Excel (Realtime)</span>
        <span v-if="currentPreviewSheet" class="opacity-60 text-[11px]">
          — {{ currentPreviewSheet.name }}
        </span>
      </div>

      <div class="flex items-center gap-2 text-[11px] opacity-70">
        <span>{{ fullGrid.length.toLocaleString('vi-VN') }} dòng</span>
        <span>•</span>
        <span>{{ columnCount }} cột</span>
        <span>Calibri 11</span>
      </div>
    </div>

    <!-- ── Spreadsheet Grid ────────────────────────────────────────────────── -->
    <div class="flex-1 overflow-auto scrollbar-thin excel-grid bg-base-100 relative">
      <div v-if="!currentPreviewSheet || columnCount === 0" class="flex flex-col items-center justify-center h-48 opacity-50 gap-2">
        <FileSpreadsheet class="w-10 h-10 stroke-1" />
        <span class="text-xs">Không có dữ liệu sheet để xem trước.</span>
      </div>

      <table v-else class="excel-sheet border-collapse w-max min-w-full bg-white text-black">
        <thead>
          <tr>
            <th class="excel-header-col sticky left-0 top-0 z-30 bg-[#f3f3f3]" />
            <th
              v-for="colIdx in columnCount"
              :key="colIdx"
              class="excel-header-col sticky top-0 z-20 bg-[#f3f3f3] text-center font-normal"
            >
              {{ columnIndexToLetter(colIdx - 1) }}
            </th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="line in pagedGrid" :key="line.index">
            <td class="excel-header-col sticky left-0 z-10 bg-[#f3f3f3] text-center font-normal">
              {{ line.index + 1 }}
            </td>
            <td
              v-for="(value, colIdx) in line.cells"
              :key="colIdx"
              class="excel-cell"
              :class="isHeaderRow(line.index) ? 'font-bold' : ''"
              :style="isHeaderRow(line.index) ? { backgroundColor: '#217346', color: '#ffffff' } : undefined"
            >
              {{ cellText(line.index, colIdx, value) }}
            </td>
          </tr>
        </tbody>
      </table>
    </div>

    <div
      v-if="previewPageCount > 1"
      class="flex items-center justify-end gap-1 border-t border-base-300 px-2 py-1 text-xs"
    >
      <button type="button" class="btn btn-ghost btn-xs" :disabled="previewPage <= 1" @click="previewPage = 1">Trang đầu</button>
      <button type="button" class="btn btn-ghost btn-xs" :disabled="previewPage <= 1" @click="previewPage -= 1">Trang trước</button>
      <span class="px-1">Trang {{ previewPage }} / {{ previewPageCount }}</span>
      <button type="button" class="btn btn-ghost btn-xs" :disabled="previewPage >= previewPageCount" @click="previewPage += 1">Trang sau</button>
      <button type="button" class="btn btn-ghost btn-xs" :disabled="previewPage >= previewPageCount" @click="previewPage = previewPageCount">Trang cuối</button>
    </div>

    <!-- ── Preview Sheet Tabs ──────────────────────────────────────────────── -->
    <PreviewSheetTabs
      v-if="workbookStore.workbook"
      :sheets="workbookStore.sheetList"
      :active-sheet-id="previewSheetId"
      @select="onSelectSheet"
    />
  </div>
</template>
