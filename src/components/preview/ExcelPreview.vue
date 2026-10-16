<script setup lang="ts">
import { ref, computed, watch } from 'vue'
import { Eye, FileSpreadsheet } from 'lucide-vue-next'
import { useWorkbookStore } from '@/stores/workbookStore'
import PreviewSheetTabs from './PreviewSheetTabs.vue'
import { columnIndexToLetter } from '@/types/excel'
import type { CellValue, RowData } from '@/types/excel'
import { formatDateValue } from '@/utils/dateFormat'
import { formatNumber } from '@/utils/numberFormat'
import { columnPaint } from '@/utils/columnStyle'
import { readLogicalRange } from '@/composables/rowPageStore'
import { cellToDisplayString } from '@/utils/sheetGrid'
import { useVirtualRows } from '@/composables/useVirtualRows'

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

const previewScrollRef = ref<HTMLElement | null>(null)
const previewRows = ref<Record<number, RowData>>({})
const VIEW_BUFFER = 80
let previewToken = 0
let heldKey = ''

const columnCount = computed(() => {
  const sheet = currentPreviewSheet.value
  if (!sheet) return 0
  const sourceWidth = sheet.sourceGrid?.reduce((max, row) => Math.max(max, row.length), 0) ?? 0
  return Math.max(sourceWidth, sheet.headers.length, sheet.paged ? 0 : 1)
})

const smallGrid = computed(() => {
  void workbookStore.dataRevision
  const sheet = currentPreviewSheet.value
  if (!sheet || sheet.paged) return [] as CellValue[][]
  const source = sheet.sourceGrid?.length ? sheet.sourceGrid : null
  const width = Math.max(columnCount.value, 1)
  const grid: CellValue[][] = (source ?? []).map((row) => {
    const line = row.slice()
    while (line.length < width) line.push(null)
    return line
  })
  const headerAt = Math.min(Math.max(0, sheet.headerRowIndex), Math.max(grid.length - 1, 0))
  while (grid.length <= headerAt) grid.push(Array.from({ length: width }, () => null))
  while (grid[headerAt].length < width) grid[headerAt].push(null)
  sheet.headers.forEach((header, column) => {
    grid[headerAt][column] = header.label
  })
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

const previewCount = computed(() => {
  const sheet = currentPreviewSheet.value
  if (!sheet) return 0
  if (!sheet.paged) return smallGrid.value.length
  return sheet.headerRowIndex + 1 + (sheet.rowCount ?? 0)
})

const previewWindow = useVirtualRows({
  count: previewCount,
  estimateSize: 22,
  scrollElement: previewScrollRef,
  measure: true,
})

function cellsFor(index: number): CellValue[] {
  const sheet = currentPreviewSheet.value
  if (!sheet) return []
  if (!sheet.paged) return smallGrid.value[index] ?? []
  const width = Math.max(columnCount.value, 1)
  const line: CellValue[] = Array.from({ length: width }, () => null)
  const headerAt = sheet.headerRowIndex
  if (index <= headerAt) {
    const source = sheet.sourceGrid?.[index] ?? []
    for (let column = 0; column < width; column++) line[column] = source[column] ?? null
    if (index === headerAt) {
      sheet.headers.forEach((header, column) => {
        line[column] = header.label
      })
    }
    return line
  }
  const dataIndex = index - headerAt - 1
  const row = previewRows.value[dataIndex] ?? {}
  sheet.headers.forEach((header, column) => {
    line[column] = row[header.key] ?? null
  })
  return line
}

watch(
  () => {
    const rows = previewWindow.items.value
    const sheet = currentPreviewSheet.value
    return [
      sheet?.id ?? '',
      sheet?.paged ? 1 : 0,
      workbookStore.dataRevision,
      rows[0]?.index ?? -1,
      rows[rows.length - 1]?.index ?? -1,
      previewCount.value,
    ].join('|')
  },
  async () => {
    const sheet = currentPreviewSheet.value
    const rows = previewWindow.items.value
    const token = ++previewToken
    if (!sheet?.paged || rows.length === 0) {
      previewRows.value = {}
      heldKey = ''
      return
    }
    const headerAt = sheet.headerRowIndex
    const dataIndexes = rows
      .map((item) => item.index - headerAt - 1)
      .filter((index) => index >= 0)
    if (dataIndexes.length === 0) return
    const lastRow = Math.max(0, (sheet.rowCount ?? 1) - 1)
    const from = Math.max(0, Math.min(...dataIndexes) - VIEW_BUFFER)
    const to = Math.min(lastRow, Math.max(...dataIndexes) + VIEW_BUFFER)
    const key = `${sheet.id}:${workbookStore.dataRevision}:${from}:${to}`
    if (key === heldKey) return
    const loaded = await readLogicalRange(sheet.id, from, to - from + 1)
    if (token !== previewToken) return
    const next: Record<number, RowData> = {}
    for (let offset = 0; offset < loaded.length; offset++) next[from + offset] = loaded[offset] ?? {}
    previewRows.value = next
    heldKey = key
  },
  { immediate: true },
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
    return Number.isNaN(num) ? cellToDisplayString(value) : formatNumber(num, header.numFmt)
  }
  if (header?.dataType === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (header?.dataType === 'date') return formatDateValue(value, header.dateFormat)
  return cellToDisplayString(value)
}

function isHeaderRow(rowIndex: number): boolean {
  return rowIndex === (currentPreviewSheet.value?.headerRowIndex ?? 0)
}

function previewCellStyle(rowIndex: number, columnIndex: number) {
  if (isHeaderRow(rowIndex)) return { backgroundColor: '#217346', color: '#ffffff' }
  const sheet = currentPreviewSheet.value
  if (!sheet || rowIndex <= sheet.headerRowIndex) return undefined
  const paint = columnPaint(sheet.headers[columnIndex]?.style)
  return { ...paint.box, ...paint.text }
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
        <span>{{ previewCount.toLocaleString('vi-VN') }} dòng</span>
        <span>•</span>
        <span>{{ columnCount }} cột</span>
        <span>Calibri 11</span>
      </div>
    </div>

    <!-- ── Spreadsheet Grid ────────────────────────────────────────────────── -->
    <div ref="previewScrollRef" class="flex-1 overflow-auto scrollbar-thin excel-grid bg-base-100 relative">
      <div v-if="!currentPreviewSheet || previewCount === 0" class="flex flex-col items-center justify-center h-48 opacity-50 gap-2">
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
          <tr v-if="previewWindow.paddingStart.value > 0" aria-hidden="true">
            <td :colspan="columnCount + 1" class="border-0 p-0" :style="{ height: `${previewWindow.paddingStart.value}px` }" />
          </tr>
          <tr
            v-for="item in previewWindow.items.value"
            :key="item.index"
            :ref="previewWindow.measureElement"
            :data-index="item.index"
          >
            <td class="excel-header-col sticky left-0 z-10 bg-[#f3f3f3] text-center font-normal">
              {{ item.index + 1 }}
            </td>
            <td
              v-for="(value, colIdx) in cellsFor(item.index)"
              :key="colIdx"
              class="excel-cell"
              :class="isHeaderRow(item.index) ? 'font-bold' : ''"
              :style="previewCellStyle(item.index, colIdx)"
            >
              {{ cellText(item.index, colIdx, value) }}
            </td>
          </tr>
          <tr v-if="previewWindow.paddingEnd.value > 0" aria-hidden="true">
            <td :colspan="columnCount + 1" class="border-0 p-0" :style="{ height: `${previewWindow.paddingEnd.value}px` }" />
          </tr>
        </tbody>
      </table>
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
