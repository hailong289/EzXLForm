<script setup lang="ts">
import { ref, computed, watch, nextTick, onMounted, onUnmounted } from 'vue'
import {
  Search,
  Filter,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight,
  FileX,
  X,
  PanelLeftOpen,
  PanelRightClose,
  GripVertical,
  Columns3,
  Plus,
  Trash2,
} from 'lucide-vue-next'
import { useWorkbookStore } from '@/stores/workbookStore'
import { useSelectionStore } from '@/stores/selectionStore'
import { PAGE_SIZE_OPTIONS, useAppConfigStore } from '@/stores/appConfigStore'
import { useHistory } from '@/composables/useHistory'
import RowForm from './RowForm.vue'
import TableHeaderCell from './TableHeaderCell.vue'
import BatchActionBar from './BatchActionBar.vue'
import type { AdvancedFilterRule, ColumnHeader, RowData } from '@/types/excel'
import { ensureSourceGrid, headerChoiceIndexes, previewGridLine, sheetRowCount } from '@/utils/sheetGrid'
import { DATE_FORMATS, DEFAULT_DATE_FORMAT, formatDateValue } from '@/utils/dateFormat'
import { activeFilterRules, operatorsFor, operatorNeedsValue, rowMatchesFilters } from '@/utils/rowFilter'
import { readRange, scanRows } from '@/composables/rowPageStore'

const workbookStore = useWorkbookStore()
const selectionStore = useSelectionStore()
const appConfigStore = useAppConfigStore()
const { push: pushHistory } = useHistory()

// ─── Kéo Co Dãn Độ Rộng Form (Resizable Form Panel) ─────────────────────────
const editorContainerRef = ref<HTMLDivElement | null>(null)
const formWidth = ref<number | null>(null)
const isResizingForm = ref(false)

function startResizeForm(e: MouseEvent) {
  e.preventDefault()
  if (formWidth.value == null && editorContainerRef.value) {
    formWidth.value = editorContainerRef.value.clientWidth * 0.5
  }
  isResizingForm.value = true
  document.addEventListener('mousemove', onResizeMouseMove)
  document.addEventListener('mouseup', stopResizeForm)
}

function maxFormWidth() {
  const width = editorContainerRef.value?.clientWidth ?? 800
  return Math.max(340, width * 0.5)
}

function onResizeMouseMove(e: MouseEvent) {
  if (!isResizingForm.value || !editorContainerRef.value) return
  const rect = editorContainerRef.value.getBoundingClientRect()
  const newWidth = e.clientX - rect.left
  formWidth.value = Math.max(340, Math.min(newWidth, maxFormWidth()))
}

function stopResizeForm() {
  if (isResizingForm.value) {
    isResizingForm.value = false
    document.removeEventListener('mousemove', onResizeMouseMove)
    document.removeEventListener('mouseup', stopResizeForm)
  }
}

function clampFormWidth() {
  if (formWidth.value == null) return
  formWidth.value = Math.max(340, Math.min(formWidth.value, maxFormWidth()))
}

const tableScrollRef = ref<HTMLElement | null>(null)
const sumScrollRef = ref<HTMLElement | null>(null)
const columnWidths = ref<number[]>([])
let editorResizeObserver: ResizeObserver | null = null
let headerSizeObserver: ResizeObserver | null = null
let syncingTableScroll = false

function measureColumns() {
  const cells = tableScrollRef.value?.querySelectorAll('thead th')
  if (!cells?.length) return
  columnWidths.value = [...cells].map((cell) => Math.round(cell.getBoundingClientRect().width))
  const head = tableScrollRef.value?.querySelector('thead')
  if (head && !headerSizeObserver) {
    headerSizeObserver = new ResizeObserver(() => measureColumns())
    headerSizeObserver.observe(head)
  }
}

function onTableScroll() {
  if (syncingTableScroll || !tableScrollRef.value || !sumScrollRef.value) return
  syncingTableScroll = true
  sumScrollRef.value.scrollLeft = tableScrollRef.value.scrollLeft
  syncingTableScroll = false
}

function onSumScroll() {
  if (syncingTableScroll || !tableScrollRef.value || !sumScrollRef.value) return
  syncingTableScroll = true
  tableScrollRef.value.scrollLeft = sumScrollRef.value.scrollLeft
  syncingTableScroll = false
}

onMounted(() => {
  clampFormWidth()
  if (editorContainerRef.value) {
    editorResizeObserver = new ResizeObserver(() => clampFormWidth())
    editorResizeObserver.observe(editorContainerRef.value)
  }
  nextTick(() => {
    measureColumns()
    if (!tableScrollRef.value || headerSizeObserver) return
    headerSizeObserver = new ResizeObserver(() => measureColumns())
    headerSizeObserver.observe(tableScrollRef.value)
  })
})

onUnmounted(() => {
  editorResizeObserver?.disconnect()
  headerSizeObserver?.disconnect()
  document.removeEventListener('mousemove', onResizeMouseMove)
  document.removeEventListener('mouseup', stopResizeForm)
})

// Active sheet & rows
const activeSheet = computed(() => workbookStore.activeSheet)
const headers = computed(() => activeSheet.value?.headers ?? [])
const allRows = computed(() => activeSheet.value?.rows ?? [])
const rowTotal = computed(() => (activeSheet.value ? sheetRowCount(activeSheet.value) : 0))
const visibleRows = ref<Record<number, RowData>>({})
const pagedMatches = ref<number[] | null>(null)
const searchScanning = ref(false)

function rowAt(index: number): RowData | undefined {
  if (activeSheet.value?.paged) return visibleRows.value[index]
  return allRows.value[index]
}

const selectedRowIndex = ref<number | null>(rowTotal.value > 0 ? 0 : null)
const selectedRow = computed(() =>
  selectedRowIndex.value !== null ? rowAt(selectedRowIndex.value) ?? null : null,
)

watch(
  () => workbookStore.workbook,
  () => {
    selectedRowIndex.value = rowTotal.value > 0 ? 0 : null
  },
)

watch(
  () => `${activeSheet.value?.id ?? ''}:${activeSheet.value?.headerRowIndex ?? ''}`,
  async () => {
    selectedRowIndex.value = null
    await nextTick()
    selectedRowIndex.value = rowTotal.value > 0 ? 0 : null
  },
)

watch(selectedRowIndex, (index) => {
  if (!activeSheet.value?.paged || index == null) return
  const page = Math.floor(index / appConfigStore.rowsPerPage) + 1
  if (page !== appConfigStore.currentPage) appConfigStore.setPage(page)
})

const headerChoices = computed(() => {
  const sheet = activeSheet.value
  if (!sheet) return []
  const grid = ensureSourceGrid(sheet)
  return headerChoiceIndexes(grid.length, sheet.headerRowIndex).map((index) => ({
    index,
    label: `Dòng ${index + 1} — ${previewGridLine(grid, index)}`,
  }))
})

function onChangeHeaderRow(event: Event) {
  if (!activeSheet.value) return
  if (activeSheet.value.paged) {
    appConfigStore.addToast({
      type: 'info',
      message: 'File lớn giữ dòng tiêu đề đã chọn lúc mở. Muốn đổi, hãy mở lại file và chọn dòng tiêu đề khác.',
    })
    return
  }
  const index = Number((event.target as HTMLSelectElement).value)
  workbookStore.setHeaderRow(activeSheet.value.id, index)
}

// Trạng thái thu gọn/mở rộng Bảng danh sách dòng
const showTableList = ref(true)
const showForm = ref(true)

function hideForm() {
  if (!showTableList.value) return
  showForm.value = false
}

function hideTable() {
  if (!showForm.value) return
  showTableList.value = false
}

// Local Search / Filter
const searchInput = ref(appConfigStore.searchQuery)
const filterCol = ref(appConfigStore.filterColumn)

function onSearchChange() {
  appConfigStore.setSearch(searchInput.value, filterCol.value)
}

watch(
  () => appConfigStore.searchQuery,
  (value) => {
    if (searchInput.value !== value) searchInput.value = value
  },
)

watch(
  () => appConfigStore.filterColumn,
  (value) => {
    if (filterCol.value !== value) filterCol.value = value
  },
)

function clearSearch() {
  searchInput.value = ''
  filterCol.value = ''
  draftJoin.value = 'and'
  draftRules.value = [blankRule()]
  showAdvancedFilter.value = false
  appConfigStore.clearSearch()
}

const showAdvancedFilter = ref(false)
const draftJoin = ref(appConfigStore.advancedJoin)
const draftRules = ref<AdvancedFilterRule[]>([])

function blankRule(): AdvancedFilterRule {
  return {
    id: `rule_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    columnKey: headers.value[0]?.key ?? '',
    operator: 'contains',
    value: '',
  }
}

function openAdvancedFilter() {
  draftJoin.value = appConfigStore.advancedJoin
  draftRules.value = appConfigStore.advancedRules.length
    ? appConfigStore.advancedRules.map((rule) => ({ ...rule }))
    : [blankRule()]
  showAdvancedFilter.value = !showAdvancedFilter.value
}

function addFilterRule() {
  draftRules.value.push(blankRule())
}

function removeFilterRule(id: string) {
  draftRules.value = draftRules.value.filter((rule) => rule.id !== id)
  if (draftRules.value.length === 0) draftRules.value = [blankRule()]
}

function onFilterColumnChange(rule: AdvancedFilterRule) {
  const header = headers.value.find((item) => item.key === rule.columnKey)
  const operators = operatorsFor(header)
  if (!operators.some((item) => item.id === rule.operator)) {
    rule.operator = operators[0]?.id ?? 'contains'
  }
  if (header?.dataType === 'boolean' && operatorNeedsValue(rule.operator) && rule.value !== 'true' && rule.value !== 'false') {
    rule.value = 'true'
  }
}

watch(
  draftRules,
  (rules) => {
    rules.forEach((rule) => onFilterColumnChange(rule))
  },
  { deep: true },
)

function applyAdvancedFilter() {
  appConfigStore.setAdvancedFilter(draftRules.value.map((rule) => ({ ...rule })), draftJoin.value)
}

function clearAdvancedFilter() {
  draftJoin.value = 'and'
  draftRules.value = [blankRule()]
  appConfigStore.setAdvancedFilter([], 'and')
}

const appliedRuleCount = computed(() => activeFilterRules(appConfigStore.advancedRules).length)

function filtersAreActive(): boolean {
  return Boolean(appConfigStore.searchQuery.trim()) || appliedRuleCount.value > 0
}

// Filtered Rows
const filteredRowIndices = computed(() => {
  if (activeSheet.value?.paged) return pagedMatches.value ?? []
  if (!filtersAreActive()) return allRows.value.map((_, index) => index)

  const indices: number[] = []
  allRows.value.forEach((row, index) => {
    if (
      rowMatchesFilters(
        row,
        headers.value,
        appConfigStore.searchQuery,
        appConfigStore.filterColumn,
        appConfigStore.advancedRules,
        appConfigStore.advancedJoin,
      )
    ) {
      indices.push(index)
    }
  })
  return indices
})

const totalFilteredRows = computed(() => {
  if (activeSheet.value?.paged && !filtersAreActive()) return rowTotal.value
  return filteredRowIndices.value.length
})
const totalPages = computed(() =>
  Math.max(1, Math.ceil(totalFilteredRows.value / appConfigStore.rowsPerPage)),
)

const paginatedIndices = computed(() => {
  const start = (appConfigStore.currentPage - 1) * appConfigStore.rowsPerPage
  const end = start + appConfigStore.rowsPerPage
  if (activeSheet.value?.paged && !filtersAreActive()) {
    const last = Math.min(end, rowTotal.value)
    const indexes: number[] = []
    for (let index = start; index < last; index++) indexes.push(index)
    return indexes
  }
  return filteredRowIndices.value.slice(start, end)
})

let searchToken = 0
watch(
  () =>
    `${activeSheet.value?.id ?? ''}|${appConfigStore.searchQuery}|${appConfigStore.filterColumn}|${appConfigStore.advancedJoin}|${JSON.stringify(appConfigStore.advancedRules)}`,
  async () => {
    const sheet = activeSheet.value
    if (!sheet?.paged || !filtersAreActive()) {
      pagedMatches.value = null
      searchScanning.value = false
      return
    }
    const token = ++searchToken
    searchScanning.value = true
    const matches: number[] = []
    await scanRows(sheet.id, sheet.rowCount ?? 0, (index, row) => {
      if (token !== searchToken) return false
      if (
        rowMatchesFilters(
          row,
          sheet.headers,
          appConfigStore.searchQuery,
          appConfigStore.filterColumn,
          appConfigStore.advancedRules,
          appConfigStore.advancedJoin,
        )
      ) {
        matches.push(index)
      }
      return matches.length < 20000
    })
    if (token !== searchToken) return
    pagedMatches.value = matches
    searchScanning.value = false
    if (matches.length >= 20000) {
      appConfigStore.addToast({
        type: 'info',
        message: 'Bộ lọc file lớn dừng ở 20.000 dòng khớp đầu tiên.',
      })
    }
  },
)

watch(
  () =>
    [
      activeSheet.value?.id,
      activeSheet.value?.paged,
      appConfigStore.currentPage,
      appConfigStore.rowsPerPage,
      workbookStore.dataRevision,
      paginatedIndices.value.join(','),
    ].join('|'),
  async () => {
    const sheet = activeSheet.value
    const indexes = paginatedIndices.value
    if (!sheet?.paged || indexes.length === 0) {
      visibleRows.value = {}
      return
    }
    const start = indexes[0]
    const rows = await readRange(sheet.id, start, indexes[indexes.length - 1] - start + 1)
    const next: Record<number, RowData> = {}
    indexes.forEach((index) => {
      next[index] = rows[index - start] ?? {}
    })
    visibleRows.value = next
  },
  { immediate: true },
)

function prevPage() {
  if (appConfigStore.currentPage > 1) {
    appConfigStore.setPage(appConfigStore.currentPage - 1)
  }
}

function nextPage() {
  if (appConfigStore.currentPage < totalPages.value) {
    appConfigStore.setPage(appConfigStore.currentPage + 1)
  }
}

function firstPage() {
  appConfigStore.setPage(1)
}

function lastPage() {
  appConfigStore.setPage(totalPages.value)
}

function onChangePageSize(event: Event) {
  appConfigStore.setRowsPerPage(Number((event.target as HTMLSelectElement).value))
}

// Select row to edit in Form
function selectRowForEdit(rowIndex: number) {
  selectedRowIndex.value = rowIndex
}

// Form Handlers: Tạo, Sửa, Xóa từng dòng
async function handleCreateRow(newRowData: RowData) {
  if (!activeSheet.value) return
  const sheetId = activeSheet.value.id

  await workbookStore.addRow(sheetId, newRowData)
  const newIndex = sheetRowCount(activeSheet.value) - 1
  selectedRowIndex.value = newIndex

  const newPage = Math.ceil(sheetRowCount(activeSheet.value) / appConfigStore.rowsPerPage)
  appConfigStore.setPage(newPage)

  pushHistory({
    type: 'ADD_ROW',
    description: `Tạo dòng mới #${newIndex + 1}`,
    sheetId,
    before: {},
    after: { rows: [newRowData] },
  })

  appConfigStore.addToast({
    type: 'success',
    message: `Đã tạo thành công dòng #${newIndex + 1}!`,
  })
}

let debouncedHistoryTimer: ReturnType<typeof setTimeout> | null = null
let pendingEditBefore: RowData[] | null = null

async function handleUpdateRow(rowIndex: number, updatedData: RowData) {
  if (!activeSheet.value || rowIndex < 0 || rowIndex >= rowTotal.value) return
  const sheetId = activeSheet.value.id
  if (!pendingEditBefore) {
    pendingEditBefore = JSON.parse(JSON.stringify(activeSheet.value.rows)) as RowData[]
  }

  await workbookStore.replaceRow(sheetId, rowIndex, updatedData)

  if (debouncedHistoryTimer) clearTimeout(debouncedHistoryTimer)
  debouncedHistoryTimer = setTimeout(() => {
    const before = pendingEditBefore
    pendingEditBefore = null
    const after = JSON.parse(JSON.stringify(activeSheet.value?.rows ?? [])) as RowData[]
    if (!before) return
    pushHistory({
      type: 'EDIT_CELL',
      description: `Sửa dòng #${rowIndex + 1}`,
      sheetId,
      before: { rows: before },
      after: { rows: after },
    })
  }, 1000)
}

async function handleDeleteRow(rowIndex: number) {
  if (!activeSheet.value || rowIndex < 0 || rowIndex >= rowTotal.value) return
  const sheetId = activeSheet.value.id
  const oldRow = { ...(rowAt(rowIndex) ?? {}) }

  await workbookStore.deleteRows(sheetId, [rowIndex])

  pushHistory({
    type: 'DELETE_ROWS',
    description: `Xóa dòng #${rowIndex + 1}`,
    sheetId,
    before: { rows: [oldRow] },
    after: {},
  })

  // Điều chỉnh selectedRowIndex sau khi xóa
  const remaining = sheetRowCount(activeSheet.value)
  if (remaining === 0) {
    selectedRowIndex.value = null
  } else if (rowIndex >= remaining) {
    selectedRowIndex.value = remaining - 1
  } else {
    selectedRowIndex.value = rowIndex
  }

  appConfigStore.addToast({
    type: 'info',
    message: `Đã xóa dòng #${rowIndex + 1}.`,
  })
}

function handleClearForm() {
  selectedRowIndex.value = null
}

function handleSelectPrev() {
  if (selectedRowIndex.value !== null && selectedRowIndex.value > 0) {
    selectedRowIndex.value--
  }
}

function handleSelectNext() {
  if (
    selectedRowIndex.value !== null &&
    selectedRowIndex.value < rowTotal.value - 1
  ) {
    selectedRowIndex.value++
  }
}

// Header Actions
function onHeaderRename(key: string, newLabel: string) {
  if (!activeSheet.value) return
  const sheetId = activeSheet.value.id
  const h = headers.value.find((item) => item.key === key)
  const oldLabel = h ? h.label : ''

  workbookStore.updateHeader(sheetId, key, newLabel)

  pushHistory({
    type: 'EDIT_HEADER',
    description: `Đổi tên cột "${oldLabel}" thành "${newLabel}"`,
    sheetId,
    before: { headers: [{ ...h!, label: oldLabel }] },
    after: { headers: [{ ...h!, label: newLabel }] },
  })
}

function onHeaderAddLeft(key: string) {
  if (!activeSheet.value) return
  const sheetId = activeSheet.value.id
  const idx = headers.value.findIndex((h) => h.key === key)
  const newKey = `col_${Date.now()}`
  const newCol: ColumnHeader = {
    key: newKey,
    label: `Cột mới`,
    originalIndex: idx,
    dataType: 'text',
  }

  const prevKey = idx > 0 ? headers.value[idx - 1].key : null
  workbookStore.addColumn(sheetId, prevKey, newCol)

  pushHistory({
    type: 'ADD_COLUMN',
    description: `Thêm cột mới bên trái`,
    sheetId,
    before: {},
    after: { headers: [newCol] },
  })
}

function onHeaderAddRight(key: string) {
  if (!activeSheet.value) return
  const sheetId = activeSheet.value.id
  const newKey = `col_${Date.now()}`
  const newCol: ColumnHeader = {
    key: newKey,
    label: `Cột mới`,
    originalIndex: headers.value.length,
    dataType: 'text',
  }

  workbookStore.addColumn(sheetId, key, newCol)

  pushHistory({
    type: 'ADD_COLUMN',
    description: `Thêm cột mới bên phải`,
    sheetId,
    before: {},
    after: { headers: [newCol] },
  })
}

function onHeaderDelete(key: string) {
  if (!activeSheet.value) return
  if (headers.value.length <= 1) {
    appConfigStore.addToast({
      type: 'warning',
      message: 'Không thể xóa cột duy nhất còn lại.',
    })
    return
  }

  const sheetId = activeSheet.value.id
  const colBackup = headers.value.find((h) => h.key === key)
  if (!colBackup) return

  workbookStore.deleteColumn(sheetId, key)

  pushHistory({
    type: 'DELETE_COLUMN',
    description: `Xóa cột "${colBackup.label}"`,
    sheetId,
    before: { headers: [colBackup] },
    after: {},
  })

  appConfigStore.addToast({
    type: 'info',
    message: `Đã xóa cột "${colBackup.label}".`,
  })
}

const draggingHeaderKey = ref<string | null>(null)
const dropHeaderKey = ref<string | null>(null)

watch(
  () => headers.value.map((header) => header.key).join('|'),
  () => nextTick(measureColumns),
  { flush: 'post', immediate: true },
)

function onHeaderDragStart(key: string) {
  draggingHeaderKey.value = key
  dropHeaderKey.value = null
}

function onHeaderDragOver(key: string) {
  if (!draggingHeaderKey.value || draggingHeaderKey.value === key) {
    dropHeaderKey.value = null
    return
  }
  dropHeaderKey.value = key
}

function onHeaderDrop(targetKey: string) {
  const fromKey = draggingHeaderKey.value
  draggingHeaderKey.value = null
  dropHeaderKey.value = null
  if (!fromKey || fromKey === targetKey || !activeSheet.value) return
  const toIndex = headers.value.findIndex((header) => header.key === targetKey)
  if (toIndex < 0) return
  workbookStore.reorderColumn(activeSheet.value.id, fromKey, toIndex)
}

function onHeaderDragEnd() {
  draggingHeaderKey.value = null
  dropHeaderKey.value = null
}

function columnSum(header: ColumnHeader): string {
  void workbookStore.dataRevision
  if (header.dataType === 'date' || header.dataType === 'boolean') return ''
  if (activeSheet.value?.paged) {
    const stat = activeSheet.value.columnSums?.[header.key]
    if (!stat || stat.count === 0) return header.dataType === 'number' ? '0' : ''
    return Number.isInteger(stat.sum)
      ? stat.sum.toLocaleString()
      : stat.sum.toLocaleString(undefined, { maximumFractionDigits: 2 })
  }
  let sum = 0
  let count = 0
  for (const index of filteredRowIndices.value) {
    const raw = allRows.value[index]?.[header.key]
    if (raw === null || raw === undefined || raw === '' || typeof raw === 'boolean') continue
    const numeric = typeof raw === 'number' ? raw : Number(String(raw).replace(/,/g, ''))
    if (Number.isNaN(numeric)) continue
    sum += numeric
    count += 1
  }
  if (count === 0) return header.dataType === 'number' ? '0' : ''
  return Number.isInteger(sum)
    ? sum.toLocaleString()
    : sum.toLocaleString(undefined, { maximumFractionDigits: 2 })
}

function toggleHeaderSelection() {
  const indexes = activeSheet.value?.paged ? paginatedIndices.value : filteredRowIndices.value
  const allOn = indexes.length > 0 && indexes.every((idx) => selectionStore.isRowSelected(idx))
  if (activeSheet.value?.paged) {
    selectionStore.selectedRowIndexes = allOn ? [] : [...indexes]
    return
  }
  if (allOn) selectionStore.clearSelection()
  else selectionStore.selectAll(allRows.value.length)
}

function onHeaderMoveLeft(key: string) {
  if (!activeSheet.value) return
  const idx = headers.value.findIndex((h) => h.key === key)
  if (idx <= 0) return
  const targetKey = headers.value[idx - 1].key
  workbookStore.moveColumn(activeSheet.value.id, key, targetKey)
}

function onHeaderMoveRight(key: string) {
  if (!activeSheet.value) return
  const idx = headers.value.findIndex((h) => h.key === key)
  if (idx === -1 || idx >= headers.value.length - 1) return
  const targetKey = headers.value[idx + 1].key
  workbookStore.moveColumn(activeSheet.value.id, key, targetKey)
}

function onHeaderSort(key: string, order: 'asc' | 'desc') {
  if (!activeSheet.value) return
  if (activeSheet.value.paged) {
    appConfigStore.addToast({
      type: 'info',
      message: 'File lớn đang xem theo trang nên không sắp xếp toàn bộ cột trong trình duyệt.',
    })
    return
  }
  const rows = [...activeSheet.value.rows]
  const col = headers.value.find((h) => h.key === key)

  rows.sort((a, b) => {
    const valA = a[key] ?? ''
    const valB = b[key] ?? ''

    if (col?.dataType === 'number') {
      const numA = Number(valA) || 0
      const numB = Number(valB) || 0
      return order === 'asc' ? numA - numB : numB - numA
    }

    const strA = String(valA).toLowerCase()
    const strB = String(valB).toLowerCase()
    return order === 'asc' ? strA.localeCompare(strB) : strB.localeCompare(strA)
  })

  activeSheet.value.rows = rows
  workbookStore.isDirty = true

  appConfigStore.addToast({
    type: 'info',
    message: `Đã sắp xếp cột ${col?.label} (${order === 'asc' ? 'A→Z' : 'Z→A'}).`,
  })
}

function formatListCell(
  val: RowData[string] | undefined,
  header: ColumnHeader,
  _revision: number,
): string {
  void _revision
  if (val === null || val === undefined || val === '') return ''
  if (header.dataType === 'boolean') return val ? '✓' : '✗'
  if (header.dataType === 'number') {
    const num = Number(val)
    return Number.isNaN(num) ? String(val) : num.toLocaleString()
  }
  if (header.dataType === 'date') return formatDateValue(val, header.dateFormat)
  return String(val)
}

function onHeaderSetDateFormat(key: string, format: string) {
  if (!activeSheet.value) return
  const header = headers.value.find((item) => item.key === key)
  if (!header) return
  header.dateFormat = format
  header.dataType = 'date'
  workbookStore.isDirty = true
  workbookStore.dataRevision += 1
}

function onHeaderSetType(key: string, type: ColumnHeader['dataType']) {
  if (!activeSheet.value) return
  const h = headers.value.find((item) => item.key === key)
  if (h) {
    h.dataType = type
    if (type === 'date' && !h.dateFormat) h.dateFormat = DEFAULT_DATE_FORMAT
    workbookStore.isDirty = true
    workbookStore.dataRevision += 1
    appConfigStore.addToast({
      type: 'info',
      message: `Đã đặt kiểu cột "${h.label}" thành ${type}.`,
    })
  }
}

const showColumnEditor = ref(false)
const newColumnLabel = ref('')

function addNamedColumn() {
  if (!activeSheet.value) return
  const label = newColumnLabel.value.trim() || 'Cột mới'
  const sheetId = activeSheet.value.id
  const newCol: ColumnHeader = {
    key: `col_${Date.now()}`,
    label,
    originalIndex: headers.value.length,
    dataType: 'text',
  }
  const afterKey = headers.value.length > 0 ? headers.value[headers.value.length - 1].key : null
  workbookStore.addColumn(sheetId, afterKey, newCol)
  pushHistory({
    type: 'ADD_COLUMN',
    description: `Thêm cột "${label}"`,
    sheetId,
    before: {},
    after: { headers: [newCol] },
  })
  newColumnLabel.value = ''
}
</script>

<template>
  <div ref="editorContainerRef" class="flex h-full w-full bg-base-100 overflow-hidden select-none relative">
    <!-- ── 1. FORM NHẬP LIỆU (KÉO CO DÃN ĐƯỢC ĐỘ RỘNG) ─────────────────────── -->
    <div
      v-show="showForm"
      class="h-full min-w-0 border-r border-base-300"
      :class="showTableList && formWidth == null ? 'flex-1' : 'flex-shrink-0'"
      :style="!showTableList ? { width: '100%', maxWidth: '100%' } : formWidth == null ? undefined : { width: `${formWidth}px`, maxWidth: '50%' }"
    >
      <RowForm
        :headers="headers"
        :sheet-id="activeSheet?.id"
        :selected-row-index="selectedRowIndex"
        :selected-row="selectedRow"
        :total-rows="rowTotal"
        :table-hidden="!showTableList"
        @create="handleCreateRow"
        @edit-columns="showColumnEditor = true"
        @update="handleUpdateRow"
        @delete="handleDeleteRow"
        @clear="handleClearForm"
        @select-prev="handleSelectPrev"
        @select-next="handleSelectNext"
        @jump-to-row="selectRowForEdit"
        @set-date-format="onHeaderSetDateFormat"
        @hide="hideForm"
        @show-table="showTableList = true"
      />
    </div>

    <!-- ── THANH RESIZER KÉO CO DÃN ĐỘ RỘNG FORM ───────────────────────────── -->
    <div
      v-show="showTableList"
      class="w-1.5 hover:w-2 bg-base-300 hover:bg-primary cursor-col-resize flex-shrink-0 transition-colors z-20 relative group select-none flex items-center justify-center"
      :class="{ 'bg-primary w-2': isResizingForm }"
      title="Kéo sang trái/phải để co dãn độ rộng Form"
      @mousedown="startResizeForm"
    >
      <div class="h-8 w-0.5 bg-base-content/30 group-hover:bg-primary-content rounded-full" />
    </div>

    <!-- ── 2. BẢNG DANH SÁCH CHỌN DÒNG & QUẢN LÝ HEADER ─────────────────────── -->
    <div v-show="showTableList" class="flex-1 flex flex-col h-full overflow-hidden min-w-0">
      <!-- Toolbar tìm kiếm & Toggle -->
      <div
        class="flex items-center justify-between p-2 bg-base-200/70 border-b border-base-300 gap-2 flex-wrap text-xs"
      >
        <!-- Search & Filter -->
        <div class="flex items-center gap-2 flex-1 min-w-[200px]">
          <div class="flex items-center gap-1 flex-1 max-w-xs">
            <div class="relative flex-1">
              <input
                v-model="searchInput"
                type="text"
                placeholder="Tìm kiếm dòng..."
                class="input input-xs input-bordered w-full pl-7"
                @input="onSearchChange"
                @keydown.escape.prevent="clearSearch"
              />
              <Search class="w-3.5 h-3.5 absolute left-2 top-1/2 -translate-y-1/2 opacity-40 pointer-events-none" />
            </div>
            <button
              v-if="searchInput || filterCol || appliedRuleCount > 0"
              type="button"
              class="btn btn-ghost btn-xs gap-1 px-1.5"
              title="Xóa bộ lọc"
              @click="clearSearch"
            >
              <X class="w-3.5 h-3.5" />
              Xóa
            </button>
          </div>

          <div class="flex items-center gap-1">
            <Filter class="w-3.5 h-3.5 opacity-50 flex-shrink-0" />
            <select
              v-model="filterCol"
              class="select select-xs select-bordered text-xs"
              @change="onSearchChange"
            >
              <option value="">Tất cả cột</option>
              <option v-for="h in headers" :key="h.key" :value="h.key">
                {{ h.label }}
              </option>
            </select>
          </div>

          <button
            type="button"
            class="btn btn-xs gap-1"
            :class="showAdvancedFilter || appliedRuleCount > 0 ? 'btn-primary' : 'btn-ghost'"
            @click="openAdvancedFilter"
          >
            <Filter class="w-3.5 h-3.5" />
            Nâng cao
            <span v-if="appliedRuleCount > 0" class="badge badge-xs">{{ appliedRuleCount }}</span>
          </button>
        </div>

        <label class="flex items-center gap-1 flex-shrink-0">
          <span class="opacity-60 text-[11px] whitespace-nowrap">Header</span>
          <select
            class="select select-xs select-bordered max-w-[180px] text-xs"
            :value="activeSheet?.headerRowIndex ?? 0"
            :disabled="!!activeSheet?.paged"
            title="Chọn dòng làm header cho sheet đang mở"
            @change="onChangeHeaderRow"
          >
            <option v-for="choice in headerChoices" :key="choice.index" :value="choice.index">
              {{ choice.label }}
            </option>
          </select>
        </label>

        <!-- Nút thu gọn / mở rộng bảng -->
        <button
          v-if="!showForm"
          class="btn btn-ghost btn-xs gap-1 text-xs"
          title="Hiện form"
          @click="showForm = true"
        >
          <PanelLeftOpen class="w-3.5 h-3.5" />
          <span class="hidden sm:inline">Hiện form</span>
        </button>
        <button
          class="btn btn-ghost btn-xs gap-1 text-xs"
          title="Ẩn bảng"
          :disabled="!showForm"
          @click="hideTable"
        >
          <PanelRightClose class="w-3.5 h-3.5" />
          <span class="hidden sm:inline">Ẩn bảng</span>
        </button>
      </div>

      <div
        v-if="showAdvancedFilter"
        class="shrink-0 border-b border-base-300 bg-base-200/50 px-2 py-2 space-y-1.5"
      >
        <div class="flex flex-wrap items-center gap-1.5 text-xs">
          <select v-model="draftJoin" class="select select-xs select-bordered">
            <option value="and">Khớp tất cả</option>
            <option value="or">Khớp một điều kiện</option>
          </select>
          <button type="button" class="btn btn-xs btn-ghost gap-1" @click="addFilterRule">
            <Plus class="w-3.5 h-3.5" />
            Thêm điều kiện
          </button>
          <button type="button" class="btn btn-xs btn-primary" @click="applyAdvancedFilter">Áp dụng</button>
          <button type="button" class="btn btn-xs btn-ghost" @click="clearAdvancedFilter">Xóa điều kiện</button>
        </div>
        <div v-for="rule in draftRules" :key="rule.id" class="flex flex-wrap items-center gap-1">
          <select
            v-model="rule.columnKey"
            class="select select-xs select-bordered max-w-[140px]"
            @change="onFilterColumnChange(rule)"
          >
            <option v-for="header in headers" :key="header.key" :value="header.key">
              {{ header.label }}
            </option>
          </select>
          <select v-model="rule.operator" class="select select-xs select-bordered">
            <option
              v-for="item in operatorsFor(headers.find((header) => header.key === rule.columnKey))"
              :key="item.id"
              :value="item.id"
            >
              {{ item.label }}
            </option>
          </select>
          <select
            v-if="
              operatorNeedsValue(rule.operator) &&
              headers.find((header) => header.key === rule.columnKey)?.dataType === 'boolean'
            "
            v-model="rule.value"
            class="select select-xs select-bordered"
          >
            <option value="true">Đúng</option>
            <option value="false">Sai</option>
          </select>
          <input
            v-else-if="operatorNeedsValue(rule.operator)"
            v-model="rule.value"
            type="text"
            class="input input-xs input-bordered w-40"
            :placeholder="rule.operator === 'like' || rule.operator === 'not_like' ? '%abc hoặc NV___' : 'Giá trị'"
            :title="rule.operator === 'like' || rule.operator === 'not_like' ? '% là nhiều ký tự, _ là một ký tự. Ví dụ: Nguyễn% hoặc NV___' : ''"
            @keydown.enter.prevent="applyAdvancedFilter"
          />
          <button
            type="button"
            class="btn btn-ghost btn-xs btn-square"
            title="Bỏ điều kiện"
            @click="removeFilterRule(rule.id)"
          >
            <Trash2 class="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      <!-- Bảng danh sách dòng -->
      <div ref="tableScrollRef" class="flex-1 overflow-auto scrollbar-thin relative" @scroll="onTableScroll">
        <table class="editor-grid table table-xs w-full">
          <thead>
            <tr>
              <th class="col-check border-r border-b border-base-300">
                <div class="cell-center">
                  <input
                    type="checkbox"
                    class="checkbox checkbox-xs checkbox-primary"
                    :checked="
                      (activeSheet?.paged ? paginatedIndices : filteredRowIndices).length > 0 &&
                      (activeSheet?.paged ? paginatedIndices : filteredRowIndices).every((idx) =>
                        selectionStore.isRowSelected(idx),
                      )
                    "
                    @change="toggleHeaderSelection"
                  />
                </div>
              </th>

              <!-- Column Headers (Hỗ trợ chỉnh sửa header trực tiếp) -->
              <TableHeaderCell
                v-for="(h, idx) in headers"
                :key="h.key"
                :header="h"
                :index="idx"
                :is-first="idx === 0"
                :is-last="idx === headers.length - 1"
                :is-dragging="draggingHeaderKey === h.key"
                :is-drop-target="dropHeaderKey === h.key"
                @rename="onHeaderRename"
                @delete="onHeaderDelete"
                @add-left="onHeaderAddLeft"
                @add-right="onHeaderAddRight"
                @move-left="onHeaderMoveLeft"
                @move-right="onHeaderMoveRight"
                @sort-asc="(k) => onHeaderSort(k, 'asc')"
                @sort-desc="(k) => onHeaderSort(k, 'desc')"
                @set-type="onHeaderSetType"
                @drag-start="onHeaderDragStart"
                @drag-over="onHeaderDragOver"
                @drop="onHeaderDrop"
                @drag-end="onHeaderDragEnd"
              />
            </tr>
          </thead>

          <tbody>
            <tr
              v-for="realIdx in paginatedIndices"
              :key="realIdx"
              class="data-row border-b border-base-300 transition-colors cursor-pointer"
              :class="{
                'row-active font-medium': selectedRowIndex === realIdx,
                'row-checked': selectionStore.isRowSelected(realIdx),
              }"
              @click="selectRowForEdit(realIdx)"
            >
              <!-- Checkbox -->
              <td class="col-check border-r border-b border-base-300" @click.stop>
                <div class="cell-center">
                  <input
                    type="checkbox"
                    class="checkbox checkbox-xs checkbox-primary"
                    :checked="selectionStore.isRowSelected(realIdx)"
                    @change="selectionStore.toggleRowSelection(realIdx)"
                  />
                </div>
              </td>

              <!-- Data Cells (Click để chọn dòng vào Form) -->
              <td
                v-for="h in headers"
                :key="h.key"
                class="px-2 py-1.5 border-r border-b border-base-300 text-xs overflow-hidden text-ellipsis whitespace-nowrap max-w-[150px]"
                :title="String(rowAt(realIdx)?.[h.key] ?? '')"
              >
                {{ formatListCell(rowAt(realIdx)?.[h.key], h, workbookStore.dataRevision) }}
              </td>
            </tr>
          </tbody>
        </table>

        <!-- Empty State -->
        <div
          v-if="totalFilteredRows === 0"
          class="flex flex-col items-center justify-center h-48 opacity-50 gap-2"
        >
          <FileX class="w-10 h-10 stroke-1" />
          <span class="text-xs">{{ searchScanning ? 'Đang lọc dữ liệu lớn...' : 'Không có dữ liệu phù hợp.' }}</span>
        </div>
      </div>

      <div
        ref="sumScrollRef"
        class="sum-scroll shrink-0 overflow-x-auto overflow-y-hidden border-t border-base-300 bg-base-200"
        @scroll="onSumScroll"
      >
        <div class="flex h-8 items-stretch text-xs">
          <div
            v-for="(width, index) in columnWidths"
            :key="index"
            class="sum-cell shrink-0 border-r border-base-300 flex items-center px-2 font-mono whitespace-nowrap"
            :class="{
              'sum-check justify-center px-0 text-[11px] font-bold': index === 0,
              'font-semibold': index >= 1,
            }"
            :style="{ width: `${width}px` }"
            :title="index >= 1 && headers[index - 1] ? `Tổng ${headers[index - 1].label}` : ''"
          >
            <template v-if="index === 0">Tổng</template>
            <template v-else-if="headers[index - 1]">
              {{ columnSum(headers[index - 1]) }}
            </template>
          </div>
        </div>
      </div>

      <!-- Phân trang -->
      <div
        class="flex items-center justify-between p-2 bg-base-200/70 border-t border-base-300 text-xs flex-wrap gap-2"
      >
        <div class="opacity-70 text-[11px]">
          Hiển thị
          <span class="font-bold">
            {{
              totalFilteredRows === 0
                ? 0
                : (appConfigStore.currentPage - 1) * appConfigStore.rowsPerPage + 1
            }}
            -
            {{
              Math.min(
                appConfigStore.currentPage * appConfigStore.rowsPerPage,
                totalFilteredRows,
              )
            }}
          </span>
          / <span class="font-bold">{{ totalFilteredRows }}</span> dòng
        </div>

        <label class="flex items-center gap-1">
          <span class="opacity-60 text-[11px] whitespace-nowrap">Số dòng</span>
          <select
            class="select select-xs select-bordered w-20 text-xs"
            :value="appConfigStore.rowsPerPage"
            title="Số dòng mỗi trang"
            @change="onChangePageSize"
          >
            <option v-for="size in PAGE_SIZE_OPTIONS" :key="size" :value="size">
              {{ size }}
            </option>
          </select>
        </label>

        <div class="flex items-center gap-1">
          <button
            class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0"
            title="Trang đầu"
            :disabled="appConfigStore.currentPage <= 1"
            @click="firstPage"
          >
            <ChevronsLeft class="w-3.5 h-3.5" />
          </button>
          <button
            class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0"
            title="Trang trước"
            :disabled="appConfigStore.currentPage <= 1"
            @click="prevPage"
          >
            <ChevronLeft class="w-3.5 h-3.5" />
          </button>
          <span class="px-1 text-[11px]">
            {{ appConfigStore.currentPage }} / {{ totalPages }}
          </span>
          <button
            class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0"
            title="Trang sau"
            :disabled="appConfigStore.currentPage >= totalPages"
            @click="nextPage"
          >
            <ChevronRight class="w-3.5 h-3.5" />
          </button>
          <button
            class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0"
            title="Trang cuối"
            :disabled="appConfigStore.currentPage >= totalPages"
            @click="lastPage"
          >
            <ChevronsRight class="w-3.5 h-3.5" />
          </button>
        </div>
      </div>
    </div>

    <BatchActionBar :headers="headers" />

    <div
      v-if="showColumnEditor"
      class="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      @click.self="showColumnEditor = false"
    >
      <div class="bg-base-100 border border-base-300 rounded-xl shadow-xl w-full max-w-lg max-h-[80vh] flex flex-col">
        <div class="flex items-center justify-between px-4 py-3 border-b border-base-300">
          <div class="flex items-center gap-2">
            <Columns3 class="w-4 h-4 text-primary" />
            <h3 class="font-bold text-sm">Cột header</h3>
          </div>
          <button class="btn btn-ghost btn-xs btn-circle" title="Đóng" @click="showColumnEditor = false">
            <X class="w-4 h-4" />
          </button>
        </div>

        <div class="flex-1 overflow-y-auto p-3 space-y-2 scrollbar-thin">
          <p v-if="headers.length === 0" class="text-xs opacity-60 text-center py-6">
            Sheet chưa có cột. Thêm cột bên dưới.
          </p>
          <div
            v-for="(h, idx) in headers"
            :key="h.key"
            class="flex items-center gap-2"
          >
            <span class="text-[11px] opacity-50 w-5 text-right font-mono">{{ idx + 1 }}</span>
            <input
              class="input input-sm input-bordered flex-1 text-xs"
              :value="h.label"
              :title="h.label"
              @change="onHeaderRename(h.key, ($event.target as HTMLInputElement).value)"
            />
            <select
              class="select select-sm select-bordered text-xs w-24"
              :value="h.dataType || 'text'"
              @change="onHeaderSetType(h.key, ($event.target as HTMLSelectElement).value as ColumnHeader['dataType'])"
            >
              <option value="text">Chữ</option>
              <option value="number">Số</option>
              <option value="date">Ngày</option>
              <option value="boolean">Đúng/Sai</option>
            </select>
            <select
              v-if="h.dataType === 'date'"
              class="select select-sm select-bordered text-xs w-28 font-mono"
              :value="h.dateFormat || DEFAULT_DATE_FORMAT"
              title="Định dạng ngày"
              @change="onHeaderSetDateFormat(h.key, ($event.target as HTMLSelectElement).value)"
            >
              <option v-for="pattern in DATE_FORMATS" :key="pattern" :value="pattern">
                {{ pattern }}
              </option>
            </select>
            <button
              class="btn btn-ghost btn-sm btn-square text-error"
              title="Xóa cột"
              @click="onHeaderDelete(h.key)"
            >
              <Trash2 class="w-3.5 h-3.5" />
            </button>
          </div>
        </div>

        <div class="flex items-center gap-2 p-3 border-t border-base-300">
          <input
            v-model="newColumnLabel"
            class="input input-sm input-bordered flex-1 text-xs"
            placeholder="Tên cột mới"
            @keydown.enter.prevent="addNamedColumn"
          />
          <button class="btn btn-primary btn-sm gap-1" @click="addNamedColumn">
            <Plus class="w-3.5 h-3.5" /> Thêm cột
          </button>
        </div>
      </div>
    </div>
  </div>
</template>
