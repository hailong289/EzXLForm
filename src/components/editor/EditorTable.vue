<script setup lang="ts">
import { ref, computed, watch, nextTick, onMounted, onUnmounted } from 'vue'
import {
  Search,
  Filter,
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
import { useAppConfigStore } from '@/stores/appConfigStore'
import { useProgressTaskStore } from '@/stores/progressTaskStore'
import { useVirtualRows } from '@/composables/useVirtualRows'
import { useHistory } from '@/composables/useHistory'
import RowForm from './RowForm.vue'
import FormPresetDialog from './FormPresetDialog.vue'
import TableHeaderCell from './TableHeaderCell.vue'
import TableRowCell from './TableRowCell.vue'
import BatchActionBar from './BatchActionBar.vue'
import type { AdvancedFilterRule, CellValue, ColumnHeader, RowData } from '@/types/excel'
import { ensureSourceGrid, headerChoiceIndexes, previewGridLine, sheetRowCount } from '@/utils/sheetGrid'
import { DATE_FORMATS, DEFAULT_DATE_FORMAT } from '@/utils/dateFormat'
import { activeFilterRules, operatorsFor, operatorNeedsValue, rowMatchesFilters } from '@/utils/rowFilter'
import { hasSheetOrder, readLogicalRange, scanLogicalRows } from '@/composables/rowPageStore'
import { useRowMarks, MARK_COLORS, type MarkColor } from '@/composables/rowMarks'
import { useFormPresets } from '@/composables/useFormPresets'
import type { PresetDraftError } from '@/utils/formPresets'
import { useLocale } from '@/composables/useLocale'

const { t } = useLocale()

const workbookStore = useWorkbookStore()
const selectionStore = useSelectionStore()
const appConfigStore = useAppConfigStore()
const progressStore = useProgressTaskStore()
const rowMarks = useRowMarks()
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
  window.addEventListener('pagehide', onEditorPageHide)
})

onUnmounted(() => {
  cancelSearchTask()
  appConfigStore.setListWindow(null)
  editorResizeObserver?.disconnect()
  headerSizeObserver?.disconnect()
  window.removeEventListener('pagehide', onEditorPageHide)
  document.removeEventListener('mousemove', onResizeMouseMove)
  document.removeEventListener('mouseup', stopResizeForm)
})

function onEditorPageHide() {
  void flushAllPagedEdits()
}

// Active sheet & rows
const activeSheet = computed(() => workbookStore.activeSheet)
const headers = computed(() => activeSheet.value?.headers ?? [])
const formPresets = useFormPresets()
const presetDialogOpen = ref(false)
const presetError = ref<PresetDraftError | null>(null)
const manageableHeaders = computed(() => headers.value.filter((header) => header.visible !== false))
const shownHeaderKeys = computed(() => {
  if (!formPresets.selectedPresetId.value) return null
  return formPresets.presets.value.find((preset) => preset.id === formPresets.selectedPresetId.value)?.visibleKeys ?? []
})

watch(
  () => [activeSheet.value?.id, headers.value.map((header) => `${header.key}:${header.visible}`).join('|')] as const,
  () => {
    if (!activeSheet.value) return
    void formPresets.loadForSheet(activeSheet.value.id, headers.value)
  },
  { immediate: true },
)

async function onCreatePreset(name: string, visibleKeys: string[]) {
  presetError.value = await formPresets.createPreset(name, visibleKeys)
}

async function onUpdatePreset(id: string, name: string, visibleKeys: string[]) {
  presetError.value = await formPresets.updatePreset(id, name, visibleKeys)
}

async function onRemovePreset(id: string) {
  presetError.value = null
  await formPresets.deletePreset(id)
}
const allRows = computed(() => activeSheet.value?.rows ?? [])
const rowTotal = computed(() => (activeSheet.value ? sheetRowCount(activeSheet.value) : 0))
const visibleRows = ref<Record<number, RowData>>({})
const pendingEdits = new Map<number, RowData>()
const pendingEditBefore = new Map<number, RowData>()
const pendingEditTimers = new Map<number, ReturnType<typeof setTimeout>>()
const pagedMatches = ref<number[] | null>(null)
const matchRevision = ref(0)
const searchScanning = ref(false)
let visibleLoadToken = 0
let searchTaskId: string | null = null
const VIEW_BUFFER = 80
const ROW_HEIGHT = 32

function rowAt(index: number): RowData | undefined {
  if (activeSheet.value?.paged) {
    return pendingEdits.get(index) ?? visibleRows.value[index]
  }
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

watch(
  () => activeSheet.value?.id,
  () => {
    pendingEdits.clear()
    pendingEditBefore.clear()
    for (const timer of pendingEditTimers.values()) clearTimeout(timer)
    pendingEditTimers.clear()
  },
)

watch(selectedRowIndex, (index) => {
  if (index == null) return
  jumpRowInput.value = index + 1
  nextTick(() => scrollToLogical(index))
})

const headerChoices = computed(() => {
  const sheet = activeSheet.value
  if (!sheet) return []
  const grid = ensureSourceGrid(sheet)
  return headerChoiceIndexes(grid.length, sheet.headerRowIndex).map((index) => ({
    index,
    label: t('common.row.n.preview', { n: index + 1, preview: previewGridLine(grid, index) }),
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
  void matchRevision.value
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

const listCount = computed(() => {
  if (filtersAreActive()) return filteredRowIndices.value.length
  return rowTotal.value
})

const rowWindow = useVirtualRows({
  count: listCount,
  estimateSize: ROW_HEIGHT,
  scrollElement: tableScrollRef,
})

function logicalAt(viewIndex: number): number {
  if (filtersAreActive()) return filteredRowIndices.value[viewIndex] ?? viewIndex
  return viewIndex
}

function scrollToLogical(index: number) {
  if (listCount.value <= 0) return
  let view = index
  if (filtersAreActive()) {
    const found = filteredRowIndices.value.indexOf(index)
    if (found < 0) return
    view = found
  }
  rowWindow.scrollToIndex(view)
}

const jumpRowInput = ref<number | ''>('')
const markColor = ref<MarkColor>('amber')
const showMarkList = ref(false)
const markPreviews = ref<Record<number, string>>({})

function jumpToRowNumber() {
  if (rowTotal.value <= 0) return
  const value = Number(jumpRowInput.value)
  if (!Number.isFinite(value)) return
  const logical = Math.min(Math.max(1, Math.floor(value)), rowTotal.value) - 1
  jumpRowInput.value = logical + 1
  selectRowForEdit(logical)
}

function jumpToLastRow() {
  if (listCount.value <= 0) return
  const logical = logicalAt(listCount.value - 1)
  jumpRowInput.value = logical + 1
  selectRowForEdit(logical)
}

const currentMarked = computed(() => {
  const sheetId = activeSheet.value?.id
  if (!sheetId || selectedRowIndex.value === null) return false
  return rowMarks.isMarked(sheetId, selectedRowIndex.value)
})

const markCount = computed(() => (activeSheet.value ? rowMarks.count(activeSheet.value.id) : 0))

function toggleRowMark(index: number) {
  const sheetId = activeSheet.value?.id
  if (!sheetId || index < 0 || index >= rowTotal.value) return
  rowMarks.toggle(sheetId, index, markColor.value)
}

function toggleCurrentMark() {
  if (selectedRowIndex.value === null) return
  toggleRowMark(selectedRowIndex.value)
}

function setCurrentMarkColor(color: MarkColor) {
  markColor.value = color
  const sheetId = activeSheet.value?.id
  if (!sheetId || selectedRowIndex.value === null) return
  if (rowMarks.isMarked(sheetId, selectedRowIndex.value)) {
    rowMarks.setMark(sheetId, selectedRowIndex.value, color)
  }
}

function markPreviewText(row: RowData): string {
  const bits = headers.value
    .slice(0, 4)
    .map((header) => {
      const value = row[header.key]
      if (value == null || value === '') return ''
      return String(value)
    })
    .filter(Boolean)
  return bits.join(' · ') || '(trống)'
}

async function openMarkList() {
  const sheet = activeSheet.value
  if (!sheet) return
  showMarkList.value = true
  const indexes = rowMarks.list(sheet.id).map((item) => item.index)
  const previews: Record<number, string> = {}
  try {
    if (sheet.paged && indexes.length > 0) {
      const loaded = await readLogicalIndexes(sheet.id, indexes)
      for (const index of indexes) previews[index] = markPreviewText(loaded.get(index) ?? {})
    } else {
      for (const index of indexes) previews[index] = markPreviewText(allRows.value[index] ?? {})
    }
    markPreviews.value = previews
  } finally {
    const current = selectedRowIndex.value
    const focus = current !== null && indexes.includes(current) ? current : indexes[0]
    if (focus !== undefined) selectRowForEdit(focus)
  }
}

function closeMarkList() {
  showMarkList.value = false
}

function removeMarkedRow(index: number) {
  const sheetId = activeSheet.value?.id
  if (!sheetId) return
  rowMarks.removeMark(sheetId, index)
  const next = { ...markPreviews.value }
  delete next[index]
  markPreviews.value = next
}

function recolorMarkedRow(index: number, color: MarkColor) {
  const sheetId = activeSheet.value?.id
  if (!sheetId) return
  rowMarks.setMark(sheetId, index, color)
}

function markClass(color: MarkColor): string {
  return `mark-swatch mark-swatch-${color}`
}

function dataRowClass(viewIndex: number): Record<string, boolean> {
  const logical = logicalAt(viewIndex)
  const sheetId = activeSheet.value?.id
  const color = sheetId ? rowMarks.colorOf(sheetId, logical) : null
  return {
    'row-active font-medium': selectedRowIndex.value === logical,
    'row-checked': selectionStore.isRowSelected(logical),
    'row-fresh': !!sheetId && workbookStore.isFreshRow(sheetId, logical),
    'row-drop': dropRowIndex.value === logical,
    ...(color ? { [`mark-${color}`]: true } : {}),
  }
}

function jumpToNextMark() {
  const sheetId = activeSheet.value?.id
  if (!sheetId) return
  const next = rowMarks.nextAfter(sheetId, selectedRowIndex.value ?? -1)
  if (next === null) return
  selectRowForEdit(next)
}

function cancelSearchTask() {
  if (!searchTaskId) return
  progressStore.remove(searchTaskId)
  searchTaskId = null
}

let searchToken = 0
watch(
  () =>
    `${activeSheet.value?.id ?? ''}|${appConfigStore.searchQuery}|${appConfigStore.filterColumn}|${appConfigStore.advancedJoin}|${JSON.stringify(appConfigStore.advancedRules)}`,
  async () => {
    const sheet = activeSheet.value
    const token = ++searchToken
    cancelSearchTask()
    if (!sheet?.paged || !filtersAreActive()) {
      pagedMatches.value = null
      matchRevision.value += 1
      searchScanning.value = false
      return
    }
    const matches: number[] = []
    pagedMatches.value = matches
    matchRevision.value += 1
    searchScanning.value = true
    const taskId = progressStore.enqueue({
      title: 'Lọc dòng',
      detail: sheet.name,
      sheetId: sheet.id,
      kind: 'filter',
    })
    searchTaskId = taskId
    progressStore.setStatus(taskId, 'running')
    const signal = progressStore.signal(taskId)
    const total = sheet.rowCount ?? 0
    let stopped = false
    await scanLogicalRows(sheet.id, total, async (index, row) => {
      if (token !== searchToken) return false
      if (signal && !(await signal.checkpoint())) {
        stopped = true
        return false
      }
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
        if (matches.length % 200 === 0) matchRevision.value += 1
      }
      if (index > 0 && index % 2000 === 0) {
        progressStore.setPercent(taskId, Math.round((index / Math.max(total, 1)) * 100))
        progressStore.setDetail(taskId, `${sheet.name} · ${matches.length.toLocaleString('vi-VN')} dòng khớp`)
      }
      return true
    })
    if (token !== searchToken) return
    matchRevision.value += 1
    searchScanning.value = false
    if (searchTaskId === taskId) searchTaskId = null
    if (!progressStore.tasks.some((task) => task.id === taskId)) return
    progressStore.setPercent(taskId, stopped ? Math.round((matches.length > 0 ? 100 : 0)) : 100)
    progressStore.setStatus(taskId, stopped ? 'cancelled' : 'done')
    progressStore.setDetail(
      taskId,
      stopped
        ? `${sheet.name} · đã dừng, ${matches.length.toLocaleString('vi-VN')} dòng khớp`
        : `${sheet.name} · ${matches.length.toLocaleString('vi-VN')} dòng khớp`,
    )
  },
)

let heldSheet = ''
let heldRevision = -1
let heldFrom = -1
let heldTo = -1

async function readLogicalIndexes(sheetId: string, indexes: number[]): Promise<Map<number, RowData>> {
  const sorted = [...new Set(indexes)].sort((left, right) => left - right)
  const loaded = new Map<number, RowData>()
  let cursor = 0
  while (cursor < sorted.length) {
    let end = cursor
    while (end + 1 < sorted.length && sorted[end + 1] === sorted[end] + 1) end += 1
    const start = sorted[cursor]
    const rows = await readLogicalRange(sheetId, start, sorted[end] - start + 1)
    for (let offset = 0; offset < rows.length; offset++) loaded.set(start + offset, rows[offset] ?? {})
    cursor = end + 1
  }
  return loaded
}

watch(
  () => {
    const rows = rowWindow.items.value
    const first = rows[0]?.index ?? -1
    const last = rows[rows.length - 1]?.index ?? -1
    return [
      activeSheet.value?.id ?? '',
      activeSheet.value?.paged ? 1 : 0,
      workbookStore.dataRevision,
      filtersAreActive() ? 1 : 0,
      matchRevision.value,
      first,
      last,
      listCount.value,
    ].join('|')
  },
  async () => {
    const sheet = activeSheet.value
    const rows = rowWindow.items.value
    const token = ++visibleLoadToken
    if (!sheet?.paged || rows.length === 0 || listCount.value === 0) {
      visibleRows.value = {}
      heldFrom = -1
      heldTo = -1
      publishListWindow()
      return
    }
    const viewFirst = rows[0].index
    const viewLast = rows[rows.length - 1].index
    const from = Math.max(0, viewFirst - VIEW_BUFFER)
    const to = Math.min(listCount.value - 1, viewLast + VIEW_BUFFER)
    publishListWindow()
    if (
      sheet.id === heldSheet &&
      workbookStore.dataRevision === heldRevision &&
      from >= heldFrom &&
      to <= heldTo
    ) {
      return
    }
    const logicals: number[] = []
    for (let index = from; index <= to; index++) logicals.push(logicalAt(index))
    const loaded = await readLogicalIndexes(sheet.id, logicals)
    if (token !== visibleLoadToken) return
    const next: Record<number, RowData> = {}
    for (const index of logicals) {
      next[index] = pendingEdits.get(index) ?? loaded.get(index) ?? {}
    }
    visibleRows.value = next
    heldSheet = sheet.id
    heldRevision = workbookStore.dataRevision
    heldFrom = from
    heldTo = to
  },
  { immediate: true },
)

function publishListWindow() {
  const rows = rowWindow.items.value
  if (rows.length === 0 || listCount.value === 0) {
    appConfigStore.setListWindow(null)
    return
  }
  const filtering = filtersAreActive()
  appConfigStore.setListWindow({
    filtering,
    from: rows[0].index + 1,
    to: rows[rows.length - 1].index + 1,
    total: filtering ? listCount.value : rowTotal.value,
  })
}

// Select row to edit in Form
function selectRowForEdit(rowIndex: number) {
  selectedRowIndex.value = rowIndex
}

function startCellEdit(rowIndex: number, columnKey: string) {
  selectedRowIndex.value = rowIndex
  selectionStore.setEditingCell(rowIndex, columnKey)
}

function sameCell(left: CellValue, right: CellValue): boolean {
  const empty = (value: CellValue) => value == null || value === ''
  if (empty(left) && empty(right)) return true
  return left === right
}

async function saveCell(rowIndex: number, columnKey: string, value: CellValue) {
  if (!activeSheet.value) return
  if (activeSheet.value.paged && rowAt(rowIndex) == null) return
  if (sameCell(rowAt(rowIndex)?.[columnKey] ?? null, value)) return
  const next = { ...(rowAt(rowIndex) ?? {}), [columnKey]: value }
  await handleUpdateRow(rowIndex, next)
}

// Form Handlers: Tạo, Sửa, Xóa từng dòng
function emptyRow(): RowData {
  const row: RowData = {}
  for (const header of headers.value) row[header.key] = header.dataType === 'boolean' ? false : null
  return row
}

async function flushAllPagedEdits() {
  for (const timer of pendingEditTimers.values()) clearTimeout(timer)
  pendingEditTimers.clear()
  const sheetId = activeSheet.value?.id
  if (!sheetId) return
  for (const index of [...pendingEdits.keys()]) {
    await flushPagedEdit(sheetId, index)
  }
}

async function handleCreateRow(
  newRowData: RowData,
  placement: 'after' | 'end' = 'end',
  anchor: number | null = selectedRowIndex.value,
  kind: 'copy' | 'blank' = 'copy',
) {
  if (!activeSheet.value) return
  await flushAllPagedEdits()
  const sheetId = activeSheet.value.id
  const afterSelected = placement === 'after' && anchor !== null
  const insertAt = afterSelected ? anchor + 1 : undefined
  const row = kind === 'blank' ? emptyRow() : { ...newRowData }

  await workbookStore.addRow(sheetId, row, insertAt)
  if (insertAt !== undefined) rowMarks.noteRowInserted(sheetId, insertAt)
  const newIndex = insertAt ?? sheetRowCount(activeSheet.value) - 1
  selectedRowIndex.value = newIndex
  await nextTick()
  scrollToLogical(newIndex)

  pushHistory({
    type: 'ADD_ROW',
    description: `Tạo dòng mới #${newIndex + 1}`,
    sheetId,
    before: {},
    after: { rows: [row] },
  })

  const afterLabel = anchor === null ? '' : ` sau dòng #${anchor + 1}`
  appConfigStore.addToast({
    type: 'success',
    message: kind === 'copy'
      ? `Đã sao chép dữ liệu thành dòng #${newIndex + 1}${afterLabel}.`
      : `Đã thêm dòng trống #${newIndex + 1}${afterSelected ? afterLabel : ' ở cuối bảng'}.`,
  })
}

function handleCommitFreshRow() {
  if (!activeSheet.value || selectedRowIndex.value === null) return
  workbookStore.commitFreshRow(activeSheet.value.id, selectedRowIndex.value)
  appConfigStore.addToast({
    type: 'success',
    message: `Đã lưu dòng #${selectedRowIndex.value + 1}.`,
  })
}

const memoryEditBefore = new Map<number, RowData>()
const memoryEditTimers = new Map<number, ReturnType<typeof setTimeout>>()

function schedulePagedEdit(sheetId: string, rowIndex: number, updatedData: RowData) {
  if (!pendingEditBefore.has(rowIndex)) {
    const current = pendingEdits.get(rowIndex) ?? visibleRows.value[rowIndex] ?? {}
    pendingEditBefore.set(rowIndex, JSON.parse(JSON.stringify(current)) as RowData)
  }
  pendingEdits.set(rowIndex, updatedData)
  visibleRows.value = { ...visibleRows.value, [rowIndex]: updatedData }

  const existing = pendingEditTimers.get(rowIndex)
  if (existing) clearTimeout(existing)
  pendingEditTimers.set(
    rowIndex,
    setTimeout(() => {
      void flushPagedEdit(sheetId, rowIndex)
    }, 180),
  )
}

async function flushPagedEdit(sheetId: string, rowIndex: number) {
  pendingEditTimers.delete(rowIndex)
  const data = pendingEdits.get(rowIndex)
  const before = pendingEditBefore.get(rowIndex)
  if (!data) return
  await workbookStore.replaceRow(sheetId, rowIndex, data)
  if (pendingEdits.get(rowIndex) !== data) return
  pendingEdits.delete(rowIndex)
  pendingEditBefore.delete(rowIndex)
  if (!before) return
  pushHistory({
    type: 'EDIT_CELL',
    description: `Sửa dòng #${rowIndex + 1}`,
    sheetId,
    rowIndex,
    before: { rows: [before] },
    after: { rows: [JSON.parse(JSON.stringify(data)) as RowData] },
  })
}

function scheduleMemoryEdit(sheetId: string, rowIndex: number, updatedData: RowData) {
  if (!memoryEditBefore.has(rowIndex)) {
    memoryEditBefore.set(rowIndex, JSON.parse(JSON.stringify(rowAt(rowIndex) ?? {})) as RowData)
  }
  const existing = memoryEditTimers.get(rowIndex)
  if (existing) clearTimeout(existing)
  memoryEditTimers.set(
    rowIndex,
    setTimeout(() => {
      memoryEditTimers.delete(rowIndex)
      const before = memoryEditBefore.get(rowIndex)
      memoryEditBefore.delete(rowIndex)
      if (!before) return
      pushHistory({
        type: 'EDIT_CELL',
        description: `Sửa dòng #${rowIndex + 1}`,
        sheetId,
        rowIndex,
        before: { rows: [before] },
        after: { rows: [JSON.parse(JSON.stringify(updatedData)) as RowData] },
      })
    }, 400),
  )
}

async function handleUpdateRow(rowIndex: number, updatedData: RowData) {
  if (!activeSheet.value || rowIndex < 0 || rowIndex >= rowTotal.value) return
  const sheetId = activeSheet.value.id
  if (activeSheet.value.paged) {
    schedulePagedEdit(sheetId, rowIndex, updatedData)
    return
  }
  scheduleMemoryEdit(sheetId, rowIndex, updatedData)
  await workbookStore.replaceRow(sheetId, rowIndex, updatedData)
}

async function handleDeleteRow(rowIndex: number) {
  if (!activeSheet.value || rowIndex < 0 || rowIndex >= rowTotal.value) return
  await flushAllPagedEdits()
  const sheetId = activeSheet.value.id
  const oldRow = { ...(rowAt(rowIndex) ?? {}) }

  const swapTail = Boolean(activeSheet.value.paged) && !hasSheetOrder(sheetId)
  rowMarks.noteRowsDeleted(sheetId, [rowIndex], swapTail, rowTotal.value)
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
const draggingRowIndex = ref<number | null>(null)
const dropRowIndex = ref<number | null>(null)

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

function shiftedIndex(index: number, from: number, to: number): number {
  if (index === from) return to
  if (from < to && index > from && index <= to) return index - 1
  if (to < from && index >= to && index < from) return index + 1
  return index
}

function onRowDragStart(index: number, event: DragEvent) {
  draggingRowIndex.value = index
  dropRowIndex.value = null
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = 'move'
    event.dataTransfer.setData('text/plain', String(index))
  }
}

function onRowDragOver(index: number) {
  if (draggingRowIndex.value === null || draggingRowIndex.value === index) {
    dropRowIndex.value = null
    return
  }
  dropRowIndex.value = index
}

function onRowDragEnd() {
  draggingRowIndex.value = null
  dropRowIndex.value = null
}

async function onRowDrop(targetIndex: number) {
  const from = draggingRowIndex.value
  draggingRowIndex.value = null
  dropRowIndex.value = null
  const sheet = activeSheet.value
  if (from === null || from === targetIndex || !sheet) return
  await flushAllPagedEdits()
  await workbookStore.moveRow(sheet.id, from, targetIndex)
  rowMarks.noteRowMoved(sheet.id, from, targetIndex)
  if (selectedRowIndex.value !== null) {
    selectedRowIndex.value = shiftedIndex(selectedRowIndex.value, from, targetIndex)
  }
  selectionStore.selectedRowIndexes = selectionStore.selectedRowIndexes.map((index) =>
    shiftedIndex(index, from, targetIndex),
  )
  if (selectionStore.editingCell) {
    selectionStore.editingCell = {
      ...selectionStore.editingCell,
      rowIndex: shiftedIndex(selectionStore.editingCell.rowIndex, from, targetIndex),
    }
  }
  pushHistory({
    type: 'MOVE_ROW',
    description: `Sắp xếp dòng #${from + 1} sang vị trí #${targetIndex + 1}`,
    sheetId: sheet.id,
    before: {},
    after: {},
    rowIndex: from,
    toIndex: targetIndex,
  })
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

const headerAllSelected = computed(() => {
  const selected = new Set(selectionStore.selectedRowIndexes)
  if (filtersAreActive()) {
    const indexes = filteredRowIndices.value
    return indexes.length > 0 && indexes.every((index) => selected.has(index))
  }
  return rowTotal.value > 0 && selected.size === rowTotal.value
})

function toggleHeaderSelection() {
  if (filtersAreActive()) {
    const indexes = filteredRowIndices.value
    const selected = new Set(selectionStore.selectedRowIndexes)
    const allOn = indexes.length > 0 && indexes.every((index) => selected.has(index))
    if (allOn) {
      const drop = new Set(indexes)
      selectionStore.selectedRowIndexes = selectionStore.selectedRowIndexes.filter((index) => !drop.has(index))
      return
    }
    indexes.forEach((index) => selected.add(index))
    selectionStore.selectedRowIndexes = [...selected]
    return
  }
  if (headerAllSelected.value) selectionStore.clearSelection()
  else selectionStore.selectAll(rowTotal.value)
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
      v-show="showForm && !showMarkList"
      class="h-full min-w-0 border-r border-base-300"
      :class="showTableList && formWidth == null ? 'flex-1' : 'flex-shrink-0'"
      :style="!showTableList ? { width: '100%', maxWidth: '100%' } : formWidth == null ? undefined : { width: `${formWidth}px`, maxWidth: '50%' }"
    >
      <RowForm
        v-if="!showMarkList"
        :headers="headers"
        :sheet-id="activeSheet?.id"
        :selected-row-index="selectedRowIndex"
        :selected-row="selectedRow"
        :data-revision="workbookStore.dataRevision"
        :total-rows="rowTotal"
        :table-hidden="!showTableList"
        :preset-options="formPresets.presets.value"
        :selected-preset-id="formPresets.selectedPresetId.value"
        :shown-header-keys="shownHeaderKeys"
        :preset-empty="formPresets.presetEmpty.value"
        :row-is-new="
          selectedRowIndex !== null &&
          !!activeSheet &&
          workbookStore.isFreshRow(activeSheet.id, selectedRowIndex)
        "
        @create="handleCreateRow"
        @select-preset="formPresets.selectPreset"
        @manage-presets="presetDialogOpen = true"
        @commit="handleCommitFreshRow"
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
      :title="t('table.drag.resize.form')"
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
                :placeholder="t('table.search.rows')"
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
              :title="t('table.clear.filter')"
              @click="clearSearch"
            >
              <X class="w-3.5 h-3.5" />
              {{ t('common.remove') }}
            </button>
          </div>

          <div class="flex items-center gap-1">
            <Filter class="w-3.5 h-3.5 opacity-50 flex-shrink-0" />
            <select
              v-model="filterCol"
              class="select select-xs select-bordered text-xs"
              @change="onSearchChange"
            >
              <option value="">{{ t('common.all.columns') }}</option>
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
            {{ t('table.advanced') }}
            <span v-if="appliedRuleCount > 0" class="badge badge-xs">{{ appliedRuleCount }}</span>
          </button>
        </div>

        <label class="flex items-center gap-1 flex-shrink-0">
          <span class="opacity-60 text-[11px] whitespace-nowrap">Header</span>
          <select
            class="select select-xs select-bordered max-w-[180px] text-xs"
            :value="activeSheet?.headerRowIndex ?? 0"
            :disabled="!!activeSheet?.paged"
            :title="t('table.choose.header.row.sheet')"
            @change="onChangeHeaderRow"
          >
            <option v-for="choice in headerChoices" :key="choice.index" :value="choice.index">
              {{ choice.label }}
            </option>
          </select>
        </label>

        <form v-if="rowTotal > 0" class="flex items-center gap-1" @submit.prevent="jumpToRowNumber">
          <span class="opacity-60 whitespace-nowrap">{{ t('table.go.row') }}</span>
          <input
            v-model.number="jumpRowInput"
            type="number"
            min="1"
            :max="rowTotal"
            class="input input-xs input-bordered w-24 text-center"
            :title="t('common.type.row.number.press.enter')"
            :aria-label="t('table.go.row')"
          />
          <button type="submit" class="btn btn-xs btn-ghost">{{ t('table.go') }}</button>
          <button type="button" class="btn btn-xs btn-ghost" :title="t('common.go.last.row')" @click="jumpToLastRow">
            {{ t('common.last') }}
          </button>
          <div class="flex items-center gap-0.5" :title="t('table.mark.color')">
            <button
              v-for="item in MARK_COLORS"
              :key="item.id"
              type="button"
              class="mark-swatch"
              :class="[markClass(item.id), { 'is-selected': markColor === item.id }]"
              :title="t(item.label)"
              :aria-label="t('mark.color', { name: t(item.label) })"
              @click="setCurrentMarkColor(item.id)"
            />
          </div>
          <button
            type="button"
            class="btn btn-xs"
            :class="currentMarked ? 'btn-warning' : 'btn-ghost'"
            :title="t('table.mark.unmark.current.row')"
            :disabled="selectedRowIndex === null"
            @click="toggleCurrentMark"
          >
            {{ t('table.mark') }}
          </button>
          <button
            type="button"
            class="btn btn-xs btn-ghost"
            :title="t('table.go.next.marked.row')"
            :disabled="markCount === 0"
            @click="jumpToNextMark"
          >
            {{ t('table.next.mark') }}<span v-if="markCount > 0"> ({{ markCount }})</span>
          </button>
          <button
            type="button"
            class="btn btn-xs btn-ghost"
            :title="t('table.view.edit.marked.rows')"
            :disabled="markCount === 0"
            @click="openMarkList"
          >
            {{ t('table.list') }}
          </button>
        </form>

        <!-- Nút thu gọn / mở rộng bảng -->
        <button
          v-if="!showForm"
          class="btn btn-ghost btn-xs gap-1 text-xs"
          :title="t('table.show.form')"
          @click="showForm = true"
        >
          <PanelLeftOpen class="w-3.5 h-3.5" />
          <span class="hidden sm:inline">{{ t('table.show.form') }}</span>
        </button>
        <button
          class="btn btn-ghost btn-xs gap-1 text-xs"
          :title="t('table.hide.table')"
          :disabled="!showForm"
          @click="hideTable"
        >
          <PanelRightClose class="w-3.5 h-3.5" />
          <span class="hidden sm:inline">{{ t('table.hide.table') }}</span>
        </button>
      </div>

      <div
        v-if="showAdvancedFilter"
        class="shrink-0 border-b border-base-300 bg-base-200/50 px-2 py-2 space-y-1.5"
      >
        <div class="flex flex-wrap items-center gap-1.5 text-xs">
          <select v-model="draftJoin" class="select select-xs select-bordered">
            <option value="and">{{ t('table.match.all') }}</option>
            <option value="or">{{ t('table.match.any') }}</option>
          </select>
          <button type="button" class="btn btn-xs btn-ghost gap-1" @click="addFilterRule">
            <Plus class="w-3.5 h-3.5" />
            Thêm điều kiện
          </button>
          <button type="button" class="btn btn-xs btn-primary" @click="applyAdvancedFilter">{{ t('common.apply') }}</button>
          <button type="button" class="btn btn-xs btn-ghost" @click="clearAdvancedFilter">{{ t('table.clear.filters') }}</button>
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
            <option value="true">{{ t('table.true') }}</option>
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
                    :checked="headerAllSelected"
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
            <tr v-if="rowWindow.paddingStart.value > 0" aria-hidden="true">
              <td
                :colspan="headers.length + 1"
                class="border-0 p-0"
                :style="{ height: `${rowWindow.paddingStart.value}px` }"
              />
            </tr>
            <tr
              v-for="item in rowWindow.items.value"
              :key="logicalAt(item.index)"
              class="data-row border-b border-base-300 transition-colors cursor-pointer"
              :style="{ height: `${ROW_HEIGHT}px` }"
              :class="dataRowClass(item.index)"
              @click="selectRowForEdit(logicalAt(item.index))"
              @dragover.prevent="onRowDragOver(logicalAt(item.index))"
              @drop.prevent="onRowDrop(logicalAt(item.index))"
            >
              <!-- Checkbox -->
              <td class="col-check border-r border-b border-base-300" @click.stop>
                <div class="cell-center gap-0.5">
                  <button
                    type="button"
                    class="row-grip"
                    draggable="true"
                    :title="t('table.drag.reorder.row')"
                    :aria-label="t('table.drag.reorder.row')"
                    @dragstart.stop="onRowDragStart(logicalAt(item.index), $event)"
                    @dragend.stop="onRowDragEnd"
                  >
                    <GripVertical class="w-3 h-3" />
                  </button>
                  <input
                    type="checkbox"
                    class="checkbox checkbox-xs checkbox-primary"
                    :checked="selectionStore.isRowSelected(logicalAt(item.index))"
                    @change="selectionStore.toggleRowSelection(logicalAt(item.index))"
                  />
                  <button
                    type="button"
                    class="row-index"
                    :class="{ 'is-marked': !!activeSheet && rowMarks.isMarked(activeSheet.id, logicalAt(item.index)) }"
                    :title="t('table.mark.row.selected.color')"
                    @click.stop="toggleRowMark(logicalAt(item.index))"
                  >
                    {{ logicalAt(item.index) + 1 }}
                  </button>
                </div>
              </td>

              <TableRowCell
                v-for="h in headers"
                :key="h.key"
                :value="rowAt(logicalAt(item.index))?.[h.key] ?? null"
                :column-key="h.key"
                :row-index="logicalAt(item.index)"
                :data-type="h.dataType"
                :date-format="h.dateFormat"
                :is-editing="
                  selectionStore.editingCell?.rowIndex === logicalAt(item.index) &&
                  selectionStore.editingCell?.columnKey === h.key
                "
                @start-edit="startCellEdit(logicalAt(item.index), h.key)"
                @end-edit="selectionStore.setEditingCell(null, null)"
                @update="(value) => saveCell(logicalAt(item.index), h.key, value)"
              />
            </tr>
            <tr v-if="rowWindow.paddingEnd.value > 0" aria-hidden="true">
              <td
                :colspan="headers.length + 1"
                class="border-0 p-0"
                :style="{ height: `${rowWindow.paddingEnd.value}px` }"
              />
            </tr>
          </tbody>
        </table>

        <!-- Empty State -->
        <div
          v-if="listCount === 0"
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
            <template v-if="index === 0">{{ t('table.total') }}</template>
            <template v-else-if="headers[index - 1]">
              {{ columnSum(headers[index - 1]) }}
            </template>
          </div>
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
            <h3 class="font-bold text-sm">{{ t('table.header.columns') }}</h3>
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
              <option value="text">{{ t('table.text') }}</option>
              <option value="number">{{ t('table.number') }}</option>
              <option value="date">{{ t('table.date') }}</option>
              <option value="boolean">{{ t('table.yes.no') }}</option>
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

    <dialog v-if="showMarkList" class="modal modal-open" @keydown.esc.prevent="closeMarkList">
      <div class="modal-box flex h-[min(820px,90vh)] w-[min(1080px,94vw)] max-w-none flex-col border border-base-300 p-0">
        <div class="flex items-center justify-between gap-3 border-b border-base-300 px-4 py-3">
          <div>
            <h3 class="text-sm font-bold">{{ t('table.marked.rows') }}</h3>
            <p class="text-xs opacity-60">{{ t('table.pick.row.edit.form.delete') }}</p>
          </div>
          <button type="button" class="btn btn-ghost btn-xs btn-square" title="Đóng" @click="closeMarkList">
            <X class="w-4 h-4" />
          </button>
        </div>
        <div class="flex min-h-0 flex-1">
          <div class="flex w-72 flex-shrink-0 flex-col border-r border-base-300">
            <p
              v-if="!activeSheet || rowMarks.list(activeSheet.id).length === 0"
              class="px-4 py-8 text-center text-sm opacity-60"
            >
              Chưa có dòng nào được đánh dấu.
            </p>
            <div v-else class="min-h-0 flex-1 overflow-auto">
              <div
                v-for="item in rowMarks.list(activeSheet.id)"
                :key="item.index"
                class="border-b border-base-300"
                :class="selectedRowIndex === item.index ? 'bg-base-200' : ''"
              >
                <button
                  type="button"
                  class="flex w-full items-start gap-2 px-3 py-2 text-left"
                  @click="selectRowForEdit(item.index)"
                >
                  <span class="mark-swatch mt-0.5 flex-shrink-0" :class="markClass(item.color)" />
                  <span class="min-w-0">
                    <span class="block font-mono text-xs font-bold">Dòng {{ item.index + 1 }}</span>
                    <span class="block truncate text-[11px] opacity-60">{{ markPreviews[item.index] || '(trống)' }}</span>
                  </span>
                </button>
                <div class="flex items-center justify-between gap-2 px-3 pb-2">
                  <div class="flex items-center gap-1">
                    <button
                      v-for="color in MARK_COLORS"
                      :key="color.id"
                      type="button"
                      class="mark-swatch"
                      :class="[markClass(color.id), { 'is-selected': item.color === color.id }]"
                      :title="t(color.label)"
                      :aria-label="t('mark.switch', { name: t(color.label) })"
                      @click="recolorMarkedRow(item.index, color.id)"
                    />
                  </div>
                  <button type="button" class="btn btn-ghost btn-xs text-error" @click="removeMarkedRow(item.index)">
                    Bỏ dấu
                  </button>
                </div>
              </div>
            </div>
          </div>
          <div class="min-h-0 min-w-0 flex-1">
            <RowForm
              :headers="headers"
              :sheet-id="activeSheet?.id"
              :selected-row-index="selectedRowIndex"
              :selected-row="selectedRow"
              :data-revision="workbookStore.dataRevision"
              :total-rows="rowTotal"
              :table-hidden="!showTableList"
              :preset-options="formPresets.presets.value"
              :selected-preset-id="formPresets.selectedPresetId.value"
              :shown-header-keys="shownHeaderKeys"
              :preset-empty="formPresets.presetEmpty.value"
              :row-is-new="
                selectedRowIndex !== null &&
                !!activeSheet &&
                workbookStore.isFreshRow(activeSheet.id, selectedRowIndex)
              "
              @create="handleCreateRow"
              @select-preset="formPresets.selectPreset"
              @manage-presets="presetDialogOpen = true"
              @commit="handleCommitFreshRow"
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
        </div>
      </div>
      <div class="modal-backdrop bg-black/40" @click="closeMarkList" />
    </dialog>
    <FormPresetDialog
      :open="presetDialogOpen"
      :headers="manageableHeaders"
      :presets="formPresets.presets.value"
      :error="presetError"
      @close="presetDialogOpen = false"
      @create="onCreatePreset"
      @update="onUpdatePreset"
      @remove="onRemovePreset"
    />
  </div>
</template>
