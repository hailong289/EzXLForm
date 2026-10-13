<script setup lang="ts">
import { ref, watch, computed, onUnmounted, nextTick } from 'vue'
import {
  Plus,
  Trash2,
  RotateCcw,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Check,
  FilePlus,
  FileEdit,
  Hash,
  Calendar,
  Type,
  CheckSquare,
  CheckCircle2,
  RefreshCw,
  Copy,
  Columns3,
  PanelLeftClose,
  PanelRightOpen,
  AlertCircle,
  Search,
  X,
  Maximize2,
} from 'lucide-vue-next'
import ConfirmDialog from '@/components/common/ConfirmDialog.vue'
import type { ColumnHeader, RowData } from '@/types/excel'
import { DATE_FORMATS, DEFAULT_DATE_FORMAT, formatDateValue, parseDateToIso } from '@/utils/dateFormat'
import { coerceCellValue, coerceRowData } from '@/utils/cellValue'
import { useLocale } from '@/composables/useLocale'

const { t } = useLocale()

const props = defineProps<{
  headers: ColumnHeader[]
  selectedRowIndex: number | null
  selectedRow: RowData | null
  dataRevision?: number
  totalRows: number
  sheetId?: string
  tableHidden?: boolean
  rowIsNew?: boolean
  presetOptions: { id: string; name: string }[]
  selectedPresetId: string
  shownHeaderKeys: string[] | null
  presetEmpty: boolean
}>()

const emit = defineEmits<{
  create: [data: RowData, placement: 'after' | 'end', anchor: number | null, kind: 'copy' | 'blank']
  commit: []
  update: [rowIndex: number, data: RowData]
  delete: [rowIndex: number]
  clear: []
  'select-prev': []
  'select-next': []
  'jump-to-row': [index: number]
  'edit-columns': []
  'set-date-format': [key: string, format: string]
  'select-preset': [presetId: string]
  'manage-presets': []
  hide: []
  'show-table': []
}>()

// Dữ liệu trong form
const formData = ref<Record<string, string>>({})
const showDeleteConfirm = ref(false)
const jumpInput = ref<number | ''>('')
type SyncUiState = 'idle' | 'syncing' | 'done'
const syncUi = ref<SyncUiState>('idle')
let saveIndicatorTimeout: ReturnType<typeof setTimeout> | null = null
let syncFallbackTimeout: ReturnType<typeof setTimeout> | null = null
let syncGeneration = 0
let awaitingSaveAck = false

function clearSyncTimers() {
  if (saveIndicatorTimeout) clearTimeout(saveIndicatorTimeout)
  if (syncFallbackTimeout) clearTimeout(syncFallbackTimeout)
  saveIndicatorTimeout = null
  syncFallbackTimeout = null
}

function resetSyncUi() {
  clearSyncTimers()
  awaitingSaveAck = false
  syncUi.value = 'idle'
}

function markSyncDone() {
  if (!awaitingSaveAck) return
  awaitingSaveAck = false
  syncUi.value = 'done'
  clearSyncTimers()
  saveIndicatorTimeout = setTimeout(() => {
    syncUi.value = 'idle'
  }, 1500)
}

const formFocused = ref(false)

const fieldErrors = ref<Record<string, string>>({})
const touchedFields = ref<Set<string>>(new Set())
const originalData = ref<Record<string, string>>({})

function syncFormFromSelection(force = false) {
  if (!force && formFocused.value) return
  if (props.selectedRowIndex !== null && props.selectedRow == null) return
  const newRow = props.selectedRow
  if (newRow && props.selectedRowIndex !== null) {
    const data: Record<string, string> = {}
    props.headers.forEach((h) => {
      const val = newRow[h.key]
      data[h.key] = h.dataType === 'date' ? formatDateValue(val, h.dateFormat) : val != null ? String(val) : ''
    })
    formData.value = data
    originalData.value = { ...data }
    fieldErrors.value = {}
    touchedFields.value = new Set()
    jumpInput.value = props.selectedRowIndex + 1
  } else {
    clearForm()
    jumpInput.value = ''
  }
}

watch(
  () => [props.sheetId, props.selectedRowIndex] as const,
  () => {
    resetSyncUi()
    formFocused.value = false
    syncFormFromSelection(true)
  },
  { immediate: true },
)

watch(
  () => props.selectedRow,
  () => {
    syncFormFromSelection(false)
  },
  { immediate: true },
)

watch(
  () => props.dataRevision,
  () => {
    if (awaitingSaveAck) markSyncDone()
    syncFormFromSelection(false)
  },
)

// Khi headers thay đổi, đảm bảo các keys tồn tại
watch(
  () => props.headers,
  (newHeaders) => {
    newHeaders.forEach((h) => {
      if (!(h.key in formData.value)) {
        formData.value[h.key] = ''
      }
    })
  },
  { deep: true },
)

function clearForm() {
  const data: Record<string, string> = {}
  props.headers.forEach((h) => {
    data[h.key] = ''
  })
  formData.value = data
  originalData.value = {}
  fieldErrors.value = {}
  touchedFields.value = new Set()
}

function parseFormData(): RowData {
  return coerceRowData(formData.value, props.headers).row
}

function headerOf(key: string): ColumnHeader | undefined {
  return props.headers.find((h) => h.key === key)
}

function refreshFieldError(key: string): void {
  const h = headerOf(key)
  const { error } = coerceCellValue(formData.value[key] ?? '', h?.dataType, h?.dateFormat)
  const next = { ...fieldErrors.value }
  if (error) next[key] = error
  else delete next[key]
  fieldErrors.value = next
}

function markTouched(key: string): void {
  const next = new Set(touchedFields.value)
  next.add(key)
  touchedFields.value = next
}

function visibleFieldError(key: string): string | undefined {
  return touchedFields.value.has(key) ? fieldErrors.value[key] : undefined
}


function datePattern(header: ColumnHeader): string {
  return header.dateFormat || DEFAULT_DATE_FORMAT
}

function onDateInput(header: ColumnHeader, event: Event) {
  const raw = (event.target as HTMLInputElement).value
  onFieldInput(header.key, raw)
}

function onDateBlur(header: ColumnHeader) {
  markTouched(header.key)
  const raw = formData.value[header.key] ?? ''
  const formatted = formatDateValue(raw, datePattern(header))
  if (!formatted || formatted === raw) return
  onFieldInput(header.key, formatted)
}

const datePickerRefs = new Map<string, HTMLInputElement>()

function setDatePickerRef(key: string, el: unknown) {
  if (el instanceof HTMLInputElement) datePickerRefs.set(key, el)
  else datePickerRefs.delete(key)
}

function openDatePicker(header: ColumnHeader) {
  const picker = datePickerRefs.get(header.key)
  if (!picker) return
  const iso = parseDateToIso(formData.value[header.key] ?? '', datePattern(header))
  if (iso) picker.value = iso
  try {
    picker.showPicker()
  } catch {
    picker.click()
  }
}

function onDatePicked(header: ColumnHeader, event: Event) {
  const iso = (event.target as HTMLInputElement).value
  onFieldInput(header.key, formatDateValue(iso, datePattern(header)))
}

function onTextAreaInput(key: string, event: Event) {
  const el = event.target as HTMLTextAreaElement
  onFieldInput(key, el.value)
  nextTick(fitRowTextAreas)
}

function onFieldInput(key: string, raw: string) {
  formFocused.value = true
  formData.value = { ...formData.value, [key]: raw }
  refreshFieldError(key)
  if (props.selectedRowIndex === null || props.selectedRow == null) return

  const generation = ++syncGeneration
  awaitingSaveAck = true
  syncUi.value = 'syncing'
  clearSyncTimers()
  emit('update', props.selectedRowIndex, parseFormData())
  syncFallbackTimeout = setTimeout(() => {
    if (generation !== syncGeneration || !awaitingSaveAck) return
    markSyncDone()
  }, 2500)
}

function handleCreate(placement: 'after' | 'end' = 'end', kind: 'copy' | 'blank' = 'copy') {
  if (kind === 'copy' && props.selectedRowIndex !== null && props.selectedRow == null) return
  if (kind === 'copy') {
    const { errors } = coerceRowData(formData.value, props.headers)
    if (Object.keys(errors).length > 0) {
      fieldErrors.value = errors
      touchedFields.value = new Set(Object.keys(errors))
      return
    }
  }
  const where = placement === 'after' && props.selectedRowIndex !== null ? 'after' : 'end'
  emit('create', parseFormData(), where, props.selectedRowIndex, kind)
}

function handleCommit() {
  emit('commit')
}

function handleDeleteClick() {
  if (props.selectedRowIndex === null) return
  showDeleteConfirm.value = true
}

function confirmDelete() {
  if (props.selectedRowIndex === null) return
  emit('delete', props.selectedRowIndex)
  showDeleteConfirm.value = false
}

function fieldControls(): HTMLElement[] {
  const host = gridHostRef.value
  if (!host) return []
  return [...host.querySelectorAll<HTMLElement>('.field-body')]
    .map((body) => body.querySelector<HTMLElement>('textarea, input:not([type="date"])'))
    .filter((control): control is HTMLElement => control != null)
}

function caretAtBoundary(control: HTMLElement, direction: 'left' | 'right' | 'up' | 'down'): boolean {
  if (!(control instanceof HTMLInputElement || control instanceof HTMLTextAreaElement)) return true
  if (control.type === 'checkbox' || control.type === 'number') {
    return direction === 'up' || direction === 'down' || direction === 'left' || direction === 'right'
  }
  const start = control.selectionStart ?? 0
  const end = control.selectionEnd ?? 0
  if (start !== end) return false
  const value = control.value
  if (direction === 'left') return start === 0
  if (direction === 'right') return end === value.length
  if (control instanceof HTMLInputElement) return true
  const lineStart = value.lastIndexOf('\n', Math.max(0, start - 1)) + 1
  const breakAt = value.indexOf('\n', start)
  const lineEnd = breakAt === -1 ? value.length : breakAt
  if (direction === 'up') return start === lineStart
  return start === lineEnd
}

function focusNeighbor(current: HTMLElement, direction: 'left' | 'right' | 'up' | 'down') {
  const controls = fieldControls()
  const index = controls.indexOf(current)
  if (index < 0) return
  const columns = Math.max(1, columnCount.value)
  const column = index % columns
  let next = -1
  if (direction === 'left' && column > 0) next = index - 1
  if (direction === 'right' && column < columns - 1) next = index + 1
  if (direction === 'up') next = index - columns
  if (direction === 'down') next = index + columns
  const target = controls[next]
  if (!target || next < 0 || next >= controls.length) return
  target.focus()
  if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
    if (target.type === 'checkbox' || target.type === 'number') return
    const pos = direction === 'left' || direction === 'up' ? target.value.length : 0
    try {
      target.setSelectionRange(pos, pos)
    } catch {
      /* number input không đặt được con trỏ */
    }
  }
}

function onFormFocusOut(event: FocusEvent) {
  const next = event.relatedTarget
  const root = event.currentTarget
  if (next instanceof Node && root instanceof Node && root.contains(next)) return
  formFocused.value = false
  syncFormFromSelection(false)
}

function onFormKeydown(event: KeyboardEvent) {
  if (event.ctrlKey && event.key === 'Enter' && !event.altKey) {
    if (!isEditMode.value) return
    event.preventDefault()
    if (event.shiftKey) emit('select-prev')
    else emit('select-next')
    return
  }
  if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return
  const direction = event.key === 'ArrowLeft'
    ? 'left'
    : event.key === 'ArrowRight'
      ? 'right'
      : event.key === 'ArrowUp'
        ? 'up'
        : event.key === 'ArrowDown'
          ? 'down'
          : null
  if (!direction) return
  const host = gridHostRef.value
  const target = event.target
  if (!host || !(target instanceof HTMLElement) || !host.contains(target)) return
  const control = target.closest('.field-body')?.querySelector<HTMLElement>('textarea, input:not([type="date"])')
  if (!control || (target !== control && !control.contains(target))) return
  if (!caretAtBoundary(control, direction)) return
  event.preventDefault()
  focusNeighbor(control, direction)
}

function handleNewRowMode() {
  clearForm()
  emit('clear')
  focusFirstField()
}

function addRowAfterWithPreset(presetId: string) {
  emit('select-preset', presetId)
  handleCreate('after', 'blank')
  ;(document.activeElement as HTMLElement | null)?.blur()
}

function focusFirstField() {
  nextTick(() => {
    const control = fieldControls()[0]
    if (!control) return
    control.focus()
    if (control instanceof HTMLInputElement || control instanceof HTMLTextAreaElement) {
      try {
        control.setSelectionRange(control.value.length, control.value.length)
      } catch {
        /* number/checkbox không đặt được con trỏ */
      }
    }
  })
}

function onJumpSubmit() {
  if (props.totalRows < 1) return
  const value = Number(jumpInput.value)
  if (!Number.isFinite(value)) return
  const rowNumber = Math.min(Math.max(1, Math.floor(value)), props.totalRows)
  jumpInput.value = rowNumber
  emit('jump-to-row', rowNumber - 1)
}

function jumpToLastRow() {
  if (props.totalRows < 1) return
  jumpInput.value = props.totalRows
  emit('jump-to-row', props.totalRows - 1)
}

const isEditMode = computed(() => props.selectedRowIndex !== null)
const hasRows = computed(() => props.totalRows > 0)

const formHeaders = computed(() => {
  const visible = props.headers.filter((header) => header.visible !== false)
  if (props.shownHeaderKeys == null) return visible
  const allowed = new Set(props.shownHeaderKeys)
  return visible.filter((header) => allowed.has(header.key))
})
const fieldSearch = ref('')

const filteredHeaders = computed(() => {
  const query = fieldSearch.value.trim().toLowerCase()
  if (!query) return formHeaders.value
  return formHeaders.value.filter((header) => header.label.toLowerCase().includes(query))
})

const filledCount = computed(() =>
  formHeaders.value.reduce(
    (count, header) => count + ((formData.value[header.key] ?? '').trim() !== '' ? 1 : 0),
    0,
  ),
)

const gridHostRef = ref<HTMLElement | null>(null)
const gridWidth = ref(520)
let gridObserver: ResizeObserver | null = null

/** Bước co giãn: 4 cột khi rộng, 2 khi vừa, 1 khi hẹp */
const MIN_FIELD_PX = 180
const GRID_GAP_PX = 10

function widthFor(count: number): number {
  return count * MIN_FIELD_PX + (count - 1) * GRID_GAP_PX
}

function optimalColumns(width: number): number {
  if (width >= widthFor(4)) return 4
  if (width >= widthFor(2)) return 2
  return 1
}

const columnCount = computed(() => optimalColumns(gridWidth.value))

const FIELD_MIN_PX = 32
const FIELD_MAX_PX = 220

/** Các ô văn bản cùng hàng cao bằng nội dung dài nhất trên hàng đó. */
function fitRowTextAreas() {
  const host = gridHostRef.value
  if (!host) return
  const cols = Math.max(1, columnCount.value)
  const cards = [...host.querySelectorAll<HTMLElement>('[data-field-card]')]
  const areas = cards
    .map((card) => card.querySelector<HTMLTextAreaElement>('textarea.form-text-input'))
    .filter((area): area is HTMLTextAreaElement => area != null)
  if (areas.length === 0) return

  const rows: HTMLTextAreaElement[][] = []
  for (let index = 0; index < areas.length; index += cols) {
    rows.push(areas.slice(index, index + cols))
  }

  rows.forEach((row) => {
    row.forEach((area) => {
      area.style.height = '0px'
      area.style.minHeight = '0px'
    })
    const tallest = Math.max(FIELD_MIN_PX, ...row.map((area) => area.scrollHeight))
    const next = Math.min(tallest, FIELD_MAX_PX)
    row.forEach((area) => {
      area.style.minHeight = `${next}px`
      area.style.height = `${next}px`
    })
  })
}

// ── Modal chỉnh sửa chi tiết ô văn bản ─────────────────────────────────────
const modalKey = ref<string | null>(null)
const modalDraft = ref('')

const modalHeader = computed(() =>
  modalKey.value ? props.headers.find((h) => h.key === modalKey.value) : null,
)

function openModal(key: string) {
  modalKey.value = key
  modalDraft.value = formData.value[key] ?? ''
}

function closeModal() {
  modalKey.value = null
  modalDraft.value = ''
}

function saveModal() {
  if (!modalKey.value) return
  onFieldInput(modalKey.value, modalDraft.value)
  closeModal()
}

watch(
  () => props.headers.map((header) => `${header.key}:${header.dateFormat ?? ''}`).join('|'),
  () => {
    const next = { ...formData.value }
    let changed = false
    props.headers.forEach((header) => {
      if (header.dataType !== 'date') return
      const formatted = formatDateValue(next[header.key], header.dateFormat)
      if (formatted && formatted !== next[header.key]) {
        next[header.key] = formatted
        changed = true
      }
    })
    if (changed) formData.value = next
  },
)

watch([formData, columnCount], () => nextTick(fitRowTextAreas), { deep: true })

watch(gridHostRef, (host) => {
  gridObserver?.disconnect()
  gridObserver = null
  if (!host) return
  gridObserver = new ResizeObserver((entries) => {
    gridWidth.value = entries[0]?.contentRect.width ?? gridWidth.value
    nextTick(fitRowTextAreas)
  })
  gridObserver.observe(host)
  nextTick(fitRowTextAreas)
})

onUnmounted(() => {
  gridObserver?.disconnect()
  clearSyncTimers()
})
</script>

<template>
  <div
    class="flex flex-col h-full bg-base-100 select-none overflow-hidden"
    @keydown="onFormKeydown"
    @focusin="formFocused = true"
    @focusout="onFormFocusOut"
  >
    <!-- ── Header Panel Gọn Gàng ───────────────────────────────────────────── -->
    <div class="p-3 bg-base-200/90 border-b border-base-300 flex items-center justify-between gap-3">
      <!-- Title & Mode Badge -->
      <div class="flex items-center gap-2.5">
        <div
          class="p-2 rounded-lg flex-shrink-0"
          :class="isEditMode ? 'bg-info/15 text-info' : 'bg-primary/15 text-primary'"
        >
          <FileEdit v-if="isEditMode" class="w-4 h-4" />
          <FilePlus v-else class="w-4 h-4" />
        </div>
        <div>
          <div class="flex items-center gap-2">
            <h3 class="font-bold text-sm leading-none">
              {{ !hasRows ? t('form.no.rows') : isEditMode ? t('form.row.n', { n: (props.selectedRowIndex ?? 0) + 1 }) : t('form.new.row') }}
            </h3>

            <span
              v-if="isEditMode && syncUi !== 'idle'"
              class="inline-flex items-center justify-center rounded-md p-1"
              :class="syncUi === 'done' ? 'bg-success/15 text-success' : 'bg-info/15 text-info'"
              :title="syncUi === 'done' ? t('form.saved') : t('form.saving')"
            >
              <CheckCircle2 v-if="syncUi === 'done'" class="w-3.5 h-3.5" />
              <RefreshCw v-else class="w-3.5 h-3.5 animate-spin" />
            </span>
            <span v-else-if="hasRows" class="badge badge-xs badge-primary font-semibold py-0.5 px-2">
              {{ t('form.new.mode') }}
            </span>
          </div>
          <span class="text-[11px] opacity-60 mt-0.5 block">
            {{
              !hasRows
                ? t('form.click.add.row.start.entering')
                : isEditMode
                  ? t('form.arrow.keys.move.between.fields')
                  : t('form.fill.fields.below.click.new')
            }}
          </span>
        </div>
      </div>

      <!-- Quick Reset / New Row Button -->
      <div class="flex items-center gap-1.5 flex-shrink-0">
        <button type="button" class="btn btn-xs btn-outline" @click="emit('manage-presets')">
          {{ t('common.form.types') }}
        </button>
        <button
          class="btn btn-xs btn-outline gap-1"
          :title="t('form.hide.form')"
          :disabled="props.tableHidden"
          @click="emit('hide')"
        >
          <PanelLeftClose class="w-3.5 h-3.5" /> {{ t('form.hide.form') }}
        </button>
        <button
          v-if="props.tableHidden"
          class="btn btn-xs btn-outline btn-primary gap-1"
          :title="t('form.show.table')"
          @click="emit('show-table')"
        >
          <PanelRightOpen class="w-3.5 h-3.5" /> {{ t('form.show.table') }}
        </button>
        <button
          class="btn btn-xs btn-outline gap-1"
          :title="t('form.add.rename.header.columns')"
          @click="emit('edit-columns')"
        >
          <Columns3 class="w-3.5 h-3.5" /> {{ t('form.columns') }}
        </button>
        <button
          v-if="isEditMode"
          class="btn btn-xs btn-outline btn-primary gap-1"
          :title="t('form.switch.new.row.entry')"
          @click="handleNewRowMode"
        >
          <Plus class="w-3.5 h-3.5" /> {{ t('form.create.row') }}
        </button>
      </div>
    </div>

    <!-- ── Thanh Điều Hướng Dòng & Nhảy Dòng ───────────────────────────────── -->
    <div
      v-if="props.totalRows > 0"
      class="px-3 py-1.5 bg-base-200/50 border-b border-base-300 flex items-center justify-between gap-2 text-xs"
    >
      <!-- Jump to row input -->
      <div class="flex items-center gap-1.5">
        <span class="opacity-70 text-[11px]">{{ t('form.viewing') }}</span>
        <input
          v-model.number="jumpInput"
          type="number"
          min="1"
          :max="props.totalRows"
          class="input input-xs input-bordered w-24 text-center font-bold text-xs h-6 px-1"
          :placeholder="t('form.row.number')"
          :title="t('common.type.row.number.press.enter')"
          @keydown.enter="onJumpSubmit"
        />
        <span class="opacity-60 text-[11px]">{{ t('form.n.rows', { n: props.totalRows }) }}</span>
        <button
          type="button"
          class="btn btn-ghost btn-xs h-6 px-2 text-[11px]"
          :title="t('common.go.last.row')"
          @click="jumpToLastRow"
        >
          {{ t('common.last') }}
        </button>
      </div>

      <!-- Prev / Next Buttons -->
      <div class="flex items-center gap-1">
        <button
          class="btn btn-ghost btn-xs gap-1 h-6 px-2 text-[11px]"
          :disabled="props.selectedRowIndex === null || props.selectedRowIndex <= 0"
          :title="t('form.previous.row.ctrl.shift.enter')"
          @click="emit('select-prev')"
        >
          <ChevronLeft class="w-3.5 h-3.5" /> {{ t('form.previous') }}
        </button>
        <button
          class="btn btn-ghost btn-xs gap-1 h-6 px-2 text-[11px]"
          :disabled="props.selectedRowIndex === null || props.selectedRowIndex >= props.totalRows - 1"
          :title="t('form.next.row.ctrl.enter')"
          @click="emit('select-next')"
        >
          {{ t('form.next') }} <ChevronRight class="w-3.5 h-3.5" />
        </button>
      </div>
    </div>

    <!-- ── Thanh tiến độ nhập liệu (chế độ tạo mới) ──────────────────────────── -->
    <div
      v-if="hasRows && !isEditMode && formHeaders.length > 0"
      class="px-3 py-1.5 bg-base-200/50 border-b border-base-300 flex items-center gap-2"
    >
      <span class="text-[11px] opacity-70 flex-shrink-0">{{ t('form.filled') }}</span>
      <progress
        class="progress progress-primary flex-1 h-2"
        :value="filledCount"
        :max="formHeaders.length"
      />
      <span class="text-[11px] opacity-60 font-semibold flex-shrink-0">
        {{ filledCount }}/{{ formHeaders.length }}
      </span>
    </div>

    <div
      v-if="!hasRows"
      class="flex-1 flex flex-col items-center justify-center gap-3 p-6 text-center"
    >
      <p class="text-sm text-base-content/70">{{ t('form.sheet.has.no.data.rows') }}</p>
      <button
        type="button"
        class="btn btn-primary btn-sm gap-1.5"
        :title="t('form.add.blank.row')"
        @click="handleCreate('end', 'blank')"
      >
        <Plus class="w-4 h-4" />
        {{ t('form.add.row') }}
      </button>
    </div>

    <!-- ── Form Inputs (Bố cục Lưới Gọn Gàng & Rộng Rãi) ────────────────────── -->
    <div v-else ref="gridHostRef" class="flex-1 overflow-y-auto p-4 scrollbar-thin">
      <div v-if="formHeaders.length > 8" class="mb-3">
        <label class="input input-bordered input-sm flex items-center gap-2 w-full">
          <Search class="w-3.5 h-3.5 opacity-50" />
          <input
            v-model="fieldSearch"
            type="text"
            class="grow text-xs bg-transparent"
            :placeholder="t('form.find.column.form')"
            :aria-label="t('form.find.column.label')"
          />
          <button
            v-if="fieldSearch"
            type="button"
            class="btn btn-ghost btn-xs btn-circle"
            :title="t('form.clear.search')"
            @click="fieldSearch = ''"
          >
            <X class="w-3 h-3" />
          </button>
        </label>
      </div>

      <div v-if="presetEmpty" class="flex flex-col items-center justify-center h-48 px-6 text-center text-xs text-base-content/70">
        {{ t('common.form.type.has.no.columns') }}
      </div>
      <div v-else-if="formHeaders.length === 0" class="flex flex-col items-center justify-center h-48 opacity-50 text-xs">
        {{ t('form.sheet.has.no.data.columns') }}
      </div>
      <div
        v-else-if="filteredHeaders.length === 0"
        class="flex flex-col items-center justify-center h-48 opacity-50 text-xs"
      >
        {{ t('form.no.column.matches.q', { q: fieldSearch }) }}
      </div>

      <div
        v-else
        data-field-grid
        class="grid gap-3 items-stretch"
        :style="{ gridTemplateColumns: `repeat(${columnCount}, minmax(0, 1fr))` }"
      >
        <div
          v-for="(h, index) in filteredHeaders"
          :key="h.key"
          data-field-card
          class="bg-base-100 hover:bg-base-200/40 transition-colors p-3 rounded-xl border border-base-300 shadow-sm flex flex-col"
          :class="{ 'border-error/60': !!visibleFieldError(h.key) }"
        >
          <div class="flex items-center justify-between mb-1 gap-2 shrink-0">
            <label :for="`field-${index}`" class="text-xs font-bold text-base-content/90 leading-tight flex items-center gap-1.5 min-w-0" :title="h.label">
              <Type v-if="h.dataType === 'text' || !h.dataType" class="w-3.5 h-3.5 text-primary opacity-70 flex-shrink-0" />
              <Hash v-else-if="h.dataType === 'number'" class="w-3.5 h-3.5 text-info opacity-70 flex-shrink-0" />
              <Calendar v-else-if="h.dataType === 'date'" class="w-3.5 h-3.5 text-warning opacity-70 flex-shrink-0" />
              <CheckSquare v-else class="w-3.5 h-3.5 text-success opacity-70 flex-shrink-0" />
              <span class="truncate min-w-0">{{ h.label }}</span>
              <span
                v-if="formData[h.key] !== originalData[h.key]"
                class="w-1.5 h-1.5 rounded-full bg-warning flex-shrink-0"
                :title="t('form.changed')"
              />
            </label>

            <div class="flex items-center gap-1 flex-shrink-0">
              <span
                v-if="h.dataType !== 'date'"
                class="badge badge-ghost badge-xs text-[10px] uppercase font-mono opacity-50"
              >
                {{ h.dataType || 'text' }}
              </span>
              <select
                v-else
                class="select select-xs select-bordered h-6 min-h-0 w-[6.5rem] text-[10px] font-mono"
                :value="datePattern(h)"
                :title="t('form.date.format')"
                @change="emit('set-date-format', h.key, ($event.target as HTMLSelectElement).value)"
              >
                <option v-for="pattern in DATE_FORMATS" :key="pattern" :value="pattern">
                  {{ pattern }}
                </option>
              </select>
              <button
                v-if="h.dataType === 'text' || !h.dataType"
                type="button"
                class="inline-flex h-5 w-5 items-center justify-center rounded text-base-content/40 hover:bg-base-200 hover:text-base-content"
                :title="t('form.expand.editor')"
                @click="openModal(h.key)"
              >
                <Maximize2 class="w-3 h-3" />
              </button>
            </div>
          </div>

          <div class="field-body mt-1.5 min-h-0 flex-1">
            <textarea
              v-if="h.dataType === 'text' || !h.dataType"
              :id="`field-${index}`"
              :key="`${props.sheetId ?? ''}-${props.selectedRowIndex ?? 'new'}-${h.key}`"
              :value="formData[h.key]"
              class="form-text-input textarea textarea-bordered w-full text-xs bg-base-100 rounded-lg resize-none scrollbar-thin overflow-y-auto !min-h-0 !py-1.5 !leading-snug"
              :class="{ 'textarea-error': !!visibleFieldError(h.key) }"
              :aria-label="h.label"
              :aria-invalid="!!visibleFieldError(h.key)"
              :placeholder="t('form.enter.label', { label: h.label })"
              @input="onTextAreaInput(h.key, $event)"
              @blur="markTouched(h.key)"
            />
            <div v-else-if="h.dataType === 'date'" class="relative w-full">
              <input
                type="text"
                :id="`field-${index}`"
                class="form-text-input input input-bordered w-full text-xs bg-base-100 rounded-lg pr-8"
                :class="{ 'input-error': !!visibleFieldError(h.key) }"
                :value="formData[h.key]"
                :aria-label="h.label"
                :aria-invalid="!!visibleFieldError(h.key)"
                :placeholder="datePattern(h)"
                :title="t('form.date.pattern.format', { pattern: datePattern(h) })"
                @input="onDateInput(h, $event)"
                @blur="onDateBlur(h)"
                @keydown.enter="!isEditMode && handleCreate()"
              />
              <button
                type="button"
                class="absolute right-1 top-1/2 -translate-y-1/2 btn btn-ghost btn-xs btn-square min-h-0 h-6 w-6"
                :title="t('form.pick.date.label', { label: h.label })"
                :aria-label="t('form.pick.date.label', { label: h.label })"
                @click="openDatePicker(h)"
              >
                <Calendar class="w-3.5 h-3.5" />
              </button>
              <input
                :ref="(el) => setDatePickerRef(h.key, el)"
                type="date"
                class="sr-only"
                tabindex="-1"
                @change="onDatePicked(h, $event)"
              />
            </div>
            <input
              v-else-if="h.dataType !== 'boolean'"
              :id="`field-${h.key}`"
              :value="formData[h.key]"
              type="number"
              class="form-text-input input input-bordered w-full text-xs bg-base-100 rounded-lg"
              :class="{ 'input-error': !!visibleFieldError(h.key) }"
              :aria-label="h.label"
              :aria-invalid="!!visibleFieldError(h.key)"
              :placeholder="t('form.enter.label', { label: h.label })"
              :title="t('form.enter.number.use.dot.decimal')"
              @input="onFieldInput(h.key, ($event.target as HTMLInputElement).value)"
              @blur="markTouched(h.key)"
              @keydown.enter="!isEditMode && handleCreate()"
            />

            <label
              v-else
              class="form-text-input flex items-center gap-2.5 cursor-pointer bg-base-100 rounded-lg border border-base-300"
            >
              <input
                type="checkbox"
                class="checkbox checkbox-sm checkbox-primary"
                :checked="formData[h.key] === 'true' || formData[h.key] === '1'"
                :aria-label="h.label"
                @change="
                  onFieldInput(h.key, ($event.target as HTMLInputElement).checked ? 'true' : 'false')
                "
              />
              <span class="text-xs font-semibold">
                {{ formData[h.key] === 'true' || formData[h.key] === '1' ? t('form.true.yes') : t('form.false.no') }}
              </span>
            </label>
          </div>

          <p
            v-if="visibleFieldError(h.key)"
            class="flex items-start gap-1 text-[10px] text-error mt-1.5 leading-tight"
          >
            <AlertCircle class="w-3 h-3 flex-shrink-0 mt-px" />
            {{ visibleFieldError(h.key) }}
          </p>
        </div>
      </div>
    </div>

    <!-- ── Cụm Nút Hành Động [Tạo Dòng Mới] [Xóa Dòng] [Làm Mới] ─────────────── -->
    <div
      v-if="hasRows"
      class="p-3 bg-base-200/90 border-t border-base-300 flex items-center justify-between gap-2 flex-wrap"
    >
      <!-- Nút Tạo dòng mới -->
      <button
        v-if="props.rowIsNew"
        type="button"
        class="btn btn-warning btn-sm gap-1.5 text-xs font-bold"
        :title="t('form.save.new.row.clear.highlight')"
        @click="handleCommit"
      >
        <CheckCircle2 class="w-4 h-4" /> {{ t('form.save.new.row') }}
      </button>
      <button
        v-if="isEditMode"
        type="button"
        class="btn btn-primary btn-sm flex-1 gap-1.5 text-xs font-bold shadow-xs min-w-[140px]"
        :title="t('form.copy.row.n.row.just', { n: (props.selectedRowIndex ?? 0) + 1 })"
        :disabled="!props.selectedRow"
        @click="handleCreate('after', 'copy')"
      >
        <Copy class="w-4 h-4" />
        {{ t('form.copy.below') }}
      </button>
      <div v-if="isEditMode" class="dropdown dropdown-top min-w-[140px] flex-1">
        <button
          type="button"
          tabindex="0"
          class="btn btn-secondary btn-sm w-full gap-1.5 text-xs font-bold"
          :title="t('form.add.blank.row.right.after', { n: (props.selectedRowIndex ?? 0) + 1 })"
        >
          <Plus class="w-4 h-4" />
          {{ t('form.add.row.after.n', { n: (props.selectedRowIndex ?? 0) + 1 }) }}
          <ChevronDown class="w-3.5 h-3.5 opacity-80" />
        </button>
        <ul
          tabindex="0"
          class="dropdown-content menu z-50 mb-1 w-56 rounded-box border border-base-300 bg-base-100 p-1 shadow-lg"
        >
          <li class="menu-title px-2 py-1 text-[10px]">{{ t('form.choose.form.type') }}</li>
          <li>
            <button type="button" class="text-xs" @click="addRowAfterWithPreset('')">
              <Check class="h-3.5 w-3.5" :class="selectedPresetId ? 'opacity-0' : ''" />
              {{ t('common.all.columns') }}
            </button>
          </li>
          <li v-for="preset in presetOptions" :key="preset.id">
            <button type="button" class="text-xs" @click="addRowAfterWithPreset(preset.id)">
              <Check class="h-3.5 w-3.5" :class="selectedPresetId === preset.id ? '' : 'opacity-0'" />
              <span class="truncate">{{ preset.name }}</span>
            </button>
          </li>
        </ul>
      </div>
      <button
        v-if="!isEditMode"
        type="button"
        class="btn btn-primary btn-sm flex-1 gap-1.5 text-xs font-bold shadow-xs min-w-[140px]"
        :title="t('form.add.current.entry.end.table')"
        @click="handleCreate('end', 'copy')"
      >
        <Plus class="w-4 h-4" />
        {{ t('form.create.row') }}
      </button>
      <button
        v-if="isEditMode"
        type="button"
        class="btn btn-outline btn-sm gap-1 text-xs"
        :title="t('form.add.blank.row.end.table')"
        @click="handleCreate('end', 'blank')"
      >
        <Plus class="w-3.5 h-3.5" /> {{ t('form.add.end') }}
      </button>

      <!-- Nhóm nút phụ: Xóa & Làm mới -->
      <div class="flex items-center gap-2">
        <button
          v-if="isEditMode"
          class="btn btn-error btn-outline btn-sm gap-1 text-xs"
          :title="t('form.delete.selected.row.table')"
          @click="handleDeleteClick"
        >
          <Trash2 class="w-3.5 h-3.5" /> {{ t('form.delete.row') }}
        </button>

        <button
          class="btn btn-ghost btn-sm gap-1 text-xs"
          :title="t('form.clear.form.new.row')"
          @click="handleNewRowMode"
        >
          <RotateCcw class="w-3.5 h-3.5" /> {{ t('form.clear') }}
        </button>
      </div>
    </div>

    <!-- ── Confirm Dialog Xóa Dòng ─────────────────────────────────────────── -->
    <ConfirmDialog
      v-model="showDeleteConfirm"
      :title="t('form.confirm.delete')"
      :message="t('form.delete.row.n.you.can', { n: (props.selectedRowIndex ?? 0) + 1 })"
      type="error"
      :confirm-text="t('form.delete.this.row')"
      @confirm="confirmDelete"
    />

    <!-- ── Modal Chỉnh Sửa Chi Tiết Ô Văn Bản ───────────────────────────────── -->
    <dialog
      v-if="modalKey !== null"
      class="modal modal-open"
      @keydown.esc.prevent="closeModal"
    >
      <div class="modal-box max-w-xl border border-base-300 p-0 flex flex-col max-h-[85vh] overflow-hidden">
        <!-- Modal Header -->
        <div class="px-5 py-3.5 border-b border-base-300 bg-base-200/60 flex items-center justify-between gap-3 flex-shrink-0">
          <div class="flex items-center gap-2 min-w-0">
            <Maximize2 class="w-4 h-4 opacity-60 flex-shrink-0" />
            <h3 class="font-bold text-sm truncate">{{ modalHeader?.label }}</h3>
            <span class="badge badge-ghost badge-xs text-[10px] uppercase font-mono opacity-50 flex-shrink-0">
              {{ modalHeader?.dataType || 'text' }}
            </span>
          </div>
          <button
            type="button"
            class="btn btn-ghost btn-xs btn-square flex-shrink-0"
            :title="t('common.close')"
            @click="closeModal"
          >
            <X class="w-4 h-4" />
          </button>
        </div>

        <!-- Modal Body -->
        <div class="p-5 flex-1 overflow-y-auto flex flex-col gap-3 scrollbar-thin">
          <textarea
            v-model="modalDraft"
            rows="10"
            autofocus
            class="textarea textarea-bordered w-full text-sm bg-base-100 rounded-lg resize-y scrollbar-thin leading-relaxed"
            :placeholder="t('form.enter.label', { label: modalHeader?.label ?? '' })"
          />
          <div class="flex items-center justify-between text-[11px] text-base-content/50">
            <span>{{ t('form.changes.write.straight.into.matching') }}</span>
            <span class="font-semibold text-base-content/70">{{ t('form.n.characters', { n: modalDraft.length }) }}</span>
          </div>
        </div>

        <!-- Modal Footer -->
        <div class="px-5 py-3 border-t border-base-300 bg-base-200/60 flex items-center justify-end gap-2 flex-shrink-0">
          <button type="button" class="btn btn-sm btn-ghost" @click="closeModal">{{ t('common.cancel') }}</button>
          <button type="button" class="btn btn-sm btn-primary gap-1.5" @click="saveModal">
            <CheckCircle2 class="w-3.5 h-3.5" />
            {{ t('form.save.changes') }}
          </button>
        </div>
      </div>
      <div class="modal-backdrop bg-black/40" @click="closeModal" />
    </dialog>
  </div>
</template>
