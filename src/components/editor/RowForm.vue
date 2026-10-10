<script setup lang="ts">
import { ref, watch, computed, onMounted, onUnmounted, nextTick } from 'vue'
import {
  Plus,
  Trash2,
  RotateCcw,
  ChevronLeft,
  ChevronRight,
  FilePlus,
  FileEdit,
  Hash,
  Calendar,
  Type,
  CheckSquare,
  CheckCircle2,
  Copy,
  Columns3,
  PanelLeftClose,
  PanelRightOpen,
  AlertCircle,
  Search,
  X,
} from 'lucide-vue-next'
import ConfirmDialog from '@/components/common/ConfirmDialog.vue'
import type { ColumnHeader, RowData } from '@/types/excel'
import { DATE_FORMATS, DEFAULT_DATE_FORMAT, formatDateValue, parseDateToIso } from '@/utils/dateFormat'
import { coerceCellValue, coerceRowData } from '@/utils/cellValue'

const props = defineProps<{
  headers: ColumnHeader[]
  selectedRowIndex: number | null
  selectedRow: RowData | null
  dataRevision?: number
  totalRows: number
  sheetId?: string
  tableHidden?: boolean
  rowIsNew?: boolean
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
  hide: []
  'show-table': []
}>()

// Dữ liệu trong form
const formData = ref<Record<string, string>>({})
const showDeleteConfirm = ref(false)
const jumpInput = ref<number | ''>('')
const isRealtimeSaved = ref(false)
let saveIndicatorTimeout: ReturnType<typeof setTimeout> | null = null

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

/**
 * Ghi giá trị vừa gõ rồi đẩy xuống bảng ngay.
 * Đọc từ event, vì @input có thể chạy trước khi v-model cập nhật formData.
 */
function fitAllTextAreas() {
  const host = gridHostRef.value
  if (!host) return
  host.querySelectorAll('textarea.form-text-input').forEach((node) => {
    const area = node as HTMLTextAreaElement
    area.style.height = '0px'
    area.style.minHeight = '0px'
    const next = Math.max(area.scrollHeight, 32)
    area.style.minHeight = `${next}px`
    area.style.height = `${next}px`
  })
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
  fitAllTextAreas()
}

function onFieldInput(key: string, raw: string) {
  formFocused.value = true
  formData.value = { ...formData.value, [key]: raw }
  refreshFieldError(key)
  if (props.selectedRowIndex === null || props.selectedRow == null) return

  emit('update', props.selectedRowIndex, parseFormData())

  isRealtimeSaved.value = true
  if (saveIndicatorTimeout) clearTimeout(saveIndicatorTimeout)
  saveIndicatorTimeout = setTimeout(() => {
    isRealtimeSaved.value = false
  }, 1500)
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
  if (typeof jumpInput.value === 'number' && jumpInput.value >= 1 && jumpInput.value <= props.totalRows) {
    emit('jump-to-row', jumpInput.value - 1)
  }
}

const isEditMode = computed(() => props.selectedRowIndex !== null)

const formHeaders = computed(() => props.headers.filter((h) => h.visible !== false))

const fieldSearch = ref('')

const filteredHeaders = computed(() => {
  const q = fieldSearch.value.trim().toLowerCase()
  if (!q) return formHeaders.value
  return formHeaders.value.filter((h) => h.label.toLowerCase().includes(q))
})

const filledCount = computed(() =>
  formHeaders.value.reduce(
    (n, h) => n + ((formData.value[h.key] ?? '').trim() !== '' ? 1 : 0),
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

watch([formData, columnCount], () => nextTick(fitAllTextAreas), { deep: true })

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

onMounted(() => {
  if (!gridHostRef.value) return
  gridObserver = new ResizeObserver((entries) => {
    gridWidth.value = entries[0]?.contentRect.width ?? gridWidth.value
  })
  gridObserver.observe(gridHostRef.value)
  nextTick(fitAllTextAreas)
})

onUnmounted(() => {
  gridObserver?.disconnect()
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
              {{ isEditMode ? `Dòng #${(props.selectedRowIndex ?? 0) + 1}` : 'Tạo Dòng Mới' }}
            </h3>

            <!-- Realtime indicator -->
            <span
              v-if="isEditMode"
              class="badge badge-xs font-semibold py-0.5 px-2 gap-1"
              :class="isRealtimeSaved ? 'badge-success text-success-content' : 'badge-info'"
            >
              <CheckCircle2 v-if="isRealtimeSaved" class="w-3 h-3" />
              {{ isRealtimeSaved ? 'Đã lưu Realtime' : 'Sửa Realtime' }}
            </span>
            <span v-else class="badge badge-xs badge-primary font-semibold py-0.5 px-2">
              Chế độ tạo mới
            </span>
          </div>
          <span class="text-[11px] opacity-60 mt-0.5 block">
            {{ isEditMode ? 'Mũi tên sang ô bên cạnh, Ctrl+Enter sang dòng kế' : 'Nhập thông tin bên dưới và bấm "Tạo dòng mới"' }}
          </span>
        </div>
      </div>

      <!-- Quick Reset / New Row Button -->
      <div class="flex items-center gap-1.5 flex-shrink-0">
        <button
          class="btn btn-xs btn-outline gap-1"
          title="Ẩn form"
          :disabled="props.tableHidden"
          @click="emit('hide')"
        >
          <PanelLeftClose class="w-3.5 h-3.5" /> Ẩn form
        </button>
        <button
          v-if="props.tableHidden"
          class="btn btn-xs btn-outline btn-primary gap-1"
          title="Hiện bảng"
          @click="emit('show-table')"
        >
          <PanelRightOpen class="w-3.5 h-3.5" /> Hiện bảng
        </button>
        <button
          class="btn btn-xs btn-outline gap-1"
          title="Thêm hoặc sửa tên cột header"
          @click="emit('edit-columns')"
        >
          <Columns3 class="w-3.5 h-3.5" /> Cột
        </button>
        <button
          v-if="isEditMode"
          class="btn btn-xs btn-outline btn-primary gap-1"
          title="Chuyển sang chế độ tạo dòng mới"
          @click="handleNewRowMode"
        >
          <Plus class="w-3.5 h-3.5" /> Tạo dòng mới
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
        <span class="opacity-70 text-[11px]">Đang xem:</span>
        <input
          v-model.number="jumpInput"
          type="number"
          min="1"
          :max="props.totalRows"
          class="input input-xs input-bordered w-14 text-center font-bold text-xs h-6 px-1"
          placeholder="Số..."
          @keydown.enter="onJumpSubmit"
          @blur="onJumpSubmit"
        />
        <span class="opacity-60 text-[11px]">/ {{ props.totalRows }} dòng</span>
      </div>

      <!-- Prev / Next Buttons -->
      <div class="flex items-center gap-1">
        <button
          class="btn btn-ghost btn-xs gap-1 h-6 px-2 text-[11px]"
          :disabled="props.selectedRowIndex === null || props.selectedRowIndex <= 0"
          title="Dòng trước (Ctrl+Shift+Enter)"
          @click="emit('select-prev')"
        >
          <ChevronLeft class="w-3.5 h-3.5" /> Trước
        </button>
        <button
          class="btn btn-ghost btn-xs gap-1 h-6 px-2 text-[11px]"
          :disabled="props.selectedRowIndex === null || props.selectedRowIndex >= props.totalRows - 1"
          title="Dòng kế tiếp (Ctrl+Enter)"
          @click="emit('select-next')"
        >
          Sau <ChevronRight class="w-3.5 h-3.5" />
        </button>
      </div>
    </div>

    <!-- ── Thanh tiến độ nhập liệu (chế độ tạo mới) ──────────────────────────── -->
    <div
      v-if="!isEditMode && formHeaders.length > 0"
      class="px-3 py-1.5 bg-base-200/50 border-b border-base-300 flex items-center gap-2"
    >
      <span class="text-[11px] opacity-70 flex-shrink-0">Nhập liệu:</span>
      <progress
        class="progress progress-primary flex-1 h-2"
        :value="filledCount"
        :max="formHeaders.length"
      />
      <span class="text-[11px] opacity-60 font-semibold flex-shrink-0">
        {{ filledCount }}/{{ formHeaders.length }}
      </span>
    </div>

    <!-- ── Form Inputs (Bố cục Lưới Gọn Gàng & Rộng Rãi) ────────────────────── -->
    <div ref="gridHostRef" class="flex-1 overflow-y-auto p-4 scrollbar-thin">
      <div v-if="formHeaders.length > 8" class="mb-3">
        <label class="input input-bordered input-sm flex items-center gap-2 w-full">
          <Search class="w-3.5 h-3.5 opacity-50" />
          <input
            v-model="fieldSearch"
            type="text"
            class="grow text-xs bg-transparent"
            placeholder="Tìm cột trong form..."
            aria-label="Tìm cột trong form"
          />
          <button
            v-if="fieldSearch"
            type="button"
            class="btn btn-ghost btn-xs btn-circle"
            title="Xóa tìm kiếm"
            @click="fieldSearch = ''"
          >
            <X class="w-3 h-3" />
          </button>
        </label>
      </div>

      <div v-if="formHeaders.length === 0" class="flex flex-col items-center justify-center h-48 opacity-50 text-xs">
        Sheet này chưa có cột dữ liệu nào.
      </div>
      <div
        v-else-if="filteredHeaders.length === 0"
        class="flex flex-col items-center justify-center h-48 opacity-50 text-xs"
      >
        Không tìm thấy cột khớp với "{{ fieldSearch }}".
      </div>

      <div
        v-else
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
          <!-- Label & Data Type Badge -->
          <div class="flex items-center justify-between mb-1 gap-2">
            <label :for="`field-${index}`" class="text-xs font-bold text-base-content/90 leading-tight flex items-center gap-1.5 min-w-0" :title="h.label">
              <Type v-if="h.dataType === 'text' || !h.dataType" class="w-3.5 h-3.5 text-primary opacity-70 flex-shrink-0" />
              <Hash v-else-if="h.dataType === 'number'" class="w-3.5 h-3.5 text-info opacity-70 flex-shrink-0" />
              <Calendar v-else-if="h.dataType === 'date'" class="w-3.5 h-3.5 text-warning opacity-70 flex-shrink-0" />
              <CheckSquare v-else class="w-3.5 h-3.5 text-success opacity-70 flex-shrink-0" />
              <span class="line-clamp-2 min-w-0">{{ h.label }}</span>
              <span
                v-if="formData[h.key] !== originalData[h.key]"
                class="w-1.5 h-1.5 rounded-full bg-warning flex-shrink-0"
                title="Đã thay đổi"
              />
            </label>

            <span
              v-if="h.dataType !== 'date'"
              class="badge badge-ghost badge-xs text-[10px] uppercase font-mono opacity-50 flex-shrink-0"
            >
              {{ h.dataType || 'text' }}
            </span>
            <select
              v-else
              class="select select-xs select-bordered h-6 min-h-0 w-[6.5rem] text-[10px] font-mono"
              :value="datePattern(h)"
              title="Định dạng ngày"
              @change="emit('set-date-format', h.key, ($event.target as HTMLSelectElement).value)"
            >
              <option v-for="pattern in DATE_FORMATS" :key="pattern" :value="pattern">
                {{ pattern }}
              </option>
            </select>
          </div>

          <!-- Input Fields (Realtime on input) -->
          <div class="field-body mt-1.5">
            <textarea
              v-if="h.dataType !== 'boolean' && h.dataType !== 'number' && h.dataType !== 'date'"
              :id="`field-${index}`"
              :key="`${props.sheetId ?? ''}-${props.selectedRowIndex ?? 'new'}-${h.key}`"
              :value="formData[h.key]"
              rows="1"
              class="form-text-input textarea textarea-bordered w-full text-xs bg-base-100 rounded-lg resize-y"
              :class="{ 'textarea-error': !!visibleFieldError(h.key) }"
              :aria-label="h.label"
              :aria-invalid="!!visibleFieldError(h.key)"
              :placeholder="`Nhập ${h.label}...`"
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
                :title="`Ngày theo định dạng ${datePattern(h)}`"
                @input="onDateInput(h, $event)"
                @blur="onDateBlur(h)"
                @keydown.enter="!isEditMode && handleCreate()"
              />
              <button
                type="button"
                class="absolute right-1 top-1/2 -translate-y-1/2 btn btn-ghost btn-xs btn-square min-h-0 h-6 w-6"
                :title="`Chọn ngày cho ${h.label}`"
                :aria-label="`Chọn ngày cho ${h.label}`"
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
              :id="`field-${index}`"
              :value="formData[h.key]"
              :type="h.dataType === 'number' ? 'number' : 'text'"
              class="form-text-input input input-bordered w-full text-xs bg-base-100 rounded-lg"
              :class="{ 'input-error': !!visibleFieldError(h.key) }"
              :aria-label="h.label"
              :aria-invalid="!!visibleFieldError(h.key)"
              :placeholder="`Nhập ${h.label}...`"
              :title="'Nhập số, dùng dấu chấm cho phần thập phân'"
              @input="onFieldInput(h.key, ($event.target as HTMLInputElement).value)"
              @blur="markTouched(h.key)"
              @keydown.enter="!isEditMode && handleCreate()"
            />

            <!-- Boolean Toggle Checkbox -->
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
                {{ formData[h.key] === 'true' || formData[h.key] === '1' ? 'Đúng (True / Yes)' : 'Sai (False / No)' }}
              </span>
            </label>
          </div>

          <!-- Error message -->
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
    <div class="p-3 bg-base-200/90 border-t border-base-300 flex items-center justify-between gap-2 flex-wrap">
      <!-- Nút Tạo dòng mới -->
      <button
        v-if="props.rowIsNew"
        type="button"
        class="btn btn-warning btn-sm gap-1.5 text-xs font-bold"
        title="Lưu dòng mới và bỏ màu đánh dấu"
        @click="handleCommit"
      >
        <CheckCircle2 class="w-4 h-4" /> Lưu dòng mới
      </button>
      <button
        v-if="isEditMode"
        type="button"
        class="btn btn-primary btn-sm flex-1 gap-1.5 text-xs font-bold shadow-xs min-w-[140px]"
        :title="`Sao chép dữ liệu dòng #${(props.selectedRowIndex ?? 0) + 1} thành dòng ngay bên dưới`"
        :disabled="!props.selectedRow"
        @click="handleCreate('after', 'copy')"
      >
        <Copy class="w-4 h-4" />
        Sao chép xuống dưới
      </button>
      <button
        v-if="isEditMode"
        type="button"
        class="btn btn-secondary btn-sm flex-1 gap-1.5 text-xs font-bold min-w-[140px]"
        :title="`Thêm một dòng trống ngay sau dòng #${(props.selectedRowIndex ?? 0) + 1}`"
        @click="handleCreate('after', 'blank')"
      >
        <Plus class="w-4 h-4" />
        Thêm dòng mới sau #{{ (props.selectedRowIndex ?? 0) + 1 }}
      </button>
      <button
        v-if="!isEditMode"
        type="button"
        class="btn btn-primary btn-sm flex-1 gap-1.5 text-xs font-bold shadow-xs min-w-[140px]"
        title="Thêm dòng đang nhập vào cuối bảng"
        @click="handleCreate('end', 'copy')"
      >
        <Plus class="w-4 h-4" />
        Tạo dòng mới
      </button>
      <button
        v-if="isEditMode"
        type="button"
        class="btn btn-outline btn-sm gap-1 text-xs"
        title="Thêm một dòng trống ở cuối bảng"
        @click="handleCreate('end', 'blank')"
      >
        <Plus class="w-3.5 h-3.5" /> Thêm ở cuối
      </button>

      <!-- Nhóm nút phụ: Xóa & Làm mới -->
      <div class="flex items-center gap-2">
        <button
          v-if="isEditMode"
          class="btn btn-error btn-outline btn-sm gap-1 text-xs"
          title="Xóa dòng đang chọn khỏi bảng"
          @click="handleDeleteClick"
        >
          <Trash2 class="w-3.5 h-3.5" /> Xóa dòng
        </button>

        <button
          class="btn btn-ghost btn-sm gap-1 text-xs"
          title="Xóa trắng form để chuẩn bị tạo dòng mới"
          @click="handleNewRowMode"
        >
          <RotateCcw class="w-3.5 h-3.5" /> Làm mới
        </button>
      </div>
    </div>

    <!-- ── Confirm Dialog Xóa Dòng ─────────────────────────────────────────── -->
    <ConfirmDialog
      v-model="showDeleteConfirm"
      title="Xác nhận xóa dòng"
      :message="`Bạn có chắc chắn muốn xóa dòng #${(props.selectedRowIndex ?? 0) + 1}? Bạn có thể hoàn tác bằng Ctrl+Z bất cứ lúc nào.`"
      type="error"
      confirm-text="Xóa dòng này"
      @confirm="confirmDelete"
    />
  </div>
</template>
