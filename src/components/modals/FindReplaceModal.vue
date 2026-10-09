<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight, Loader2, Replace, Search, X } from 'lucide-vue-next'
import { useWorkbookStore } from '@/stores/workbookStore'
import { PAGE_SIZE_OPTIONS, useAppConfigStore } from '@/stores/appConfigStore'
import { useHistory } from '@/composables/useHistory'
import { useProgressTaskStore } from '@/stores/progressTaskStore'
import { scanRows } from '@/composables/rowPageStore'
import { rowMatchesFind } from '@/utils/rowFilter'
import { formatDateValue } from '@/utils/dateFormat'
import type { ColumnHeader, FindReplaceOptions, RowData } from '@/types/excel'

const props = defineProps<{
  modelValue: boolean
}>()

const emit = defineEmits<{
  'update:modelValue': [val: boolean]
}>()

const workbookStore = useWorkbookStore()
const appConfigStore = useAppConfigStore()
const { push: pushHistory } = useHistory()
const progressStore = useProgressTaskStore()

const PREVIEW_LIMIT = 200

const step = ref<'form' | 'list'>('form')
const findText = ref('')
const replaceText = ref('')
const selectedColumn = ref<string>('all')
const caseSensitive = ref(false)
const exactMatch = ref(false)
const useRegex = ref(false)
const matchEmpty = ref(false)
const scanning = ref(false)
const statusText = ref('')
const matchCount = ref(0)
const matchedRows = ref<Array<{ index: number; cells: Record<string, string> }>>([])
const listPage = ref(1)
const listPageSize = ref<number>(20)
const activeOptions = ref<FindReplaceOptions | null>(null)
const jobs = new Map<string, FindReplaceOptions>()
const runningTaskIds = new Set<string>()

function rememberStoredJobs() {
  for (const task of progressStore.tasks) {
    if (task.replaceOptions) jobs.set(task.id, task.replaceOptions)
  }
}

rememberStoredJobs()

watch(
  () => progressStore.tasks.map((task) => `${task.id}:${task.status}`).join('|'),
  () => {
    rememberStoredJobs()
    void pump()
  },
)

const availableColumns = computed(() => workbookStore.activeSheet?.headers ?? [])
const listHeaders = computed(() => workbookStore.activeSheet?.headers ?? [])
const busy = computed(() => scanning.value)
const listPageCount = computed(() => Math.max(1, Math.ceil(matchedRows.value.length / listPageSize.value)))
const visibleMatches = computed(() => {
  const start = (listPage.value - 1) * listPageSize.value
  return matchedRows.value.slice(start, start + listPageSize.value)
})

function onListPageSize(event: Event) {
  const next = Number((event.target as HTMLSelectElement).value)
  if (!Number.isFinite(next) || next <= 0) return
  listPageSize.value = next
  listPage.value = 1
}

watch(
  () => props.modelValue,
  (open) => {
    if (!open) return
    step.value = 'form'
    scanning.value = false
    matchedRows.value = []
    matchCount.value = 0
    listPage.value = 1
    activeOptions.value = null
  },
)

function close() {
  if (busy.value) return
  emit('update:modelValue', false)
}

function currentOptions(): FindReplaceOptions {
  const empty = matchEmpty.value || findText.value.trim() === ''
  return {
    findText: empty ? '' : findText.value,
    replaceText: replaceText.value,
    columns: selectedColumn.value === 'all' ? [] : [selectedColumn.value],
    caseSensitive: caseSensitive.value,
    exactMatch: exactMatch.value,
    useRegex: useRegex.value,
    matchEmpty: empty,
  }
}

function previewValue(value: RowData[string], header: ColumnHeader): string {
  if (value == null || value === '') return ''
  if (header.dataType === 'date') return formatDateValue(value, header.dateFormat)
  if (header.dataType === 'boolean') return value ? 'Có' : 'Không'
  if (header.dataType === 'number') {
    const numeric = Number(value)
    return Number.isNaN(numeric) ? String(value) : numeric.toLocaleString('vi-VN')
  }
  return String(value)
}

function rememberMatch(index: number, row: RowData, headers: ColumnHeader[]) {
  if (matchedRows.value.length >= PREVIEW_LIMIT) return
  const cells: Record<string, string> = {}
  for (const header of headers) cells[header.key] = previewValue(row[header.key], header)
  matchedRows.value.push({ index, cells })
}

async function handleFindOnly() {
  const sheet = workbookStore.activeSheet
  if (!sheet || busy.value) return
  const options = currentOptions()
  activeOptions.value = options
  step.value = 'list'
  scanning.value = true
  statusText.value = 'Đang lọc danh sách...'
  matchedRows.value = []
  matchCount.value = 0
  listPage.value = 1
  await nextTick()
  await new Promise((resolve) => setTimeout(resolve, 30))

  const headers = sheet.headers
  let count = 0
  try {
    if (sheet.paged) {
      await scanRows(sheet.id, sheet.rowCount ?? 0, (rowIndex, row) => {
        if (rowMatchesFind(row, headers, options)) {
          count += 1
          rememberMatch(rowIndex, row, headers)
        }
        if (rowIndex > 0 && rowIndex % 8000 === 0) {
          statusText.value = `Đang lọc ${rowIndex.toLocaleString('vi-VN')} dòng...`
        }
        return true
      })
    } else {
      for (let index = 0; index < sheet.rows.length; index++) {
        const row = sheet.rows[index]
        if (rowMatchesFind(row, headers, options)) {
          count += 1
          rememberMatch(index, row, headers)
        }
        if (index > 0 && index % 8000 === 0) {
          statusText.value = `Đang lọc ${index.toLocaleString('vi-VN')} dòng...`
          await new Promise((resolve) => setTimeout(resolve, 0))
        }
      }
    }
    matchCount.value = count
  } finally {
    scanning.value = false
    statusText.value = ''
  }
}

function backToForm() {
  if (busy.value) return
  step.value = 'form'
}

async function executeTask(taskId: string): Promise<boolean> {
  const options = jobs.get(taskId)
  const signal = progressStore.signal(taskId)
  const sheet = workbookStore.workbook?.sheets.find((item) => item.id === progressStore.tasks.find((task) => task.id === taskId)?.sheetId)
  if (!options || !signal) return true
  if (!sheet) {
    if (progressStore.tasks.some((task) => task.id === taskId)) progressStore.setStatus(taskId, 'paused')
    return false
  }
  progressStore.setStatus(taskId, 'running')
  const report = (done: number, total: number, stats?: { matchedRows: number; changedRows: number }) => {
    const percent = total <= 0 ? 100 : Math.min(99, Math.round((done / total) * 100))
    progressStore.setPercent(taskId, percent)
    if (stats) progressStore.setDetail(taskId, replaceProgressDetail(options, stats.changedRows, stats.matchedRows))
  }
  try {
    const result = sheet.paged
      ? await workbookStore.findAndReplacePaged(sheet.id, options, report, signal)
      : await replaceInMemory(sheet.id, options, report, signal)
    if (!progressStore.tasks.some((task) => task.id === taskId)) return true
    progressStore.setPercent(taskId, 100)
    progressStore.setDetail(
      taskId,
      replaceProgressDetail(options, result.affectedRowCount ?? 0, result.matchedRowCount ?? 0),
    )
    progressStore.setStatus(taskId, 'done')
    appConfigStore.addToast({
      type: result.replacedCount > 0 ? 'success' : 'info',
      message: result.replacedCount > 0
        ? `Đã thay thế ${result.replacedCount.toLocaleString('vi-VN')} kết quả.`
        : 'Không có ô nào được thay thế.',
    })
  } catch (error) {
    if (progressStore.tasks.some((task) => task.id === taskId)) {
      progressStore.setStatus(taskId, 'cancelled')
    }
    appConfigStore.addToast({
      type: 'error',
      message: error instanceof Error ? error.message : 'Thay thế thất bại.',
    })
  }
  return true
}

async function replaceInMemory(
  sheetId: string,
  options: FindReplaceOptions,
  report: (done: number, total: number) => void,
  signal: { checkpoint: () => Promise<boolean> },
) {
  const sheet = workbookStore.workbook?.sheets.find((item) => item.id === sheetId)
  const rowsBackup = sheet ? JSON.parse(JSON.stringify(sheet.rows)) as RowData[] : []
  const result = await workbookStore.findAndReplace(sheetId, options, report, signal)
  if (result.replacedCount > 0 && sheet) {
    pushHistory({
      type: 'FIND_REPLACE',
      description: options.matchEmpty
        ? `Điền ô trống bằng "${options.replaceText}" (${result.replacedCount} ô)`
        : `Thay thế "${options.findText}" bằng "${options.replaceText}" (${result.replacedCount} ô)`,
      sheetId,
      before: { rows: rowsBackup },
      after: { rows: JSON.parse(JSON.stringify(sheet.rows)) as RowData[] },
    })
  }
  return result
}

async function pump() {
  const pending = progressStore.tasks.filter((task) => {
    if (runningTaskIds.has(task.id) || !jobs.has(task.id)) return false
    return task.status === 'queued' || task.status === 'running'
  })
  for (const task of pending) {
    runningTaskIds.add(task.id)
    void runReplaceTask(task.id)
  }
}

async function runReplaceTask(taskId: string) {
  let finished = false
  try {
    finished = await executeTask(taskId)
  } finally {
    runningTaskIds.delete(taskId)
    if (finished) {
      jobs.delete(taskId)
      progressStore.release(taskId)
    }
    void pump()
  }
}

function replaceProgressDetail(options: FindReplaceOptions, changedRows: number, matchedRows: number): string {
  const action = options.matchEmpty
    ? `Ô trống → "${options.replaceText}"`
    : `"${options.findText}" → "${options.replaceText}"`
  const changed = changedRows.toLocaleString('vi-VN')
  const matched = matchedRows.toLocaleString('vi-VN')
  return `${action} · ${changed} / ${matched} dòng`
}

async function replaceAll() {
  const sheet = workbookStore.activeSheet
  const options = activeOptions.value
  if (!sheet || !options || busy.value || matchCount.value === 0) return
  const savedOptions: FindReplaceOptions = {
    findText: options.findText,
    replaceText: options.replaceText,
    columns: [...options.columns],
    caseSensitive: options.caseSensitive,
    exactMatch: options.exactMatch,
    useRegex: options.useRegex,
    matchEmpty: options.matchEmpty,
  }
  const detail = replaceProgressDetail(savedOptions, 0, matchCount.value)
  const taskId = progressStore.enqueue({
    title: 'Thay thế',
    detail,
    sheetId: sheet.id,
    replaceOptions: savedOptions,
  })
  jobs.set(taskId, savedOptions)
  emit('update:modelValue', false)
  await nextTick()
  void pump()
}
</script>

<template>
  <dialog v-if="props.modelValue" class="modal modal-open select-none">
    <div
      class="modal-box relative border border-base-300 p-5 shadow-2xl"
      :class="step === 'list' ? 'flex max-h-[85vh] w-11/12 max-w-5xl flex-col' : 'max-w-md'"
    >
      <div
        v-if="busy"
        class="absolute inset-0 z-20 flex flex-col items-center justify-center gap-2 rounded-2xl bg-base-100/80"
      >
        <Loader2 class="h-6 w-6 animate-spin text-primary" />
        <span class="text-sm font-medium">{{ statusText || 'Đang xử lý...' }}</span>
      </div>

      <template v-if="step === 'form'">
        <div class="mb-4 flex items-center justify-between border-b border-base-300 pb-3">
          <div class="flex items-center gap-2">
            <Replace class="h-5 w-5 text-primary" />
            <h3 class="text-base font-bold">Tìm kiếm & Thay thế</h3>
          </div>
          <button class="btn btn-ghost btn-xs btn-circle" @click="close">
            <X class="h-4 w-4" />
          </button>
        </div>

        <div class="flex flex-col gap-3">
          <div>
            <label class="label py-1 text-xs font-semibold">Tìm kiếm:</label>
            <div class="relative">
              <input
                v-model="findText"
                type="text"
                placeholder="Để trống để lọc ô trống"
                class="input input-sm input-bordered w-full pr-8"
                :disabled="matchEmpty"
                @keydown.enter="handleFindOnly"
              />
              <Search class="pointer-events-none absolute right-2.5 top-2.5 h-4 w-4 opacity-40" />
            </div>
          </div>

          <div>
            <label class="label py-1 text-xs font-semibold">Thay thế bằng:</label>
            <div class="relative">
              <input
                v-model="replaceText"
                type="text"
                placeholder="Nhập giá trị mới..."
                class="input input-sm input-bordered w-full pr-8"
                @keydown.enter="handleFindOnly"
              />
              <Replace class="pointer-events-none absolute right-2.5 top-2.5 h-4 w-4 opacity-40" />
            </div>
          </div>

          <div>
            <label class="label py-1 text-xs font-semibold">Phạm vi cột:</label>
            <select v-model="selectedColumn" class="select select-sm select-bordered w-full text-xs">
              <option value="all">Tất cả các cột</option>
              <option v-for="col in availableColumns" :key="col.key" :value="col.key">
                {{ col.label }}
              </option>
            </select>
          </div>

          <div class="mt-1 grid grid-cols-2 gap-2 rounded-lg border border-base-300 bg-base-200/50 p-2.5">
            <label class="flex cursor-pointer items-center gap-2 text-xs">
              <input v-model="caseSensitive" type="checkbox" class="checkbox checkbox-xs checkbox-primary" />
              <span>Phân biệt hoa/thường</span>
            </label>
            <label class="flex cursor-pointer items-center gap-2 text-xs">
              <input v-model="exactMatch" type="checkbox" class="checkbox checkbox-xs checkbox-primary" />
              <span>Khớp toàn bộ ô</span>
            </label>
            <label class="col-span-2 flex cursor-pointer items-center gap-2 text-xs">
              <input v-model="matchEmpty" type="checkbox" class="checkbox checkbox-xs checkbox-primary" />
              <span>Ô trống</span>
            </label>
            <label class="col-span-2 flex cursor-pointer items-center gap-2 text-xs" :class="matchEmpty ? 'pointer-events-none opacity-40' : ''">
              <input v-model="useRegex" type="checkbox" class="checkbox checkbox-xs checkbox-primary" :disabled="matchEmpty" />
              <span>Sử dụng Biểu thức chính quy (Regex)</span>
            </label>
          </div>
        </div>

        <div class="modal-action mt-5 flex justify-end gap-2 border-t border-base-300 pt-3">
          <button class="btn btn-sm btn-ghost" @click="close">Đóng</button>
          <button class="btn btn-sm btn-primary gap-1" @click="handleFindOnly">
            <Search class="h-3.5 w-3.5" /> Lọc
          </button>
        </div>
      </template>

      <template v-else>
        <div class="mb-3 flex items-center justify-between gap-2 border-b border-base-300 pb-3">
          <div class="min-w-0">
            <h3 class="text-base font-bold">Danh sách đã lọc</h3>
            <p class="truncate text-xs text-base-content/70">
              {{ matchCount.toLocaleString('vi-VN') }} dòng
              {{ activeOptions?.matchEmpty ? 'có ô trống' : `khớp "${activeOptions?.findText}"` }}
              <template v-if="matchCount > matchedRows.length">
                · đang hiện {{ matchedRows.length.toLocaleString('vi-VN') }} dòng đầu
              </template>
            </p>
          </div>
          <button class="btn btn-ghost btn-xs btn-circle" :disabled="busy" @click="close">
            <X class="h-4 w-4" />
          </button>
        </div>

        <div class="min-h-0 flex-1 overflow-auto rounded-lg border border-base-300">
          <table class="table table-xs table-pin-rows">
            <thead>
              <tr>
                <th v-for="header in listHeaders" :key="header.key">{{ header.label }}</th>
              </tr>
            </thead>
            <tbody>
              <tr v-if="!scanning && matchedRows.length === 0">
                <td :colspan="Math.max(listHeaders.length, 1)" class="py-8 text-center text-sm text-base-content/60">
                  Không có dòng nào khớp.
                </td>
              </tr>
              <tr v-for="item in visibleMatches" :key="item.index">
                <td v-for="header in listHeaders" :key="header.key" class="max-w-48 truncate">
                  {{ item.cells[header.key] }}
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        <div v-if="matchedRows.length > 0" class="mt-2 flex flex-wrap items-center justify-end gap-2 text-xs">
          <label class="flex items-center gap-1">
            <span class="whitespace-nowrap opacity-60">Số dòng</span>
            <select
              class="select select-xs select-bordered w-20"
              :value="listPageSize"
              title="Số dòng mỗi trang"
              :disabled="busy"
              @change="onListPageSize"
            >
              <option v-for="size in PAGE_SIZE_OPTIONS" :key="size" :value="size">{{ size }}</option>
            </select>
          </label>
          <div class="flex items-center gap-1">
            <button
              type="button"
              class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0"
              title="Trang đầu"
              :disabled="listPage <= 1 || busy"
              @click="listPage = 1"
            >
              <ChevronsLeft class="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0"
              title="Trang trước"
              :disabled="listPage <= 1 || busy"
              @click="listPage -= 1"
            >
              <ChevronLeft class="h-3.5 w-3.5" />
            </button>
            <span class="px-1">Trang {{ listPage }} / {{ listPageCount }}</span>
            <button
              type="button"
              class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0"
              title="Trang sau"
              :disabled="listPage >= listPageCount || busy"
              @click="listPage += 1"
            >
              <ChevronRight class="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0"
              title="Trang cuối"
              :disabled="listPage >= listPageCount || busy"
              @click="listPage = listPageCount"
            >
              <ChevronsRight class="h-3.5 w-3.5" />
            </button>
          </div>
        </div>

        <div class="modal-action mt-3 flex flex-wrap items-center justify-between gap-2 border-t border-base-300 pt-3">
          <button class="btn btn-sm btn-ghost" :disabled="busy" @click="backToForm">Sửa bộ lọc</button>
          <button
            class="btn btn-sm btn-primary gap-1"
            :disabled="busy || matchCount === 0"
            @click="replaceAll"
          >
            <Replace class="h-3.5 w-3.5" />
            Thay thế tất cả bằng "{{ activeOptions?.replaceText }}"
          </button>
        </div>
      </template>
    </div>
    <div class="modal-backdrop bg-black/40" @click="close" />
  </dialog>
</template>
