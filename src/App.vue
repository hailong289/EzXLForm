<script setup lang="ts">
import { ref, watch, onMounted, onUnmounted } from 'vue'
import { FileSpreadsheet, Upload, Sparkles } from 'lucide-vue-next'
import AppNavbar from '@/components/layout/AppNavbar.vue'
import SheetTabsBar from '@/components/sheets/SheetTabsBar.vue'
import StatusBar from '@/components/layout/StatusBar.vue'
import EditorTable from '@/components/editor/EditorTable.vue'
import ExcelPreview from '@/components/preview/ExcelPreview.vue'
import FindReplaceModal from '@/components/modals/FindReplaceModal.vue'
import ImportOptionsModal from '@/components/modals/ImportOptionsModal.vue'
import ToastContainer from '@/components/common/ToastContainer.vue'
import ProgressTaskDock from '@/components/common/ProgressTaskDock.vue'
import IntroPage from '@/components/layout/IntroPage.vue'
import { useWorkbookStore } from '@/stores/workbookStore'
import { useAppConfigStore } from '@/stores/appConfigStore'
import { useProgressTaskStore } from '@/stores/progressTaskStore'
import { useExcelParser } from '@/composables/useExcelParser'
import { LARGE_FILE_BYTES, isImportStopped } from '@/composables/importLargeWorkbook'
import { useTheme } from '@/composables/useTheme'
import { useHistory } from '@/composables/useHistory'
import { useSqliteStorage } from '@/composables/useSqliteStorage'
import { clearAllRowMarks, rowMarksRevision } from '@/composables/rowMarks'
import { syncBodyRows } from '@/utils/sheetGrid'
import { debounce } from 'lodash-es'
import type { ImportOptions, WorkbookData } from '@/types/excel'

const workbookStore = useWorkbookStore()
const appConfigStore = useAppConfigStore()
const progressStore = useProgressTaskStore()
const { readExcelFile, exportToExcel, exportToCsv, isLoading, importProgress } = useExcelParser()
const { loadSavedTheme } = useTheme()
const { undo, redo, canUndo, canRedo } = useHistory()
const {
  saveSessionToSqlite,
  loadSessionFromSqlite,
  loadSessionById,
  downloadSqliteDatabase,
  deleteFileSession,
  adoptSessionForFile,
  setActiveSessionId,
  activeSessionId,
  sessions,
} = useSqliteStorage()

// Tự động lưu phiên làm việc vào SQLite mỗi khi dữ liệu thay đổi
const debouncedSaveSqlite = debounce(async () => {
  if (workbookStore.workbook && activeSessionId.value) {
    await saveSessionToSqlite(workbookStore.workbook, activeSessionId.value)
  }
}, 800)

async function flushCurrentSession() {
  debouncedSaveSqlite.cancel()
  if (workbookStore.workbook && activeSessionId.value) {
    await saveSessionToSqlite(workbookStore.workbook, activeSessionId.value)
  }
}

function onPageHide() {
  void flushCurrentSession()
}

watch(
  () => workbookStore.workbook,
  () => {
    if (workbookStore.workbook) {
      debouncedSaveSqlite()
    }
  },
  { deep: true },
)

watch(
  () => workbookStore.dataRevision,
  () => {
    if (workbookStore.workbook) debouncedSaveSqlite()
  },
)

watch(rowMarksRevision, () => {
  if (workbookStore.workbook) debouncedSaveSqlite()
})

// Modal states
const showFindReplace = ref(false)
const showImportOptions = ref(false)
const pendingFile = ref<File | null>(null)
const isDraggingOver = ref(false)
const fileInputRef = ref<HTMLInputElement | null>(null)
const showIntro = ref(true)

// ─── File Handling ──────────────────────────────────────────────────────────

function triggerFileInput() {
  fileInputRef.value?.click()
}

function onFileSelected(e: Event) {
  const target = e.target as HTMLInputElement
  const file = target.files?.[0]
  if (file) {
    promptImportOptions(file)
  }
  target.value = ''
}

function handleDragOver(e: DragEvent) {
  e.preventDefault()
  isDraggingOver.value = true
}

function handleDragLeave() {
  isDraggingOver.value = false
}

function handleDrop(e: DragEvent) {
  e.preventDefault()
  isDraggingOver.value = false
  const file = e.dataTransfer?.files?.[0]
  if (file) {
    promptImportOptions(file)
  }
}

function promptImportOptions(file: File) {
  pendingFile.value = file
  showImportOptions.value = true
}

async function handleImportConfirm(options: ImportOptions) {
  const file = pendingFile.value
  if (!file) return
  pendingFile.value = null

  const extension = file.name.split('.').pop()?.toLowerCase()
  const large = (extension === 'xlsx' || extension === 'csv') && file.size >= LARGE_FILE_BYTES
  if (large) {
    const taskId = progressStore.enqueue({
      title: 'Nhập file',
      detail: file.name,
      sheetId: '',
      kind: 'import',
    })
    importJobs.set(taskId, { file, options })
    void pumpImports()
    return
  }

  try {
    const wb = await readExcelFile(file, options)
    await openImportedWorkbook(file.name, wb)
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Lỗi khi mở file'
    appConfigStore.addToast({ type: 'error', message: msg })
  }
}

const importJobs = new Map<string, { file: File; options: ImportOptions }>()
const runningImports = new Set<string>()

async function openImportedWorkbook(fileName: string, wb: WorkbookData) {
  await flushCurrentSession()
  adoptSessionForFile(fileName)
  workbookStore.setWorkbook(wb)
  clearAllRowMarks()
  appConfigStore.clearSearch()
  const bigRows = wb.sheets.reduce(
    (sum, sheet) => sum + (sheet.paged ? sheet.rowCount ?? 0 : sheet.rows.length),
    0,
  )
  appConfigStore.addToast({
    type: 'success',
    message: wb.sheets.some((sheet) => sheet.paged)
      ? `Đã mở "${fileName}" theo trang (${bigRows.toLocaleString('vi-VN')} dòng)`
      : `Đã mở thành công tệp "${fileName}" (${wb.sheets.length} sheets)`,
  })
}

async function pumpImports() {
  for (const [id, job] of importJobs) {
    if (runningImports.has(id)) continue
    const task = progressStore.tasks.find((item) => item.id === id)
    if (!task || task.status !== 'queued') continue
    runningImports.add(id)
    void runLargeImport(id, job).finally(() => {
      runningImports.delete(id)
      importJobs.delete(id)
      progressStore.release(id)
      void pumpImports()
    })
  }
}

async function runLargeImport(taskId: string, job: { file: File; options: ImportOptions }) {
  const signal = progressStore.signal(taskId)
  if (!signal) return
  progressStore.setStatus(taskId, 'running')
  let lastPercent = -1
  try {
    const wb = await readExcelFile(job.file, job.options, {
      checkpoint: () => signal.checkpoint(),
      onRatio: (ratio) => {
        const percent = Math.min(99, Math.round(ratio * 100))
        if (percent === lastPercent) return
        lastPercent = percent
        progressStore.setPercent(taskId, percent)
      },
      onProgress: (message) => {
        progressStore.setDetail(taskId, `${job.file.name} · ${message}`)
      },
    })
    if (!progressStore.tasks.some((task) => task.id === taskId)) return
    progressStore.setPercent(taskId, 100)
    progressStore.setStatus(taskId, 'done')
    await openImportedWorkbook(job.file.name, wb)
  } catch (err: unknown) {
    if (isImportStopped(err)) {
      if (progressStore.tasks.some((task) => task.id === taskId)) {
        progressStore.setStatus(taskId, 'cancelled')
        progressStore.setDetail(taskId, `${job.file.name} · đã dừng`)
      }
      return
    }
    if (progressStore.tasks.some((task) => task.id === taskId)) {
      progressStore.setStatus(taskId, 'cancelled')
    }
    const msg = err instanceof Error ? err.message : 'Lỗi khi mở file'
    appConfigStore.addToast({ type: 'error', message: msg })
  }
}

async function handleExportXlsx() {
  if (!workbookStore.workbook) return
  try {
    await exportToExcel(workbookStore.workbook)
    workbookStore.isDirty = false
    appConfigStore.addToast({
      type: 'success',
      message: 'Đã xuất file .xlsx thành công!',
    })
  } catch (err: unknown) {
    if (err instanceof DOMException && err.name === 'AbortError') return
    const msg = err instanceof Error ? err.message : 'Lỗi xuất file'
    appConfigStore.addToast({ type: 'error', message: msg })
  }
}

async function handleExportCsv() {
  if (!workbookStore.activeSheet) return
  try {
    await exportToCsv(workbookStore.activeSheet)
    appConfigStore.addToast({
      type: 'success',
      message: `Đã xuất sheet "${workbookStore.activeSheet.name}" ra .csv`,
    })
  } catch (err: unknown) {
    if (err instanceof DOMException && err.name === 'AbortError') return
    const msg = err instanceof Error ? err.message : 'Lỗi xuất CSV'
    appConfigStore.addToast({ type: 'error', message: msg })
  }
}

// Tạo dữ liệu mẫu nhanh để test mà không cần sẵn file Excel
function loadSampleData() {
  const sampleWb: WorkbookData = {
    meta: {
      fileName: 'Danh_Sach_Nhan_Vien_Mau.xlsx',
      fileSize: 12400,
      lastModified: new Date(),
      sheetCount: 2,
    },
    sheets: [
      {
        id: 'sheet_sample_1',
        name: 'Nhân Sự',
        headers: [
          { key: 'id', label: 'Mã NV', originalIndex: 0, dataType: 'text' },
          { key: 'name', label: 'Họ và Tên', originalIndex: 1, dataType: 'text' },
          { key: 'department', label: 'Phòng Ban', originalIndex: 2, dataType: 'text' },
          { key: 'salary', label: 'Mức Lương', originalIndex: 3, dataType: 'number' },
          { key: 'status', label: 'Đang làm việc', originalIndex: 4, dataType: 'boolean' },
          { key: 'joinDate', label: 'Ngày vào làm', originalIndex: 5, dataType: 'date' },
        ],
        rows: [
          { id: 'NV001', name: 'Nguyễn Văn An', department: 'Công nghệ', salary: 25000000, status: true, joinDate: '2023-01-15' },
          { id: 'NV002', name: 'Trần Thị Bình', department: 'Kinh doanh', salary: 18000000, status: true, joinDate: '2023-03-20' },
          { id: 'NV003', name: 'Lê Hoàng Cường', department: 'Marketing', salary: 16500000, status: true, joinDate: '2023-05-10' },
          { id: 'NV004', name: 'Phạm Minh Đức', department: 'Công nghệ', salary: 28000000, status: true, joinDate: '2022-11-01' },
          { id: 'NV005', name: 'Vũ Thu Hà', department: 'Nhân sự', salary: 15000000, status: false, joinDate: '2023-08-12' },
          { id: 'NV006', name: 'Đặng Tuấn Kiệt', department: 'Kinh doanh', salary: 20000000, status: true, joinDate: '2024-02-01' },
        ],
        headerRowIndex: 0,
      },
      {
        id: 'sheet_sample_2',
        name: 'Dự Án',
        headers: [
          { key: 'code', label: 'Mã Dự Án', originalIndex: 0, dataType: 'text' },
          { key: 'title', label: 'Tên Dự Án', originalIndex: 1, dataType: 'text' },
          { key: 'budget', label: 'Ngân Sách', originalIndex: 2, dataType: 'number' },
          { key: 'progress', label: 'Tiến độ (%)', originalIndex: 3, dataType: 'number' },
        ],
        rows: [
          { code: 'PRJ-01', title: 'Nâng cấp Website E-commerce', budget: 150000000, progress: 85 },
          { code: 'PRJ-02', title: 'Phát triển Mobile App Flutter', budget: 300000000, progress: 45 },
          { code: 'PRJ-03', title: 'Triển khai Hạ tầng Cloud AWS', budget: 220000000, progress: 100 },
        ],
        headerRowIndex: 0,
      },
    ],
    activeSheetId: 'sheet_sample_1',
  }

  void flushCurrentSession().then(() => {
    adoptSessionForFile(sampleWb.meta.fileName)
    workbookStore.setWorkbook(sampleWb)
    appConfigStore.addToast({
      type: 'success',
      message: 'Đã nạp dữ liệu mẫu thành công!',
    })
  })
}

// ─── Keyboard Shortcuts ──────────────────────────────────────────────────────

function handleKeydown(e: KeyboardEvent) {
  if (e.ctrlKey || e.metaKey) {
    if (e.key === 'z' || e.key === 'Z') {
      if (e.shiftKey) {
        if (canRedo.value) {
          e.preventDefault()
          handleHistoryRedo()
        }
      } else {
        if (canUndo.value) {
          e.preventDefault()
          handleHistoryUndo()
        }
      }
    } else if (e.key === 'y' || e.key === 'Y') {
      if (canRedo.value) {
        e.preventDefault()
        handleHistoryRedo()
      }
    } else if (e.key === 'h' || e.key === 'H' || e.key === 'f' || e.key === 'F') {
      // Find & Replace
      if (workbookStore.workbook) {
        e.preventDefault()
        showFindReplace.value = !showFindReplace.value
      }
    }
  }
}

async function handleGoHome() {
  await flushCurrentSession()
  clearAllRowMarks()
  workbookStore.resetWorkbook()
  showIntro.value = true
}

function openWebEditor() {
  showIntro.value = false
}

async function handleSwitchSession(id: string) {
  if (id === activeSessionId.value && workbookStore.workbook) return
  if (workbookStore.workbook) await flushCurrentSession()
  setActiveSessionId(id)
  const wb = await loadSessionById(id)
  if (wb) workbookStore.setWorkbook(wb)
}

async function handleDeleteSession(id: string) {
  const wasActive = activeSessionId.value === id
  if (!wasActive) {
    await deleteFileSession(id)
    return
  }
  debouncedSaveSqlite.cancel()
  await deleteFileSession(id)
  if (activeSessionId.value) {
    const wb = await loadSessionById(activeSessionId.value)
    if (wb) {
      workbookStore.setWorkbook(wb)
      return
    }
  }
  clearAllRowMarks()
  workbookStore.resetWorkbook()
}

function handleHistoryUndo() {
  const entry = undo()
  if (!entry || !workbookStore.workbook) return
  const sheet = workbookStore.workbook.sheets.find((s) => s.id === entry.sheetId)
  if (!sheet) return

  if (entry.type === 'MOVE_ROW' && typeof entry.rowIndex === 'number' && typeof entry.toIndex === 'number') {
    void workbookStore.moveRow(sheet.id, entry.toIndex, entry.rowIndex)
    appConfigStore.addToast({ type: 'info', message: `Đã hoàn tác: ${entry.description}` })
    return
  }

  if (typeof entry.rowIndex === 'number' && entry.before.rows?.[0]) {
    void workbookStore.replaceRow(sheet.id, entry.rowIndex, entry.before.rows[0])
    appConfigStore.addToast({
      type: 'info',
      message: `Đã hoàn tác: ${entry.description}`,
    })
    return
  }

  if (entry.before.rows && !sheet.paged) sheet.rows = JSON.parse(JSON.stringify(entry.before.rows))
  if (entry.before.headers) sheet.headers = JSON.parse(JSON.stringify(entry.before.headers))
  if (entry.before.name) sheet.name = entry.before.name
  if (entry.before.rows && !sheet.paged) syncBodyRows(sheet)
  workbookStore.dataRevision += 1

  appConfigStore.addToast({
    type: 'info',
    message: `Đã hoàn tác: ${entry.description}`,
  })
}

function handleHistoryRedo() {
  const entry = redo()
  if (!entry || !workbookStore.workbook) return
  const sheet = workbookStore.workbook.sheets.find((s) => s.id === entry.sheetId)
  if (!sheet) return

  if (entry.type === 'MOVE_ROW' && typeof entry.rowIndex === 'number' && typeof entry.toIndex === 'number') {
    void workbookStore.moveRow(sheet.id, entry.rowIndex, entry.toIndex)
    appConfigStore.addToast({ type: 'info', message: `Đã làm lại: ${entry.description}` })
    return
  }

  if (typeof entry.rowIndex === 'number' && entry.after.rows?.[0]) {
    void workbookStore.replaceRow(sheet.id, entry.rowIndex, entry.after.rows[0])
    appConfigStore.addToast({
      type: 'info',
      message: `Đã làm lại: ${entry.description}`,
    })
    return
  }

  if (entry.after.rows && !sheet.paged) {
    sheet.rows = JSON.parse(JSON.stringify(entry.after.rows))
    syncBodyRows(sheet)
  }
  if (entry.after.headers) sheet.headers = JSON.parse(JSON.stringify(entry.after.headers))
  if (entry.after.name) sheet.name = entry.after.name
  workbookStore.dataRevision += 1

  appConfigStore.addToast({
    type: 'info',
    message: `Đã làm lại: ${entry.description}`,
  })
}

// Lifecycle
  onMounted(async () => {
    loadSavedTheme()
    window.addEventListener('keydown', handleKeydown)
    window.addEventListener('pagehide', onPageHide)

    try {
      const savedSession = await loadSessionFromSqlite()
      if (savedSession) {
        await workbookStore.setWorkbook(savedSession)
        appConfigStore.addToast({
          type: 'info',
          message: `Đã khôi phục dữ liệu từ SQLite (${savedSession.meta.fileName})`,
        })
      }
    } catch (err) {
      console.error('Không thể tải phiên SQLite:', err)
    }
  })

  onUnmounted(() => {
    window.removeEventListener('keydown', handleKeydown)
    window.removeEventListener('pagehide', onPageHide)
  })
</script>

<template>
  <div class="flex flex-col h-screen w-screen overflow-hidden bg-base-100 text-base-content select-none">
    <!-- Toast Notifications -->
    <ToastContainer />
    <ProgressTaskDock />

    <!-- Navbar -->
    <AppNavbar
      @file-open="promptImportOptions"
      @export-xlsx="handleExportXlsx"
      @export-csv="handleExportCsv"
      @export-sqlite="() => downloadSqliteDatabase()"
      @toggle-find-replace="showFindReplace = !showFindReplace"
      @switch-session="handleSwitchSession"
      @delete-session="handleDeleteSession"
      @go-home="handleGoHome"
      @undo="handleHistoryUndo"
      @redo="handleHistoryRedo"
    />

    <!-- Main Content Area -->
    <main class="flex-1 overflow-hidden relative">
      <IntroPage v-if="showIntro && !workbookStore.workbook" @open-web="openWebEditor" />

      <!-- 1. Landing View khi chưa mở file -->
      <div
        v-else-if="!workbookStore.workbook"
        class="flex flex-col items-center justify-center h-full p-6 text-center"
        @dragover="handleDragOver"
        @dragleave="handleDragLeave"
        @drop="handleDrop"
      >
        <div
          class="max-w-md w-full border-2 border-dashed rounded-2xl p-8 flex flex-col items-center gap-4 transition-all duration-200 cursor-pointer"
          :class="
            isDraggingOver
              ? 'border-primary bg-primary/10 scale-102 shadow-xl'
              : 'border-base-300 hover:border-primary/50 hover:bg-base-200/50'
          "
          @click="triggerFileInput"
        >
          <div class="p-4 bg-primary/10 rounded-full text-primary">
            <FileSpreadsheet class="w-12 h-12" />
          </div>

          <div>
            <h2 class="text-lg font-bold">Kéo thả file Excel vào đây</h2>
            <p class="text-xs opacity-60 mt-1">
              Hỗ trợ các định dạng .xlsx, .xls, .csv với nhiều sheets
            </p>
          </div>

          <button class="btn btn-primary btn-sm gap-2">
            <Upload class="w-4 h-4" /> Chọn tệp từ máy tính
          </button>
        </div>

        <div v-if="sessions.length" class="mt-6 w-full max-w-md text-left" @click.stop>
          <h3 class="text-xs font-semibold uppercase tracking-wide opacity-60 mb-2">Phiên đang có</h3>
          <ul class="flex flex-col gap-1 max-h-64 overflow-y-auto scrollbar-thin">
            <li v-for="session in sessions" :key="session.id">
              <button
                type="button"
                class="btn btn-ghost btn-sm w-full justify-start h-auto py-2 px-3"
                @click="handleSwitchSession(session.id)"
              >
                <span class="min-w-0 text-left">
                  <span class="block text-sm truncate">{{ session.fileName }}</span>
                  <span class="block text-[11px] opacity-50 font-normal">
                    {{ session.rowCount }} dòng · {{ session.sheetCount }} sheet
                  </span>
                </span>
              </button>
            </li>
          </ul>
        </div>

        <!-- Nút nạp dữ liệu mẫu -->
        <div class="mt-6 flex flex-col items-center gap-2">
          <span class="text-xs opacity-50">hoặc thử nghiệm ngay không cần tải file:</span>
          <button class="btn btn-outline btn-xs gap-1.5" @click="loadSampleData">
            <Sparkles class="w-3.5 h-3.5 text-warning" /> Nạp dữ liệu mẫu (Nhân sự & Dự án)
          </button>
        </div>
      </div>

      <!-- 2. Workspace View khi đã có Workbook -->
      <div v-else class="flex flex-col h-full overflow-hidden">
        <!-- Main Area: Editor hoặc Preview toàn màn hình -->
        <div class="flex-1 overflow-hidden relative">
          <EditorTable v-show="appConfigStore.viewMode === 'editor'" />
          <ExcelPreview v-show="appConfigStore.viewMode === 'preview'" />
        </div>

        <!-- Sheet Tabs Bar -->
        <SheetTabsBar />

        <!-- Status Bar -->
        <StatusBar />
      </div>

      <!-- Loading Overlay -->
      <div
        v-if="isLoading"
        class="absolute inset-0 bg-base-100/60 backdrop-blur-xs flex flex-col items-center justify-center z-50 gap-3"
      >
        <span class="loading loading-spinner loading-lg text-primary" />
        <span class="text-sm font-medium">{{ importProgress || 'Đang xử lý file Excel...' }}</span>
      </div>
    </main>

    <!-- Hidden file input -->
    <input
      ref="fileInputRef"
      type="file"
      accept=".xlsx,.xls,.csv"
      class="hidden"
      @change="onFileSelected"
    />

    <!-- Find & Replace Modal -->
    <FindReplaceModal v-model="showFindReplace" />

    <!-- Import Options Modal -->
    <ImportOptionsModal
      v-model="showImportOptions"
      :file-name="pendingFile?.name || ''"
      @confirm="handleImportConfirm"
    />
  </div>
</template>
