<script setup lang="ts">
/**
 * AppNavbar.vue
 * Thanh điều hướng chính của ứng dụng Excel Editor.
 * Bao gồm: Logo, thao tác file (Mở/Xuất), chuyển View Mode,
 * nút Undo/Redo, tìm kiếm, và ThemeDropdown.
 */
import { ref } from 'vue'
import {
  Upload,
  Download,
  Undo2,
  Redo2,
  Search,
  List,
  Pause,
  Play,
  Trash2,
  Table2,
  Eye,
  FileSpreadsheet,
  ChevronDown,
  Files,
  CircleHelp,
  X,
} from 'lucide-vue-next'
import { useWorkbookStore } from '@/stores/workbookStore'
import { useAppConfigStore } from '@/stores/appConfigStore'
import { useProgressTaskStore, canPauseTask, canRemoveTask, canResumeTask, taskStatusBadgeClass, taskStatusLabel } from '@/stores/progressTaskStore'
import { useHistory } from '@/composables/useHistory'
import { useSqliteStorage } from '@/composables/useSqliteStorage'
import ThemeDropdown from '@/components/common/ThemeDropdown.vue'
import HelpGuide from '@/components/modals/HelpGuide.vue'

// ─── Emits ──────────────────────────────────────────────────────────────────
const emit = defineEmits<{
  /** Người dùng đã chọn file để mở */
  'file-open': [file: File]
  /** Yêu cầu xuất file .xlsx */
  'export-xlsx': []
  /** Yêu cầu xuất file .csv */
  'export-csv': []
  /** Yêu cầu xuất file .sqlite */
  'export-sqlite': []
  /** Bật/tắt modal Find & Replace */
  'toggle-find-replace': []
  /** Mở một phiên file đã lưu */
  'switch-session': [id: string]
  /** Đóng một phiên file */
  'delete-session': [id: string]
  /** Về màn hình trang chủ, giữ các phiên đã lưu */
  'go-home': []
  /** Hoàn tác thao tác gần nhất */
  'undo': []
  /** Làm lại thao tác vừa hoàn tác */
  'redo': []
}>()

// ─── Stores & Composables ────────────────────────────────────────────────────
const workbookStore = useWorkbookStore()
const appConfigStore = useAppConfigStore()
const progressStore = useProgressTaskStore()
const { canUndo, canRedo, undoDescription, redoDescription } = useHistory()
const { sessions, activeSessionId } = useSqliteStorage()

// ─── Refs ────────────────────────────────────────────────────────────────────
/** Hidden file input element */
const fileInputRef = ref<HTMLInputElement | null>(null)
const showHelp = ref(false)

// ─── View Mode Options ───────────────────────────────────────────────────────
interface ViewModeOption {
  key: 'split' | 'editor' | 'preview'
  label: string
  title: string
}

const viewModeOptions: ViewModeOption[] = [
  { key: 'split',   label: 'Split',   title: 'Chia đôi màn hình (Split View)' },
  { key: 'editor',  label: 'Editor',  title: 'Chỉ hiển thị bảng chỉnh sửa' },
  { key: 'preview', label: 'Preview', title: 'Chỉ hiển thị xem trước' },
]

function formatSessionTime(ts: number): string {
  const diff = Date.now() - ts
  const minute = 60_000
  const hour = 60 * minute
  const day = 24 * hour
  if (diff < minute) return 'vừa xong'
  if (diff < hour) return `${Math.floor(diff / minute)} phút trước`
  if (diff < day) return `${Math.floor(diff / hour)} giờ trước`
  return new Date(ts).toLocaleDateString('vi-VN')
}

function closeMenus() {
  ;(document.activeElement as HTMLElement | null)?.blur()
}

// ─── Handlers ────────────────────────────────────────────────────────────────

/** Kích hoạt file input ẩn để người dùng chọn file */
function openFilePicker() {
  fileInputRef.value?.click()
}

/** Xử lý khi người dùng chọn file từ input */
function onFileSelected(event: Event) {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (file) {
    emit('file-open', file)
  }
  // Reset value để cho phép chọn lại cùng file
  target.value = ''
}
</script>

<template>
  <nav class="navbar bg-base-200 border-b border-base-300 px-4 min-h-12 h-12 flex items-center gap-2">

    <!-- ── Left: Logo & File Actions ─────────────────────────────────────── -->
    <div class="flex items-center gap-2 flex-none">

      <button
        type="button"
        class="btn btn-ghost btn-xs gap-1.5 px-1.5"
        title="Về trang chủ"
        @click="emit('go-home')"
      >
        <FileSpreadsheet class="w-4 h-4 text-primary" />
        <span class="font-semibold text-sm hidden md:inline">Excel Editor</span>
      </button>

      <!-- Nút Mở File -->
      <button
        class="btn btn-ghost btn-xs gap-1"
        title="Mở file Excel (.xlsx, .xls, .csv)"
        @click="openFilePicker"
      >
        <Upload class="w-3.5 h-3.5" />
        <span class="hidden md:inline">Mở File</span>
      </button>

      <!-- Menu phiên các file đã mở -->
      <div class="dropdown dropdown-bottom">
        <button
          tabindex="0"
          class="btn btn-ghost btn-xs gap-1"
          title="Các file đang giữ trong phiên làm việc"
        >
          <Files class="w-3.5 h-3.5" />
          <span class="hidden md:inline">Phiên</span>
          <span v-if="sessions.length" class="badge badge-xs badge-primary">{{ sessions.length }}</span>
          <ChevronDown class="w-3 h-3" />
        </button>

        <div
          tabindex="0"
          class="dropdown-content z-50 mt-1 w-80 rounded-box border border-base-300 bg-base-100 p-2 shadow-lg"
        >
          <div class="mb-1 flex items-center justify-between px-1">
            <span class="text-xs font-semibold">Phiên đang mở</span>
            <span class="text-[10px] opacity-50">{{ sessions.length }} file</span>
          </div>

          <p v-if="sessions.length === 0" class="px-2 py-3 text-xs opacity-60">
            Chưa có phiên nào. Mở một file để lưu vào đây.
          </p>

          <ul v-else class="max-h-72 space-y-1 overflow-y-auto">
            <li v-for="session in sessions" :key="session.id">
              <div
                class="flex items-center gap-1 rounded-lg px-1.5 py-1"
                :class="session.id === activeSessionId ? 'bg-primary/10' : 'hover:bg-base-200'"
              >
                <button
                  type="button"
                  class="flex min-w-0 flex-1 items-center gap-2 rounded-md px-1 py-0.5 text-left"
                  :title="session.fileName"
                  @click="emit('switch-session', session.id); closeMenus()"
                >
                  <FileSpreadsheet
                    class="h-4 w-4 shrink-0"
                    :class="session.id === activeSessionId ? 'text-primary' : 'opacity-40'"
                  />
                  <span class="min-w-0">
                    <span class="flex items-center gap-1.5">
                      <span class="truncate text-sm">{{ session.fileName }}</span>
                      <span
                        v-if="session.id === activeSessionId"
                        class="shrink-0 rounded bg-primary/15 px-1 py-px text-[9px] font-medium text-primary"
                      >
                        Đang mở
                      </span>
                    </span>
                    <span class="block truncate text-[10px] opacity-50">
                      {{ session.rowCount }} dòng · {{ session.sheetCount }} sheet · {{ formatSessionTime(session.updatedAt) }}
                    </span>
                  </span>
                </button>
                <button
                  type="button"
                  class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0 text-base-content/50 hover:text-error"
                  title="Đóng phiên này"
                  @click.stop="emit('delete-session', session.id)"
                >
                  <X class="h-3 w-3" />
                </button>
              </div>
            </li>
          </ul>
        </div>
      </div>

      <!-- Hidden file input -->
      <input
        ref="fileInputRef"
        type="file"
        accept=".xlsx,.xls,.csv"
        class="hidden"
        @change="onFileSelected"
      />

      <!-- Nút Xuất (Dropdown) -->
      <div class="dropdown dropdown-bottom">
        <button
          tabindex="0"
          class="btn btn-ghost btn-xs gap-1"
          :disabled="!workbookStore.workbook"
          title="Xuất file"
        >
          <Download class="w-3.5 h-3.5" />
          <span class="hidden md:inline">Xuất</span>
          <ChevronDown class="w-3 h-3" />
        </button>

        <ul
          tabindex="0"
          class="dropdown-content menu bg-base-100 rounded-box z-50 w-36 p-1 shadow-lg border border-base-300"
        >
          <li>
            <button class="text-sm" @click="emit('export-xlsx')">
              Xuất .xlsx
            </button>
          </li>
          <li>
            <button class="text-sm" @click="emit('export-csv')">
              Xuất .csv
            </button>
          </li>
          <div class="divider my-0.5" />
          <li>
            <button class="text-sm text-primary font-medium" @click="emit('export-sqlite')">
              Xuất .sqlite
            </button>
          </li>
        </ul>
      </div>

      <!-- Tên file & dirty indicator -->
      <div v-if="workbookStore.workbook" class="flex items-center gap-1.5 ml-1">
        <!-- Dấu * khi có thay đổi chưa lưu -->
        <span
          v-if="workbookStore.isDirty"
          class="text-error font-bold text-sm leading-none"
          title="Có thay đổi chưa được lưu"
        >*</span>

        <span
          class="text-sm opacity-70 max-w-[140px] truncate"
          :title="workbookStore.workbook.meta.fileName"
        >
          {{ workbookStore.workbook.meta.fileName }}
        </span>

        <!-- Badge SQLite Auto-Saved -->
        <span
          class="badge badge-xs badge-success badge-outline opacity-80 text-[10px] hidden xl:inline-flex py-1 px-1.5 font-mono"
          title="Đã lưu phiên làm việc tự động vào SQLite (chống mất khi reload)"
        >
          SQLite Saved
        </span>
      </div>
    </div>

    <!-- ── Center: View Mode (Editor vs Preview) ───────────────────────── -->
    <div class="flex-1 flex justify-center">
      <div class="join bg-base-300/40 p-0.5 rounded-lg border border-base-300">
        <!-- Editor View -->
        <button
          class="join-item btn btn-xs px-3 gap-1.5"
          :class="appConfigStore.viewMode === 'editor' ? 'btn-primary font-bold shadow-xs' : 'btn-ghost'"
          title="Chế độ Chỉnh sửa dữ liệu & Form"
          @click="appConfigStore.setViewMode('editor')"
        >
          <Table2 class="w-3.5 h-3.5" />
          <span>Chỉnh sửa</span>
        </button>

        <!-- Preview View -->
        <button
          class="join-item btn btn-xs px-3 gap-1.5"
          :class="appConfigStore.viewMode === 'preview' ? 'btn-primary font-bold shadow-xs' : 'btn-ghost'"
          title="Chế độ Xem trước Excel (Spreadsheet Preview)"
          @click="appConfigStore.setViewMode('preview')"
        >
          <Eye class="w-3.5 h-3.5" />
          <span>Xem trước Excel</span>
        </button>
      </div>
    </div>

    <!-- ── Right: Undo/Redo, Search, Theme ───────────────────────────────── -->
    <div class="flex items-center gap-1 flex-none">

      <!-- Nút Undo -->
      <button
        class="btn btn-ghost btn-xs"
        :disabled="!canUndo"
        :title="canUndo ? `Hoàn tác: ${undoDescription}` : 'Không có thao tác để hoàn tác'"
        @click="emit('undo')"
      >
        <Undo2 class="w-3.5 h-3.5" />
      </button>

      <!-- Nút Redo -->
      <button
        class="btn btn-ghost btn-xs"
        :disabled="!canRedo"
        :title="canRedo ? `Làm lại: ${redoDescription}` : 'Không có thao tác để làm lại'"
        @click="emit('redo')"
      >
        <Redo2 class="w-3.5 h-3.5" />
      </button>

      <!-- Divider -->
      <div class="w-px h-5 bg-base-300 mx-1" />

      <!-- Nút Find & Replace -->
      <button
        class="btn btn-ghost btn-xs"
        title="Tìm & Thay thế (Ctrl+H)"
        @click="emit('toggle-find-replace')"
      >
        <Search class="w-3.5 h-3.5" />
      </button>

      <div class="dropdown dropdown-bottom dropdown-end">
        <button
          tabindex="0"
          class="btn btn-ghost btn-xs gap-1"
          title="Danh sách tiến trình"
          @click="progressStore.showDock()"
        >
          <List class="w-3.5 h-3.5" />
          <span class="hidden md:inline">Tiến trình</span>
          <span v-if="progressStore.tasks.length" class="badge badge-xs badge-primary">
            {{ progressStore.tasks.length }}
          </span>
          <ChevronDown class="w-3 h-3" />
        </button>
        <div
          tabindex="0"
          class="dropdown-content z-50 mt-1 w-80 rounded-box border border-base-300 bg-base-100 p-2 shadow-lg"
        >
          <div class="mb-1 flex items-center justify-between px-1">
            <span class="text-xs font-semibold">Danh sách tiến trình</span>
            <button
              type="button"
              class="btn btn-ghost btn-xs h-6 min-h-0"
              @click="progressStore.openManager(); closeMenus()"
            >
              Quản lý
            </button>
          </div>
          <p v-if="progressStore.tasks.length === 0" class="px-2 py-3 text-xs opacity-60">
            Không có tiến trình.
          </p>
          <ul v-else class="max-h-72 space-y-1 overflow-y-auto">
            <li
              v-for="task in progressStore.tasks"
              :key="task.id"
              class="rounded-lg px-2 py-1.5 hover:bg-base-200"
            >
              <div class="flex items-start justify-between gap-2">
                <div class="min-w-0">
                  <div class="flex items-center gap-1">
                    <p class="min-w-0 truncate text-sm">{{ task.title }}</p>
                    <span class="badge badge-xs shrink-0" :class="taskStatusBadgeClass(task.status)">
                      {{ taskStatusLabel(task.status) }}
                    </span>
                  </div>
                  <p class="mt-0.5 line-clamp-2 text-xs leading-snug opacity-70 break-words">{{ task.detail }}</p>
                </div>
                <span class="shrink-0 font-mono text-[11px]">{{ task.percent }}%</span>
              </div>
              <div class="mt-1 flex justify-end gap-1">
                <button
                  v-if="canPauseTask(task)"
                  type="button"
                  class="btn btn-ghost btn-xs h-6 min-h-0 gap-1"
                  @click="progressStore.pause(task.id)"
                >
                  <Pause class="h-3 w-3" /> Dừng
                </button>
                <button
                  v-if="canResumeTask(task)"
                  type="button"
                  class="btn btn-ghost btn-xs h-6 min-h-0 gap-1"
                  @click="progressStore.resume(task.id)"
                >
                  <Play class="h-3 w-3" /> Tiếp tục
                </button>
                <button
                  v-if="canRemoveTask(task)"
                  type="button"
                  class="btn btn-ghost btn-xs h-6 min-h-0 gap-1 text-error"
                  @click="progressStore.remove(task.id)"
                >
                  <Trash2 class="h-3 w-3" /> Xóa
                </button>
              </div>
            </li>
          </ul>
        </div>
      </div>

      <!-- Divider -->
      <div class="w-px h-5 bg-base-300 mx-1" />

      <button
        type="button"
        class="btn btn-ghost btn-xs"
        title="Hướng dẫn (F1)"
        @click="showHelp = true"
      >
        <CircleHelp class="w-3.5 h-3.5" />
      </button>

      <!-- Theme Dropdown -->
      <!-- Tách thành component riêng để dễ mở rộng -->
      <ThemeDropdown />
    </div>
  </nav>
  <HelpGuide v-model="showHelp" />
</template>
