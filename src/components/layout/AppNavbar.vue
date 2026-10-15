<script setup lang="ts">
/**
 * AppNavbar.vue
 * Thanh điều hướng chính của ứng dụng Excel Editor.
 * Bao gồm: Logo, thao tác file (Mở/Xuất), chuyển View Mode,
 * nút Undo/Redo, tìm kiếm, và ThemeDropdown.
 */
import { computed, nextTick, onMounted, onUnmounted, ref } from 'vue'
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
  FileEdit,
  FileSpreadsheet,
  Menu,
  ChevronDown,
  Files,
  CircleHelp,
  X,
  HardDrive,
  Cloud,
} from 'lucide-vue-next'
import { useWorkbookStore } from '@/stores/workbookStore'
import { useAppConfigStore } from '@/stores/appConfigStore'
import { useProgressTaskStore, canPauseTask, canRemoveTask, canResumeTask, taskStatusBadgeClass, taskStatusLabel } from '@/stores/progressTaskStore'
import { useHistory } from '@/composables/useHistory'
import { useSqliteStorage } from '@/composables/useSqliteStorage'
import { useCloudBindings } from '@/cloud/bindings'
import { googleConfigured, microsoftConfigured } from '@/cloud/config'
import ThemeDropdown from '@/components/common/ThemeDropdown.vue'
import LanguageDropdown from '@/components/common/LanguageDropdown.vue'
import HelpGuide from '@/components/modals/HelpGuide.vue'
import { useLocale } from '@/composables/useLocale'
import { menuWidthPx } from '@/utils/workspaceFrame'
import { startPageTour } from '@/composables/usePageTour'

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
  /** Mở Google Drive hoặc OneDrive */
  'open-cloud': [provider: 'google' | 'microsoft']
  /** Ghi workbook hiện tại lên file cloud đang gắn */
  'save-cloud': []
}>()

// ─── Stores & Composables ────────────────────────────────────────────────────
const workbookStore = useWorkbookStore()
const appConfigStore = useAppConfigStore()
const progressStore = useProgressTaskStore()
const { canUndo, canRedo, undoDescription, redoDescription } = useHistory()
const { sessions, activeSessionId } = useSqliteStorage()
const { bindingFor, cloudSyncing } = useCloudBindings()
const cloudBinding = computed(() => bindingFor(activeSessionId.value))

// ─── Refs ────────────────────────────────────────────────────────────────────
/** Hidden file input element */
const fileInputRef = ref<HTMLInputElement | null>(null)
const showHelp = ref(false)
const mobileOpen = ref(false)
const { t, locale } = useLocale()
const viewportWidth = ref(window.innerWidth)
const menuWidth = computed(() => `${menuWidthPx(viewportWidth.value, 320)}px`)

function openTour() {
  closeMenus()
  nextTick(() => startPageTour(t))
}

function onViewportResize() {
  viewportWidth.value = window.innerWidth
  if (viewportWidth.value >= 1024) mobileOpen.value = false
  if (viewportWidth.value < 768) appConfigStore.setViewMode('editor')
}

onMounted(() => window.addEventListener('resize', onViewportResize))
onUnmounted(() => window.removeEventListener('resize', onViewportResize))

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
  if (diff < minute) return t('nav.just.now')
  if (diff < hour) return t('nav.n.min.ago', { n: Math.floor(diff / minute) })
  if (diff < day) return t('nav.n.hr.ago', { n: Math.floor(diff / hour) })
  return new Date(ts).toLocaleDateString(locale.value === 'en' ? 'en-US' : 'vi-VN')
}

function closeMenus() {
  ;(document.activeElement as HTMLElement | null)?.blur()
  mobileOpen.value = false
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
  <nav class="border-b border-base-300 bg-base-200">
    <div class="flex h-12 items-center gap-2 px-2 lg:hidden">
      <button
        type="button"
        class="btn btn-ghost btn-sm h-9 min-h-9 shrink-0 px-2"
        :title="t('nav.home')"
        @click="emit('go-home'); mobileOpen = false"
      >
        <FileSpreadsheet class="h-4 w-4 text-primary" />
      </button>
      <div
        v-if="workbookStore.workbook && viewportWidth < 768"
        class="flex h-9 min-w-0 flex-1 rounded-full bg-base-300/80 p-0.5"
        role="tablist"
        data-tour="screen"
      >
        <button
          type="button"
          role="tab"
          class="flex h-8 flex-1 items-center justify-center gap-1 rounded-full text-xs font-medium motion-safe:active:scale-[0.98]"
          :class="appConfigStore.phoneScreen === 'table' ? 'bg-primary text-primary-content' : 'text-base-content/70'"
          :aria-selected="appConfigStore.phoneScreen === 'table'"
          @click="appConfigStore.setPhoneScreen('table')"
        >
          <Table2 class="h-3.5 w-3.5" /> {{ t('nav.screen.table') }}
        </button>
        <button
          type="button"
          role="tab"
          class="flex h-8 flex-1 items-center justify-center gap-1 rounded-full text-xs font-medium motion-safe:active:scale-[0.98]"
          :class="appConfigStore.phoneScreen === 'form' ? 'bg-primary text-primary-content' : 'text-base-content/70'"
          :aria-selected="appConfigStore.phoneScreen === 'form'"
          @click="appConfigStore.setPhoneScreen('form')"
        >
          <FileEdit class="h-3.5 w-3.5" /> {{ t('nav.screen.form') }}
        </button>
      </div>
      <span
        v-else-if="workbookStore.workbook"
        class="min-w-0 flex-1 truncate text-xs opacity-70"
        :title="workbookStore.workbook.meta.fileName"
      >
        <span v-if="workbookStore.isDirty" class="mr-1 font-bold text-error">*</span>
        {{ workbookStore.workbook.meta.fileName }}
      </span>
      <span v-else class="flex-1" />
      <div v-if="!(workbookStore.workbook && viewportWidth < 768)" class="join shrink-0" data-tour="screen">
        <button
          type="button"
          class="btn join-item btn-sm h-11 min-h-11 px-3"
          :class="appConfigStore.viewMode === 'editor' ? 'btn-primary' : 'btn-ghost'"
          :aria-label="t('nav.edit')"
          @click="appConfigStore.setViewMode('editor')"
        >
          <Table2 class="h-4 w-4" />
        </button>
        <button
          type="button"
          class="btn join-item btn-sm h-11 min-h-11 px-3"
          :class="appConfigStore.viewMode === 'preview' ? 'btn-primary' : 'btn-ghost'"
          :aria-label="t('nav.excel.preview')"
          @click="appConfigStore.setViewMode('preview')"
        >
          <Eye class="h-4 w-4" />
        </button>
      </div>
      <button
        type="button"
        class="btn btn-ghost btn-sm h-9 min-h-9 shrink-0 px-2"
        data-tour="menu"
        :aria-expanded="mobileOpen"
        :aria-label="t('nav.menu')"
        @click="mobileOpen = !mobileOpen"
      >
        <X v-if="mobileOpen" class="h-5 w-5" />
        <Menu v-else class="h-5 w-5" />
      </button>
    </div>

    <div
      v-if="mobileOpen"
      class="fixed inset-0 z-[80] flex flex-col bg-base-100 lg:hidden"
      role="dialog"
      aria-modal="true"
      :aria-label="t('nav.menu')"
    >
      <div class="flex h-14 shrink-0 items-center justify-between border-b border-base-300 px-4">
        <span class="text-sm font-semibold">{{ t('nav.menu') }}</span>
        <button
          type="button"
          class="btn btn-ghost btn-sm h-11 min-h-11 px-3"
          :aria-label="t('common.close')"
          @click="mobileOpen = false"
        >
          <X class="h-5 w-5" />
        </button>
      </div>

      <div class="flex-1 overflow-y-auto px-4 py-4">
        <p class="mb-2 text-xs font-medium text-base-content/55">{{ t('nav.open') }}</p>
        <div class="overflow-hidden rounded-xl border border-base-300">
          <button type="button" class="flex h-12 w-full items-center gap-3 px-3 text-left text-sm hover:bg-base-200" @click="openFilePicker(); closeMenus()">
            <HardDrive class="h-4 w-4 shrink-0 opacity-70" /> {{ t('nav.computer') }}
          </button>
          <button v-if="googleConfigured()" type="button" class="flex h-12 w-full items-center gap-3 border-t border-base-300 px-3 text-left text-sm hover:bg-base-200" @click="emit('open-cloud', 'google'); closeMenus()">
            <Cloud class="h-4 w-4 shrink-0 opacity-70" /> Google Drive
          </button>
          <button v-if="microsoftConfigured()" type="button" class="flex h-12 w-full items-center gap-3 border-t border-base-300 px-3 text-left text-sm hover:bg-base-200" @click="emit('open-cloud', 'microsoft'); closeMenus()">
            <FileSpreadsheet class="h-4 w-4 shrink-0 opacity-70" /> OneDrive
          </button>
        </div>

        <p class="mb-2 mt-5 text-xs font-medium text-base-content/55">{{ t('nav.export') }}</p>
        <div class="grid grid-cols-3 gap-2">
          <button type="button" class="btn btn-outline h-12 min-h-12 text-xs" :disabled="!workbookStore.workbook" @click="emit('export-xlsx'); closeMenus()">
            <Download class="h-3.5 w-3.5" /> .xlsx
          </button>
          <button type="button" class="btn btn-outline h-12 min-h-12 text-xs" :disabled="!workbookStore.workbook" @click="emit('export-csv'); closeMenus()">
            <Download class="h-3.5 w-3.5" /> .csv
          </button>
          <button type="button" class="btn btn-outline h-12 min-h-12 text-xs" :disabled="!workbookStore.workbook" @click="emit('export-sqlite'); closeMenus()">
            <Download class="h-3.5 w-3.5" /> .sqlite
          </button>
        </div>

        <p class="mb-2 mt-5 text-xs font-medium text-base-content/55">{{ t('nav.sessions') }}</p>
        <p v-if="sessions.length === 0" class="rounded-xl border border-dashed border-base-300 px-3 py-4 text-xs opacity-60">
          {{ t('nav.no.sessions.yet.open.file') }}
        </p>
        <div v-else class="overflow-hidden rounded-xl border border-base-300">
          <div v-for="(session, index) in sessions" :key="session.id" class="flex items-center" :class="index > 0 ? 'border-t border-base-300' : ''">
            <button
              type="button"
              class="flex h-12 min-w-0 flex-1 items-center px-3 text-left text-sm hover:bg-base-200"
              @click="emit('switch-session', session.id); closeMenus()"
            >
              <span class="truncate">{{ session.fileName }}</span>
            </button>
            <button
              type="button"
              class="btn btn-ghost h-12 min-h-12 px-3"
              :aria-label="t('nav.close.session')"
              @click="emit('delete-session', session.id)"
            >
              <X class="h-4 w-4" />
            </button>
          </div>
        </div>

        <div class="mt-5 grid grid-cols-2 gap-2">
          <button type="button" class="btn btn-ghost h-12 min-h-12 justify-start gap-2 border border-base-300 px-3" :disabled="!canUndo" @click="emit('undo'); closeMenus()">
            <Undo2 class="h-4 w-4 shrink-0" />
            <span class="truncate">{{ canUndo ? t('nav.undo.desc', { desc: undoDescription }) : t('nav.nothing.undo') }}</span>
          </button>
          <button type="button" class="btn btn-ghost h-12 min-h-12 justify-start gap-2 border border-base-300 px-3" :disabled="!canRedo" @click="emit('redo'); closeMenus()">
            <Redo2 class="h-4 w-4 shrink-0" />
            <span class="truncate">{{ canRedo ? t('nav.redo.desc', { desc: redoDescription }) : t('nav.nothing.redo') }}</span>
          </button>
        </div>
        <button type="button" class="btn btn-ghost mt-2 h-12 min-h-12 w-full justify-start gap-2 border border-base-300" @click="emit('toggle-find-replace'); closeMenus()">
          <Search class="h-4 w-4" /> {{ t('nav.find.replace.ctrl.h') }}
        </button>
        <button
          v-if="cloudBinding"
          type="button"
          class="btn btn-ghost mt-2 h-12 min-h-12 w-full justify-start gap-2 border border-base-300"
          :disabled="cloudSyncing"
          @click="emit('save-cloud'); closeMenus()"
        >
          <Cloud class="h-4 w-4" /> {{ t('nav.sync') }}
        </button>
        <button type="button" class="btn btn-ghost mt-2 h-12 min-h-12 w-full justify-start gap-2 border border-base-300" @click="progressStore.openManager(); closeMenus()">
          <List class="h-4 w-4" /> {{ t('nav.tasks') }}
          <span v-if="progressStore.tasks.length" class="badge badge-sm badge-primary">{{ progressStore.tasks.length }}</span>
        </button>
        <button type="button" class="btn btn-ghost mt-2 h-12 min-h-12 w-full justify-start gap-2 border border-base-300" @click="openTour()">
          <CircleHelp class="h-4 w-4" /> {{ t('nav.help.f1') }}
        </button>

        <div class="mt-5 flex items-center gap-2 border-t border-base-300 pt-4">
          <LanguageDropdown />
          <ThemeDropdown />
        </div>
      </div>
    </div>

    <div class="navbar hidden h-12 min-h-12 items-center gap-2 overflow-x-auto px-4 lg:flex lg:overflow-x-visible">

    <!-- ── Left: Logo & File Actions ─────────────────────────────────────── -->
    <div class="flex items-center gap-2 flex-none">

      <button
        type="button"
        class="btn btn-ghost btn-xs gap-1.5 px-1.5"
        :title="t('nav.home')"
        @click="emit('go-home')"
      >
        <FileSpreadsheet class="w-4 h-4 text-primary" />
        <span class="font-semibold text-sm hidden md:inline">Excel Editor</span>
      </button>

      <div class="dropdown dropdown-bottom">
        <button
          tabindex="0"
          class="btn btn-ghost btn-xs gap-1"
          data-tour="open"
          :title="t('nav.open.file.computer.cloud')"
        >
          <Upload class="w-3.5 h-3.5" />
          <span class="hidden md:inline">{{ t('nav.open') }}</span>
          <ChevronDown class="w-3 h-3" />
        </button>
        <ul
          tabindex="0"
          class="dropdown-content menu bg-base-100 rounded-box z-50 mt-1 w-52 p-1 shadow-lg border border-base-300"
        >
          <li>
            <button class="text-sm gap-2" @click="openFilePicker(); closeMenus()">
              <HardDrive class="w-3.5 h-3.5" />
              {{ t('nav.computer') }}
            </button>
          </li>
          <li v-if="googleConfigured()">
            <button class="text-sm gap-2" @click="emit('open-cloud', 'google'); closeMenus()">
              <Cloud class="w-3.5 h-3.5" />
              Google Drive
            </button>
          </li>
          <li v-if="microsoftConfigured()">
            <button class="text-sm gap-2" @click="emit('open-cloud', 'microsoft'); closeMenus()">
              <FileSpreadsheet class="w-3.5 h-3.5" />
              OneDrive
            </button>
          </li>
        </ul>
      </div>

      <!-- Menu phiên các file đã mở -->
      <div class="dropdown dropdown-bottom">
        <button
          tabindex="0"
          class="btn btn-ghost btn-xs gap-1"
          :title="t('nav.files.kept.session')"
        >
          <Files class="w-3.5 h-3.5" />
          <span class="hidden md:inline">{{ t('nav.sessions') }}</span>
          <span v-if="sessions.length" class="badge badge-xs badge-primary">{{ sessions.length }}</span>
          <ChevronDown class="w-3 h-3" />
        </button>

        <div
          tabindex="0"
          class="dropdown-content z-50 mt-1 rounded-box border border-base-300 bg-base-100 p-2 shadow-lg"
          :style="{ width: menuWidth }"
        >
          <div class="mb-1 flex items-center justify-between px-1">
            <span class="text-xs font-semibold">{{ t('nav.open.sessions') }}</span>
            <span class="text-[10px] opacity-50">{{ t('nav.n.files', { n: sessions.length }) }}</span>
          </div>

          <p v-if="sessions.length === 0" class="px-2 py-3 text-xs opacity-60">
            {{ t('nav.no.sessions.yet.open.file') }}
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
                        {{ t('nav.session.current') }}
                      </span>
                    </span>
                    <span class="block truncate text-[10px] opacity-50">
                      {{ t('nav.rows.rows.sheets.sheets.time', { rows: session.rowCount, sheets: session.sheetCount, time: formatSessionTime(session.updatedAt) }) }}
                    </span>
                  </span>
                </button>
                <button
                  type="button"
                  class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0 text-base-content/50 hover:text-error"
                  :title="t('nav.close.session')"
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
          data-tour="export"
          :title="t('nav.export.file')"
        >
          <Download class="w-3.5 h-3.5" />
          <span class="hidden md:inline">{{ t('nav.export') }}</span>
          <ChevronDown class="w-3 h-3" />
        </button>

        <ul
          tabindex="0"
          class="dropdown-content menu bg-base-100 rounded-box z-50 w-36 p-1 shadow-lg border border-base-300"
        >
          <li>
            <button class="text-sm" @click="emit('export-xlsx')">
              {{ t('nav.export.xlsx') }}
            </button>
          </li>
          <li>
            <button class="text-sm" @click="emit('export-csv')">
              {{ t('nav.export.csv') }}
            </button>
          </li>
          <div class="divider my-0.5" />
          <li>
            <button class="text-sm text-primary font-medium" @click="emit('export-sqlite')">
              {{ t('nav.export.sqlite') }}
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
          :title="t('nav.unsaved.changes')"
        >*</span>

        <span
          class="hidden text-sm opacity-70 max-w-[140px] truncate md:inline"
          :title="workbookStore.workbook.meta.fileName"
        >
          {{ workbookStore.workbook.meta.fileName }}
        </span>

        <!-- Badge SQLite Auto-Saved -->
        <span
          class="badge badge-xs badge-success badge-outline opacity-80 text-[10px] hidden xl:inline-flex py-1 px-1.5 font-mono"
          :title="t('nav.session.saved.automatically.sqlite.kept')"
        >
          SQLite Saved
        </span>
        <span
          v-if="cloudBinding"
          class="badge badge-xs badge-info badge-outline text-[10px] hidden lg:inline-flex"
          :title="cloudBinding.mode === 'live' ? t('nav.linked.live.cell.read.write') : t('nav.linked.cloud.file')"
        >
          {{ cloudBinding.provider === 'google' ? 'Google' : 'OneDrive' }}
        </span>
      </div>
    </div>

    <!-- ── Center: View Mode (Editor vs Preview) ───────────────────────── -->
    <div class="flex flex-1 justify-center max-lg:flex-none">
      <div class="join bg-base-300/40 p-0.5 rounded-lg border border-base-300" data-tour="screen">
        <!-- Editor View -->
        <button
          class="join-item btn btn-xs px-3 gap-1.5"
          :class="appConfigStore.viewMode === 'editor' ? 'btn-primary font-bold shadow-xs' : 'btn-ghost'"
          :title="t('nav.edit.data.form')"
          @click="appConfigStore.setViewMode('editor')"
        >
          <Table2 class="w-3.5 h-3.5" />
          <span class="hidden md:inline">{{ t('nav.edit') }}</span>
        </button>

        <!-- Preview View -->
        <button
          class="join-item btn btn-xs px-3 gap-1.5"
          :class="appConfigStore.viewMode === 'preview' ? 'btn-primary font-bold shadow-xs' : 'btn-ghost'"
          :title="t('nav.excel.spreadsheet.preview')"
          @click="appConfigStore.setViewMode('preview')"
        >
          <Eye class="w-3.5 h-3.5" />
          <span class="hidden md:inline">{{ t('nav.excel.preview') }}</span>
        </button>
      </div>
    </div>

    <!-- ── Right: Undo/Redo, Search, Theme ───────────────────────────────── -->
    <div class="flex items-center gap-1 flex-none">

      <!-- Nút Undo -->
      <button
        class="btn btn-ghost btn-xs"
        :disabled="!canUndo"
        data-tour="undo"
        :title="canUndo ? t('nav.undo.desc', { desc: undoDescription }) : t('nav.nothing.undo')"
        @click="emit('undo')"
      >
        <Undo2 class="w-3.5 h-3.5" />
      </button>

      <!-- Nút Redo -->
      <button
        class="btn btn-ghost btn-xs"
        :disabled="!canRedo"
        :title="canRedo ? t('nav.redo.desc', { desc: redoDescription }) : t('nav.nothing.redo')"
        @click="emit('redo')"
      >
        <Redo2 class="w-3.5 h-3.5" />
      </button>

      <!-- Divider -->
      <div class="w-px h-5 bg-base-300 mx-1" />

      <button
        v-if="cloudBinding"
        class="btn btn-ghost btn-xs gap-1"
        :disabled="cloudSyncing"
        :title="t('nav.sync.cloud')"
        @click="emit('save-cloud')"
      >
        <span v-if="cloudSyncing" class="loading loading-spinner loading-xs" />
        <Cloud v-else class="w-3.5 h-3.5" />
        <span class="hidden md:inline">{{ t('nav.sync') }}</span>
      </button>

      <!-- Nút Find & Replace -->
      <button
        class="btn btn-ghost btn-xs"
        data-tour="find"
        :title="t('nav.find.replace.ctrl.h')"
        @click="emit('toggle-find-replace')"
      >
        <Search class="w-3.5 h-3.5" />
      </button>

      <div class="dropdown dropdown-bottom dropdown-end">
        <button
          tabindex="0"
          class="btn btn-ghost btn-xs gap-1"
          :title="t('nav.tasks')"
        >
          <List class="w-3.5 h-3.5" />
          <span class="hidden md:inline">{{ t('nav.tasks.menu') }}</span>
          <span v-if="progressStore.tasks.length" class="badge badge-xs badge-primary">
            {{ progressStore.tasks.length }}
          </span>
          <ChevronDown class="w-3 h-3" />
        </button>
        <div
          tabindex="0"
          class="dropdown-content z-50 mt-1 rounded-box border border-base-300 bg-base-100 p-2 shadow-lg"
          :style="{ width: menuWidth }"
        >
          <div class="mb-1 flex items-center justify-between px-1">
            <span class="text-xs font-semibold">{{ t('nav.tasks') }}</span>
            <button
              type="button"
              class="btn btn-ghost btn-xs h-6 min-h-0"
              @click="progressStore.openManager(); closeMenus()"
            >
              {{ t('nav.manage') }}
            </button>
          </div>
          <p v-if="progressStore.tasks.length === 0" class="px-2 py-3 text-xs opacity-60">
            {{ t('nav.no.tasks') }}
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
                      {{ t(taskStatusLabel(task.status)) }}
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
                  <Pause class="h-3 w-3" /> {{ t('nav.pause') }}
                </button>
                <button
                  v-if="canResumeTask(task)"
                  type="button"
                  class="btn btn-ghost btn-xs h-6 min-h-0 gap-1"
                  @click="progressStore.resume(task.id)"
                >
                  <Play class="h-3 w-3" /> {{ t('nav.resume') }}
                </button>
                <button
                  v-if="canRemoveTask(task)"
                  type="button"
                  class="btn btn-ghost btn-xs h-6 min-h-0 gap-1 text-error"
                  @click="progressStore.remove(task.id)"
                >
                  <Trash2 class="h-3 w-3" /> {{ t('common.remove') }}
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
        :title="t('nav.help.f1')"
        @click="openTour"
      >
        <CircleHelp class="w-3.5 h-3.5" />
      </button>

      <!-- Theme Dropdown -->
      <!-- Tách thành component riêng để dễ mở rộng -->
      <LanguageDropdown />
      <ThemeDropdown />
    </div>
    </div>
  </nav>
  <HelpGuide v-model="showHelp" @tour="openTour" />
</template>
