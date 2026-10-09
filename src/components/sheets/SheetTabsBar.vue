<script setup lang="ts">
import { ref, nextTick, onMounted, onUnmounted } from 'vue'
import { Plus, MoreVertical, Copy, Trash2, Edit2, Table } from 'lucide-vue-next'
import { useWorkbookStore } from '@/stores/workbookStore'
import { useHistory } from '@/composables/useHistory'
import { useAppConfigStore } from '@/stores/appConfigStore'
import type { SheetData } from '@/types/excel'
import { ensureSourceGrid, headerChoiceIndexes, previewGridLine } from '@/utils/sheetGrid'

const workbookStore = useWorkbookStore()
const appConfigStore = useAppConfigStore()
const { push: pushHistory } = useHistory()

const editingSheetId = ref<string | null>(null)
const editingName = ref('')
const renameInputRef = ref<HTMLInputElement | null>(null)
const openMenuId = ref<string | null>(null)
const menuStyle = ref<Record<string, string>>({})

function headerChoices(sheet: SheetData) {
  const grid = ensureSourceGrid(sheet)
  return headerChoiceIndexes(grid.length, sheet.headerRowIndex).map((index) => ({
    index,
    label: `Dòng ${index + 1} — ${previewGridLine(grid, index)}`,
  }))
}

function toggleSheetMenu(sheetId: string, event: MouseEvent) {
  event.stopPropagation()
  if (openMenuId.value === sheetId) {
    openMenuId.value = null
    return
  }
  const button = event.currentTarget as HTMLElement
  const rect = button.getBoundingClientRect()
  const menuWidth = 260
  const menuHeight = 320
  const left = Math.max(8, Math.min(rect.right - menuWidth, window.innerWidth - menuWidth - 8))
  const openUp = rect.top > menuHeight || rect.bottom + menuHeight > window.innerHeight
  menuStyle.value = openUp
    ? { top: `${rect.top - 6}px`, left: `${left}px`, transform: 'translateY(-100%)' }
    : { top: `${rect.bottom + 6}px`, left: `${left}px` }
  openMenuId.value = sheetId
}

function closeSheetMenu() {
  openMenuId.value = null
}

function onDocumentClick() {
  closeSheetMenu()
}

onMounted(() => window.addEventListener('click', onDocumentClick))
onUnmounted(() => window.removeEventListener('click', onDocumentClick))

function onPickHeader(sheet: SheetData, headerRowIndex: number) {
  workbookStore.setHeaderRow(sheet.id, headerRowIndex)
  appConfigStore.addToast({
    type: 'info',
    message: `Sheet "${sheet.name}" dùng dòng ${headerRowIndex + 1} làm header`,
  })
}

function selectSheet(sheetId: string) {
  workbookStore.setActiveSheet(sheetId)
}

function startRename(sheet: SheetData) {
  editingSheetId.value = sheet.id
  editingName.value = sheet.name
  nextTick(() => {
    renameInputRef.value?.focus()
    renameInputRef.value?.select()
  })
}

function saveRename(sheetId: string) {
  if (!editingSheetId.value) return
  const trimmed = editingName.value.trim()
  if (trimmed && workbookStore.workbook) {
    const sheet = workbookStore.workbook.sheets.find((s) => s.id === sheetId)
    if (sheet && sheet.name !== trimmed) {
      const oldName = sheet.name
      workbookStore.renameSheet(sheetId, trimmed)
      pushHistory({
        type: 'RENAME_SHEET',
        description: `Đổi tên sheet "${oldName}" thành "${trimmed}"`,
        sheetId,
        before: { name: oldName },
        after: { name: trimmed },
      })
      appConfigStore.addToast({
        type: 'info',
        message: `Đã đổi tên sheet thành "${trimmed}"`,
      })
    }
  }
  editingSheetId.value = null
}

function cancelRename() {
  editingSheetId.value = null
}

function onAddNewSheet() {
  if (!workbookStore.workbook) return
  const nextNum = workbookStore.workbook.sheets.length + 1
  const newSheetName = `Sheet ${nextNum}`
  const newSheetId = `sheet_${Date.now()}`

  const newSheet: SheetData = {
    id: newSheetId,
    name: newSheetName,
    headers: [
      { key: 'col_0', label: 'Cột 1', originalIndex: 0, dataType: 'text' },
      { key: 'col_1', label: 'Cột 2', originalIndex: 1, dataType: 'text' },
      { key: 'col_2', label: 'Cột 3', originalIndex: 2, dataType: 'text' },
    ],
    rows: [
      { col_0: '', col_1: '', col_2: '' },
    ],
    headerRowIndex: 0,
  }

  workbookStore.addSheet(newSheet)
  workbookStore.setActiveSheet(newSheetId)

  pushHistory({
    type: 'ADD_SHEET',
    description: `Thêm sheet mới "${newSheetName}"`,
    sheetId: newSheetId,
    before: {},
    after: newSheet,
  })

  appConfigStore.addToast({
    type: 'success',
    message: `Đã tạo sheet mới "${newSheetName}"`,
  })
}

function onDuplicateSheet(sheet: SheetData) {
  if (!workbookStore.workbook) return
  const newSheetId = `sheet_clone_${Date.now()}`
  const clonedName = `${sheet.name} (Copy)`

  const clonedSheet: SheetData = {
    id: newSheetId,
    name: clonedName,
    headers: sheet.headers.map((h) => ({ ...h })),
    rows: sheet.rows.map((r) => ({ ...r })),
    headerRowIndex: sheet.headerRowIndex,
    sourceGrid: sheet.sourceGrid?.map((line) => [...line]),
  }

  workbookStore.addSheet(clonedSheet)
  workbookStore.setActiveSheet(newSheetId)

  pushHistory({
    type: 'ADD_SHEET',
    description: `Nhân bản sheet "${sheet.name}"`,
    sheetId: newSheetId,
    before: {},
    after: clonedSheet,
  })

  appConfigStore.addToast({
    type: 'success',
    message: `Đã nhân bản sheet "${sheet.name}"`,
  })
}

function onDeleteSheet(sheet: SheetData) {
  if (!workbookStore.workbook || workbookStore.workbook.sheets.length <= 1) {
    appConfigStore.addToast({
      type: 'warning',
      message: 'Không thể xóa sheet duy nhất còn lại.',
    })
    return
  }

  const sheetBackup = { ...sheet }
  workbookStore.deleteSheet(sheet.id)

  pushHistory({
    type: 'DELETE_SHEET',
    description: `Xóa sheet "${sheet.name}"`,
    sheetId: sheet.id,
    before: sheetBackup,
    after: {},
  })

  appConfigStore.addToast({
    type: 'info',
    message: `Đã xóa sheet "${sheet.name}"`,
  })
}
</script>

<template>
  <div
    v-if="workbookStore.workbook"
    class="flex items-center bg-base-200 border-t border-base-300 h-9 px-2 overflow-x-auto scrollbar-thin select-none"
  >
    <!-- Add Sheet Button -->
    <button
      class="btn btn-ghost btn-xs btn-square mr-2 hover:bg-base-300 flex-shrink-0"
      title="Thêm Sheet mới"
      @click="onAddNewSheet"
    >
      <Plus class="w-4 h-4 text-primary" />
    </button>

    <!-- Sheets List -->
    <div class="flex items-center gap-1">
      <div
        v-for="sheet in workbookStore.sheetList"
        :key="sheet.id"
        class="sheet-tab group relative flex items-center h-7 px-2.5 rounded-t text-xs cursor-pointer border-t-2 transition-colors"
        :class="
          sheet.id === workbookStore.activeSheet?.id
            ? 'bg-base-100 border-primary font-bold shadow-sm'
            : 'bg-base-300/40 hover:bg-base-300 border-transparent text-base-content/70'
        "
        @click="selectSheet(sheet.id)"
      >
        <Table class="w-3.5 h-3.5 mr-1.5 opacity-60 flex-shrink-0" />

        <!-- Inline rename or Display text -->
        <input
          v-if="editingSheetId === sheet.id"
          ref="renameInputRef"
          v-model="editingName"
          type="text"
          class="input input-xs h-5 px-1 py-0 w-24 text-xs font-normal"
          @blur="saveRename(sheet.id)"
          @keydown.enter="saveRename(sheet.id)"
          @keydown.esc="cancelRename"
          @click.stop
        />
        <span
          v-else
          class="truncate max-w-[120px]"
          :title="sheet.name"
          @dblclick.stop="startRename(sheet)"
        >
          {{ sheet.name }}
        </span>

        <button
          class="btn btn-ghost btn-xs btn-circle h-4 w-4 min-h-0 p-0 ml-1 opacity-70 hover:opacity-100"
          title="Tùy chọn sheet"
          @click="toggleSheetMenu(sheet.id, $event)"
        >
          <MoreVertical class="w-3 h-3 opacity-60" />
        </button>

        <Teleport to="body">
          <div
            v-if="openMenuId === sheet.id"
            class="fixed z-[80] w-[260px] rounded-box border border-base-300 bg-base-100 p-1 shadow-xl text-xs"
            :style="menuStyle"
            @click.stop
          >
            <div class="px-2 py-1.5 border-b border-base-300 mb-1">
              <label class="block text-[10px] font-semibold uppercase opacity-60 mb-1">
                Dòng header của sheet này
              </label>
              <select
                class="select select-xs select-bordered w-full text-xs"
                :value="sheet.headerRowIndex"
                @change="onPickHeader(sheet, Number(($event.target as HTMLSelectElement).value))"
              >
                <option v-for="choice in headerChoices(sheet)" :key="choice.index" :value="choice.index">
                  {{ choice.label }}
                </option>
              </select>
            </div>
            <ul class="menu p-0">
              <li>
                <a class="py-1 px-2" @click="closeSheetMenu(); startRename(sheet)">
                  <Edit2 class="w-3 h-3" /> Đổi tên
                </a>
              </li>
              <li>
                <a class="py-1 px-2" @click="closeSheetMenu(); onDuplicateSheet(sheet)">
                  <Copy class="w-3 h-3" /> Nhân bản
                </a>
              </li>
              <li v-if="workbookStore.sheetList.length > 1">
                <a class="py-1 px-2 text-error" @click="closeSheetMenu(); onDeleteSheet(sheet)">
                  <Trash2 class="w-3 h-3" /> Xóa sheet
                </a>
              </li>
            </ul>
          </div>
        </Teleport>
      </div>
    </div>
  </div>
</template>
