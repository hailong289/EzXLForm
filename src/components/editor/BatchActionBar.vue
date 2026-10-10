<script setup lang="ts">
import { ref } from 'vue'
import {
  CheckSquare,
  Trash2,
  Type,
  X,
  Edit2,
  ChevronDown,
  Sparkles,
  ArrowRight,
} from 'lucide-vue-next'
import { useSelectionStore } from '@/stores/selectionStore'
import { useWorkbookStore } from '@/stores/workbookStore'
import { useHistory } from '@/composables/useHistory'
import { useAppConfigStore } from '@/stores/appConfigStore'
import { useBatchOperations } from '@/composables/useBatchOperations'
import { physicalOf, getRow, hasSheetOrder } from '@/composables/rowPageStore'
import { useRowMarks } from '@/composables/rowMarks'
import { sheetRowCount } from '@/utils/sheetGrid'
import type { ColumnHeader, BatchOperationType } from '@/types/excel'

const props = defineProps<{
  headers: ColumnHeader[]
}>()

const selectionStore = useSelectionStore()
const rowMarks = useRowMarks()
const workbookStore = useWorkbookStore()
const appConfigStore = useAppConfigStore()
const { push: pushHistory } = useHistory()
const { applyBatchOperation } = useBatchOperations()

// Inline dialog states for batch actions
const activeAction = ref<BatchOperationType | null>(null)
const selectedColKey = ref<string>('')
const actionValue = ref<string>('')

function openAction(action: BatchOperationType) {
  activeAction.value = action
  selectedColKey.value = props.headers[0]?.key || ''
  actionValue.value = ''
}

function closeAction() {
  activeAction.value = null
  actionValue.value = ''
}

async function executeBatchAction() {
  if (!activeAction.value || !selectedColKey.value || !workbookStore.activeSheet) return

  const sheetId = workbookStore.activeSheet.id
  const targetIndexes = [...selectionStore.selectedRowIndexes]
  const operation = {
    type: activeAction.value,
    columnKey: selectedColKey.value,
    value: actionValue.value,
    targetRowIndexes: targetIndexes,
  }

  if (workbookStore.activeSheet.paged) {
    for (const index of targetIndexes) {
      const row = (await getRow(sheetId, physicalOf(sheetId, index))) ?? {}
      const [next] = applyBatchOperation([row], { ...operation, targetRowIndexes: [0] })
      await workbookStore.replaceRow(sheetId, index, next)
    }
    appConfigStore.addToast({
      type: 'success',
      message: `Đã cập nhật ${targetIndexes.length} dòng.`,
    })
    closeAction()
    return
  }

  const rowsBackup = JSON.parse(JSON.stringify(workbookStore.activeSheet.rows))

  const updatedRows = applyBatchOperation(workbookStore.activeSheet.rows, {
    type: activeAction.value,
    columnKey: selectedColKey.value,
    value: actionValue.value,
    targetRowIndexes: targetIndexes,
  })

  // Cập nhật rows trong store
  workbookStore.activeSheet.rows = updatedRows
  workbookStore.isDirty = true

  const colLabel = props.headers.find((h) => h.key === selectedColKey.value)?.label || selectedColKey.value
  const desc = `Sửa hàng loạt: ${activeAction.value} trên cột "${colLabel}" (${targetIndexes.length} dòng)`

  pushHistory({
    type: 'BATCH_EDIT',
    description: desc,
    sheetId,
    before: { rows: rowsBackup },
    after: { rows: JSON.parse(JSON.stringify(updatedRows)) },
  })

  appConfigStore.addToast({
    type: 'success',
    message: `Đã cập nhật ${targetIndexes.length} dòng.`,
  })

  closeAction()
}

async function handleDeleteSelected() {
  if (!workbookStore.activeSheet || selectionStore.selectedCount === 0) return

  const sheetId = workbookStore.activeSheet.id
  const count = selectionStore.selectedCount
  const targetIndexes = [...selectionStore.selectedRowIndexes]
  const swapTail = Boolean(workbookStore.activeSheet.paged) && !hasSheetOrder(sheetId)
  rowMarks.noteRowsDeleted(sheetId, targetIndexes, swapTail, sheetRowCount(workbookStore.activeSheet))
  if (workbookStore.activeSheet.paged) {
    await workbookStore.deleteRows(sheetId, targetIndexes)
    selectionStore.clearSelection()
    appConfigStore.addToast({
      type: 'info',
      message: `Đã xóa ${count} dòng. Dòng cuối được đưa lên chỗ trống để khỏi dịch cả file.`,
    })
    return
  }
  const rowsBackup = JSON.parse(JSON.stringify(workbookStore.activeSheet.rows))

  workbookStore.deleteRows(sheetId, targetIndexes)
  selectionStore.clearSelection()

  pushHistory({
    type: 'DELETE_ROWS',
    description: `Xóa ${count} dòng đã chọn`,
    sheetId,
    before: { rows: rowsBackup },
    after: { rows: JSON.parse(JSON.stringify(workbookStore.activeSheet.rows)) },
  })

  appConfigStore.addToast({
    type: 'info',
    message: `Đã xóa ${count} dòng.`,
  })
}
</script>

<template>
  <Transition name="batch-bar">
    <div
      v-if="selectionStore.selectedCount > 0"
      class="fixed bottom-12 left-1/2 -translate-x-1/2 z-40 bg-neutral text-neutral-content shadow-2xl rounded-2xl px-4 py-2.5 flex items-center gap-3 border border-neutral-focus/50 select-none max-w-[90vw] flex-wrap"
    >
      <!-- Left: Counter -->
      <div class="flex items-center gap-1.5 font-bold text-xs pr-2 border-r border-neutral-content/20">
        <CheckSquare class="w-4 h-4 text-primary" />
        <span>{{ selectionStore.selectedCount }} dòng đã chọn</span>
      </div>

      <!-- Action Buttons -->
      <div v-if="!activeAction" class="flex items-center gap-1.5 flex-wrap">
        <!-- Gán giá trị -->
        <button
          class="btn btn-xs btn-ghost hover:bg-neutral-focus gap-1 text-xs"
          @click="openAction('SET_VALUE')"
        >
          <Edit2 class="w-3.5 h-3.5 text-info" /> Gán giá trị
        </button>

        <!-- Thêm tiền tố -->
        <button
          class="btn btn-xs btn-ghost hover:bg-neutral-focus gap-1 text-xs"
          @click="openAction('ADD_PREFIX')"
        >
          <Sparkles class="w-3.5 h-3.5 text-warning" /> Thêm tiền tố
        </button>

        <!-- Thêm hậu tố -->
        <button
          class="btn btn-xs btn-ghost hover:bg-neutral-focus gap-1 text-xs"
          @click="openAction('ADD_SUFFIX')"
        >
          <Sparkles class="w-3.5 h-3.5 text-warning" /> Thêm hậu tố
        </button>

        <!-- Dropdown kiểu chữ -->
        <div class="dropdown dropdown-top">
          <label tabindex="0" class="btn btn-xs btn-ghost hover:bg-neutral-focus gap-1 text-xs">
            <Type class="w-3.5 h-3.5 text-success" /> Đổi chữ <ChevronDown class="w-3 h-3" />
          </label>
          <ul
            tabindex="0"
            class="dropdown-content menu p-1 shadow-lg bg-base-100 text-base-content rounded-box w-40 border border-base-300 text-xs z-50 mb-1"
          >
            <li><a @click="openAction('UPPERCASE')">VIẾT HOA</a></li>
            <li><a @click="openAction('LOWERCASE')">viết thường</a></li>
            <li><a @click="openAction('CAPITALIZE')">Viết Hoa Đầu Từ</a></li>
            <li><a @click="openAction('TRIM')">Xóa khoảng trắng</a></li>
            <li><a class="text-error" @click="openAction('CLEAR')">Xóa ô rỗng</a></li>
          </ul>
        </div>

        <!-- Xóa các dòng -->
        <button
          class="btn btn-xs btn-error btn-outline gap-1 text-xs ml-1"
          @click="handleDeleteSelected"
        >
          <Trash2 class="w-3.5 h-3.5" /> Xóa
        </button>
      </div>

      <!-- Mini-Form khi đang thực hiện Action -->
      <div v-else class="flex items-center gap-2 flex-wrap">
        <span class="text-xs font-semibold text-primary">
          {{
            activeAction === 'SET_VALUE'
              ? 'Gán giá trị'
              : activeAction === 'ADD_PREFIX'
              ? 'Thêm tiền tố'
              : activeAction === 'ADD_SUFFIX'
              ? 'Thêm hậu tố'
              : activeAction
          }}:
        </span>

        <!-- Chọn Cột -->
        <select v-model="selectedColKey" class="select select-xs select-bordered bg-base-100 text-base-content text-xs">
          <option v-for="h in props.headers" :key="h.key" :value="h.key">
            {{ h.label }}
          </option>
        </select>

        <!-- Nhập giá trị nếu cần -->
        <input
          v-if="['SET_VALUE', 'ADD_PREFIX', 'ADD_SUFFIX'].includes(activeAction)"
          v-model="actionValue"
          type="text"
          class="input input-xs input-bordered bg-base-100 text-base-content text-xs w-32"
          placeholder="Nhập giá trị..."
          @keydown.enter="executeBatchAction"
        />

        <button class="btn btn-xs btn-primary gap-1" @click="executeBatchAction">
          <ArrowRight class="w-3 h-3" /> Thực hiện
        </button>
        <button class="btn btn-xs btn-ghost" @click="closeAction">
          Hủy
        </button>
      </div>

      <!-- Clear selection button -->
      <button
        class="btn btn-ghost btn-xs btn-circle ml-2 hover:bg-neutral-focus"
        title="Bỏ chọn tất cả"
        @click="selectionStore.clearSelection"
      >
        <X class="w-4 h-4" />
      </button>
    </div>
  </Transition>
</template>
