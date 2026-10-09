<script setup lang="ts">
import { ref } from 'vue'
import { FileSpreadsheet, X, Check } from 'lucide-vue-next'
import type { ImportOptions } from '@/types/excel'

const props = defineProps<{
  modelValue: boolean
  fileName: string
}>()

const emit = defineEmits<{
  'update:modelValue': [val: boolean]
  confirm: [options: ImportOptions]
}>()

const headerRowIndex = ref<number>(1) // 1-based in UI, convert to 0-based
const skipRows = ref<number>(0)

function close() {
  emit('update:modelValue', false)
}

function handleConfirm() {
  emit('confirm', {
    headerRowIndex: Math.max(0, headerRowIndex.value - 1),
    skipRows: Math.max(0, skipRows.value),
  })
  close()
}
</script>

<template>
  <dialog v-if="props.modelValue" class="modal modal-open select-none">
    <div class="modal-box max-w-sm p-5 border border-base-300 shadow-2xl">
      <!-- Header -->
      <div class="flex items-center justify-between pb-3 mb-4 border-b border-base-300">
        <div class="flex items-center gap-2">
          <FileSpreadsheet class="w-5 h-5 text-primary" />
          <h3 class="font-bold text-base">Tùy chọn nhập File</h3>
        </div>
        <button class="btn btn-ghost btn-xs btn-circle" @click="close">
          <X class="w-4 h-4" />
        </button>
      </div>

      <!-- File info -->
      <div class="p-2.5 bg-base-200/50 rounded-lg border border-base-300 mb-4 text-xs">
        <span class="opacity-60 block mb-0.5">Tệp đã chọn:</span>
        <span class="font-bold truncate block text-primary">{{ props.fileName }}</span>
      </div>

      <!-- Form settings -->
      <div class="flex flex-col gap-3">
        <div>
          <label class="label py-1 text-xs font-semibold">
            Dòng làm tiêu đề cột (Header Row):
          </label>
          <input
            v-model.number="headerRowIndex"
            type="number"
            min="1"
            max="20"
            class="input input-sm input-bordered w-full text-xs"
          />
          <span class="text-[11px] opacity-60 mt-0.5 block">
            Đây là dòng mặc định khi mở file. Mỗi sheet có thể chọn một dòng header khác trong menu của sheet đó.
          </span>
        </div>

        <div>
          <label class="label py-1 text-xs font-semibold">
            Số dòng bỏ qua sau Header:
          </label>
          <input
            v-model.number="skipRows"
            type="number"
            min="0"
            max="10"
            class="input input-sm input-bordered w-full text-xs"
          />
          <span class="text-[11px] opacity-60 mt-0.5 block">
            Bỏ qua các dòng trống hoặc ghi chú đầu bảng dữ liệu.
          </span>
        </div>
      </div>

      <!-- Action Buttons -->
      <div class="modal-action mt-5 pt-3 border-t border-base-300 flex justify-end gap-2">
        <button class="btn btn-sm btn-ghost" @click="close">Hủy</button>
        <button class="btn btn-sm btn-primary gap-1" @click="handleConfirm">
          <Check class="w-3.5 h-3.5" /> Nhập dữ liệu
        </button>
      </div>
    </div>
    <div class="modal-backdrop bg-black/40" @click="close" />
  </dialog>
</template>
