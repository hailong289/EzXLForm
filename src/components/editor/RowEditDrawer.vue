<script setup lang="ts">
import { ref, watch } from 'vue'
import { X, Check, Edit3 } from 'lucide-vue-next'
import type { ColumnHeader, RowData, CellValue } from '@/types/excel'

const props = defineProps<{
  modelValue: boolean
  rowIndex: number | null
  row: RowData | null
  headers: ColumnHeader[]
}>()

const emit = defineEmits<{
  'update:modelValue': [val: boolean]
  save: [updatedRow: RowData]
}>()

const formData = ref<Record<string, string>>({})

watch(
  () => props.row,
  (newRow) => {
    if (newRow) {
      const data: Record<string, string> = {}
      props.headers.forEach((h) => {
        const val = newRow[h.key]
        data[h.key] = val != null ? String(val) : ''
      })
      formData.value = data
    } else {
      formData.value = {}
    }
  },
  { immediate: true },
)

function close() {
  emit('update:modelValue', false)
}

function handleSave() {
  const updated: RowData = {}
  props.headers.forEach((h) => {
    const raw = formData.value[h.key] ?? ''
    let val: CellValue = raw
    if (h.dataType === 'number') {
      val = raw.trim() === '' ? null : Number(raw)
    } else if (h.dataType === 'boolean') {
      val = raw === 'true' || raw === '1'
    }
    updated[h.key] = val
  })

  emit('save', updated)
  close()
}
</script>

<template>
  <div v-if="props.modelValue" class="fixed inset-0 z-50 flex justify-end select-none">
    <!-- Backdrop -->
    <div class="fixed inset-0 bg-black/40 backdrop-blur-xs transition-opacity" @click="close" />

    <!-- Drawer Panel -->
    <div
      class="relative w-96 max-w-full bg-base-100 shadow-2xl h-full flex flex-col border-l border-base-300 z-10"
    >
      <!-- Header -->
      <div class="flex items-center justify-between p-4 border-b border-base-300 bg-base-200/50">
        <div class="flex items-center gap-2">
          <Edit3 class="w-5 h-5 text-primary" />
          <div>
            <h3 class="font-bold text-sm">Chỉnh sửa Dòng #{{ (props.rowIndex ?? 0) + 1 }}</h3>
            <span class="text-[11px] opacity-60">Chi tiết tất cả các cột dữ liệu</span>
          </div>
        </div>
        <button class="btn btn-ghost btn-xs btn-circle" @click="close">
          <X class="w-4 h-4" />
        </button>
      </div>

      <!-- Fields List -->
      <div class="flex-1 overflow-y-auto p-4 space-y-3.5 scrollbar-thin">
        <div v-for="h in props.headers" :key="h.key" class="form-control">
          <label class="label py-1">
            <span class="label-text text-xs font-semibold flex items-center gap-1.5">
              {{ h.label }}
              <span class="text-[10px] opacity-50 uppercase font-mono">({{ h.dataType || 'text' }})</span>
            </span>
          </label>

          <!-- Input by type -->
          <input
            v-if="h.dataType !== 'boolean'"
            v-model="formData[h.key]"
            :type="h.dataType === 'number' ? 'number' : h.dataType === 'date' ? 'date' : 'text'"
            class="input input-sm input-bordered w-full text-xs"
            :placeholder="`Nhập ${h.label}...`"
          />

          <!-- Checkbox for boolean -->
          <div v-else class="flex items-center gap-2 mt-1">
            <input
              type="checkbox"
              class="checkbox checkbox-sm checkbox-primary"
              :checked="formData[h.key] === 'true' || formData[h.key] === '1'"
              @change="formData[h.key] = ($event.target as HTMLInputElement).checked ? 'true' : 'false'"
            />
            <span class="text-xs">{{ formData[h.key] === 'true' ? 'Đúng (True)' : 'Sai (False)' }}</span>
          </div>
        </div>
      </div>

      <!-- Footer Actions -->
      <div class="p-4 border-t border-base-300 bg-base-200/50 flex justify-end gap-2">
        <button class="btn btn-sm btn-ghost" @click="close">Hủy bỏ</button>
        <button class="btn btn-sm btn-primary gap-1" @click="handleSave">
          <Check class="w-3.5 h-3.5" /> Lưu thay đổi
        </button>
      </div>
    </div>
  </div>
</template>
