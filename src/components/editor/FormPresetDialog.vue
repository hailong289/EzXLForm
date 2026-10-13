<script setup lang="ts">
import { ref, watch } from 'vue'
import type { ColumnHeader } from '@/types/excel'
import type { FormPreset, PresetDraftError } from '@/utils/formPresets'

const props = defineProps<{
  open: boolean
  headers: ColumnHeader[]
  presets: FormPreset[]
  error: PresetDraftError | null
}>()

const emit = defineEmits<{
  close: []
  create: [name: string, visibleKeys: string[]]
  update: [id: string, name: string, visibleKeys: string[]]
  remove: [id: string]
}>()

const editingId = ref<string | null>(null)
const name = ref('')
const checked = ref<string[]>([])

watch(
  () => props.open,
  (open) => {
    if (!open) return
    editingId.value = null
    name.value = ''
    checked.value = []
  },
)

function beginEdit(preset: FormPreset) {
  editingId.value = preset.id
  name.value = preset.name
  checked.value = props.headers.map((header) => header.key).filter((key) => preset.visibleKeys.includes(key))
}

function chosenKeys(): string[] {
  const allowed = new Set(checked.value)
  return props.headers.map((header) => header.key).filter((key) => allowed.has(key))
}

function submit() {
  if (editingId.value) emit('update', editingId.value, name.value, chosenKeys())
  else emit('create', name.value, chosenKeys())
}

const errorText: Record<PresetDraftError, string> = {
  'empty-name': 'Nhập tên loại form.',
  'duplicate-name': 'Tên này đã có.',
  'no-columns': 'Chọn ít nhất một cột.',
}
</script>

<template>
  <dialog v-if="open" class="modal modal-open" @keydown.esc.prevent="emit('close')">
    <div class="modal-box max-w-lg">
      <h3 class="font-bold text-base">Loại form</h3>
      <ul class="mt-3 flex flex-col gap-1">
        <li v-for="preset in presets" :key="preset.id" class="flex items-center justify-between gap-2">
          <span class="truncate text-sm">{{ preset.name }}</span>
          <span class="flex gap-1">
            <button type="button" class="btn btn-ghost btn-xs" @click="beginEdit(preset)">Sửa</button>
            <button type="button" class="btn btn-ghost btn-xs text-error" @click="emit('remove', preset.id)">Xóa</button>
          </span>
        </li>
      </ul>
      <label class="mt-4 flex flex-col gap-1 text-xs">
        Tên loại form
        <input v-model="name" class="input input-bordered input-sm" placeholder="Tên loại form" />
      </label>
      <div class="mt-3 flex max-h-48 flex-col gap-1 overflow-y-auto">
        <label v-for="header in headers" :key="header.key" class="flex items-center gap-2 text-sm">
          <input v-model="checked" type="checkbox" class="checkbox checkbox-xs" :value="header.key" />
          {{ header.label }}
        </label>
      </div>
      <p v-if="error" class="mt-2 text-xs text-error">{{ errorText[error] }}</p>
      <div class="modal-action">
        <button type="button" class="btn btn-ghost btn-sm" @click="emit('close')">Đóng</button>
        <button type="button" class="btn btn-primary btn-sm" @click="submit">
          {{ editingId ? 'Lưu' : 'Thêm loại' }}
        </button>
      </div>
    </div>
  </dialog>
</template>
