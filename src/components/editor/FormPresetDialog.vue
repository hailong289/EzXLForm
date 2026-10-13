<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Calendar, Check, CheckSquare, Hash, LayoutTemplate, Search, Type, X } from 'lucide-vue-next'
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
const columnQuery = ref('')

watch(
  () => props.open,
  (open) => {
    if (!open) return
    editingId.value = null
    name.value = ''
    checked.value = []
    columnQuery.value = ''
  },
)

const checkedSet = computed(() => new Set(checked.value))

const filteredHeaders = computed(() => {
  const query = columnQuery.value.trim().toLowerCase()
  if (!query) return props.headers
  return props.headers.filter((header) => header.label.toLowerCase().includes(query))
})

function beginEdit(preset: FormPreset) {
  editingId.value = preset.id
  name.value = preset.name
  checked.value = props.headers.map((header) => header.key).filter((key) => preset.visibleKeys.includes(key))
  columnQuery.value = ''
}

function toggle(key: string) {
  if (checkedSet.value.has(key)) checked.value = checked.value.filter((item) => item !== key)
  else checked.value = [...checked.value, key]
}

function selectAll() {
  checked.value = props.headers.map((header) => header.key)
}

function clearAll() {
  checked.value = []
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
    <div class="modal-box flex max-h-[min(40rem,calc(100vh-2rem))] max-w-xl flex-col gap-0 overflow-hidden border border-base-300 p-0">
      <div class="flex items-center justify-between gap-3 border-b border-base-300 px-4 py-3">
        <div class="flex min-w-0 items-center gap-3">
          <span class="grid h-9 w-9 shrink-0 place-items-center rounded-xl bg-primary/10 text-primary">
            <LayoutTemplate class="h-4 w-4" />
          </span>
          <div class="min-w-0">
            <h3 class="text-sm font-semibold leading-tight">Loại form</h3>
            <p class="text-[11px] opacity-60">Chọn cột sẽ hiện khi nhập hoặc sửa dòng</p>
          </div>
        </div>
        <button type="button" class="btn btn-ghost btn-xs btn-circle" title="Đóng" @click="emit('close')">
          <X class="h-4 w-4" />
        </button>
      </div>

      <div class="flex min-h-0 flex-1 flex-col gap-3 overflow-y-auto px-4 py-3">
        <div v-if="presets.length" class="flex flex-col gap-1.5">
          <p class="text-[11px] font-medium opacity-60">Đã có</p>
          <ul class="flex flex-col gap-1">
            <li
              v-for="preset in presets"
              :key="preset.id"
              class="flex items-center justify-between gap-2 rounded-lg border border-base-300 bg-base-100 px-2.5 py-1.5"
              :class="editingId === preset.id ? 'border-primary/50 bg-primary/5' : ''"
            >
              <span class="truncate text-sm">{{ preset.name }}</span>
              <span class="flex shrink-0 gap-1">
                <button type="button" class="btn btn-ghost btn-xs" @click="beginEdit(preset)">Sửa</button>
                <button type="button" class="btn btn-ghost btn-xs text-error" @click="emit('remove', preset.id)">Xóa</button>
              </span>
            </li>
          </ul>
        </div>

        <label class="flex flex-col gap-1 text-xs">
          <span class="font-medium opacity-70">{{ editingId ? 'Sửa loại' : 'Tên loại mới' }}</span>
          <input v-model="name" class="input input-bordered input-sm" placeholder="Ví dụ: Nhập hàng hóa" />
        </label>

        <div class="flex flex-col gap-2">
          <div class="flex items-center justify-between gap-2">
            <p class="text-xs font-medium">
              Cột trên form
              <span class="font-normal opacity-60">{{ checked.length }}/{{ headers.length }}</span>
            </p>
            <div class="flex gap-1">
              <button type="button" class="btn btn-ghost btn-xs" @click="selectAll">Chọn hết</button>
              <button type="button" class="btn btn-ghost btn-xs" @click="clearAll">Bỏ chọn</button>
            </div>
          </div>
          <label class="input input-bordered input-sm flex items-center gap-2">
            <Search class="h-3.5 w-3.5 shrink-0 opacity-50" />
            <input v-model="columnQuery" class="grow bg-transparent text-xs" placeholder="Tìm cột..." />
          </label>
          <p v-if="filteredHeaders.length === 0" class="py-6 text-center text-xs opacity-60">
            Không có cột khớp.
          </p>
          <div v-else class="grid grid-cols-2 gap-2">
            <button
              v-for="header in filteredHeaders"
              :key="header.key"
              type="button"
              class="flex items-center gap-2 rounded-xl border px-2.5 py-2 text-left transition-colors"
              :class="checkedSet.has(header.key)
                ? 'border-primary bg-primary/10'
                : 'border-base-300 bg-base-100 hover:bg-base-200/60'"
              @click="toggle(header.key)"
            >
              <Type v-if="header.dataType === 'text' || !header.dataType" class="h-3.5 w-3.5 shrink-0 text-primary opacity-70" />
              <Hash v-else-if="header.dataType === 'number'" class="h-3.5 w-3.5 shrink-0 text-info opacity-70" />
              <Calendar v-else-if="header.dataType === 'date'" class="h-3.5 w-3.5 shrink-0 text-warning opacity-70" />
              <CheckSquare v-else class="h-3.5 w-3.5 shrink-0 text-success opacity-70" />
              <span class="min-w-0 flex-1 truncate text-xs font-medium">{{ header.label }}</span>
              <span
                class="grid h-4 w-4 shrink-0 place-items-center rounded-full border"
                :class="checkedSet.has(header.key) ? 'border-primary bg-primary text-primary-content' : 'border-base-300'"
              >
                <Check v-if="checkedSet.has(header.key)" class="h-3 w-3" />
              </span>
            </button>
          </div>
        </div>

        <p v-if="error" class="text-xs text-error">{{ errorText[error] }}</p>
      </div>

      <div class="flex items-center justify-end gap-2 border-t border-base-300 px-4 py-3">
        <button type="button" class="btn btn-ghost btn-sm" @click="emit('close')">Đóng</button>
        <button type="button" class="btn btn-primary btn-sm" @click="submit">
          {{ editingId ? 'Lưu' : 'Thêm loại' }}
        </button>
      </div>
    </div>
    <div class="modal-backdrop bg-black/40" @click="emit('close')" />
  </dialog>
</template>
