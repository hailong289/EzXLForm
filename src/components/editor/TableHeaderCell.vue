<script setup lang="ts">
import { ref, nextTick } from 'vue'
import {
  MoreVertical,
  ArrowUpAZ,
  ArrowDownZA,
  ArrowLeft,
  ArrowRight,
  Plus,
  Trash2,
  FileText,
  Binary,
  Calendar,
  CheckSquare,
  GripVertical,
} from 'lucide-vue-next'
import type { ColumnHeader } from '@/types/excel'

const props = defineProps<{
  header: ColumnHeader
  isFirst: boolean
  isLast: boolean
  index: number
  isDropTarget?: boolean
  isDragging?: boolean
}>()

const emit = defineEmits<{
  rename: [key: string, newLabel: string]
  delete: [key: string]
  'add-left': [key: string]
  'add-right': [key: string]
  'move-left': [key: string]
  'move-right': [key: string]
  'sort-asc': [key: string]
  'sort-desc': [key: string]
  'set-type': [key: string, type: 'text' | 'number' | 'date' | 'boolean']
  'drag-start': [key: string]
  'drag-over': [key: string]
  drop: [key: string]
  'drag-end': []
}>()

const isEditing = ref(false)
const editValue = ref('')
const inputRef = ref<HTMLInputElement | null>(null)

function startEdit() {
  isEditing.value = true
  editValue.value = props.header.label
  nextTick(() => {
    inputRef.value?.focus()
    inputRef.value?.select()
  })
}

function saveEdit() {
  if (isEditing.value) {
    const trimmed = editValue.value.trim()
    if (trimmed && trimmed !== props.header.label) {
      emit('rename', props.header.key, trimmed)
    }
    isEditing.value = false
  }
}

function cancelEdit() {
  isEditing.value = false
}

function onGripDragStart(event: DragEvent) {
  event.dataTransfer?.setData('text/plain', props.header.key)
  if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move'
  emit('drag-start', props.header.key)
}
</script>

<template>
  <th
    class="relative p-0 h-9 font-semibold text-xs border-r border-base-300 bg-base-200 group select-none"
    :class="{
      'opacity-40': isDragging,
      'outline outline-2 outline-primary': isDropTarget,
    }"
    @dragover.prevent="emit('drag-over', header.key)"
    @drop.prevent="emit('drop', header.key)"
  >
    <div class="flex items-center justify-between h-full px-1 gap-1">
      <span
        class="cursor-grab active:cursor-grabbing text-base-content/40 hover:text-base-content px-0.5"
        title="Kéo để đổi chỗ cột"
        draggable="true"
        @dragstart="onGripDragStart"
        @dragend="emit('drag-end')"
        @click.stop
      >
        <GripVertical class="w-3.5 h-3.5" />
      </span>

      <div class="flex-1 truncate min-w-0">
        <input
          v-if="isEditing"
          ref="inputRef"
          v-model="editValue"
          type="text"
          class="input input-xs h-6 px-1.5 py-0 w-full text-xs font-semibold"
          @blur="saveEdit"
          @keydown.enter="saveEdit"
          @keydown.esc="cancelEdit"
          @click.stop
        />
        <span
          v-else
          class="cursor-text truncate block hover:text-primary transition-colors"
          :title="`Bấm để đổi tên: ${header.label}`"
          @click.stop="startEdit"
        >
          {{ header.label }}
        </span>
      </div>

      <!-- Type badge icon -->
      <span class="opacity-40 text-[10px] uppercase font-mono mr-0.5">
        {{ header.dataType === 'number' ? '#' : header.dataType === 'date' ? '📅' : '' }}
      </span>

      <!-- Action Dropdown Menu -->
      <div class="dropdown dropdown-end dropdown-bottom">
        <label
          tabindex="0"
          class="btn btn-ghost btn-xs btn-circle h-5 w-5 min-h-0 p-0 opacity-0 group-hover:opacity-100 transition-opacity"
          @click.stop
        >
          <MoreVertical class="w-3.5 h-3.5 opacity-60" />
        </label>

        <ul
          tabindex="0"
          class="dropdown-content menu p-1 shadow-xl bg-base-100 rounded-box w-44 border border-base-300 text-xs z-50 font-normal"
        >
          <li>
            <a class="py-1" @click="startEdit">
              <FileText class="w-3.5 h-3.5" /> Đổi tên
            </a>
          </li>
          <li>
            <a class="py-1" @click="emit('sort-asc', header.key)">
              <ArrowUpAZ class="w-3.5 h-3.5 text-primary" /> Sắp xếp A → Z
            </a>
          </li>
          <li>
            <a class="py-1" @click="emit('sort-desc', header.key)">
              <ArrowDownZA class="w-3.5 h-3.5 text-primary" /> Sắp xếp Z → A
            </a>
          </li>

          <div class="divider my-0.5" />

          <li>
            <a class="py-1" @click="emit('add-left', header.key)">
              <Plus class="w-3.5 h-3.5" /> Chèn cột bên trái
            </a>
          </li>
          <li>
            <a class="py-1" @click="emit('add-right', header.key)">
              <Plus class="w-3.5 h-3.5" /> Chèn cột bên phải
            </a>
          </li>

          <li :class="{ disabled: isFirst }">
            <a class="py-1" @click="!isFirst && emit('move-left', header.key)">
              <ArrowLeft class="w-3.5 h-3.5" /> Chuyển sang trái
            </a>
          </li>
          <li :class="{ disabled: isLast }">
            <a class="py-1" @click="!isLast && emit('move-right', header.key)">
              <ArrowRight class="w-3.5 h-3.5" /> Chuyển sang phải
            </a>
          </li>

          <div class="divider my-0.5" />

          <!-- Submenu Kiểu dữ liệu -->
          <li class="menu-title py-0.5 text-[10px] uppercase opacity-60">Kiểu dữ liệu</li>
          <li>
            <a
              class="py-0.5"
              :class="{ 'active font-bold': header.dataType === 'text' }"
              @click="emit('set-type', header.key, 'text')"
            >
              <FileText class="w-3 h-3" /> Văn bản (Text)
            </a>
          </li>
          <li>
            <a
              class="py-0.5"
              :class="{ 'active font-bold': header.dataType === 'number' }"
              @click="emit('set-type', header.key, 'number')"
            >
              <Binary class="w-3 h-3" /> Số (Number)
            </a>
          </li>
          <li>
            <a
              class="py-0.5"
              :class="{ 'active font-bold': header.dataType === 'date' }"
              @click="emit('set-type', header.key, 'date')"
            >
              <Calendar class="w-3 h-3" /> Ngày (Date)
            </a>
          </li>
          <li>
            <a
              class="py-0.5"
              :class="{ 'active font-bold': header.dataType === 'boolean' }"
              @click="emit('set-type', header.key, 'boolean')"
            >
              <CheckSquare class="w-3 h-3" /> Đúng/Sai (Boolean)
            </a>
          </li>

          <div class="divider my-0.5" />

          <li>
            <a class="py-1 text-error" @click="emit('delete', header.key)">
              <Trash2 class="w-3.5 h-3.5" /> Xóa cột này
            </a>
          </li>
        </ul>
      </div>
    </div>
  </th>
</template>
