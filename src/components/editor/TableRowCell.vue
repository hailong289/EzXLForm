<script setup lang="ts">
import { ref, watch, nextTick } from 'vue'
import type { CellValue, ColumnHeader } from '@/types/excel'

const props = defineProps<{
  value: CellValue
  columnKey: string
  rowIndex: number
  dataType?: ColumnHeader['dataType']
  isEditing: boolean
}>()

const emit = defineEmits<{
  update: [newVal: CellValue]
  'start-edit': []
  'end-edit': []
}>()

const localVal = ref<string>('')
const inputRef = ref<HTMLInputElement | null>(null)

watch(
  () => props.isEditing,
  (editing) => {
    if (editing) {
      localVal.value = props.value != null ? String(props.value) : ''
      nextTick(() => {
        inputRef.value?.focus()
        inputRef.value?.select()
      })
    }
  },
  { immediate: true },
)

function save() {
  if (!props.isEditing) return

  let finalVal: CellValue = localVal.value
  if (props.dataType === 'number') {
    finalVal = localVal.value.trim() === '' ? null : Number(localVal.value)
  } else if (props.dataType === 'boolean') {
    finalVal = localVal.value === 'true' || localVal.value === '1'
  }

  emit('update', finalVal)
  emit('end-edit')
}

function cancel() {
  emit('end-edit')
}

function formatDisplay(val: CellValue): string {
  if (val === null || val === undefined || val === '') return ''
  if (props.dataType === 'number') {
    const num = Number(val)
    return isNaN(num) ? String(val) : num.toLocaleString()
  }
  if (props.dataType === 'boolean') {
    return val ? '✓' : '✗'
  }
  return String(val)
}
</script>

<template>
  <td
    class="p-0 h-8 border-r border-base-300 relative text-xs"
    :class="{ 'cell-editing': props.isEditing }"
    @dblclick="emit('start-edit')"
  >
    <!-- Inline Edit Input -->
    <div v-if="props.isEditing" class="h-full w-full flex items-center p-0.5">
      <input
        ref="inputRef"
        v-model="localVal"
        :type="props.dataType === 'number' ? 'number' : props.dataType === 'date' ? 'date' : 'text'"
        class="input input-xs w-full h-7 px-1.5 py-0 text-xs rounded border-primary outline-none focus:ring-1 focus:ring-primary"
        @blur="save"
        @keydown.enter="save"
        @keydown.esc="cancel"
        @click.stop
      />
    </div>

    <!-- Display Value -->
    <div
      v-else
      class="px-2 py-1 h-full flex items-center overflow-hidden text-ellipsis whitespace-nowrap cursor-text"
      :class="{
        'text-right justify-end font-mono': props.dataType === 'number',
        'text-center justify-center font-bold text-primary': props.dataType === 'boolean',
      }"
      :title="props.value != null ? String(props.value) : ''"
    >
      {{ formatDisplay(props.value) }}
    </div>
  </td>
</template>
