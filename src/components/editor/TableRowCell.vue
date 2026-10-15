<script setup lang="ts">
import { computed, ref, watch, nextTick } from 'vue'
import type { CellValue, ColumnHeader } from '@/types/excel'
import { useImageQuickLook } from '@/composables/imageQuickLook'
import { useLocale } from '@/composables/useLocale'
import { DEFAULT_DATE_FORMAT, formatDateValue, parseDateToIso } from '@/utils/dateFormat'
import { formatNumber } from '@/utils/numberFormat'
import { imageLabel, imageSources } from '@/utils/imageSource'

const props = defineProps<{
  value: CellValue
  columnKey: string
  rowIndex: number
  dataType?: ColumnHeader['dataType']
  dateFormat?: string
  numFmt?: string
  isEditing: boolean
}>()

const emit = defineEmits<{
  update: [newVal: CellValue]
  'start-edit': []
  'end-edit': []
}>()

const localVal = ref<string>('')
const inputRef = ref<HTMLInputElement | null>(null)
const failedThumbs = ref<Set<string>>(new Set())
const { t } = useLocale()
const quickLook = useImageQuickLook()
const previews = computed(() => imageSources(props.value))

watch(previews, () => {
  failedThumbs.value = new Set()
})

function openPreview(source: string) {
  quickLook.open(source, imageLabel(source))
}

function markThumbFailed(source: string) {
  failedThumbs.value = new Set(failedThumbs.value).add(source)
}

watch(
  () => props.isEditing,
  (editing) => {
    if (editing) {
      if (props.dataType === 'date') {
        localVal.value = formatDateValue(props.value, props.dateFormat ?? DEFAULT_DATE_FORMAT)
      } else if (props.dataType === 'boolean') {
        localVal.value = props.value ? 'true' : 'false'
      } else {
        localVal.value = props.value != null ? String(props.value) : ''
      }
      nextTick(() => {
        inputRef.value?.focus()
        if (props.dataType !== 'boolean') inputRef.value?.select()
      })
    }
  },
  { immediate: true },
)

function save() {
  if (!props.isEditing) return

  let finalVal: CellValue = localVal.value
  if (props.dataType === 'number') {
    const raw = localVal.value.trim()
    if (raw === '') finalVal = null
    else {
      const num = Number(raw)
      finalVal = Number.isNaN(num) ? raw : num
    }
  } else if (props.dataType === 'boolean') {
    finalVal = localVal.value === 'true'
  } else if (props.dataType === 'date') {
    const raw = localVal.value.trim()
    finalVal = raw === '' ? null : (parseDateToIso(raw, props.dateFormat ?? DEFAULT_DATE_FORMAT) ?? raw)
  }

  emit('update', finalVal)
  emit('end-edit')
}

function cancel() {
  emit('end-edit')
}

function onBooleanChange(event: Event) {
  localVal.value = (event.target as HTMLInputElement).checked ? 'true' : 'false'
}

function formatDisplay(val: CellValue): string {
  if (val === null || val === undefined || val === '') return ''
  if (props.dataType === 'number') {
    const num = Number(val)
    return isNaN(num) ? String(val) : formatNumber(num, props.numFmt)
  }
  if (props.dataType === 'boolean') {
    return val ? '✓' : '✗'
  }
  if (props.dataType === 'date') return formatDateValue(val, props.dateFormat ?? DEFAULT_DATE_FORMAT)
  return String(val)
}
</script>

<template>
  <td
    class="p-0 h-8 max-w-[150px] border-r border-b border-base-300 relative text-xs"
    :class="{ 'cell-editing': props.isEditing }"
    @dblclick.stop="emit('start-edit')"
  >
    <div v-if="props.isEditing" class="h-full w-full flex items-center p-0.5">
      <input
        v-if="props.dataType === 'boolean'"
        ref="inputRef"
        type="checkbox"
        class="checkbox checkbox-xs checkbox-primary mx-auto"
        :checked="localVal === 'true'"
        @change="onBooleanChange"
        @blur="save"
        @keydown.enter.prevent="save"
        @keydown.esc="cancel"
        @click.stop
      />
      <input
        v-else
        ref="inputRef"
        v-model="localVal"
        :type="props.dataType === 'number' ? 'number' : 'text'"
        class="input input-xs select-text w-full h-7 px-1.5 py-0 text-xs rounded border-primary outline-none focus:ring-1 focus:ring-primary"
        @blur="save"
        @keydown.enter.prevent="save"
        @keydown.esc="cancel"
        @click.stop
      />
    </div>

    <div
      v-else-if="previews.length"
      class="flex h-full w-full items-center gap-1 overflow-x-auto overscroll-x-contain px-1"
      :title="props.value != null && props.value !== '' ? String(props.value) : ''"
    >
      <button
        v-for="(source, imageIndex) in previews"
        v-show="!failedThumbs.has(source)"
        :key="`${imageIndex}-${source}`"
        type="button"
        class="h-6 w-6 shrink-0 overflow-hidden rounded border border-base-300 bg-base-200"
        :aria-label="t('image.quicklook')"
        :title="t('image.quicklook')"
        @click.stop="openPreview(source)"
        @dblclick.stop
      >
        <img :src="source" alt="" class="h-full w-full object-cover" @error="markThumbFailed(source)" />
      </button>
    </div>

    <!-- Display Value -->
    <div
      v-else
      class="px-2 py-1 h-full flex items-center overflow-hidden text-ellipsis whitespace-nowrap cursor-text"
      :class="{
        'text-right justify-end font-mono': props.dataType === 'number',
        'text-center justify-center font-bold text-primary': props.dataType === 'boolean',
      }"
      :title="props.value != null && props.value !== '' ? String(props.value) : 'Bấm đúp để sửa trong ô'"
    >
      <span class="truncate">{{ formatDisplay(props.value) }}</span>
    </div>
  </td>
</template>
