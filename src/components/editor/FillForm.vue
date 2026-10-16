<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { appendSharedRow, assertCloudEditor, fillFile, readSharedForm, uniquePeers } from '@/cloud/formCloud'
import type { SharedForm, SharedFormColumn } from '@/cloud/formShare'
import type { CloudProviderId } from '@/cloud/types'
import { useLocale } from '@/composables/useLocale'
import type { ColumnHeader, RowData } from '@/types/excel'
import { coerceRowData } from '@/utils/cellValue'
import { columnPaint } from '@/utils/columnStyle'
import { applyFormulas } from '@/utils/columnRules'

const props = defineProps<{ provider: CloudProviderId; fileId: string }>()
const emit = defineEmits<{ close: [] }>()
const { t } = useLocale()

const form = ref<SharedForm | null>(null)
const fileName = ref('')
const values = ref<Record<string, string>>({})
const errors = ref<Record<string, string>>({})
const busy = ref(false)
const errorText = ref('')
const sent = ref(false)

const headers = computed(() => (form.value?.columns ?? []).map(asHeader))

function asHeader(column: SharedFormColumn): ColumnHeader {
  return {
    key: column.key,
    label: column.label,
    originalIndex: column.index,
    dataType: column.dataType,
    dateFormat: column.dateFormat,
    required: column.required,
    unique: column.unique,
    pattern: column.pattern,
    options: column.options,
    formula: column.formula,
    style: column.style,
  }
}

function labelStyle(header: ColumnHeader) {
  return columnPaint(header.style).text
}

function fieldStyle(header: ColumnHeader) {
  const paint = columnPaint(header.style)
  const style: Record<string, string> = { ...paint.box, ...paint.text }
  const border = header.style?.border
  if (!border) return style
  const color = header.style?.borderColor || '#000000'
  style.boxShadow = 'none'
  style.outline = 'none'
  if (border === 'bottom') {
    style.border = 'none'
    style.borderBottom = `1px solid ${color}`
    style.borderRadius = '0'
  } else {
    style.border = `1px solid ${color}`
  }
  return style
}

async function load() {
  busy.value = true
  errorText.value = ''
  form.value = null
  try {
    const file = await fillFile(props.provider, props.fileId)
    fileName.value = file.name
    await assertCloudEditor(props.provider, file)
    form.value = await readSharedForm(props.provider, file)
    if (!form.value) errorText.value = t('fill.missing')
    values.value = Object.fromEntries((form.value?.columns ?? []).map((column) => [column.key, '']))
  } catch (error) {
    errorText.value = error instanceof Error ? error.message : t('fill.denied')
  } finally {
    busy.value = false
  }
}

watch(() => [props.provider, props.fileId], () => { void load() }, { immediate: true })

const preview = computed(() => applyFormulas(
  Object.fromEntries(headers.value.map((header) => [header.key, values.value[header.key] ?? ''])),
  headers.value,
))

async function submit() {
  if (!form.value) return
  busy.value = true
  errorText.value = ''
  sent.value = false
  try {
    const file = await fillFile(props.provider, props.fileId)
    const peers: RowData[] = []
    for (const column of form.value.columns) {
      if (!column.unique) continue
      const existing = await uniquePeers(props.provider, file, form.value, column.index)
      existing.forEach((value, index) => {
        peers[index] = { ...peers[index], [column.key]: value }
      })
    }
    const checked = coerceRowData(values.value, headers.value, { rows: peers })
    errors.value = checked.errors
    if (Object.keys(checked.errors).length) return
    const row = applyFormulas(checked.row, headers.value)
    await appendSharedRow(props.provider, file, form.value, row)
    values.value = {}
    sent.value = true
  } catch (error) {
    errorText.value = error instanceof Error ? error.message : t('fill.denied')
  } finally {
    busy.value = false
  }
}
</script>

<template>
  <div class="fixed inset-0 z-40 overflow-auto bg-base-200">
    <div class="mx-auto flex min-h-full w-full max-w-lg flex-col gap-3 p-4">
      <div class="flex items-center justify-between gap-2">
        <h1 class="truncate text-lg font-semibold">{{ fileName || t('fill.title') }}</h1>
        <button type="button" class="btn btn-ghost btn-xs" @click="emit('close')">✕</button>
      </div>
      <p v-if="busy" class="text-sm opacity-70">...</p>
      <p v-if="errorText" class="text-sm text-error">{{ errorText }}</p>
      <p v-if="sent" class="text-sm text-success">{{ t('fill.sent') }}</p>
      <form v-if="form && !busy" class="flex flex-col gap-3" @submit.prevent="submit">
        <label v-for="header in headers" :key="header.key" class="flex flex-col gap-1 text-sm">
          <span :style="labelStyle(header)">{{ header.label }}</span>
          <input v-if="header.formula" class="input input-bordered input-sm" :style="fieldStyle(header)" :value="String(preview[header.key] ?? '')" disabled />
          <select v-else-if="header.options?.length" v-model="values[header.key]" class="select select-bordered select-sm" :style="fieldStyle(header)">
            <option value="">{{ t('fill.choose') }}</option>
            <option v-for="option in header.options" :key="option" :value="option">{{ option }}</option>
          </select>
          <input
            v-else
            v-model="values[header.key]"
            class="input input-bordered input-sm"
            :style="fieldStyle(header)"
            :placeholder="header.dataType === 'date' ? header.dateFormat : ''"
          />
          <span v-if="errors[header.key]" class="text-xs text-error">{{ errors[header.key] }}</span>
        </label>
        <button type="submit" class="btn btn-primary btn-sm" :disabled="busy">{{ t('fill.send') }}</button>
      </form>
    </div>
  </div>
</template>
