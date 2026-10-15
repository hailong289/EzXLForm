<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import type { ColumnHeader } from '@/types/excel'
import { useLocale } from '@/composables/useLocale'
import {
  DATE_FORMATS,
  DEFAULT_DATE_FORMAT,
  FORMAT_CATEGORIES,
  categoryOf,
  decimalCount,
  formatFromDraft,
  formatNumber,
  type FormatCategory,
} from '@/utils/numberFormat'
import { formatDateValue } from '@/utils/dateFormat'

const props = defineProps<{ header: ColumnHeader }>()
const emit = defineEmits<{
  close: []
  apply: [format: ReturnType<typeof formatFromDraft>]
}>()

const { t } = useLocale()
const tabs = ['number', 'align', 'font', 'border', 'fill', 'protect'] as const
const tab = ref<(typeof tabs)[number]>('number')
const category = ref<FormatCategory>('general')
const decimals = ref(2)
const thousands = ref(true)
const dateFormat = ref(DEFAULT_DATE_FORMAT)
const custom = ref('')

const SAMPLE_NUMBER = 1234.5
const SAMPLE_DATE = '2026-10-15'

watch(
  () => props.header,
  (header) => {
    tab.value = 'number'
    category.value = categoryOf(header)
    decimals.value = decimalCount(header.numFmt)
    thousands.value = header.numFmt ? header.numFmt.includes(',') : true
    dateFormat.value = header.dateFormat || DEFAULT_DATE_FORMAT
    custom.value = header.numFmt || header.dateFormat || 'General'
  },
  { immediate: true },
)

const sample = computed(() => {
  if (category.value === 'date' || (category.value === 'custom' && /yyyy|yy|dd/.test(custom.value))) {
    return formatDateValue(SAMPLE_DATE, category.value === 'date' ? dateFormat.value : custom.value)
  }
  if (category.value === 'boolean') return t('form.true.yes')
  if (category.value === 'text' || category.value === 'general') return t('format.sample.text')
  const draft = formatFromDraft({
    category: category.value,
    decimals: decimals.value,
    thousands: thousands.value,
    dateFormat: dateFormat.value,
    custom: custom.value,
  })
  return formatNumber(SAMPLE_NUMBER, draft.numFmt)
})

function save() {
  emit('apply', formatFromDraft({
    category: category.value,
    decimals: decimals.value,
    thousands: thousands.value,
    dateFormat: dateFormat.value,
    custom: custom.value,
  }))
}
</script>

<template>
  <div class="fixed inset-0 z-[70] flex items-center justify-center bg-black/40 p-3" @click.self="emit('close')">
    <div class="flex max-h-[90dvh] w-full max-w-xl flex-col border border-base-300 bg-base-100 shadow-xl" role="dialog" aria-modal="true" :aria-label="t('format.title')">
      <div class="border-b border-base-300 px-3 py-2 text-sm font-semibold">{{ t('format.title') }}</div>
      <div class="flex gap-1 overflow-x-auto border-b border-base-300 px-2 pt-2">
        <button
          v-for="item in tabs"
          :key="item"
          type="button"
          class="border border-b-0 px-3 py-1 text-xs"
          :class="tab === item ? 'border-base-300 bg-base-100 font-semibold' : 'border-transparent bg-base-200 text-base-content/70'"
          @click="tab = item"
        >
          {{ t(`format.tab.${item}`) }}
        </button>
      </div>

      <div v-if="tab !== 'number'" class="px-4 py-8 text-sm text-base-content/70">
        {{ t('format.tab.later') }}
      </div>
      <div v-else class="grid min-h-0 flex-1 gap-3 overflow-auto p-3 sm:grid-cols-[11rem_1fr]">
        <div>
          <p class="mb-1 text-[11px] text-base-content/60">{{ t('format.category') }}</p>
          <div class="max-h-64 overflow-auto border border-base-300">
            <button
              v-for="item in FORMAT_CATEGORIES"
              :key="item"
              type="button"
              class="block w-full px-2 py-1 text-left text-xs"
              :class="category === item ? 'bg-primary text-primary-content' : 'hover:bg-base-200'"
              @click="category = item"
            >
              {{ t(`format.cat.${item}`) }}
            </button>
          </div>
        </div>
        <div class="min-w-0">
          <p class="mb-1 text-[11px] text-base-content/60">{{ t('format.sample') }}</p>
          <div class="mb-3 border border-base-300 bg-base-200 px-2 py-1 text-right font-mono text-sm">{{ sample }}</div>
          <p class="mb-2 text-xs text-base-content/70">{{ t(`format.hint.${category}`) }}</p>

          <div v-if="category === 'number' || category === 'percentage'" class="space-y-2 text-xs">
            <label class="flex items-center gap-2">
              {{ t('format.decimals') }}
              <input v-model.number="decimals" type="number" min="0" max="10" class="input input-xs input-bordered w-16 text-center" />
            </label>
            <label v-if="category === 'number'" class="flex items-center gap-2">
              <input v-model="thousands" type="checkbox" class="checkbox checkbox-xs checkbox-primary" />
              {{ t('format.thousands') }}
            </label>
          </div>

          <div v-else-if="category === 'date'" class="max-h-40 overflow-auto border border-base-300">
            <button
              v-for="pattern in DATE_FORMATS"
              :key="pattern"
              type="button"
              class="block w-full px-2 py-1 text-left font-mono text-xs"
              :class="dateFormat === pattern ? 'bg-primary text-primary-content' : 'hover:bg-base-200'"
              @click="dateFormat = pattern"
            >
              {{ pattern }}
            </button>
          </div>

          <label v-else-if="category === 'custom'" class="block text-xs">
            {{ t('format.code') }}
            <input v-model="custom" type="text" class="input input-sm input-bordered mt-1 w-full font-mono" />
          </label>
        </div>
      </div>

      <div class="flex justify-end gap-2 border-t border-base-300 px-3 py-2">
        <button type="button" class="btn btn-primary btn-sm" @click="save">{{ t('format.ok') }}</button>
        <button type="button" class="btn btn-ghost btn-sm" @click="emit('close')">{{ t('format.cancel') }}</button>
      </div>
    </div>
  </div>
</template>
