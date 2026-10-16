<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import type { ColumnHeader } from '@/types/excel'
import { useLocale } from '@/composables/useLocale'
import {
  CURRENCY_SYMBOLS,
  CUSTOM_FORMATS,
  DATE_FORMATS,
  DEFAULT_DATE_FORMAT,
  FORMAT_CATEGORIES,
  FRACTION_CODES,
  TIME_FORMATS,
  categoryOf,
  formatFromDraft,
  formatNumber,
  readFormatOptions,
  type FormatCategory,
  type NegativeStyle,
} from '@/utils/numberFormat'
import { formatDateValue } from '@/utils/dateFormat'
import { columnPaint, styleFromDraft, type ColumnAlign, type ColumnBorder, type ColumnVertical } from '@/utils/columnStyle'

const props = defineProps<{ header: ColumnHeader }>()
const emit = defineEmits<{
  close: []
  apply: [format: ReturnType<typeof formatFromDraft> & { style?: ReturnType<typeof styleFromDraft> }]
}>()

const { t } = useLocale()
const tabs = ['number', 'align', 'font', 'border', 'fill', 'protect'] as const
const tab = ref<(typeof tabs)[number]>('number')
const category = ref<FormatCategory>('general')
const decimals = ref(2)
const thousands = ref(true)
const dateFormat = ref(DEFAULT_DATE_FORMAT)
const custom = ref('')
const symbol = ref('₫')
const negative = ref<NegativeStyle>('minus')
const fraction = ref('# ?/?')
const timeFormat = ref('hh:mm')
const align = ref<'' | ColumnAlign>('')
const vertical = ref<'' | ColumnVertical>('')
const wrap = ref(false)
const indent = ref(0)
const bold = ref(false)
const italic = ref(false)
const underline = ref(false)
const strike = ref(false)
const fontSize = ref(0)
const fontColor = ref('')
const fontName = ref('')
const border = ref<'' | ColumnBorder>('')
const borderColor = ref('#000000')
const fill = ref('')
const locked = ref(false)
const hidden = ref(false)
const FONT_NAMES = ['Arial', 'Calibri', 'Times New Roman', 'Courier New']
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 36]
const SWATCHES = ['#000000', '#ff0000', '#0000ff', '#008000', '#ffc000', '#7030a0', '#ffffff', '#fff2cc']

const SAMPLE_NUMBER = 1234.5
const SAMPLE_DATE = '2026-10-15'
const SAMPLE_TIME = 13.5 / 24

function draft() {
  return {
    category: category.value,
    decimals: decimals.value,
    thousands: thousands.value,
    dateFormat: dateFormat.value,
    custom: custom.value,
    symbol: symbol.value,
    negative: negative.value,
    fraction: fraction.value,
    timeFormat: timeFormat.value,
  }
}

watch(
  () => props.header,
  (header) => {
    tab.value = 'number'
    category.value = categoryOf(header)
    const options = readFormatOptions(header.numFmt)
    decimals.value = header.numFmt ? options.decimals : 2
    thousands.value = header.numFmt ? options.thousands : true
    dateFormat.value = header.dateFormat || DEFAULT_DATE_FORMAT
    custom.value = header.numFmt || header.dateFormat || 'General'
    symbol.value = options.symbol
    negative.value = options.negative
    fraction.value = options.fraction
    timeFormat.value = options.timeFormat
    const paint = header.style ?? {}
    align.value = paint.align ?? ''
    vertical.value = paint.vertical ?? ''
    wrap.value = !!paint.wrap
    indent.value = paint.indent ?? 0
    bold.value = !!paint.bold
    italic.value = !!paint.italic
    underline.value = !!paint.underline
    strike.value = !!paint.strike
    fontSize.value = paint.fontSize ?? 0
    fontColor.value = paint.fontColor ?? ''
    fontName.value = paint.fontName ?? ''
    border.value = paint.border ?? ''
    borderColor.value = paint.borderColor ?? '#000000'
    fill.value = paint.fill ?? ''
    locked.value = !!paint.locked
    hidden.value = !!paint.hidden
  },
  { immediate: true },
)

const sample = computed(() => {
  if (category.value === 'custom') {
    const code = custom.value.trim()
    if (!code || code === 'General') return formatNumber(SAMPLE_NUMBER)
    if (code === '@') return t('format.sample.text')
    if (/yyyy|yy|dd/.test(code)) return formatDateValue(SAMPLE_DATE, code)
    if (/h/i.test(code) && /m/i.test(code)) return formatNumber(SAMPLE_TIME, code)
    return formatNumber(SAMPLE_NUMBER, code)
  }
  if (category.value === 'date') return formatDateValue(SAMPLE_DATE, dateFormat.value)
  if (category.value === 'time') return formatNumber(SAMPLE_TIME, timeFormat.value)
  if (category.value === 'boolean') return t('form.true.yes')
  if (category.value === 'text' || category.value === 'general') return t('format.sample.text')
  return formatNumber(SAMPLE_NUMBER, formatFromDraft(draft()).numFmt)
})

function styleDraft() {
  return styleFromDraft({
    align: align.value,
    vertical: vertical.value,
    wrap: wrap.value,
    indent: indent.value,
    bold: bold.value,
    italic: italic.value,
    underline: underline.value,
    strike: strike.value,
    fontSize: fontSize.value,
    fontColor: fontColor.value,
    fontName: fontName.value,
    border: border.value,
    borderColor: borderColor.value,
    fill: fill.value,
    locked: locked.value,
    hidden: hidden.value,
  })
}

const painted = computed(() => columnPaint(styleDraft()))

function save() {
  emit('apply', { ...formatFromDraft(draft()), style: styleDraft() })
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

      <div v-if="tab !== 'number'" class="min-h-0 flex-1 space-y-3 overflow-auto p-3 text-xs">
        <div class="border border-base-300 px-2 py-2" :style="{ ...painted.box, ...painted.text }">{{ t('format.sample.text') }}</div>

        <div v-if="tab === 'align'" class="space-y-2">
          <label class="flex items-center gap-2">
            {{ t('format.align.horizontal') }}
            <select v-model="align" class="select select-xs select-bordered">
              <option value="">{{ t('format.align.general') }}</option>
              <option value="left">{{ t('format.align.left') }}</option>
              <option value="center">{{ t('format.align.center') }}</option>
              <option value="right">{{ t('format.align.right') }}</option>
              <option value="justify">{{ t('format.align.justify') }}</option>
            </select>
          </label>
          <label class="flex items-center gap-2">
            {{ t('format.align.vertical') }}
            <select v-model="vertical" class="select select-xs select-bordered">
              <option value="">{{ t('format.align.general') }}</option>
              <option value="top">{{ t('format.align.top') }}</option>
              <option value="middle">{{ t('format.align.middle') }}</option>
              <option value="bottom">{{ t('format.align.bottom') }}</option>
            </select>
          </label>
          <label class="flex items-center gap-2">
            <input v-model="wrap" type="checkbox" class="checkbox checkbox-xs checkbox-primary" />
            {{ t('format.align.wrap') }}
          </label>
          <label class="flex items-center gap-2">
            {{ t('format.align.indent') }}
            <input v-model.number="indent" type="number" min="0" max="15" class="input input-xs input-bordered w-16 text-center" />
          </label>
        </div>

        <div v-else-if="tab === 'font'" class="space-y-2">
          <label class="flex items-center gap-2">
            {{ t('format.font.name') }}
            <select v-model="fontName" class="select select-xs select-bordered">
              <option value="">{{ t('format.align.general') }}</option>
              <option v-for="name in FONT_NAMES" :key="name" :value="name">{{ name }}</option>
            </select>
          </label>
          <label class="flex items-center gap-2">
            {{ t('format.font.size') }}
            <select v-model.number="fontSize" class="select select-xs select-bordered">
              <option :value="0">{{ t('format.align.general') }}</option>
              <option v-for="size in FONT_SIZES" :key="size" :value="size">{{ size }}</option>
            </select>
          </label>
          <label class="flex items-center gap-2"><input v-model="bold" type="checkbox" class="checkbox checkbox-xs" /> {{ t('format.font.bold') }}</label>
          <label class="flex items-center gap-2"><input v-model="italic" type="checkbox" class="checkbox checkbox-xs" /> {{ t('format.font.italic') }}</label>
          <label class="flex items-center gap-2"><input v-model="underline" type="checkbox" class="checkbox checkbox-xs" /> {{ t('format.font.underline') }}</label>
          <label class="flex items-center gap-2"><input v-model="strike" type="checkbox" class="checkbox checkbox-xs" /> {{ t('format.font.strike') }}</label>
          <div class="flex flex-wrap items-center gap-1">
            <span>{{ t('format.font.color') }}</span>
            <button type="button" class="btn btn-xs" @click="fontColor = ''">{{ t('format.fill.none') }}</button>
            <button
              v-for="color in SWATCHES"
              :key="color"
              type="button"
              class="h-5 w-5 border border-base-300"
              :style="{ backgroundColor: color }"
              :aria-label="color"
              @click="fontColor = color"
            />
            <input v-model="fontColor" type="color" class="h-6 w-8" />
          </div>
        </div>

        <div v-else-if="tab === 'border'" class="space-y-2">
          <label class="flex items-center gap-2">
            <input v-model="border" type="radio" class="radio radio-xs" value="" /> {{ t('format.border.none') }}
          </label>
          <label class="flex items-center gap-2">
            <input v-model="border" type="radio" class="radio radio-xs" value="box" /> {{ t('format.border.box') }}
          </label>
          <label class="flex items-center gap-2">
            <input v-model="border" type="radio" class="radio radio-xs" value="bottom" /> {{ t('format.border.bottom') }}
          </label>
          <div class="flex flex-wrap items-center gap-1">
            <span>{{ t('format.border.color') }}</span>
            <button
              v-for="color in SWATCHES"
              :key="`border-${color}`"
              type="button"
              class="h-5 w-5 border border-base-300"
              :style="{ backgroundColor: color }"
              :aria-label="color"
              @click="borderColor = color"
            />
            <input v-model="borderColor" type="color" class="h-6 w-8" />
          </div>
        </div>

        <div v-else-if="tab === 'fill'" class="flex flex-wrap items-center gap-1">
          <button type="button" class="btn btn-xs" @click="fill = ''">{{ t('format.fill.none') }}</button>
          <button
            v-for="color in SWATCHES"
            :key="`fill-${color}`"
            type="button"
            class="h-6 w-6 border border-base-300"
            :class="fill === color ? 'outline outline-2 outline-primary' : ''"
            :style="{ backgroundColor: color }"
            :aria-label="color"
            @click="fill = color"
          />
          <input v-model="fill" type="color" class="h-6 w-8" />
        </div>

        <div v-else class="space-y-2">
          <p class="text-base-content/70">{{ t('format.protect.hint') }}</p>
          <label class="flex items-center gap-2"><input v-model="locked" type="checkbox" class="checkbox checkbox-xs checkbox-primary" /> {{ t('format.protect.locked') }}</label>
          <label class="flex items-center gap-2"><input v-model="hidden" type="checkbox" class="checkbox checkbox-xs checkbox-primary" /> {{ t('format.protect.hidden') }}</label>
        </div>
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

          <div v-if="category === 'number' || category === 'percentage' || category === 'currency' || category === 'accounting' || category === 'scientific'" class="space-y-2 text-xs">
            <label class="flex items-center gap-2">
              {{ t('format.decimals') }}
              <input v-model.number="decimals" type="number" min="0" max="10" class="input input-xs input-bordered w-16 text-center" />
            </label>
            <label v-if="category === 'number'" class="flex items-center gap-2">
              <input v-model="thousands" type="checkbox" class="checkbox checkbox-xs checkbox-primary" />
              {{ t('format.thousands') }}
            </label>
            <label v-if="category === 'currency' || category === 'accounting'" class="flex items-center gap-2">
              {{ t('format.symbol') }}
              <select v-model="symbol" class="select select-xs select-bordered">
                <option v-for="item in CURRENCY_SYMBOLS" :key="item" :value="item">{{ item }}</option>
              </select>
            </label>
            <label v-if="category === 'number' || category === 'currency'" class="flex items-center gap-2">
              {{ t('format.negative') }}
              <select v-model="negative" class="select select-xs select-bordered">
                <option value="minus">{{ t('format.negative.minus') }}</option>
                <option value="parens">{{ t('format.negative.parens') }}</option>
              </select>
            </label>
          </div>

          <div v-else-if="category === 'fraction'" class="max-h-40 overflow-auto border border-base-300">
            <button
              v-for="code in FRACTION_CODES"
              :key="code"
              type="button"
              class="block w-full px-2 py-1 text-left font-mono text-xs"
              :class="fraction === code ? 'bg-primary text-primary-content' : 'hover:bg-base-200'"
              @click="fraction = code"
            >
              {{ formatNumber(SAMPLE_NUMBER, code) }}
            </button>
          </div>

          <div v-else-if="category === 'time'" class="max-h-40 overflow-auto border border-base-300">
            <button
              v-for="pattern in TIME_FORMATS"
              :key="pattern"
              type="button"
              class="block w-full px-2 py-1 text-left font-mono text-xs"
              :class="timeFormat === pattern ? 'bg-primary text-primary-content' : 'hover:bg-base-200'"
              @click="timeFormat = pattern"
            >
              {{ formatNumber(SAMPLE_TIME, pattern) }}
            </button>
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

          <div v-else-if="category === 'custom'" class="text-xs">
            <label class="block">
              {{ t('format.code') }}
              <input v-model="custom" type="text" class="input input-sm input-bordered mt-1 w-full font-mono" />
            </label>
            <div class="mt-2 max-h-40 overflow-auto border border-base-300">
              <button
                v-for="code in CUSTOM_FORMATS"
                :key="code"
                type="button"
                class="block w-full px-2 py-1 text-left font-mono text-xs"
                :class="custom === code ? 'bg-primary text-primary-content' : 'hover:bg-base-200'"
                @click="custom = code"
              >
                {{ code }}
              </button>
            </div>
          </div>
        </div>
      </div>

      <div class="flex justify-end gap-2 border-t border-base-300 px-3 py-2">
        <button type="button" class="btn btn-primary btn-sm" @click="save">{{ t('format.ok') }}</button>
        <button type="button" class="btn btn-ghost btn-sm" @click="emit('close')">{{ t('format.cancel') }}</button>
      </div>
    </div>
  </div>
</template>
