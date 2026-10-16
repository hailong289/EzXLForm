<script setup lang="ts">
import { ref, watch } from 'vue'
import type { ColumnHeader } from '@/types/excel'
import { useLocale } from '@/composables/useLocale'

const props = defineProps<{ header: ColumnHeader }>()
const emit = defineEmits<{
  close: []
  apply: [rules: { required: boolean; unique: boolean; pattern: string; options: string[]; formula: string }]
}>()

const { t } = useLocale()
const required = ref(false)
const unique = ref(false)
const pattern = ref('')
const options = ref('')
const formula = ref('')

watch(
  () => props.header,
  (header) => {
    required.value = !!header.required
    unique.value = !!header.unique
    pattern.value = header.pattern ?? ''
    options.value = header.options?.join('\n') ?? ''
    formula.value = header.formula ?? ''
  },
  { immediate: true },
)

function save() {
  emit('apply', {
    required: required.value,
    unique: unique.value,
    pattern: pattern.value.trim(),
    options: options.value.split('\n').map((item) => item.trim()).filter(Boolean),
    formula: formula.value.trim(),
  })
}
</script>

<template>
  <div class="fixed inset-0 z-[70] flex items-center justify-center bg-black/40 p-3" @click.self="emit('close')">
    <div class="w-full max-w-md border border-base-300 bg-base-100 shadow-xl" role="dialog" aria-modal="true" :aria-label="t('rules.title')">
      <div class="border-b border-base-300 px-3 py-2 text-sm font-semibold">{{ t('rules.title') }}</div>
      <div class="space-y-3 p-3 text-xs">
        <p class="text-base-content/70">{{ props.header.label }}</p>
        <label class="flex items-center gap-2">
          <input v-model="required" type="checkbox" class="checkbox checkbox-xs checkbox-primary" />
          {{ t('rules.required') }}
        </label>
        <label class="flex items-center gap-2">
          <input v-model="unique" type="checkbox" class="checkbox checkbox-xs checkbox-primary" />
          {{ t('rules.unique') }}
        </label>
        <label class="block">
          {{ t('rules.pattern') }}
          <input v-model="pattern" type="text" class="input input-sm input-bordered mt-1 w-full font-mono" placeholder="^A\\d+$" />
        </label>
        <label class="block">
          {{ t('rules.options') }}
          <textarea v-model="options" class="textarea textarea-bordered mt-1 w-full font-mono text-xs" rows="4" />
        </label>
        <label class="block">
          {{ t('rules.formula') }}
          <input v-model="formula" type="text" class="input input-sm input-bordered mt-1 w-full font-mono" placeholder="qty * price" />
        </label>
      </div>
      <div class="flex justify-end gap-2 border-t border-base-300 px-3 py-2">
        <button type="button" class="btn btn-primary btn-sm" @click="save">{{ t('format.ok') }}</button>
        <button type="button" class="btn btn-ghost btn-sm" @click="emit('close')">{{ t('format.cancel') }}</button>
      </div>
    </div>
  </div>
</template>
