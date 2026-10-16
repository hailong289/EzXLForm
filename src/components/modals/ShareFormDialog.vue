<script setup lang="ts">
import { ref, watch } from 'vue'
import { assertCloudEditor, readSharedForm, writeSharedForm } from '@/cloud/formCloud'
import { fillHash, formSnapshot } from '@/cloud/formShare'
import type { CloudBinding } from '@/cloud/types'
import { useLocale } from '@/composables/useLocale'
import type { ColumnHeader } from '@/types/excel'
import type { FormPreset } from '@/utils/formPresets'

const props = defineProps<{
  modelValue: boolean
  binding: CloudBinding | null
  presets: FormPreset[]
  headers: ColumnHeader[]
  sheetName: string
  headerRow: number
}>()
const emit = defineEmits<{ 'update:modelValue': [value: boolean] }>()
const { t } = useLocale()

const presetId = ref('')
const link = ref('')
const busy = ref(false)
const errorText = ref('')
const replacing = ref(false)

watch(() => props.modelValue, (open) => {
  if (!open) return
  presetId.value = props.presets[0]?.id ?? ''
  link.value = ''
  errorText.value = ''
  replacing.value = false
})

function close() {
  emit('update:modelValue', false)
}

async function publish(force: boolean) {
  if (!props.binding) return
  const preset = props.presets.find((item) => item.id === presetId.value)
  if (!preset) {
    errorText.value = t('share.need.preset')
    return
  }
  const form = formSnapshot(props.sheetName, props.headerRow, props.headers, preset.visibleKeys)
  if (form.columns.length === 0) {
    errorText.value = t('share.need.preset')
    return
  }
  busy.value = true
  errorText.value = ''
  try {
    await assertCloudEditor(props.binding.provider, props.binding.file)
    if (!force && await readSharedForm(props.binding.provider, props.binding.file)) {
      replacing.value = true
      return
    }
    await writeSharedForm(props.binding.provider, props.binding.file, form)
    replacing.value = false
    link.value = `${location.origin}${location.pathname}${fillHash(props.binding.provider, props.binding.file.id)}`
  } catch (error) {
    errorText.value = error instanceof Error ? error.message : t('share.failed')
  } finally {
    busy.value = false
  }
}

function selectLink(event: Event) {
  if (event.target instanceof HTMLInputElement) event.target.select()
}

async function copyLink() {
  if (!link.value) return
  await navigator.clipboard.writeText(link.value)
}
</script>

<template>
  <div v-if="modelValue" class="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" @click.self="close">
    <div class="w-full max-w-md rounded-box border border-base-300 bg-base-100 p-4 shadow-xl">
      <div class="mb-3 flex items-center justify-between gap-2">
        <h2 class="truncate text-base font-semibold">{{ t('share.title') }}</h2>
        <button type="button" class="btn btn-ghost btn-xs" @click="close">✕</button>
      </div>
      <p v-if="presets.length === 0" class="text-sm opacity-70">{{ t('share.none') }}</p>
      <template v-else>
        <label class="mb-3 flex flex-col gap-1 text-sm">
          {{ t('share.preset') }}
          <select v-model="presetId" class="select select-bordered select-sm">
            <option v-for="preset in presets" :key="preset.id" :value="preset.id">{{ preset.name }}</option>
          </select>
        </label>
        <p v-if="replacing" class="mb-2 text-sm">{{ t('share.replace') }}</p>
        <p v-if="errorText" class="mb-2 text-sm text-error">{{ errorText }}</p>
        <button type="button" class="btn btn-primary btn-sm" :disabled="busy" @click="publish(replacing)">
          {{ replacing ? t('share.overwrite') : t('share.create') }}
        </button>
        <label v-if="link" class="mt-3 flex gap-2">
          <input :value="link" readonly class="input input-bordered input-sm min-w-0 flex-1" @focus="selectLink" />
          <button type="button" class="btn btn-sm" @click="copyLink">{{ t('share.copy') }}</button>
        </label>
      </template>
    </div>
  </div>
</template>
