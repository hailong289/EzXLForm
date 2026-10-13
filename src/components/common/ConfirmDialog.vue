<script setup lang="ts">
import { computed } from 'vue'
import { AlertTriangle, Trash2, Info } from 'lucide-vue-next'
import { useLocale } from '@/composables/useLocale'

const { t } = useLocale()

const props = withDefaults(
  defineProps<{
    modelValue: boolean
    title?: string
    message: string
    confirmText?: string
    cancelText?: string
    type?: 'error' | 'warning' | 'info'
  }>(),
  {
    type: 'warning',
  },
)

const titleText = computed(() => props.title ?? t('dialog.confirm'))
const confirmLabel = computed(() => props.confirmText ?? t('dialog.ok'))
const cancelLabel = computed(() => props.cancelText ?? t('common.cancel'))

const emit = defineEmits<{
  'update:modelValue': [value: boolean]
  confirm: []
  cancel: []
}>()

function onCancel() {
  emit('update:modelValue', false)
  emit('cancel')
}

function onConfirm() {
  emit('update:modelValue', false)
  emit('confirm')
}
</script>

<template>
  <dialog v-if="props.modelValue" class="modal modal-open">
    <div class="modal-box max-w-sm">
      <div class="flex items-center gap-3 mb-3">
        <div
          class="p-2 rounded-full"
          :class="{
            'bg-error/20 text-error': props.type === 'error',
            'bg-warning/20 text-warning': props.type === 'warning',
            'bg-info/20 text-info': props.type === 'info',
          }"
        >
          <Trash2 v-if="props.type === 'error'" class="w-5 h-5" />
          <AlertTriangle v-else-if="props.type === 'warning'" class="w-5 h-5" />
          <Info v-else class="w-5 h-5" />
        </div>
        <h3 class="font-bold text-base">{{ titleText }}</h3>
      </div>

      <p class="text-sm opacity-80 leading-relaxed">{{ props.message }}</p>

      <div class="modal-action mt-5">
        <button class="btn btn-sm btn-ghost" @click="onCancel">
          {{ cancelLabel }}
        </button>
        <button
          class="btn btn-sm"
          :class="{
            'btn-error': props.type === 'error',
            'btn-warning': props.type === 'warning',
            'btn-info': props.type === 'info',
          }"
          @click="onConfirm"
        >
          {{ confirmLabel }}
        </button>
      </div>
    </div>
    <div class="modal-backdrop bg-black/40" @click="onCancel" />
  </dialog>
</template>
