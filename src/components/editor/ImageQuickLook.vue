<script setup lang="ts">
import { onMounted, onUnmounted, ref, watch } from 'vue'
import { useImageQuickLook } from '@/composables/imageQuickLook'
import { useLocale } from '@/composables/useLocale'

const { src, label, close } = useImageQuickLook()
const { t } = useLocale()
const failed = ref(false)

watch(src, () => {
  failed.value = false
})

function onKey(event: KeyboardEvent) {
  if (!src.value) return
  if (event.key !== 'Escape' && event.key !== ' ') return
  event.preventDefault()
  close()
}

onMounted(() => document.addEventListener('keydown', onKey))
onUnmounted(() => document.removeEventListener('keydown', onKey))
</script>

<template>
  <Teleport to="body">
    <div
      v-if="src"
      class="fixed inset-0 z-[80] flex items-center justify-center bg-black/80 p-3 sm:p-8 backdrop-blur-sm"
      role="dialog"
      aria-modal="true"
      :aria-label="t('image.quicklook')"
      @click="close"
    >
      <figure class="flex max-h-full max-w-full flex-col items-center gap-3" @click.stop>
        <button
          type="button"
          class="absolute right-2 top-2 flex h-11 w-11 items-center justify-center rounded-full text-2xl leading-none text-white/80 hover:bg-white/10"
          :aria-label="t('common.close')"
          @click="close"
        >
          ×
        </button>
        <img
          v-if="!failed"
          :src="src"
          :alt="label || t('image.quicklook')"
          class="max-h-[min(70dvh,52rem)] max-w-[min(92vw,72rem)] rounded-md object-contain shadow-2xl"
          @error="failed = true"
        />
        <p v-else class="px-4 text-center text-sm text-white/80">{{ t('image.unavailable') }}</p>
        <figcaption class="max-w-[min(92vw,40rem)] break-all text-center text-xs text-white/70">
          {{ label }}
        </figcaption>
      </figure>
    </div>
  </Teleport>
</template>
