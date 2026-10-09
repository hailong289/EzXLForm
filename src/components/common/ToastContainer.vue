<script setup lang="ts">
import { CheckCircle2, AlertCircle, AlertTriangle, Info, X } from 'lucide-vue-next'
import { useAppConfigStore } from '@/stores/appConfigStore'

const appConfigStore = useAppConfigStore()
</script>

<template>
  <div class="toast toast-top toast-end z-50 pointer-events-none p-4">
    <div
      v-for="toast in appConfigStore.toasts"
      :key="toast.id"
      class="alert shadow-lg py-2 px-3 text-xs flex items-center gap-2 pointer-events-auto transition-all duration-300"
      :class="{
        'alert-success': toast.type === 'success',
        'alert-error': toast.type === 'error',
        'alert-warning': toast.type === 'warning',
        'alert-info': toast.type === 'info',
      }"
    >
      <CheckCircle2 v-if="toast.type === 'success'" class="w-4 h-4 flex-shrink-0" />
      <AlertCircle v-else-if="toast.type === 'error'" class="w-4 h-4 flex-shrink-0" />
      <AlertTriangle v-else-if="toast.type === 'warning'" class="w-4 h-4 flex-shrink-0" />
      <Info v-else class="w-4 h-4 flex-shrink-0" />

      <span class="flex-1 font-medium">{{ toast.message }}</span>

      <button
        class="btn btn-ghost btn-xs btn-circle ml-1"
        @click="appConfigStore.removeToast(toast.id)"
      >
        <X class="w-3.5 h-3.5" />
      </button>
    </div>
  </div>
</template>
