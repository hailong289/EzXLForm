<script setup lang="ts">
import { History } from 'lucide-vue-next'
import { useHistory } from '@/composables/useHistory'
import { useLocale } from '@/composables/useLocale'

defineProps<{ block?: boolean; compact?: boolean }>()

const { activity } = useHistory()
const { t } = useLocale()

function clock(at: number): string {
  return new Date(at).toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit', second: '2-digit' })
}

function line(kind: 'do' | 'undo' | 'redo', text: string): string {
  if (kind === 'undo') return t('nav.activity.undo', { text })
  if (kind === 'redo') return t('nav.activity.redo', { text })
  return text
}
</script>

<template>
  <details v-if="block" class="rounded-lg border border-base-300 bg-base-100">
    <summary class="flex h-12 cursor-pointer list-none items-center gap-2 px-4 text-sm">
      <History class="h-4 w-4" />
      {{ t('nav.activity') }}
      <span v-if="activity.length" class="badge badge-sm badge-primary">{{ activity.length }}</span>
    </summary>
    <div class="max-h-60 overflow-y-auto border-t border-base-300 p-2">
      <p v-if="activity.length === 0" class="px-2 py-3 text-xs opacity-60">{{ t('nav.no.activity') }}</p>
      <ul v-else class="space-y-1">
        <li v-for="entry in activity" :key="entry.id" class="rounded-lg px-2 py-1.5">
          <p class="text-sm leading-snug">{{ line(entry.kind, entry.text) }}</p>
          <p v-for="(detail, index) in entry.details" :key="index" class="mt-0.5 break-words text-xs opacity-70">{{ detail }}</p>
          <p class="mt-0.5 font-mono text-[10px] opacity-50">{{ clock(entry.at) }}</p>
        </li>
      </ul>
    </div>
  </details>

  <div v-else class="dropdown dropdown-bottom dropdown-end">
    <button type="button" tabindex="0" class="btn btn-ghost btn-xs gap-1" :title="t('nav.activity')">
      <History class="h-3.5 w-3.5" />
      <span v-if="!compact" class="hidden md:inline">{{ t('nav.activity.menu') }}</span>
      <span v-if="activity.length" class="badge badge-xs badge-primary">{{ activity.length }}</span>
    </button>
    <div
      tabindex="0"
      class="dropdown-content z-50 mt-1 max-h-80 w-80 max-w-[90vw] overflow-y-auto rounded-box border border-base-300 bg-base-100 p-2 shadow-lg"
    >
      <p class="mb-1 px-1 text-xs font-semibold">{{ t('nav.activity') }}</p>
      <p v-if="activity.length === 0" class="px-2 py-3 text-xs opacity-60">{{ t('nav.no.activity') }}</p>
      <ul v-else class="space-y-1">
        <li v-for="entry in activity" :key="entry.id" class="rounded-lg px-2 py-1.5">
          <p class="text-sm leading-snug">{{ line(entry.kind, entry.text) }}</p>
          <p v-for="(detail, index) in entry.details" :key="index" class="mt-0.5 break-words text-xs opacity-70">{{ detail }}</p>
          <p class="mt-0.5 font-mono text-[10px] opacity-50">{{ clock(entry.at) }}</p>
        </li>
      </ul>
    </div>
  </div>
</template>
