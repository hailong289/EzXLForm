<script setup lang="ts">
import { Check, Languages } from 'lucide-vue-next'
import { useLocale, type AppLocale } from '@/composables/useLocale'

const { locale, localeLabel, setLocale, t } = useLocale()

function pick(next: AppLocale) {
  setLocale(next)
  ;(document.activeElement as HTMLElement | null)?.blur()
}
</script>

<template>
  <div class="dropdown dropdown-bottom dropdown-end">
    <button type="button" tabindex="0" class="btn btn-ghost btn-xs gap-1" :title="t('locale.language')">
      <Languages class="w-3.5 h-3.5" />
      <span class="hidden lg:inline text-xs font-normal">{{ localeLabel }}</span>
    </button>
    <ul
      tabindex="0"
      class="dropdown-content menu z-50 mt-1 w-40 rounded-box border border-base-300 bg-base-100 p-1 shadow-lg"
    >
      <li>
        <button type="button" class="text-xs" @click="pick('vi')">
          <Check class="h-3.5 w-3.5" :class="locale === 'vi' ? '' : 'opacity-0'" />
          Tiếng Việt
        </button>
      </li>
      <li>
        <button type="button" class="text-xs" @click="pick('en')">
          <Check class="h-3.5 w-3.5" :class="locale === 'en' ? '' : 'opacity-0'" />
          English
        </button>
      </li>
    </ul>
  </div>
</template>
