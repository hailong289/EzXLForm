import { computed, ref } from 'vue'
import vi from '@/locales/vi.json'
import en from '@/locales/en.json'

export type AppLocale = 'vi' | 'en'

const STORAGE_KEY = 'excel_editor_locale'

const messages: Record<AppLocale, Record<string, string>> = {
  vi,
  en,
}

function readInitial(): AppLocale {
  try {
    const saved = localStorage.getItem(STORAGE_KEY)
    if (saved === 'en' || saved === 'vi') return saved
  } catch {
    /* private mode */
  }
  return 'vi'
}

const locale = ref<AppLocale>(readInitial())

function applyDocumentLang(next: AppLocale) {
  if (typeof document === 'undefined') return
  document.documentElement.lang = next === 'en' ? 'en' : 'vi'
}

applyDocumentLang(locale.value)

export function useLocale() {
  const localeLabel = computed(() => (locale.value === 'en' ? 'EN' : 'VI'))

  function setLocale(next: AppLocale) {
    locale.value = next
    try {
      localStorage.setItem(STORAGE_KEY, next)
    } catch {
      /* ignore */
    }
    applyDocumentLang(next)
  }

  function t(key: string, params?: Record<string, string | number>) {
    let text = messages[locale.value][key] ?? ''
    if (params) {
      for (const [name, value] of Object.entries(params)) {
        text = text.split(`{${name}}`).join(String(value))
      }
    }
    return text
  }

  return { locale, localeLabel, setLocale, t }
}
