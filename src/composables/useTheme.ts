import { ref, computed } from 'vue'

export interface ThemeItem {
  name: string
  label: string
  isDark: boolean
  previewColor: string
}

export const AVAILABLE_THEMES: ThemeItem[] = [
  { name: 'dark', label: 'Dark', isDark: true, previewColor: '#1d232a' },
  { name: 'light', label: 'Light', isDark: false, previewColor: '#ffffff' },
  { name: 'excel-custom', label: 'Excel Green', isDark: false, previewColor: '#217346' },
  { name: 'emerald', label: 'Emerald', isDark: false, previewColor: '#66cc8a' },
  { name: 'corporate', label: 'Corporate', isDark: false, previewColor: '#4b6bfb' },
  { name: 'synthwave', label: 'Synthwave', isDark: true, previewColor: '#e779c1' },
  { name: 'retro', label: 'Retro', isDark: false, previewColor: '#ef9995' },
  { name: 'cyberpunk', label: 'Cyberpunk', isDark: false, previewColor: '#ff7598' },
  { name: 'valentine', label: 'Valentine', isDark: false, previewColor: '#e96d7b' },
  { name: 'halloween', label: 'Halloween', isDark: true, previewColor: '#f28c18' },
  { name: 'garden', label: 'Garden', isDark: false, previewColor: '#5c7f67' },
  { name: 'forest', label: 'Forest', isDark: true, previewColor: '#1eb854' },
  { name: 'aqua', label: 'Aqua', isDark: true, previewColor: '#09ecf3' },
  { name: 'lofi', label: 'Lo-Fi', isDark: false, previewColor: '#0d0d0d' },
  { name: 'pastel', label: 'Pastel', isDark: false, previewColor: '#d1c1d7' },
  { name: 'fantasy', label: 'Fantasy', isDark: false, previewColor: '#6e0b75' },
  { name: 'wireframe', label: 'Wireframe', isDark: false, previewColor: '#b8b8b8' },
  { name: 'black', label: 'Black', isDark: true, previewColor: '#000000' },
  { name: 'luxury', label: 'Luxury', isDark: true, previewColor: '#ffffff' },
  { name: 'dracula', label: 'Dracula', isDark: true, previewColor: '#ff79c6' },
  { name: 'cmyk', label: 'CMYK', isDark: false, previewColor: '#45AEEE' },
  { name: 'autumn', label: 'Autumn', isDark: false, previewColor: '#8C0327' },
  { name: 'business', label: 'Business', isDark: true, previewColor: '#1C4E80' },
  { name: 'acid', label: 'Acid', isDark: false, previewColor: '#FF00F4' },
  { name: 'lemonade', label: 'Lemonade', isDark: false, previewColor: '#519903' },
  { name: 'night', label: 'Night', isDark: true, previewColor: '#38bdf8' },
  { name: 'coffee', label: 'Coffee', isDark: true, previewColor: '#DB924B' },
  { name: 'winter', label: 'Winter', isDark: false, previewColor: '#047AFF' },
  { name: 'dim', label: 'Dim', isDark: true, previewColor: '#9FE88D' },
  { name: 'nord', label: 'Nord', isDark: false, previewColor: '#5E81AC' },
  { name: 'sunset', label: 'Sunset', isDark: true, previewColor: '#FF865B' },
  { name: 'cupcake', label: 'Cupcake', isDark: false, previewColor: '#65c3c8' },
  { name: 'bumblebee', label: 'Bumblebee', isDark: false, previewColor: '#e0a82e' },
]

const THEME_STORAGE_KEY = 'excel_editor_theme'
const THEME_CHOSEN_KEY = 'excel_editor_theme_chosen'
export const SYSTEM_THEME = 'system'
const COLORS_STORAGE_KEY = 'excel_editor_custom_colors'
const FONT_STORAGE_KEY = 'excel_editor_font_size'

export const FONT_SIZE_OPTIONS = [
  { px: 14, label: 'common.small' },
  { px: 16, label: 'common.medium' },
  { px: 18, label: 'common.large' },
  { px: 20, label: 'common.extra.large' },
]

export const FONT_SIZE_MIN = 12
export const FONT_SIZE_MAX = 32

const currentTheme = ref<string>(SYSTEM_THEME)
const prefersDark = ref(false)

function systemPrefersDark(): boolean {
  return window.matchMedia('(prefers-color-scheme: dark)').matches
}

function resolvedThemeName(themeName: string): string {
  if (themeName !== SYSTEM_THEME) return themeName
  return prefersDark.value ? 'dark' : 'light'
}
const fontSize = ref(16)
const customColors = ref<{ primary: string; secondary: string; accent: string }>({
  primary: '#217346',
  secondary: '#1D6035',
  accent: '#185C37',
})

// Chuyển hex (#RRGGBB) sang HSL format: "h s% l%"
function hexToHsl(hex: string): string {
  let c = hex.replace('#', '')
  if (c.length === 3) {
    c = c.split('').map((char) => char + char).join('')
  }
  const r = parseInt(c.substring(0, 2), 16) / 255
  const g = parseInt(c.substring(2, 4), 16) / 255
  const b = parseInt(c.substring(4, 6), 16) / 255

  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let h = 0
  let s = 0
  const l = (max + min) / 2

  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0)
        break
      case g:
        h = (b - r) / d + 2
        break
      case b:
        h = (r - g) / d + 4
        break
    }
    h /= 6
  }

  const hDeg = Math.round(h * 360)
  const sPct = Math.round(s * 100)
  const lPct = Math.round(l * 100)

  return `${hDeg} ${sPct}% ${lPct}%`
}

export function useTheme() {
  const isDarkTheme = computed(() => {
    if (currentTheme.value === SYSTEM_THEME) return prefersDark.value
    const found = AVAILABLE_THEMES.find((t) => t.name === currentTheme.value)
    return found ? found.isDark : prefersDark.value
  })

  const themeLabel = computed(() => {
    if (currentTheme.value === SYSTEM_THEME) return 'Hệ thống'
    return AVAILABLE_THEMES.find((t) => t.name === currentTheme.value)?.label ?? currentTheme.value
  })

  function applyTheme(themeName: string) {
    prefersDark.value = systemPrefersDark()
    currentTheme.value = themeName
    document.documentElement.setAttribute('data-theme', resolvedThemeName(themeName))
    try {
      localStorage.setItem(THEME_STORAGE_KEY, themeName)
    } catch {
      // ignore storage error
    }
  }

  function applyCustomColors(colors: { primary?: string; secondary?: string; accent?: string }) {
    if (colors.primary) customColors.value.primary = colors.primary
    if (colors.secondary) customColors.value.secondary = colors.secondary
    if (colors.accent) customColors.value.accent = colors.accent

    const root = document.documentElement
    if (colors.primary) {
      const hsl = hexToHsl(colors.primary)
      root.style.setProperty('--p', hsl)
    }
    if (colors.secondary) {
      const hsl = hexToHsl(colors.secondary)
      root.style.setProperty('--s', hsl)
    }
    if (colors.accent) {
      const hsl = hexToHsl(colors.accent)
      root.style.setProperty('--a', hsl)
    }

    try {
      localStorage.setItem(COLORS_STORAGE_KEY, JSON.stringify(customColors.value))
    } catch {
      // ignore
    }
  }

  function resetCustomColors() {
    customColors.value = {
      primary: '#217346',
      secondary: '#1D6035',
      accent: '#185C37',
    }
    const root = document.documentElement
    root.style.removeProperty('--p')
    root.style.removeProperty('--s')
    root.style.removeProperty('--a')
    try {
      localStorage.removeItem(COLORS_STORAGE_KEY)
    } catch {
      // ignore
    }
  }

  function applyFontSize(px: number) {
    const allowed = Math.max(FONT_SIZE_MIN, Math.min(FONT_SIZE_MAX, Math.round(px)))
    fontSize.value = allowed
    document.documentElement.style.fontSize = `${allowed}px`
    try {
      localStorage.setItem(FONT_STORAGE_KEY, String(allowed))
    } catch {
      // ignore storage error
    }
  }

  function stepFontSize(delta: number) {
    applyFontSize(fontSize.value + delta)
  }

  function chooseTheme(themeName: string) {
    try {
      localStorage.setItem(THEME_CHOSEN_KEY, '1')
    } catch {
      // ignore storage error
    }
    applyTheme(themeName)
  }

  function loadSavedTheme() {
    try {
      const chosen = localStorage.getItem(THEME_CHOSEN_KEY) === '1'
      const savedTheme = localStorage.getItem(THEME_STORAGE_KEY)
      applyTheme(chosen && savedTheme ? savedTheme : SYSTEM_THEME)

      const savedColors = localStorage.getItem(COLORS_STORAGE_KEY)
      if (savedColors) {
        const parsed = JSON.parse(savedColors)
        applyCustomColors(parsed)
      }

      const savedFont = Number(localStorage.getItem(FONT_STORAGE_KEY))
      applyFontSize(savedFont || 16)
    } catch {
      applyTheme(SYSTEM_THEME)
    }
  }

  return {
    currentTheme,
    themeLabel,
    fontSize,
    fontSizeOptions: FONT_SIZE_OPTIONS,
    customColors,
    availableThemes: AVAILABLE_THEMES,
    isDarkTheme,
    applyTheme,
    chooseTheme,
    applyFontSize,
    stepFontSize,
    fontSizeMin: FONT_SIZE_MIN,
    fontSizeMax: FONT_SIZE_MAX,
    applyCustomColors,
    resetCustomColors,
    loadSavedTheme,
    hexToHsl,
  }
}

if (typeof window !== 'undefined') {
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
    prefersDark.value = systemPrefersDark()
    if (currentTheme.value === SYSTEM_THEME) {
      document.documentElement.setAttribute('data-theme', resolvedThemeName(SYSTEM_THEME))
    }
  })
}
