import { ref } from 'vue'

const src = ref<string | null>(null)
const label = ref('')

export function useImageQuickLook() {
  function open(next: string, name = '') {
    src.value = next
    label.value = name
  }

  function close() {
    src.value = null
    label.value = ''
  }

  return { src, label, open, close }
}
