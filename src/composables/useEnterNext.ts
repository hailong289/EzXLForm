import { ref } from 'vue'

const STORAGE_KEY = 'excel_editor_enter_next'

function readInitial(): boolean {
  try {
    return localStorage.getItem(STORAGE_KEY) === '1'
  } catch {
    return false
  }
}

const enterNext = ref(readInitial())

export function useEnterNext() {
  function setEnterNext(next: boolean) {
    enterNext.value = next
    try {
      localStorage.setItem(STORAGE_KEY, next ? '1' : '0')
    } catch {
      /* private mode */
    }
  }

  return { enterNext, setEnterNext }
}
