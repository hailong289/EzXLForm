import { ref } from 'vue'
import { forgetCloudBaseline } from '@/cloud/baseline'
import type { CloudBinding } from '@/cloud/types'

const STORAGE_KEY = 'ezxl_cloud_bindings'

function readAll(): Record<string, CloudBinding> {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY)
    if (!raw) return {}
    const parsed = JSON.parse(raw) as Record<string, CloudBinding>
    return parsed && typeof parsed === 'object' ? parsed : {}
  } catch {
    return {}
  }
}

const bindings = ref<Record<string, CloudBinding>>(readAll())
const cloudSyncing = ref(false)

function persist() {
  sessionStorage.setItem(STORAGE_KEY, JSON.stringify(bindings.value))
}

export function useCloudBindings() {
  function bindingFor(sessionId: string | null): CloudBinding | null {
    if (!sessionId) return null
    return bindings.value[sessionId] ?? null
  }

  function setBinding(sessionId: string, binding: CloudBinding) {
    bindings.value = { ...bindings.value, [sessionId]: binding }
    persist()
  }

  function clearBinding(sessionId: string) {
    forgetCloudBaseline(sessionId)
    if (!bindings.value[sessionId]) return
    const next = { ...bindings.value }
    delete next[sessionId]
    bindings.value = next
    persist()
  }

  return { bindings, bindingFor, setBinding, clearBinding, cloudSyncing }
}
