<script setup lang="ts">
import { ref } from 'vue'
import { X } from 'lucide-vue-next'
import { useWorkbookStore } from '@/stores/workbookStore'
import { useAppConfigStore } from '@/stores/appConfigStore'
import { useSqliteStorage } from '@/composables/useSqliteStorage'
import { useLocale } from '@/composables/useLocale'
import { DISCORD_FILE_LIMIT, SUPPORT_FILE_MAIL, supportPayload } from '@/utils/supportRequest'

const open = defineModel<boolean>({ required: true })
const workbookStore = useWorkbookStore()
const appConfigStore = useAppConfigStore()
const { saveSessionToSqlite, sqliteSnapshot, downloadSqliteDatabase, activeSessionId } = useSqliteStorage()
const { t } = useLocale()

const kind = ref<'issue' | 'feature'>('issue')
const description = ref('')
const email = ref('')
const sending = ref(false)
const errorText = ref('')
const mailNotice = ref(false)

async function submit() {
  errorText.value = ''
  const nextTitle = t(kind.value === 'feature' ? 'support.kind.feature' : 'support.kind.issue')
  const nextDescription = description.value.trim()
  const nextEmail = email.value.trim()
  if (!nextDescription || !nextEmail.includes('@')) {
    errorText.value = t('support.invalid')
    return
  }
  sending.value = true
  try {
    if (workbookStore.workbook && activeSessionId.value) {
      await saveSessionToSqlite(workbookStore.workbook, activeSessionId.value)
    }
    const blob = await sqliteSnapshot()
    const base = (workbookStore.workbook?.meta.fileName || 'session').replace(/\.[^.]+$/, '')
    const fileName = blob ? `${base}.sqlite` : null
    const built = supportPayload({
      title: nextTitle,
      description: nextDescription,
      email: nextEmail,
      fileName,
      fileBytes: blob?.size ?? 0,
    })
    const tooLarge = Boolean(blob && fileName && blob.size > DISCORD_FILE_LIMIT)
    if (tooLarge && fileName) await downloadSqliteDatabase(fileName)
    const body = new FormData()
    body.append('payload_json', built.payload)
    if (blob && built.attach && fileName) body.append('files[0]', blob, fileName)
    const response = await fetch('/api/support', { method: 'POST', body })
    if (!response.ok) throw new Error((await response.text()) || t('support.failed'))
    appConfigStore.addToast({ type: 'success', message: t('support.sent') })
    kind.value = 'issue'
    description.value = ''
    email.value = ''
    mailNotice.value = tooLarge
    if (!tooLarge) open.value = false
  } catch (err) {
    errorText.value = err instanceof Error && err.message ? err.message : t('support.failed')
  } finally {
    sending.value = false
  }
}
</script>

<template>
  <dialog v-if="open" class="modal modal-open">
    <form class="modal-box max-w-lg border border-base-300" @submit.prevent="submit">
      <div class="mb-3 flex items-center justify-between">
        <h3 class="text-base font-bold">{{ t('nav.support') }}</h3>
        <button type="button" class="btn btn-ghost btn-xs btn-circle" :title="t('common.close')" @click="open = false">
          <X class="h-4 w-4" />
        </button>
      </div>
      <label class="form-control mb-2">
        <span class="label-text mb-1 text-xs">{{ t('support.title') }}</span>
        <select v-model="kind" class="select select-bordered select-sm">
          <option value="issue">{{ t('support.kind.issue') }}</option>
          <option value="feature">{{ t('support.kind.feature') }}</option>
        </select>
      </label>
      <label class="form-control mb-2">
        <span class="label-text mb-1 text-xs">{{ t('support.description') }}</span>
        <textarea v-model="description" class="textarea textarea-bordered textarea-sm min-h-28" maxlength="4000" required />
      </label>
      <label class="form-control mb-2">
        <span class="label-text mb-1 text-xs">{{ t('support.email') }}</span>
        <input v-model="email" type="email" class="input input-bordered input-sm" maxlength="256" required />
      </label>
      <p class="mb-3 text-xs text-base-content/60">{{ t('support.file') }}</p>
      <p v-if="mailNotice" class="mb-3 text-sm">{{ t('support.file.mail', { email: SUPPORT_FILE_MAIL }) }}</p>
      <p v-if="errorText" class="mb-3 text-xs text-error">{{ errorText }}</p>
      <div class="modal-action mt-0">
        <button type="submit" class="btn btn-primary btn-sm" :disabled="sending">
          {{ sending ? t('support.sending') : t('support.send') }}
        </button>
      </div>
    </form>
    <button type="button" class="modal-backdrop" :aria-label="t('common.close')" @click="open = false" />
  </dialog>
</template>
