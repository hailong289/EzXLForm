<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Cloud, FileSpreadsheet, Link2, LogOut, RefreshCw, Search, X } from 'lucide-vue-next'
import { googleConfigured, microsoftConfigured } from '@/cloud/config'
import { googleAccountLabel, signOutGoogle } from '@/cloud/googleAuth'
import { microsoftAccountLabel, signOutMicrosoft } from '@/cloud/microsoftAuth'
import { fileFromShareLink } from '@/cloud/shareLink'
import { listGoogleFiles, listRecentExcel, openCloudFile } from '@/cloud/sync'
import { GOOGLE_SHEET_MIME, type CloudFileRef, type CloudOpenMode, type CloudProviderId } from '@/cloud/types'
import type { CloudOpenResult } from '@/cloud/sync'

const props = defineProps<{
  modelValue: boolean
  providerId: CloudProviderId
}>()
const emit = defineEmits<{
  'update:modelValue': [open: boolean]
  opened: [result: CloudOpenResult]
}>()

const mode = ref<CloudOpenMode>('file')
const provider = ref<CloudProviderId | null>(null)
const account = ref('')
const files = ref<CloudFileRef[]>([])
const query = ref('')
const shareLink = ref('')
const busy = ref('')
const error = ref('')
const empty = ref('')
const openingId = ref('')
const linkBusy = ref(false)

const title = computed(() => (props.providerId === 'google' ? 'Google Drive' : 'OneDrive'))
const modeHint = computed(() =>
  mode.value === 'live'
    ? 'Đọc và ghi giá trị ô. Tối đa 10.000 dòng và 50 cột. Công thức và định dạng không đi theo.'
    : 'Tải file về rồi sửa như file máy. Google Sheet lưu lại chỉ ghi giá trị ô.',
)
const visibleFiles = computed(() => {
  const needle = query.value.trim().toLowerCase()
  if (!needle) return files.value
  return files.value.filter((file) => file.name.toLowerCase().includes(needle))
})

watch(
  () => props.modelValue,
  (open) => {
    if (!open) return
    error.value = ''
    empty.value = ''
    files.value = []
    query.value = ''
    shareLink.value = ''
    busy.value = ''
    openingId.value = ''
    linkBusy.value = false
    provider.value = props.providerId
    const ready = props.providerId === 'google' ? googleConfigured() : microsoftConfigured()
    if (!ready) return
    void signIn(props.providerId)
  },
)

watch(mode, () => {
  if (!props.modelValue || !provider.value || busy.value) return
  void reload()
})

function close() {
  emit('update:modelValue', false)
}

function fileKind(file: CloudFileRef): string {
  if (file.mime === GOOGLE_SHEET_MIME) return 'Google Sheet'
  return 'Excel'
}

async function run(label: string, action: () => Promise<void>) {
  error.value = ''
  empty.value = ''
  busy.value = label
  try {
    await action()
  } catch (err: unknown) {
    error.value = err instanceof Error ? err.message : 'Không kết nối được cloud'
  } finally {
    busy.value = ''
  }
}

function noteIfEmpty() {
  if (files.value.length > 0) return
  empty.value =
    provider.value === 'microsoft'
      ? 'Không thấy file .xlsx gần đây trên OneDrive.'
      : mode.value === 'live'
        ? 'Không thấy Google Sheet nào.'
        : 'Không thấy Google Sheet hoặc file .xlsx nào.'
}

async function signIn(next: CloudProviderId) {
  provider.value = next
  files.value = []
  await run(next === 'google' ? 'Đang mở Google Drive...' : 'Đang mở OneDrive...', async () => {
    if (next === 'google') {
      files.value = await listGoogleFiles(mode.value === 'live')
      account.value = googleAccountLabel() || 'Google'
      noteIfEmpty()
      return
    }
    account.value = (await microsoftAccountLabel()) || ''
    if (!account.value) {
      files.value = await listRecentExcel()
      account.value = (await microsoftAccountLabel()) || 'Microsoft'
      noteIfEmpty()
    }
  })
}

async function reload() {
  if (!provider.value) return
  await run('Đang tải danh sách...', async () => {
    files.value =
      provider.value === 'google' ? await listGoogleFiles(mode.value === 'live') : await listRecentExcel()
    if (provider.value === 'google') account.value = googleAccountLabel() || account.value
    else account.value = (await microsoftAccountLabel()) || account.value
    noteIfEmpty()
  })
}

async function openListed(file: CloudFileRef) {
  openingId.value = file.graphPath || file.id
  await run(mode.value === 'live' ? 'Đang đọc ô...' : 'Đang tải file...', async () => {
    if (!provider.value) return
    const result = await openCloudFile(provider.value, mode.value, file)
    emit('opened', result)
    close()
  })
  openingId.value = ''
}

async function openFromLink() {
  const raw = shareLink.value.trim()
  if (!raw || linkBusy.value) return
  error.value = ''
  linkBusy.value = true
  try {
    const resolved = await fileFromShareLink(raw)
    const result = await openCloudFile(resolved.provider, mode.value, resolved.file)
    emit('opened', result)
    close()
  } catch (err: unknown) {
    error.value = err instanceof Error ? err.message : 'Không mở được link'
  } finally {
    linkBusy.value = false
  }
}

async function signOut() {
  if (provider.value === 'google') signOutGoogle()
  if (provider.value === 'microsoft') await signOutMicrosoft()
  account.value = ''
  files.value = []
  empty.value = ''
}
</script>

<template>
  <dialog v-if="props.modelValue" class="modal modal-open">
    <div class="modal-box flex max-h-[min(40rem,calc(100vh-2rem))] max-w-xl flex-col gap-0 overflow-hidden p-0">
      <div class="flex items-start justify-between gap-3 border-b border-base-300 px-4 py-3">
        <div class="flex min-w-0 items-center gap-3">
          <span class="grid h-9 w-9 shrink-0 place-items-center rounded-xl bg-primary/10 text-primary">
            <Cloud class="h-4 w-4" />
          </span>
          <div class="min-w-0">
            <h3 class="text-sm font-semibold leading-tight">{{ title }}</h3>
            <p class="truncate text-[11px] opacity-60">
              {{ account || 'Chọn file để mở trong phiên này' }}
            </p>
          </div>
        </div>
        <div class="flex shrink-0 items-center gap-1">
          <button
            v-if="account"
            type="button"
            class="btn btn-ghost btn-xs"
            :disabled="!!busy"
            @click="signOut"
          >
            <LogOut class="h-3.5 w-3.5" />
            Đăng xuất
          </button>
          <button type="button" class="btn btn-ghost btn-xs btn-circle" @click="close">
            <X class="h-4 w-4" />
          </button>
        </div>
      </div>

      <div class="flex flex-col gap-3 px-4 py-3">
        <div class="join w-full bg-base-200 p-0.5">
          <button
            type="button"
            class="btn join-item btn-xs flex-1"
            :class="mode === 'file' ? 'btn-primary' : 'btn-ghost'"
            @click="mode = 'file'"
          >
            File .xlsx
          </button>
          <button
            type="button"
            class="btn join-item btn-xs flex-1"
            :class="mode === 'live' ? 'btn-primary' : 'btn-ghost'"
            @click="mode = 'live'"
          >
            Đọc / ghi ô
          </button>
        </div>
        <p class="text-[11px] leading-relaxed opacity-60">{{ modeHint }}</p>

        <p v-if="props.providerId === 'google' && !googleConfigured()" class="text-xs text-warning">
          Thiếu VITE_GOOGLE_CLIENT_ID.
        </p>
        <p v-else-if="props.providerId === 'microsoft' && !microsoftConfigured()" class="text-xs text-warning">
          Thiếu VITE_MICROSOFT_CLIENT_ID.
        </p>

        <form class="flex items-center gap-2" @submit.prevent="openFromLink">
          <label class="input input-sm input-bordered flex min-w-0 flex-1 items-center gap-2">
            <Link2 class="h-3.5 w-3.5 shrink-0 opacity-50" />
            <input
              v-model="shareLink"
              type="text"
              class="min-w-0 grow bg-transparent text-sm"
              placeholder="Dán link chia sẻ"
            />
          </label>
          <button type="submit" class="btn btn-primary btn-sm shrink-0" :disabled="linkBusy || !shareLink.trim()">
            <span v-if="linkBusy" class="loading loading-spinner loading-xs" />
            Mở link
          </button>
        </form>

        <label class="input input-sm input-bordered flex items-center gap-2">
          <Search class="h-3.5 w-3.5 opacity-50" />
          <input v-model="query" type="search" class="grow bg-transparent text-sm" placeholder="Lọc theo tên file" />
        </label>
      </div>

      <div class="min-h-48 flex-1 overflow-y-auto border-t border-base-300 px-2 py-2 scrollbar-thin">
        <div v-if="busy && files.length === 0" class="flex h-48 flex-col items-center justify-center gap-2 text-xs opacity-70">
          <span class="loading loading-spinner loading-md text-primary" />
          {{ busy }}
        </div>

        <div v-else-if="empty && !query" class="flex h-48 flex-col items-center justify-center px-6 text-center">
          <FileSpreadsheet class="mb-2 h-8 w-8 opacity-30" />
          <p class="text-sm">{{ empty }}</p>
        </div>

        <div v-else-if="visibleFiles.length === 0" class="flex h-48 items-center justify-center px-6 text-center text-sm opacity-60">
          Không có file khớp với “{{ query }}”.
        </div>

        <ul v-else class="space-y-0.5">
          <li v-for="file in visibleFiles" :key="file.graphPath || file.id">
            <button
              type="button"
              class="flex w-full items-center gap-3 rounded-lg px-2 py-2 text-left hover:bg-base-200 disabled:opacity-60"
              :disabled="!!busy"
              @click="openListed(file)"
            >
              <span class="grid h-9 w-9 shrink-0 place-items-center rounded-lg bg-success/10 text-success">
                <span
                  v-if="openingId === (file.graphPath || file.id)"
                  class="loading loading-spinner loading-xs"
                />
                <FileSpreadsheet v-else class="h-4 w-4" />
              </span>
              <span class="min-w-0 flex-1">
                <span class="block truncate text-sm font-medium">{{ file.name }}</span>
                <span class="block text-[11px] opacity-50">{{ fileKind(file) }}</span>
              </span>
            </button>
          </li>
        </ul>
      </div>

      <div class="flex items-center justify-between gap-2 border-t border-base-300 px-4 py-2">
        <p class="min-w-0 truncate text-[11px]" :class="error ? 'text-error' : 'opacity-60'">
          {{ error || busy || `${visibleFiles.length} file` }}
        </p>
        <button type="button" class="btn btn-ghost btn-xs gap-1" :disabled="!!busy || !provider" @click="reload">
          <RefreshCw class="h-3.5 w-3.5" :class="busy ? 'animate-spin' : ''" />
          Tải lại
        </button>
      </div>
    </div>
    <form method="dialog" class="modal-backdrop">
      <button type="button" @click="close">đóng</button>
    </form>
  </dialog>
</template>
