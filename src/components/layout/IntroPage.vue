<script setup lang="ts">
import { computed, onMounted, ref } from 'vue'
import { FileSpreadsheet, Globe, MonitorDown } from 'lucide-vue-next'

const REPO = 'hailong289/EzXLForm'
const RELEASES_URL = `https://github.com/${REPO}/releases/latest`

type Installer = {
  name: string
  url: string
  os: 'windows' | 'mac' | 'linux'
  label: string
}

const emit = defineEmits<{
  'open-web': []
}>()

const points = [
  'Mở và sửa file .xlsx, .xls, .csv, nhiều sheet trong một sổ.',
  'Form và bảng chia đôi màn hình, bấm đúp ô để sửa trực tiếp.',
  'Tìm và thay thế, xem trước Excel, hoàn tác và lưu phiên trên máy.',
  'Bản desktop đóng bằng Tauri cho Windows, macOS và Linux.',
]

const installers = ref<Installer[]>([])
const version = ref('')
const loading = ref(true)

function detectOs(): Installer['os'] {
  const ua = navigator.userAgent.toLowerCase()
  if (ua.includes('win')) return 'windows'
  if (ua.includes('mac')) return 'mac'
  return 'linux'
}

function classify(name: string): Pick<Installer, 'os' | 'label'> | null {
  const file = name.toLowerCase()
  if (file.endsWith('.exe')) return { os: 'windows', label: 'Windows' }
  if (file.endsWith('.msi')) return { os: 'windows', label: 'Windows (MSI)' }
  if (file.endsWith('.dmg')) return { os: 'mac', label: 'macOS' }
  if (file.endsWith('.deb')) return { os: 'linux', label: 'Linux (.deb)' }
  if (file.endsWith('.appimage')) return { os: 'linux', label: 'Linux (AppImage)' }
  return null
}

const currentOs = detectOs()

const primary = computed(() => {
  const sameOs = installers.value.filter((item) => item.os === currentOs)
  return sameOs.find((item) => item.name.toLowerCase().endsWith('.exe') || item.name.toLowerCase().endsWith('.dmg') || item.name.toLowerCase().endsWith('.deb')) ?? sameOs[0] ?? null
})

const others = computed(() => installers.value.filter((item) => item !== primary.value))

onMounted(async () => {
  try {
    const response = await fetch(`https://api.github.com/repos/${REPO}/releases/latest`)
    if (!response.ok) return
    const data = (await response.json()) as {
      tag_name?: string
      assets?: { name?: string; browser_download_url?: string }[]
    }
    version.value = data.tag_name ?? ''
    installers.value = (data.assets ?? []).flatMap((asset) => {
      if (!asset.name || !asset.browser_download_url) return []
      const kind = classify(asset.name)
      if (!kind) return []
      return [{ name: asset.name, url: asset.browser_download_url, ...kind }]
    })
  } catch {
    installers.value = []
  } finally {
    loading.value = false
  }
})
</script>

<template>
  <div class="flex h-full items-center justify-center overflow-auto p-6">
    <div class="w-full max-w-3xl">
      <div class="mb-6 flex items-center gap-3">
        <div class="rounded-2xl bg-primary/10 p-3 text-primary">
          <FileSpreadsheet class="h-8 w-8" />
        </div>
        <div>
          <h1 class="text-2xl font-bold">EzXLForm</h1>
          <p class="text-sm text-base-content/70">Excel Editor — sửa bảng tính trên web hoặc trên máy.</p>
        </div>
      </div>

      <p class="mb-4 text-sm leading-relaxed text-base-content/80">
        EzXLForm giúp xem và sửa dữ liệu Excel như một form. Chọn dòng trên bảng, gõ trên form,
        hoặc sửa ngay trong ô. File lớn được chia trang, tiến trình tìm và thay thế có thể tạm dừng.
      </p>

      <ul class="mb-6 space-y-2 text-sm">
        <li v-for="point in points" :key="point" class="flex gap-2">
          <span class="mt-1 h-1.5 w-1.5 shrink-0 rounded-full bg-primary" />
          <span>{{ point }}</span>
        </li>
      </ul>

      <div class="grid gap-3 sm:grid-cols-2">
        <button
          type="button"
          class="card card-bordered border-base-300 bg-base-100 text-left transition hover:border-primary"
          @click="emit('open-web')"
        >
          <div class="card-body gap-3 p-4">
            <div class="flex items-center gap-2 font-semibold">
              <Globe class="h-4 w-4 text-primary" />
              Bản web
            </div>
            <p class="text-xs leading-relaxed text-base-content/70">
              Mở trình sửa ngay trên trang này. Kéo thả file hoặc vào lại phiên đã lưu.
            </p>
            <span class="btn btn-primary btn-sm w-fit">Mở bản web</span>
          </div>
        </button>

        <div class="card card-bordered border-base-300 bg-base-100 text-left">
          <div class="card-body gap-3 p-4">
            <div class="flex items-center gap-2 font-semibold">
              <MonitorDown class="h-4 w-4 text-primary" />
              Bản desktop
              <span v-if="version" class="badge badge-sm badge-ghost font-mono">{{ version }}</span>
            </div>
            <p class="text-xs leading-relaxed text-base-content/70">
              Tải bộ cài mới nhất. Nút tự chọn file đúng hệ điều hành đang mở trang.
            </p>
            <a
              v-if="primary"
              class="btn btn-primary btn-sm w-fit"
              :href="primary.url"
            >
              Tải {{ primary.label }}
            </a>
            <a
              v-else
              class="btn btn-primary btn-sm w-fit"
              :href="RELEASES_URL"
              target="_blank"
              rel="noopener noreferrer"
            >
              {{ loading ? 'Đang tìm bản cài...' : 'Tải bản desktop' }}
            </a>
            <div v-if="others.length" class="flex flex-wrap gap-2">
              <a
                v-for="item in others"
                :key="item.name"
                class="btn btn-ghost btn-xs"
                :href="item.url"
              >
                {{ item.label }}
              </a>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>
