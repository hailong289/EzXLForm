<script setup lang="ts">
import { onMounted, onUnmounted, ref } from 'vue'
import { Palette, Check, RotateCcw, Type, Minus, Plus } from 'lucide-vue-next'
import { useTheme } from '@/composables/useTheme'

const {
  currentTheme,
  fontSize,
  fontSizeOptions,
  customColors,
  availableThemes,
  applyTheme,
  applyFontSize,
  stepFontSize,
  fontSizeMin,
  fontSizeMax,
  applyCustomColors,
  resetCustomColors,
} = useTheme()

const isOpen = ref(false)
const isCustomizing = ref(false)
const tempPrimary = ref(customColors.value.primary)
const tempSecondary = ref(customColors.value.secondary)
const tempAccent = ref(customColors.value.accent)
const menuStyle = ref<Record<string, string>>({})

function toggle(event: MouseEvent) {
  event.stopPropagation()
  if (isOpen.value) {
    isOpen.value = false
    return
  }
  const rect = (event.currentTarget as HTMLElement).getBoundingClientRect()
  const width = 280
  const left = Math.max(8, Math.min(rect.right - width, window.innerWidth - width - 8))
  menuStyle.value = { top: `${rect.bottom + 6}px`, left: `${left}px` }
  isOpen.value = true
}

function close() {
  isOpen.value = false
}

onMounted(() => window.addEventListener('click', close))
onUnmounted(() => window.removeEventListener('click', close))

function onApplyColors() {
  applyCustomColors({
    primary: tempPrimary.value,
    secondary: tempSecondary.value,
    accent: tempAccent.value,
  })
}

function onResetColors() {
  resetCustomColors()
  tempPrimary.value = customColors.value.primary
  tempSecondary.value = customColors.value.secondary
  tempAccent.value = customColors.value.accent
}

function pickTheme(name: string) {
  applyTheme(name)
  close()
}
</script>

<template>
  <button
    type="button"
    class="btn btn-ghost btn-xs gap-1"
    title="Giao diện và cỡ chữ"
    @click="toggle"
  >
    <Palette class="w-3.5 h-3.5 text-primary" />
    <span class="hidden lg:inline capitalize text-xs font-normal">{{ currentTheme }}</span>
  </button>

  <Teleport to="body">
    <div
      v-if="isOpen"
      class="fixed z-[80] w-[280px] max-h-[min(480px,70vh)] overflow-y-auto scrollbar-thin rounded-box border border-base-300 bg-base-100 p-3 shadow-2xl"
      :style="menuStyle"
      @click.stop
    >
      <div class="flex items-center justify-between pb-2 mb-2 border-b border-base-300">
        <span class="font-bold text-xs uppercase tracking-wider opacity-70">Giao diện</span>
        <button type="button" class="btn btn-xs btn-outline btn-primary" @click="isCustomizing = !isCustomizing">
          {{ isCustomizing ? 'Chọn theme' : 'Chỉnh màu' }}
        </button>
      </div>

      <div class="mb-3">
        <div class="flex items-center gap-1.5 mb-1.5 text-xs font-semibold">
          <Type class="w-3.5 h-3.5 opacity-70" />
          Cỡ chữ
        </div>
        <div class="flex items-center gap-1 mb-1.5">
          <button
            type="button"
            class="btn btn-xs btn-ghost btn-square"
            title="Giảm cỡ chữ"
            :disabled="fontSize <= fontSizeMin"
            @click="stepFontSize(-1)"
          >
            <Minus class="w-3.5 h-3.5" />
          </button>
          <span class="flex-1 text-center text-xs font-mono">{{ fontSize }}px</span>
          <button
            type="button"
            class="btn btn-xs btn-ghost btn-square"
            title="Tăng cỡ chữ"
            :disabled="fontSize >= fontSizeMax"
            @click="stepFontSize(1)"
          >
            <Plus class="w-3.5 h-3.5" />
          </button>
        </div>
        <div class="grid grid-cols-4 gap-1">
          <button
            v-for="option in fontSizeOptions"
            :key="option.px"
            type="button"
            class="btn btn-xs font-normal"
            :class="fontSize === option.px ? 'btn-primary' : 'btn-ghost'"
            @click="applyFontSize(option.px)"
          >
            {{ option.label }}
          </button>
        </div>
      </div>

      <div v-if="isCustomizing" class="flex flex-col gap-3 py-1">
        <div class="flex items-center justify-between">
          <label class="text-xs font-medium">Màu chính</label>
          <input v-model="tempPrimary" type="color" class="w-7 h-7 rounded border border-base-300 cursor-pointer" />
        </div>
        <div class="flex items-center justify-between">
          <label class="text-xs font-medium">Màu phụ</label>
          <input v-model="tempSecondary" type="color" class="w-7 h-7 rounded border border-base-300 cursor-pointer" />
        </div>
        <div class="flex items-center justify-between">
          <label class="text-xs font-medium">Màu nhấn</label>
          <input v-model="tempAccent" type="color" class="w-7 h-7 rounded border border-base-300 cursor-pointer" />
        </div>
        <div class="flex gap-2 pt-2 border-t border-base-300">
          <button type="button" class="btn btn-primary btn-xs flex-1 gap-1" @click="onApplyColors">
            <Check class="w-3 h-3" /> Áp dụng
          </button>
          <button type="button" class="btn btn-ghost btn-xs gap-1" @click="onResetColors">
            <RotateCcw class="w-3 h-3" /> Khôi phục
          </button>
        </div>
      </div>

      <div v-else class="flex flex-col gap-0.5">
        <button
          v-for="th in availableThemes"
          :key="th.name"
          type="button"
          class="btn btn-xs justify-start h-8 px-2 font-normal text-xs"
          :class="currentTheme === th.name ? 'btn-primary' : 'btn-ghost'"
          @click="pickTheme(th.name)"
        >
          <span
            class="w-3 h-3 rounded-full border border-base-content/20 flex-shrink-0"
            :style="{ backgroundColor: th.previewColor }"
          />
          <span class="truncate flex-1 text-left ml-1.5">{{ th.label }}</span>
          <Check v-if="currentTheme === th.name" class="w-3 h-3 flex-shrink-0" />
        </button>
      </div>
    </div>
  </Teleport>
</template>
