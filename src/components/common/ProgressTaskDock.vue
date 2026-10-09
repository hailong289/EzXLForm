<script setup lang="ts">
import { computed, onMounted, onUnmounted, ref } from 'vue'
import { List, Minus, Pause, Play, Trash2, X } from 'lucide-vue-next'
import { useProgressTaskStore, canPauseTask, canRemoveTask, canResumeTask, taskStatusBadgeClass, taskStatusLabel } from '@/stores/progressTaskStore'
import type { ProgressTaskStatus } from '@/stores/progressTaskStore'

const DOCK_POS_KEY = 'excel_editor_progress_dock_pos'
const progressStore = useProgressTaskStore()
const collapsed = ref(true)
const position = ref<{ x: number; y: number } | null>(null)
const dragMoved = ref(false)
const activeTask = computed(() => progressStore.activeTask)
const tasks = computed(() => progressStore.tasks)

function progressClass(status: ProgressTaskStatus | undefined): string {
  if (status === 'paused') return 'progress-warning'
  if (status === 'done') return 'progress-success'
  if (status === 'cancelled') return 'progress-error'
  if (status === 'queued') return ''
  return 'progress-info'
}

function expandDock() {
  if (dragMoved.value) return
  collapsed.value = false
}

function openManager() {
  progressStore.openManager()
}

function readPosition(): { x: number; y: number } | null {
  try {
    const raw = localStorage.getItem(DOCK_POS_KEY)
    if (!raw) return null
    const parsed = JSON.parse(raw) as { x?: number; y?: number }
    if (typeof parsed.x !== 'number' || typeof parsed.y !== 'number') return null
    return { x: parsed.x, y: parsed.y }
  } catch {
    return null
  }
}

function clampPosition(x: number, y: number): { x: number; y: number } {
  const maxX = Math.max(8, window.innerWidth - 340)
  const maxY = Math.max(8, window.innerHeight - 180)
  return {
    x: Math.min(Math.max(8, x), maxX),
    y: Math.min(Math.max(8, y), maxY),
  }
}

function placeDock() {
  const saved = readPosition()
  position.value = clampPosition(
    saved?.x ?? window.innerWidth - 360,
    saved?.y ?? window.innerHeight - 200,
  )
}

function onResize() {
  if (!position.value) return
  position.value = clampPosition(position.value.x, position.value.y)
}

onMounted(() => {
  placeDock()
  window.addEventListener('resize', onResize)
})

onUnmounted(() => {
  window.removeEventListener('resize', onResize)
})

function onDragStart(event: PointerEvent) {
  if (!position.value) return
  const target = event.target as HTMLElement | null
  if (target?.closest('button')) return
  const origin = position.value
  const startX = event.clientX
  const startY = event.clientY
  dragMoved.value = false
  const move = (ev: PointerEvent) => {
    if (Math.hypot(ev.clientX - startX, ev.clientY - startY) > 4) dragMoved.value = true
    position.value = clampPosition(origin.x + ev.clientX - startX, origin.y + ev.clientY - startY)
  }
  const up = () => {
    window.removeEventListener('pointermove', move)
    window.removeEventListener('pointerup', up)
    if (position.value) localStorage.setItem(DOCK_POS_KEY, JSON.stringify(position.value))
  }
  window.addEventListener('pointermove', move)
  window.addEventListener('pointerup', up)
}
</script>

<template>
  <Teleport to="body">
  <div
    v-if="progressStore.hasTasks && progressStore.dockVisible && position"
    class="fixed z-[200] cursor-grab active:cursor-grabbing"
    :style="{ left: `${position.x}px`, top: `${position.y}px` }"
    @pointerdown="onDragStart"
  >
    <div
      v-if="collapsed"
      class="w-72 rounded-xl border border-base-300 bg-base-100 p-2.5 shadow-lg"
      @click="expandDock"
    >
      <div class="flex items-start gap-2">
        <p class="min-w-0 flex-1 text-sm font-medium leading-snug">{{ activeTask?.title ?? 'Tiến trình' }}</p>
        <button
          type="button"
          class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0"
          title="Tắt popup"
          @click.stop="progressStore.hideDock()"
        >
          <X class="h-3.5 w-3.5" />
        </button>
      </div>
      <div class="mt-1 flex items-center gap-2">
        <span v-if="activeTask" class="badge badge-sm" :class="taskStatusBadgeClass(activeTask.status)">
          {{ taskStatusLabel(activeTask.status) }}
        </span>
        <span class="font-mono text-sm tabular-nums">{{ activeTask ? `${activeTask.percent}%` : tasks.length }}</span>
      </div>
      <progress
        class="progress mt-2 h-2 w-full"
        :class="progressClass(activeTask?.status)"
        :value="activeTask?.percent ?? 0"
        max="100"
      />
      <p v-if="activeTask?.detail" class="mt-1.5 line-clamp-2 text-xs leading-snug text-base-content/80 break-words">
        {{ activeTask.detail }}
      </p>
    </div>

    <div v-else class="w-80 rounded-xl border border-base-300 bg-base-100 shadow-xl">
      <div class="flex items-start gap-2 border-b border-base-300 px-3 py-2.5">
        <div class="min-w-0 flex-1">
          <p class="text-sm font-semibold leading-snug">{{ activeTask?.title ?? 'Tiến trình' }}</p>
          <div class="mt-1 flex items-center gap-2">
            <span v-if="activeTask" class="badge badge-sm" :class="taskStatusBadgeClass(activeTask.status)">
              {{ taskStatusLabel(activeTask.status) }}
            </span>
            <span class="font-mono text-sm tabular-nums">{{ activeTask?.percent ?? 0 }}%</span>
          </div>
        </div>
        <button type="button" class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0" title="Thu nhỏ" @click="collapsed = true">
          <Minus class="h-3.5 w-3.5" />
        </button>
        <button type="button" class="btn btn-ghost btn-xs btn-circle h-6 w-6 min-h-0" title="Tắt popup" @click="progressStore.hideDock()">
          <X class="h-3.5 w-3.5" />
        </button>
      </div>
      <div class="space-y-2 px-3 py-2.5">
        <progress
          class="progress h-2 w-full"
          :class="progressClass(activeTask?.status)"
          :value="activeTask?.percent ?? 0"
          max="100"
        />
        <p v-if="activeTask?.detail" class="text-xs leading-relaxed text-base-content/80 break-words">{{ activeTask.detail }}</p>
      </div>
      <div class="flex items-center justify-end gap-1 border-t border-base-300 px-2 py-1.5">
        <button
          v-if="activeTask && canPauseTask(activeTask)"
          type="button"
          class="btn btn-ghost btn-xs h-7 min-h-0 gap-1"
          @click="progressStore.pause(activeTask.id)"
        >
          <Pause class="h-3.5 w-3.5" /> Dừng
        </button>
        <button
          v-if="activeTask && canResumeTask(activeTask)"
          type="button"
          class="btn btn-ghost btn-xs h-7 min-h-0 gap-1"
          @click="progressStore.resume(activeTask.id)"
        >
          <Play class="h-3.5 w-3.5" /> Tiếp tục
        </button>
        <button type="button" class="btn btn-ghost btn-xs h-7 min-h-0 gap-1" @click="openManager">
          <List class="h-3.5 w-3.5" /> Danh sách
        </button>
      </div>
    </div>
  </div>

  <dialog v-if="progressStore.managerOpen" class="modal modal-open">
    <div class="modal-box max-w-lg border border-base-300 p-4">
      <div class="mb-3 flex items-center justify-between border-b border-base-300 pb-2">
        <h3 class="text-base font-bold">Tiến trình</h3>
        <button type="button" class="btn btn-ghost btn-xs btn-circle" @click="progressStore.closeManager()">
          <X class="h-4 w-4" />
        </button>
      </div>

      <p v-if="tasks.length === 0" class="py-6 text-center text-sm text-base-content/60">Không có tiến trình.</p>

      <ul v-else class="max-h-[60vh] space-y-2 overflow-auto">
        <li v-for="task in tasks" :key="task.id" class="rounded-xl border border-base-300 p-3">
          <div class="flex items-start justify-between gap-3">
            <p class="min-w-0 text-sm font-semibold leading-snug">{{ task.title }}</p>
            <span class="shrink-0 font-mono text-sm tabular-nums">{{ task.percent }}%</span>
          </div>
          <p v-if="task.detail" class="mt-1 text-xs leading-relaxed text-base-content/80 break-words">{{ task.detail }}</p>
          <progress class="progress mt-2 h-2 w-full" :class="progressClass(task.status)" :value="task.percent" max="100" />
          <div class="mt-2 flex items-center justify-between gap-2">
            <span class="badge badge-sm" :class="taskStatusBadgeClass(task.status)">{{ taskStatusLabel(task.status) }}</span>
            <div class="flex gap-1">
              <button
                v-if="canPauseTask(task)"
                type="button"
                class="btn btn-ghost btn-xs h-6 min-h-0 gap-1"
                @click="progressStore.pause(task.id)"
              >
                <Pause class="h-3 w-3" /> Dừng
              </button>
              <button
                v-if="canResumeTask(task)"
                type="button"
                class="btn btn-ghost btn-xs h-6 min-h-0 gap-1"
                @click="progressStore.resume(task.id)"
              >
                <Play class="h-3 w-3" /> Tiếp tục
              </button>
              <button
                v-if="canRemoveTask(task)"
                type="button"
                class="btn btn-ghost btn-xs h-6 min-h-0 gap-1 text-error"
                @click="progressStore.remove(task.id)"
              >
                <Trash2 class="h-3 w-3" /> Xóa
              </button>
            </div>
          </div>
        </li>
      </ul>
    </div>
    <div class="modal-backdrop bg-black/40" @click="progressStore.closeManager()" />
  </dialog>
  </Teleport>
</template>
