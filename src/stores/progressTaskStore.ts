import { defineStore } from 'pinia'
import type { FindReplaceOptions } from '@/types/excel'

export type ProgressTaskStatus = 'queued' | 'running' | 'paused' | 'done' | 'cancelled'
export type ProgressTaskKind = 'import' | 'replace' | 'filter' | 'cloud'

export interface ProgressTask {
  id: string
  title: string
  detail: string
  percent: number
  status: ProgressTaskStatus
  sheetId: string
  kind: ProgressTaskKind
  replaceOptions?: FindReplaceOptions
}

const TASKS_STORAGE_KEY = 'excel_editor_progress_tasks'

function readStoredTasks(): ProgressTask[] {
  try {
    const raw = localStorage.getItem(TASKS_STORAGE_KEY)
    if (!raw) return []
    const parsed = JSON.parse(raw) as ProgressTask[]
    if (!Array.isArray(parsed)) return []
    return parsed.flatMap((task) => {
      const kind: ProgressTaskKind =
        task.kind === 'import' || task.kind === 'filter' || task.kind === 'replace' || task.kind === 'cloud' ? task.kind : 'replace'
      const unfinished = task.status === 'running' || task.status === 'paused' || task.status === 'queued'
      if ((kind === 'import' || kind === 'filter' || kind === 'cloud') && unfinished) {
        return [{ ...task, kind, status: 'cancelled' as const }]
      }
      if (kind === 'replace' && unfinished) {
        if (!task.replaceOptions) return []
        return [{ ...task, kind, status: 'paused' as const }]
      }
      return [{ ...task, kind }]
    })
  } catch {
    return []
  }
}

function writeStoredTasks(tasks: ProgressTask[]) {
  localStorage.setItem(TASKS_STORAGE_KEY, JSON.stringify(tasks))
}

export function canPauseTask(task: ProgressTask): boolean {
  return task.kind !== 'import' && task.kind !== 'cloud' && task.status === 'running'
}

export function canResumeTask(task: ProgressTask): boolean {
  return task.kind !== 'import' && task.kind !== 'cloud' && task.status === 'paused'
}

export function canRemoveTask(task: ProgressTask): boolean {
  return !((task.kind === 'import' || task.kind === 'cloud') && (task.status === 'running' || task.status === 'queued'))
}

export function taskStatusLabel(status: ProgressTaskStatus): string {
  if (status === 'running') return 'common.running'
  if (status === 'paused') return 'common.paused'
  if (status === 'queued') return 'common.queued'
  if (status === 'done') return 'common.done'
  return 'common.stopped'
}

export function taskStatusBadgeClass(status: ProgressTaskStatus): string {
  if (status === 'running') return 'badge-info'
  if (status === 'paused') return 'badge-warning'
  if (status === 'queued') return 'badge-ghost'
  if (status === 'done') return 'badge-success'
  return 'badge-error'
}

export interface TaskSignal {
  checkpoint: () => Promise<boolean>
}

class TaskSignalImpl implements TaskSignal {
  status: ProgressTaskStatus = 'queued'
  private wake: (() => void) | null = null

  async checkpoint(): Promise<boolean> {
    while (this.status === 'paused') {
      await new Promise<void>((resolve) => {
        this.wake = resolve
      })
    }
    return this.status === 'running'
  }

  pause() {
    if (this.status === 'running') this.status = 'paused'
  }

  resume() {
    if (this.status !== 'paused') return
    this.status = 'running'
    this.wake?.()
    this.wake = null
  }

  cancel() {
    this.status = 'cancelled'
    this.wake?.()
    this.wake = null
  }
}

const signals = new Map<string, TaskSignalImpl>()

function restorePausedSignals(tasks: ProgressTask[]) {
  for (const task of tasks) {
    if (task.kind !== 'replace' || task.status !== 'paused' || !task.replaceOptions) continue
    const signal = new TaskSignalImpl()
    signal.status = 'paused'
    signals.set(task.id, signal)
  }
}

const storedTasks = readStoredTasks()
restorePausedSignals(storedTasks)

function nextId(): string {
  return `task_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`
}

export const useProgressTaskStore = defineStore('progressTasks', {
  state: () => ({
    tasks: storedTasks,
    managerOpen: false,
    dockVisible: false,
  }),
  getters: {
    activeTask(state): ProgressTask | null {
      return (
        state.tasks.find((task) => task.status === 'running')
        ?? state.tasks.find((task) => task.status === 'paused')
        ?? state.tasks.find((task) => task.status === 'queued')
        ?? state.tasks[state.tasks.length - 1]
        ?? null
      )
    },
    hasTasks(state): boolean {
      return state.tasks.length > 0
    },
  },
  actions: {
    enqueue(input: {
      title: string
      detail: string
      sheetId: string
      kind?: ProgressTaskKind
      replaceOptions?: FindReplaceOptions
    }): string {
      const id = nextId()
      const signal = new TaskSignalImpl()
      signals.set(id, signal)
      this.tasks.push({
        id,
        title: input.title,
        detail: input.detail,
        percent: 0,
        status: 'queued',
        sheetId: input.sheetId,
        kind: input.kind ?? 'replace',
        replaceOptions: input.replaceOptions,
      })
      writeStoredTasks(this.tasks)
      this.dockVisible = true
      return id
    },
    hideDock() {
      this.dockVisible = false
    },
    signal(id: string): TaskSignal | null {
      return signals.get(id) ?? null
    },
    setStatus(id: string, status: ProgressTaskStatus) {
      const task = this.tasks.find((item) => item.id === id)
      const signal = signals.get(id)
      if (task) task.status = status
      if (signal) signal.status = status
      writeStoredTasks(this.tasks)
    },
    setPercent(id: string, percent: number) {
      const task = this.tasks.find((item) => item.id === id)
      if (task) task.percent = Math.max(0, Math.min(100, percent))
      writeStoredTasks(this.tasks)
    },
    setDetail(id: string, detail: string) {
      const task = this.tasks.find((item) => item.id === id)
      if (task) task.detail = detail
      writeStoredTasks(this.tasks)
    },
    pause(id: string) {
      const task = this.tasks.find((item) => item.id === id)
      const signal = signals.get(id)
      if (!task || !signal || !canPauseTask(task)) return
      signal.pause()
      task.status = 'paused'
      writeStoredTasks(this.tasks)
    },
    resume(id: string) {
      const task = this.tasks.find((item) => item.id === id)
      const signal = signals.get(id)
      if (!task || !signal || !canResumeTask(task)) return
      signal.resume()
      task.status = 'running'
      writeStoredTasks(this.tasks)
    },
    remove(id: string) {
      const task = this.tasks.find((item) => item.id === id)
      if (!task || !canRemoveTask(task)) return
      signals.get(id)?.cancel()
      this.tasks = this.tasks.filter((item) => item.id !== id)
      writeStoredTasks(this.tasks)
    },
    release(id: string) {
      signals.delete(id)
    },
    openManager() {
      this.managerOpen = true
    },
    closeManager() {
      this.managerOpen = false
    },
  },
})
