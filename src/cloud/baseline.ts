import type { LiveCell } from '@/cloud/values'

const baselines = new Map<string, Map<string, LiveCell[][]>>()

export function cloudBaseline(sessionId: string): Map<string, LiveCell[][]> | null {
  return baselines.get(sessionId) ?? null
}

export function setCloudBaseline(sessionId: string, snapshot: Map<string, LiveCell[][]>): void {
  baselines.set(sessionId, snapshot)
}

export function forgetCloudBaseline(sessionId: string): void {
  baselines.delete(sessionId)
}
