export type WorkspaceTier = 'phone' | 'tablet' | 'desktop'

export type EditorPanes = { form: boolean; table: boolean }

export type EditorVisibilityEvent = 'select-row' | 'show-table' | 'show-form' | 'hide-form' | 'hide-table'

export const BATCH_BAR_BOTTOM_PX = 60

export function workspaceTier(width: number): WorkspaceTier {
  if (width < 768) return 'phone'
  if (width < 1024) return 'tablet'
  return 'desktop'
}

export function panesForTier(tier: WorkspaceTier, panes: EditorPanes): EditorPanes {
  if (tier !== 'phone') {
    if (!panes.form && !panes.table) return { form: true, table: true }
    return { form: panes.form, table: panes.table }
  }
  if (panes.form && !panes.table) return { form: true, table: false }
  return { form: false, table: true }
}

export function nextEditorVisibility(
  tier: WorkspaceTier,
  panes: EditorPanes,
  event: EditorVisibilityEvent,
): EditorPanes {
  if (tier === 'phone') {
    if (event === 'select-row' || event === 'show-form' || event === 'hide-table') {
      return panesForTier(tier, { form: true, table: false })
    }
    return panesForTier(tier, { form: false, table: true })
  }
  if (event === 'select-row') return panesForTier(tier, panes)
  if (event === 'hide-form') return panesForTier(tier, { form: false, table: true })
  if (event === 'hide-table') return panesForTier(tier, { form: true, table: false })
  if (event === 'show-form') return panesForTier(tier, { form: true, table: panes.table })
  return panesForTier(tier, { form: panes.form, table: true })
}

export function clampFormWidth(tier: WorkspaceTier, editorWidth: number, requested: number): number {
  const min = tier === 'tablet' ? 280 : 340
  const ratio = tier === 'tablet' ? 0.42 : 0.5
  const max = Math.max(min, Math.round(editorWidth * ratio))
  return Math.min(max, Math.max(min, requested))
}

export function allowInlineCellEdit(tier: WorkspaceTier): boolean {
  return tier !== 'phone'
}

export function menuWidthPx(viewport: number, preferred: number): number {
  return Math.min(preferred, Math.max(160, viewport - 16))
}
