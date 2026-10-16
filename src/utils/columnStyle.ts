export type ColumnAlign = 'left' | 'center' | 'right' | 'justify'
export type ColumnVertical = 'top' | 'middle' | 'bottom'
export type ColumnBorder = 'box' | 'bottom'

export interface ColumnStyle {
  align?: ColumnAlign
  vertical?: ColumnVertical
  wrap?: boolean
  indent?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  fontSize?: number
  fontColor?: string
  fontName?: string
  border?: ColumnBorder
  borderColor?: string
  fill?: string
  locked?: boolean
  hidden?: boolean
}

export interface ColumnStyleDraft {
  align: '' | ColumnAlign
  vertical: '' | ColumnVertical
  wrap: boolean
  indent: number
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  fontSize: number
  fontColor: string
  fontName: string
  border: '' | ColumnBorder
  borderColor: string
  fill: string
  locked: boolean
  hidden: boolean
}

export function styleFromDraft(draft: ColumnStyleDraft): ColumnStyle | undefined {
  const style: ColumnStyle = {}
  if (draft.align) style.align = draft.align
  if (draft.vertical) style.vertical = draft.vertical
  if (draft.wrap) style.wrap = true
  if (draft.indent > 0) style.indent = Math.min(15, Math.floor(draft.indent))
  if (draft.bold) style.bold = true
  if (draft.italic) style.italic = true
  if (draft.underline) style.underline = true
  if (draft.strike) style.strike = true
  if (draft.fontSize > 0) style.fontSize = draft.fontSize
  if (draft.fontColor) style.fontColor = draft.fontColor
  if (draft.fontName) style.fontName = draft.fontName
  if (draft.border) {
    style.border = draft.border
    if (draft.borderColor) style.borderColor = draft.borderColor
  }
  if (draft.fill) style.fill = draft.fill
  if (draft.locked) style.locked = true
  if (draft.hidden) style.hidden = true
  return Object.keys(style).length ? style : undefined
}

export function columnPaint(style?: ColumnStyle): { box: Record<string, string>; text: Record<string, string> } {
  const box: Record<string, string> = {}
  const text: Record<string, string> = {}
  if (!style) return { box, text }
  if (style.align) text.textAlign = style.align
  if (style.align === 'center') text.justifyContent = 'center'
  else if (style.align === 'right') text.justifyContent = 'flex-end'
  else if (style.align) text.justifyContent = 'flex-start'
  if (style.vertical === 'top') text.alignItems = 'flex-start'
  else if (style.vertical === 'bottom') text.alignItems = 'flex-end'
  else if (style.vertical === 'middle') text.alignItems = 'center'
  if (style.wrap) text.whiteSpace = 'normal'
  if (style.bold) text.fontWeight = '700'
  if (style.italic) text.fontStyle = 'italic'
  const deco = [style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean).join(' ')
  if (deco) text.textDecoration = deco
  if (style.fontSize) text.fontSize = `${style.fontSize}px`
  if (style.fontColor) text.color = style.fontColor
  if (style.fontName) text.fontFamily = style.fontName
  if (style.indent) text.paddingLeft = `${style.indent * 8}px`
  if (style.fill) box.backgroundColor = style.fill
  if (style.border) {
    const color = style.borderColor || '#000000'
    box.boxShadow = style.border === 'bottom' ? `inset 0 -1px 0 0 ${color}` : `inset 0 0 0 1px ${color}`
  }
  return { box, text }
}

export interface StoredColumnFormat {
  dateFormat?: string
  numFmt?: string
  style?: ColumnStyle
  pinned?: boolean
  required?: boolean
  unique?: boolean
  pattern?: string
  options?: string[]
  formula?: string
}

export function columnFormatPayload(header: StoredColumnFormat): string | null {
  const payload: StoredColumnFormat = {}
  if (header.dateFormat) payload.dateFormat = header.dateFormat
  if (header.numFmt) payload.numFmt = header.numFmt
  if (header.style) payload.style = header.style
  if (header.pinned) payload.pinned = true
  if (header.required) payload.required = true
  if (header.unique) payload.unique = true
  if (header.pattern) payload.pattern = header.pattern
  if (header.options?.length) payload.options = header.options
  if (header.formula) payload.formula = header.formula
  return Object.keys(payload).length ? JSON.stringify(payload) : null
}

export function readColumnFormat(json: unknown): StoredColumnFormat {
  if (typeof json !== 'string' || !json) return {}
  try {
    const value = JSON.parse(json) as Record<string, unknown>
    const options = Array.isArray(value.options) ? value.options.filter((item): item is string => typeof item === 'string') : []
    const stored: StoredColumnFormat = {}
    if (typeof value.dateFormat === 'string') stored.dateFormat = value.dateFormat
    if (typeof value.numFmt === 'string') stored.numFmt = value.numFmt
    if (value.style && typeof value.style === 'object') stored.style = value.style as ColumnStyle
    if (value.pinned === true) stored.pinned = true
    if (value.required === true) stored.required = true
    if (value.unique === true) stored.unique = true
    if (typeof value.pattern === 'string') stored.pattern = value.pattern
    if (options.length) stored.options = options
    if (typeof value.formula === 'string') stored.formula = value.formula
    return stored
  } catch {
    return {}
  }
}
