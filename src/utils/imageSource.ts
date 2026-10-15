const IMAGE_EXT = /\.(avif|bmp|gif|jpe?g|png|svg|webp)$/i

export function imageSource(value: unknown): string | null {
  if (typeof value !== 'string') return null
  const text = value.replace(/[\s\u200b]+/g, '')
  if (!text || text.length > 2_000_000) return null
  if (/^data:image\/[a-z0-9.+-]+;base64,/i.test(text)) return text
  try {
    const url = new URL(text)
    if (url.protocol !== 'http:' && url.protocol !== 'https:') return null
    if (url.pathname.includes(',')) return null
    if (!IMAGE_EXT.test(decodeURIComponent(url.pathname))) return null
    return text
  } catch {
    return null
  }
}

/** One image, or several image URLs separated by a comma. Mixed text is not an image list. */
export function imageSources(value: unknown): string[] {
  if (typeof value !== 'string') return []
  const parts = splitCommaParts(value)
  if (parts.length >= 2) {
    const sources: string[] = []
    for (const part of parts) {
      const source = imageSource(part)
      if (!source) return []
      sources.push(source)
    }
    return sources
  }
  const single = imageSource(value)
  return single ? [single] : []
}

function splitCommaParts(text: string): string[] {
  const raw = text.split(',')
  const parts: string[] = []
  for (let i = 0; i < raw.length; i++) {
    const piece = raw[i] ?? ''
    const next = raw[i + 1]
    if (/data:image\//i.test(piece) && next != null && !/^\s*https?:\/\//i.test(next)) {
      parts.push(`${piece},${next}`)
      i++
      continue
    }
    parts.push(piece)
  }
  return parts.map((part) => part.trim()).filter(Boolean)
}

export function imageLabel(source: string): string {
  if (source.startsWith('data:image/')) return 'image'
  try {
    const name = decodeURIComponent(new URL(source).pathname.split('/').pop() ?? '')
    return name || source
  } catch {
    return source
  }
}
