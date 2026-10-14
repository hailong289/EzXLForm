const IMAGE_EXT = /\.(avif|bmp|gif|jpe?g|png|svg|webp)$/i

export function imageSource(value: unknown): string | null {
  if (typeof value !== 'string') return null
  const text = value.replace(/[\s\u200b]+/g, '')
  if (!text || text.length > 2_000_000) return null
  if (/^data:image\/[a-z0-9.+-]+;base64,/i.test(text)) return text
  try {
    const url = new URL(text)
    if (url.protocol !== 'http:' && url.protocol !== 'https:') return null
    if (!IMAGE_EXT.test(decodeURIComponent(url.pathname))) return null
    return text
  } catch {
    return null
  }
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
