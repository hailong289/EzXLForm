export interface ZipEntry {
  name: string
  method: number
  compressedSize: number
  dataOffset: number
}

function u16(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8)
}

function u32(bytes: Uint8Array, offset: number): number {
  return (
    (bytes[offset] |
      (bytes[offset + 1] << 8) |
      (bytes[offset + 2] << 16) |
      (bytes[offset + 3] << 24)) >>>
    0
  )
}

function findEocd(bytes: Uint8Array): number {
  const start = Math.max(0, bytes.length - 22 - 65535)
  for (let index = bytes.length - 22; index >= start; index--) {
    if (
      bytes[index] === 0x50 &&
      bytes[index + 1] === 0x4b &&
      bytes[index + 2] === 0x05 &&
      bytes[index + 3] === 0x06
    ) {
      return index
    }
  }
  return -1
}

/** Đọc mục lục ZIP từ cuối file, không nạp cả file vào RAM */
export async function listZipEntries(file: File): Promise<ZipEntry[]> {
  const tailLen = Math.min(file.size, 66_000)
  const tail = new Uint8Array(await file.slice(file.size - tailLen).arrayBuffer())
  const eocd = findEocd(tail)
  if (eocd < 0) throw new Error('File Excel không phải ZIP hợp lệ.')

  const cdSize = u32(tail, eocd + 12)
  const cdOffset = u32(tail, eocd + 16)
  if (cdSize === 0xffffffff || cdOffset === 0xffffffff) {
    throw new Error('File vượt quá 4GB. Hãy tách sheet nhỏ hơn rồi mở lại.')
  }

  const cd = new Uint8Array(await file.slice(cdOffset, cdOffset + cdSize).arrayBuffer())
  const entries: ZipEntry[] = []
  let cursor = 0
  const decoder = new TextDecoder()

  while (cursor + 46 <= cd.length && u32(cd, cursor) === 0x02014b50) {
    const method = u16(cd, cursor + 10)
    const compressedSize = u32(cd, cursor + 20)
    const nameLen = u16(cd, cursor + 28)
    const extraLen = u16(cd, cursor + 30)
    const commentLen = u16(cd, cursor + 32)
    const localOffset = u32(cd, cursor + 42)
    if (compressedSize === 0xffffffff || localOffset === 0xffffffff) {
      throw new Error('File vượt quá 4GB. Hãy tách sheet nhỏ hơn rồi mở lại.')
    }
    const name = decoder.decode(cd.subarray(cursor + 46, cursor + 46 + nameLen))
    const local = new Uint8Array(await file.slice(localOffset, localOffset + 30).arrayBuffer())
    if (u32(local, 0) !== 0x04034b50) throw new Error('ZIP thiếu header của ' + name)
    const localNameLen = u16(local, 26)
    const localExtraLen = u16(local, 28)
    entries.push({
      name,
      method,
      compressedSize,
      dataOffset: localOffset + 30 + localNameLen + localExtraLen,
    })
    cursor += 46 + nameLen + extraLen + commentLen
  }

  return entries
}

export function entryStream(
  file: File,
  entry: ZipEntry,
  onCompressedBytes?: (byteLength: number) => void,
): ReadableStream<Uint8Array> {
  let source: ReadableStream<Uint8Array> = file
    .slice(entry.dataOffset, entry.dataOffset + entry.compressedSize)
    .stream()
  if (onCompressedBytes) {
    source = source.pipeThrough(
      new TransformStream<Uint8Array, Uint8Array>({
        transform(chunk, controller) {
          onCompressedBytes(chunk.byteLength)
          controller.enqueue(chunk)
        },
      }),
    )
  }
  if (entry.method === 0) return source
  if (entry.method === 8) return source.pipeThrough(new DecompressionStream('deflate-raw'))
  throw new Error(`Không đọc được kiểu nén ${entry.method} trong file Excel.`)
}

export async function readEntryText(file: File, entry: ZipEntry): Promise<string> {
  const response = new Response(entryStream(file, entry))
  return response.text()
}
