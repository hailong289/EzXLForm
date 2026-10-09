const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let index = 0; index < 256; index++) {
    let crc = index
    for (let bit = 0; bit < 8; bit++) {
      crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1
    }
    table[index] = crc >>> 0
  }
  return table
})()

function crcUpdate(crc: number, data: Uint8Array): number {
  let next = crc >>> 0
  for (let index = 0; index < data.length; index++) {
    next = CRC_TABLE[(next ^ data[index]) & 0xff] ^ (next >>> 8)
  }
  return next >>> 0
}

function u16(value: number): Uint8Array {
  return new Uint8Array([value & 0xff, (value >> 8) & 0xff])
}

function u32(value: number): Uint8Array {
  return new Uint8Array([
    value & 0xff,
    (value >> 8) & 0xff,
    (value >> 16) & 0xff,
    (value >> 24) & 0xff,
  ])
}

function concat(parts: Uint8Array[]): Uint8Array {
  const size = parts.reduce((sum, part) => sum + part.length, 0)
  const out = new Uint8Array(size)
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

interface ZipItem {
  name: Uint8Array
  crc: number
  size: number
  offset: number
  flags: number
}

/** Ghi ZIP kiểu stored (không nén) thẳng ra đĩa */
export class StoredZipWriter {
  private offset = 0
  private items: ZipItem[] = []
  private encoder = new TextEncoder()

  constructor(private write: (chunk: Uint8Array) => Promise<void>) {}

  async addText(name: string, text: string): Promise<void> {
    await this.addBytes(name, this.encoder.encode(text))
  }

  async addBytes(name: string, data: Uint8Array): Promise<void> {
    const crc = crcUpdate(0xffffffff, data) ^ 0xffffffff
    await this.begin(name, crc, data.length)
    await this.write(data)
    this.offset += data.length
  }

  async streamText(
    name: string,
    produce: (push: (text: string) => Promise<void>) => Promise<void>,
  ): Promise<void> {
    const encoded = this.encoder.encode(name)
    const header = concat([
      u32(0x04034b50),
      u16(20),
      u16(0x0008),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(0),
      u32(0),
      u16(encoded.length),
      u16(0),
      encoded,
    ])
    const localOffset = this.offset
    await this.write(header)
    this.offset += header.length
    let crc = 0xffffffff
    let size = 0
    await produce(async (text) => {
      const bytes = this.encoder.encode(text)
      crc = crcUpdate(crc, bytes)
      size += bytes.length
      await this.write(bytes)
      this.offset += bytes.length
    })
    const finalCrc = (crc ^ 0xffffffff) >>> 0
    const descriptor = concat([u32(0x08074b50), u32(finalCrc), u32(size), u32(size)])
    await this.write(descriptor)
    this.offset += descriptor.length
    this.items.push({ name: encoded, crc: finalCrc, size, offset: localOffset, flags: 0x0008 })
  }

  private async begin(name: string, crc: number, size: number): Promise<void> {
    const encoded = this.encoder.encode(name)
    const header = concat([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(crc),
      u32(size),
      u32(size),
      u16(encoded.length),
      u16(0),
      encoded,
    ])
    this.items.push({ name: encoded, crc, size, offset: this.offset, flags: 0 })
    await this.write(header)
    this.offset += header.length
  }

  async close(): Promise<void> {
    const cdOffset = this.offset
    for (const item of this.items) {
      const header = concat([
        u32(0x02014b50),
        u16(20),
        u16(20),
        u16(item.flags),
        u16(0),
        u16(0),
        u16(0),
        u32(item.crc),
        u32(item.size),
        u32(item.size),
        u16(item.name.length),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(0),
        u32(item.offset),
        item.name,
      ])
      await this.write(header)
      this.offset += header.length
    }
    const cdSize = this.offset - cdOffset
    const eocd = concat([
      u32(0x06054b50),
      u16(0),
      u16(0),
      u16(this.items.length),
      u16(this.items.length),
      u32(cdSize),
      u32(cdOffset),
      u16(0),
    ])
    await this.write(eocd)
  }
}
