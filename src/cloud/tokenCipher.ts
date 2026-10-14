const keyCache = new Map<string, Promise<CryptoKey>>()

function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary)
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
  return bytes
}

function aesKey(secret: string): Promise<CryptoKey> {
  const cached = keyCache.get(secret)
  if (cached) return cached
  const derived = crypto.subtle
    .digest('SHA-256', new TextEncoder().encode(`ezxl-google-token:${secret}`))
    .then((salt) =>
      crypto.subtle.importKey('raw', new TextEncoder().encode(secret), 'PBKDF2', false, ['deriveKey']).then((base) =>
        crypto.subtle.deriveKey(
          { name: 'PBKDF2', salt: new Uint8Array(salt), iterations: 100_000, hash: 'SHA-256' },
          base,
          { name: 'AES-GCM', length: 256 },
          false,
          ['encrypt', 'decrypt'],
        ),
      ),
    )
  keyCache.set(secret, derived)
  return derived
}

export async function sealToken(secret: string, plain: string): Promise<string> {
  const iv = crypto.getRandomValues(new Uint8Array(12))
  const cipher = new Uint8Array(
    await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await aesKey(secret), new TextEncoder().encode(plain)),
  )
  const packed = new Uint8Array(1 + iv.length + cipher.length)
  packed[0] = 1
  packed.set(iv, 1)
  packed.set(cipher, 1 + iv.length)
  return bytesToBase64(packed)
}

export async function openToken(secret: string, packed: string): Promise<string> {
  const bytes = base64ToBytes(packed)
  if (bytes[0] !== 1 || bytes.length < 14) throw new Error('Phiên Google không đọc được')
  const iv = bytes.slice(1, 13)
  const cipher = bytes.slice(13)
  const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, await aesKey(secret), cipher)
  return new TextDecoder().decode(plain)
}
