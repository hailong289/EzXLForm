import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { imageLabel, imageSource, imageSources } from './imageSource.ts'

describe('image source', () => {
  it('accepts image urls and data urls', () => {
    const remote = 'https://cdn.example.com/a/photo.PNG?size=2'
    assert.equal(imageSource(remote), remote)
    assert.equal(imageSource('  https://cdn.example.com/a/photo.webp#x  '), 'https://cdn.example.com/a/photo.webp#x')
    assert.equal(imageSource('data:image/png;base64,aaaa'), 'data:image/png;base64,aaaa')
    assert.equal(
      imageSource('https://cdn.example.com/upl\noads/photo.png'),
      'https://cdn.example.com/uploads/photo.png',
    )
  })

  it('rejects text that is not an image', () => {
    assert.equal(imageSource('Nguyễn Văn An'), null)
    assert.equal(imageSource('https://example.com/report.pdf'), null)
    assert.equal(imageSource('https://example.com/page'), null)
    assert.equal(imageSource('notes.png'), null)
    assert.equal(imageSource(12), null)
  })

  it('accepts a comma-separated list of images', () => {
    const a = 'https://cdn.example.com/a/photo.png'
    const b = 'https://cdn.example.com/b/other.jpg?w=1'
    assert.deepEqual(imageSources(`${a}, ${b}`), [a, b])
    assert.deepEqual(imageSources(`${a},${b}`), [a, b])
    assert.deepEqual(imageSources(`data:image/png;base64,aaaa, ${a}`), ['data:image/png;base64,aaaa', a])
    assert.deepEqual(imageSources(a), [a])
    assert.deepEqual(imageSources(`${a}, notes`), [])
    assert.deepEqual(imageSources('Nguyễn Văn An, hello'), [])
  })

  it('uses the file name as the quick look title', () => {
    assert.equal(imageLabel('https://cdn.example.com/a/photo.png'), 'photo.png')
    assert.equal(imageLabel('data:image/png;base64,aaaa'), 'image')
  })
})
