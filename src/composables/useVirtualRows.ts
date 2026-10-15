import { computed, isRef, type Ref } from 'vue'
import { useVirtualizer } from '@tanstack/vue-virtual'

export interface VirtualRow {
  index: number
  start: number
  size: number
}

/**
 * Cửa sổ cuộn dùng chung: chỉ những dòng trong vùng nhìn, cộng overscan, được vẽ.
 * Chiều cao thanh cuộn vẫn bằng tổng số dòng.
 */
export function useVirtualRows(input: {
  count: Ref<number>
  estimateSize: Ref<number> | number
  scrollElement: Ref<HTMLElement | null>
  measure?: boolean
  overscan?: number
}) {
  const virtualizer = useVirtualizer(
    computed(() => {
      const size = isRef(input.estimateSize) ? input.estimateSize.value : input.estimateSize
      return {
      count: input.count.value,
      getScrollElement: () => input.scrollElement.value,
      estimateSize: () => size,
      overscan: input.overscan ?? 8,
      measureElement: input.measure
        ? (element: HTMLElement) => element.getBoundingClientRect().height
        : undefined,
      }
    }),
  )

  const items = computed<VirtualRow[]>(() =>
    virtualizer.value.getVirtualItems().map((item) => ({
      index: item.index,
      start: item.start,
      size: item.size,
    })),
  )

  const totalSize = computed(() => virtualizer.value.getTotalSize())

  const paddingStart = computed(() => items.value[0]?.start ?? 0)

  const paddingEnd = computed(() => {
    const rows = items.value
    const total = totalSize.value
    if (rows.length === 0) return total
    const last = rows[rows.length - 1]
    return Math.max(0, total - (last.start + last.size))
  })

  function scrollToIndex(index: number) {
    const count = input.count.value
    if (count <= 0) return
    const next = Math.min(Math.max(0, index), count - 1)
    virtualizer.value.scrollToIndex(next, { align: 'auto' })
  }

  function measureElement(node: unknown) {
    if (!(node instanceof HTMLElement)) return
    virtualizer.value.measureElement(node)
  }

  return {
    items,
    paddingStart,
    paddingEnd,
    totalSize,
    scrollToIndex,
    measureElement: input.measure ? measureElement : undefined,
  }
}
