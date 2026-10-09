<script setup lang="ts">
import { Table } from 'lucide-vue-next'
import type { SheetData } from '@/types/excel'

const props = defineProps<{
  sheets: SheetData[]
  activeSheetId: string
}>()

const emit = defineEmits<{
  select: [sheetId: string]
}>()
</script>

<template>
  <div
    class="flex items-center bg-base-200 border-t border-base-300 h-7 px-2 overflow-x-auto scrollbar-thin select-none"
  >
    <div class="flex items-center gap-1">
      <button
        v-for="sheet in props.sheets"
        :key="sheet.id"
        class="sheet-tab flex items-center h-6 px-2.5 rounded-t text-xs cursor-pointer border-t-2 transition-colors whitespace-nowrap"
        :class="
          sheet.id === props.activeSheetId
            ? 'bg-base-100 border-primary font-bold shadow-xs'
            : 'bg-base-300/40 hover:bg-base-300 border-transparent text-base-content/70'
        "
        @click="emit('select', sheet.id)"
      >
        <Table class="w-3 h-3 mr-1 opacity-60 flex-shrink-0" />
        <span class="truncate max-w-[100px]">{{ sheet.name }}</span>
      </button>
    </div>
  </div>
</template>
