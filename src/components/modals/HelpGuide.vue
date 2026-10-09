<script setup lang="ts">
import { onMounted, onUnmounted } from 'vue'
import { X } from 'lucide-vue-next'

const open = defineModel<boolean>({ required: true })

const shortcutGroups = [
  {
    title: 'Chung',
    rows: [
      ['F1', 'Mở hoặc đóng hướng dẫn'],
      ['Ctrl + Z', 'Hoàn tác'],
      ['Ctrl + Y hoặc Ctrl + Shift + Z', 'Làm lại'],
      ['Ctrl + H hoặc Ctrl + F', 'Mở tìm kiếm và thay thế'],
    ],
  },
  {
    title: 'Form sửa dòng',
    rows: [
      ['← →', 'Sang ô cùng hàng. Trong ô chữ, mũi tên đi trong nội dung trước, chạm mép rồi mới sang ô kế'],
      ['↑ ↓', 'Sang ô hàng trên hoặc hàng dưới'],
      ['Ctrl + Enter', 'Sang dòng kế tiếp'],
      ['Ctrl + Shift + Enter', 'Về dòng trước'],
      ['Enter', 'Ở ô số dòng: nhảy tới dòng đó. Khi đang tạo dòng mới: tạo dòng'],
    ],
  },
  {
    title: 'Bảng và sheet',
    rows: [
      ['Bấm đúp ô', 'Sửa trực tiếp trong ô. Enter lưu, Esc hủy'],
      ['Enter', 'Lưu ô đang sửa, tên cột, tên sheet, hoặc áp dụng bộ lọc nâng cao'],
      ['Esc', 'Hủy sửa ô, hủy đổi tên, hoặc xóa ô tìm trong bảng'],
    ],
  },
]

function onKeydown(event: KeyboardEvent) {
  if (event.key !== 'F1') return
  event.preventDefault()
  open.value = !open.value
}

onMounted(() => window.addEventListener('keydown', onKeydown))
onUnmounted(() => window.removeEventListener('keydown', onKeydown))
</script>

<template>
  <dialog v-if="open" class="modal modal-open">
    <div class="modal-box max-w-2xl border border-base-300 p-0">
      <div class="flex items-center justify-between border-b border-base-300 px-4 py-3">
        <h3 class="text-base font-bold">Hướng dẫn sử dụng</h3>
        <button type="button" class="btn btn-ghost btn-xs btn-circle" title="Đóng" @click="open = false">
          <X class="h-4 w-4" />
        </button>
      </div>

      <div class="max-h-[70vh] space-y-4 overflow-auto px-4 py-3 text-sm">
        <section class="space-y-1">
          <h4 class="font-semibold">Mở file và phiên</h4>
          <p class="text-base-content/80">
            Kéo thả hoặc bấm Mở File để nạp .xlsx, .xls, .csv. File được lưu thành phiên trên máy này.
            Bấm tên phiên đang mở để vào lại màn hình sửa.
          </p>
        </section>

        <section class="space-y-1">
          <h4 class="font-semibold">Sửa dữ liệu</h4>
          <p class="text-base-content/80">
            Chọn một dòng trên bảng để đổ vào form. Bấm đúp một ô để sửa ngay trong ô đó.
            Gõ trên form là bảng cập nhật ngay.
            Có thể đổi tên cột, kéo để đổi thứ tự cột, lọc dòng và chọn dòng tiêu đề của sheet.
          </p>
        </section>

        <section class="space-y-1">
          <h4 class="font-semibold">Xem trước Excel</h4>
          <p class="text-base-content/80">
            Nút Xem trước Excel hiện toàn bộ dòng của sheet, gồm cả dòng phía trên tiêu đề.
            Chữ xuống dòng đủ nội dung và dùng font Calibri 11. Sheet dài được chia trang.
          </p>
        </section>

        <section class="space-y-1">
          <h4 class="font-semibold">Tìm, thay thế và tiến trình</h4>
          <p class="text-base-content/80">
            Lọc trước để xem danh sách khớp, rồi mới thay thế tất cả. Thay thế chạy nền, có phần trăm,
            tạm dừng và chạy tiếp. File lớn khi mở cũng vào danh sách tiến trình.
            Popup tiến trình kéo được; tắt popup không xóa tiến trình. Mở lại từ menu Tiến trình.
          </p>
        </section>

        <section class="space-y-2">
          <h4 class="font-semibold">Bảng phím tắt</h4>
          <div v-for="group in shortcutGroups" :key="group.title">
            <p class="mb-1 text-xs font-medium uppercase tracking-wide opacity-60">{{ group.title }}</p>
            <table class="table table-xs table-zebra">
              <thead>
                <tr>
                  <th class="w-56">Phím</th>
                  <th>Tác dụng</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="row in group.rows" :key="row[0]">
                  <td><kbd class="kbd kbd-xs">{{ row[0] }}</kbd></td>
                  <td>{{ row[1] }}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </div>
    <div class="modal-backdrop bg-black/40" @click="open = false" />
  </dialog>
</template>
