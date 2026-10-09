# Excel Editor

Ứng dụng chỉnh sửa file Excel hiện đại, xây dựng với Vue 3 + DaisyUI + Tauri.

## Tính năng

- 📂 **Import/Export** — Hỗ trợ `.xlsx`, `.xls`, `.csv` với giữ nguyên định dạng
- 📋 **Đa Sheet** — Quản lý nhiều sheet: thêm, xóa, đổi tên, nhân bản, sắp xếp
- ✏️ **Chỉnh sửa linh hoạt** — Inline edit, Drawer chi tiết, quản lý cột/header toàn diện
- 🔍 **Tìm & Thay thế** — Find & Replace với phạm vi cột, case-sensitive, regex
- ⚡ **Batch Operations** — Sửa hàng loạt: gán giá trị, prefix/suffix, đổi kiểu chữ, xóa hàng loạt
- 👁️ **Split View Preview** — Xem trước Excel thời gian thực với chia đôi màn hình
- ↩️ **Undo/Redo** — Lịch sử thao tác với `Ctrl+Z` / `Ctrl+Y`
- 🎨 **Theme Tùy biến** — 30+ themes DaisyUI + Color Picker tùy chỉnh Primary/Accent color
- 🖥️ **Desktop App** — Đóng gói với Tauri v2 (Windows, macOS, Linux)

## Cài đặt & Chạy

### Yêu cầu

- Node.js >= 18
- Rust (nếu muốn build desktop app)

### Web App

```bash
npm install
npm run dev
```

Mở trình duyệt tại `http://localhost:5173`

### Desktop App (Tauri)

```bash
npm install
npm run tauri dev      # Development
npm run tauri build    # Build production .exe
```

## Cấu trúc Project

```
src/
├── assets/          # CSS global (Tailwind + DaisyUI)
├── components/
│   ├── common/      # ThemeDropdown, ConfirmDialog, ToastContainer
│   ├── editor/      # EditorTable, BatchActionBar, RowEditDrawer...
│   ├── layout/      # AppNavbar, SplitContainer, StatusBar
│   ├── modals/      # FindReplaceModal, ImportOptionsModal
│   ├── preview/     # ExcelPreview, PreviewSheetTabs
│   └── sheets/      # SheetTabsBar
├── composables/     # useExcelParser, useTheme, useHistory, useBatchOperations
├── stores/          # workbookStore, selectionStore, appConfigStore (Pinia)
└── types/           # TypeScript type definitions
```

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | Vue 3 + TypeScript + Vite |
| UI | Tailwind CSS + DaisyUI |
| State | Pinia |
| Excel Engine | ExcelJS + SheetJS (xlsx) |
| Desktop | Tauri v2 |
| Icons | Lucide Vue Next |
