# Kế hoạch Triển khai: Ứng dụng Quản trị & Chỉnh sửa Excel (Vue 3 + DaisyUI + Tauri)

Ứng dụng web và desktop hiện đại cho phép import, chỉnh sửa bảng tính Excel với đầy đủ thao tác CRUD dòng, quản lý nhiều sheet, chỉnh sửa header cột, tìm kiếm & thay thế hàng loạt (Batch Edit), giao diện xem trước Excel thời gian thực theo cơ chế Split View, và hệ thống giao diện Sáng/Tối cùng Color Picker tùy biến linh hoạt.

---

## 1. Tổng quan Kiến trúc Kỹ thuật

| Thành phần | Công nghệ / Thư viện | Vai trò |
| :--- | :--- | :--- |
| **Frontend Framework** | **Vue 3** (Composition API, `<script setup>`) + **TypeScript** | Xây dựng UI linh hoạt, phản ứng nhanh (reactive), type-safe |
| **Build Tool** | **Vite** | Tốc độ dev server cực nhanh, tối ưu hóa bundle |
| **Desktop Wrapper** | **Tauri v2** | Đóng gói ứng dụng desktop Windows (.exe) siêu nhẹ, tốn ít tài nguyên |
| **Styling & Components**| **Tailwind CSS** + **DaisyUI** | Hệ thống UI hiện đại, component đẹp mắt, hỗ trợ hệ thống theme mạnh mẽ |
| **State Management** | **Pinia** | Quản lý trạng thái Workbook, Sheet hiện tại, Lịch sử Undo/Redo, Theme |
| **Excel Engine** | **ExcelJS** + **xlsx (SheetJS)** | Parse file Excel (.xlsx, .xls, .csv), giữ nguyên định dạng, xuất file chuẩn xác |
| **Icons & Utilities** | **Lucide Vue Next** + **Lodash-es** + **File-saver** | Bộ icon hiện đại, xử lý debounce tìm kiếm, xuất file |

---

## 2. Các Phân hệ Tính năng Chính

### 2.1. Quản lý File & Đa Sheet (Multi-sheet Management)
- **Import File**: Hỗ trợ kéo thả (drag & drop) hoặc click chọn file `.xlsx`, `.xls`, `.csv`.
- **Tùy chọn Header Row**: Cho phép người dùng chọn dòng làm tiêu đề (dòng 1, dòng 2,...) khi import.
- **Thanh Sheet Tabs (Dưới cùng hoặc trên cùng)**:
  - Hiển thị danh sách sheet dưới dạng các tab chuyển đổi nhanh.
  - Thêm sheet mới (+), đổi tên sheet (inline edit), nhân bản (clone), xóa sheet.
  - Đổi thứ tự sheet bằng nút di chuyển.
- **Export File**: Xuất file ra `.xlsx` hoặc `.csv` chất lượng cao, giữ màu sắc và công thức.

### 2.2. Trình chỉnh sửa Dữ liệu (Editor Table) & CRUD Dòng
- **CRUD Dòng**:
  - Thêm dòng mới (chèn ở đầu bảng, cuối bảng, hoặc vị trí đang chọn).
  - Xóa 1 dòng hoặc xóa nhiều dòng đã chọn qua checkbox.
  - Sửa dòng: Hỗ trợ cả **Inline Edit** (click đúp vào ô để sửa) và **Quick Edit Drawer/Modal** (mở form chi tiết các trường của dòng).
- **Quản lý Cột & Header**:
  - Đổi tên tiêu đề cột trực tiếp (double-click hoặc popover).
  - Thêm cột mới (chèn bên trái / bên phải cột hiện tại).
  - Xóa cột, sắp xếp thứ tự cột (Move Left / Move Right).
  - Thiết lập kiểu dữ liệu cột (Text, Number, Date, Boolean).

### 2.3. Tìm kiếm & Sửa hàng loạt (Search & Batch Operations)
- **Bộ lọc & Tìm kiếm**:
  - Ô tìm kiếm toàn cục (Global search) tìm trên tất cả các cột.
  - Lọc theo từng cột cụ thể (Column filter).
- **Tìm & Thay thế (Find & Replace)**:
  - Hộp thoại tìm từ khóa và thay thế bằng giá trị mới.
  - Tùy chọn phạm vi: Toàn bộ sheet hoặc chỉ trên các cột được chỉ định.
  - Hỗ trợ phân biệt hoa/thường (Case-sensitive) và khớp toàn bộ ô (Exact match).
- **Thanh tác vụ hàng loạt (Batch Action Bar)**:
  - Khi tích chọn nhiều dòng (checkbox), xuất hiện thanh công cụ nổi:
    - *Gán giá trị chung*: Đặt giá trị mới cho một cột cụ thể cho tất cả các dòng đã chọn.
    - *Thêm tiền tố/hậu tố (Prefix/Suffix)*: Thêm chuỗi vào đầu/cuối giá trị hiện có.
    - *Chuyển đổi kiểu chữ*: Viết hoa (UPPERCASE), viết thường (lowercase), viết hoa chữ cái đầu (Capitalize).
    - *Xóa hàng loạt*: Xóa toàn bộ dòng đã chọn với hộp thoại xác nhận.

### 2.4. Chế độ Xem trước Excel Thời gian thực (Split View Realtime Preview)
- **Bố cục 3 chế độ xem (View Modes)**:
  - **Split View (Mặc định)**: Chia đôi màn hình 50/50 (có thanh trượt resize độ rộng) giữa Editor Table và Excel Preview.
  - **Focus Editor**: Mở rộng toàn màn hình bảng chỉnh sửa.
  - **Focus Preview**: Mở rộng toàn màn hình xem trước.
- **Excel Preview Component**:
  - Hiển thị bảng tính mô phỏng định dạng Excel thực tế (chỉ số cột A, B, C..., chỉ số dòng 1, 2, 3...).
  - Tự động đồng bộ thời gian thực theo từng thay đổi từ Editor.
  - Chuyển sheet preview đồng bộ với sheet đang chỉnh sửa hoặc xem độc lập.

### 2.5. Hoàn tác & Làm lại (Undo / Redo History)
- Ngăn xếp lịch sử (History Stack) theo dõi mọi hành động: thêm dòng, xóa dòng, sửa ô, batch replace, sửa header.
- Phím tắt tiện lợi: `Ctrl + Z` (Undo) và `Ctrl + Y` / `Ctrl + Shift + Z` (Redo).
- Nút bấm Undo/Redo trực quan trên thanh điều khiển kèm số bước có thể hoàn tác.

### 2.6. Hệ thống Giao diện DaisyUI & Tùy biến Màu sắc (Theme & Color Customizer)
- **Chế độ Sáng / Tối**: Nút chuyển đổi nhanh Dark/Light tiện dụng.
- **Preset DaisyUI Themes**: Menu chọn các chủ đề yêu thích (Light, Dark, Corporate, Emerald, Synthwave, Cyberpunk, Dracula, Cupcake, Nord, Sunset,...).
- **Color Customizer (Bảng màu tùy biến)**:
  - Chọn màu chủ đạo (Primary Color, Secondary, Accent) qua Color Picker.
  - Tự động tạo mã HSL/HEX và cập nhật CSS Variables của DaisyUI theo thời gian thực.
  - Tự động lưu cấu hình giao diện vào `localStorage`, tự khôi phục khi mở lại app.

---

## 3. Cấu trúc Dự án Đề xuất

```
edit-excel/
├── src-tauri/                 # Cấu hình Tauri v2 cho Desktop App
│   ├── Cargo.toml
│   ├── tauri.conf.json
│   └── src/main.rs
├── src/
│   ├── assets/                # CSS, SVG, Icons
│   │   └── main.css           # Tailwind + DaisyUI base styling
│   ├── components/
│   │   ├── common/            # Modal, Toast, ColorPicker, ThemeSelector
│   │   │   ├── ColorPickerModal.vue
│   │   │   ├── ThemeDropdown.vue
│   │   │   └── ConfirmDialog.vue
│   │   ├── editor/            # Phân hệ bảng chỉnh sửa
│   │   │   ├── EditorTable.vue
│   │   │   ├── TableHeaderCell.vue
│   │   │   ├── TableRowCell.vue
│   │   │   ├── RowEditDrawer.vue
│   │   │   └── BatchActionBar.vue
│   │   ├── preview/           # Phân hệ xem trước Excel
│   │   │   ├── ExcelPreview.vue
│   │   │   └── PreviewSheetTabs.vue
│   │   ├── sheets/            # Quản lý Tabs đa sheet
│   │   │   └── SheetTabsBar.vue
│   │   ├── modals/            # Các hộp thoại nghiệp vụ
│   │   │   ├── FindReplaceModal.vue
│   │   │   ├── HeaderManageModal.vue
│   │   │   └── ImportOptionsModal.vue
│   │   └── layout/            # Layout chính và Navbar
│   │       ├── AppNavbar.vue
│   │       ├── SplitContainer.vue
│   │       └── StatusBar.vue
│   ├── composables/           # Logic tái sử dụng
│   │   ├── useExcelParser.ts  # Đọc, ghi và xử lý file với ExcelJS & SheetJS
│   │   ├── useBatchOperations.ts # Xử lý Find & Replace, Batch Update
│   │   ├── useHistory.ts      # Undo / Redo state tracking
│   │   └── useTheme.ts        # Quản lý theme DaisyUI & custom color variables
│   ├── stores/                # Pinia Stores
│   │   ├── workbookStore.ts   # Lưu trữ workbook, sheets, activeSheet, cell data
│   │   ├── selectionStore.ts  # Dòng/cột đang chọn, checkbox selection
│   │   └── appConfigStore.ts  # View mode (Split/Editor/Preview), split ratio
│   ├── types/                 # TypeScript interfaces
│   │   └── excel.ts
│   ├── App.vue
│   └── main.ts
├── index.html
├── package.json
├── tailwind.config.js
├── vite.config.ts
└── tsconfig.json
```

---

## 4. Kế hoạch Triển khai Từng bước

### Bước 1: Khởi tạo Project & Cài đặt Dependencies
- Khởi tạo Vue 3 + TypeScript + Vite.
- Cấu hình Tailwind CSS, DaisyUI (kèm danh sách toàn bộ các theme).
- Cài đặt các thư viện lõi: `exceljs`, `xlsx`, `pinia`, `lucide-vue-next`, `lodash-es`.
- Cấu hình template Tauri v2 (`src-tauri`) để ứng dụng sẵn sàng build Desktop.

### Bước 2: Xây dựng Module Theme & Color Customizer
- Triển khai `useTheme.ts` và component `ThemeDropdown.vue` + `ColorPickerModal.vue`.
- Cấu hình CSS Variables cho DaisyUI (`--p`, `--pf`, `--b1`, `--bc`,...) để đổi màu tự do theo color picker.
- Tích hợp lưu/khôi phục từ `localStorage`.

### Bước 3: Xây dựng Core State & Xử lý File Excel (Excel Engine)
- Thiết kế Data Model: `Workbook`, `Sheet`, `ColumnHeader`, `RowData`, `CellFormat`.
- Xây dựng `useExcelParser.ts`:
  - Hàm `readExcelFile(file)` đọc file bằng `exceljs` / `xlsx`.
  - Hàm `exportExcelFile(workbook)` xuất file giữ công thức và format.
- Xây dựng Pinia store `workbookStore.ts` với đầy đủ actions.
- Tích hợp lịch sử Undo/Redo (`useHistory.ts`) hỗ trợ phím tắt `Ctrl+Z`, `Ctrl+Y`.

### Bước 4: Xây dựng Giao diện Bảng Chỉnh sửa (Editor Table & Header Editor)
- Component `SheetTabsBar.vue`: Chuyển sheet, thêm (+), đổi tên, xóa, clone sheet.
- Component `EditorTable.vue`:
  - Hiển thị danh sách hàng theo phân trang hoặc cuộn mượt mà.
  - Tích hợp checkbox chọn nhiều dòng (Select all, Invert selection).
  - Chỉnh sửa header: Đổi tên header (double click / popover), thêm/xóa cột, đổi vị trí cột.
  - Chỉnh sửa dòng: Inline edit ô dữ liệu & thêm dòng mới.
  - `RowEditDrawer.vue`: Drawer sửa nhanh chi tiết một dòng.

### Bước 5: Xây dựng Tìm kiếm & Sửa hàng loạt (Batch Operations)
- Thanh tìm kiếm và bộ lọc cột theo thời gian thực (debounced).
- Hộp thoại `FindReplaceModal.vue`: Tìm và thay thế chuỗi trên nhiều cột hoặc toàn sheet.
- Thanh công cụ nổi `BatchActionBar.vue`: Gán giá trị chung, thêm prefix/suffix, đổi kiểu chữ, xóa nhiều dòng.

### Bước 6: Xây dựng Realtime Excel Preview & Bố cục Split View
- Component `SplitContainer.vue`: Phân chia 2 cột với thanh kéo điều chỉnh tỉ lệ (Resizer draggable), các nút chuyển chế độ Split / Full Editor / Full Preview.
- Component `ExcelPreview.vue`: Render bảng tính Excel với chỉ số cột chữ (A, B, C...) và dòng số (1, 2, 3...) phản ánh tức thì các chỉnh sửa từ Editor.

### Bước 7: Kiểm thử & Tối ưu hóa (Verification)
- Kiểm tra nhập/xuất các file Excel mẫu (.xlsx có nhiều sheet, nhiều cột).
- Kiểm tra tính năng Undo/Redo khi thao tác hàng loạt.
- Kiểm tra tính năng đổi theme DaisyUI và tùy chỉnh Color Picker.
- Kiểm tra build sản phẩm web (`npm run build`) và cấu hình Desktop Tauri.

---

## 5. Kế hoạch Kiểm thử & Xác minh (Verification Plan)

### Automated Checks
- Kiểm tra type check: `npm run type-check` hoặc `vue-tsc --noEmit`.
- Kiểm tra build production: `npm run build` không phát sinh lỗi bundle.

### Manual Verification Flows
1. **Kiểm tra File I/O**:
   - Tải file Excel có ít nhất 3 sheets lên hệ thống.
   - Thêm sheet mới, đổi tên sheet hiện có, chuyển qua lại giữa các sheet.
   - Xuất file và mở lại kiểm tra tính nguyên vẹn của dữ liệu.
2. **Kiểm tra CRUD Dòng & Header**:
   - Thêm dòng mới, sửa ô trực tiếp bằng double click, xóa dòng.
   - Đổi tên header cột, chèn thêm cột mới, xóa cột.
   - Bấm `Ctrl+Z` để kiểm tra hoàn tác.
3. **Kiểm tra Tìm kiếm & Batch Edit**:
   - Tìm kiếm từ khóa và lọc danh sách.
   - Dùng Find & Replace thay thế một từ khóa xuất hiện ở nhiều dòng.
   - Tích chọn 5 dòng, dùng Batch Action Bar để gán cùng một giá trị cho cột trạng thái.
4. **Kiểm tra Split View Preview**:
   - Kéo thanh resizer để thay đổi tỉ lệ khung nhìn.
   - Sửa 1 ô ở Editor Table và quan sát sự thay đổi ngay lập tức ở khung Excel Preview.
5. **Kiểm tra Giao diện & Theme**:
   - Đổi qua lại giữa các theme Dark, Light, Cyberpunk, Emerald,...
   - Mở Color Picker đổi màu Primary sang màu tùy ý, F5 tải lại trang để kiểm tra màu vẫn được lưu.
