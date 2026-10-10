# Chức năng EzXLForm

EzXLForm mở file bảng tính, sửa trên form hoặc ngay trong ô, rồi xuất lại. Dữ liệu đang làm được giữ thành phiên trên máy này.

Bản lưu cho lần sửa sau nằm ở [docs/superpowers/plans/2026-10-10-chuc-nang.md](docs/superpowers/plans/2026-10-10-chuc-nang.md).

## Mở file và phiên

- Kéo thả file vào trang, hoặc bấm **Mở File**. Định dạng nhận: `.xlsx`, `.xls`, `.csv`.
- Khi nhập, chọn dòng làm tiêu đề cột và số dòng bỏ qua ngay sau tiêu đề. Mỗi sheet sau đó vẫn đổi được dòng tiêu đề riêng.
- File được lưu thành phiên trên máy. Bấm tên phiên trên thanh trên để vào lại màn hình sửa. Đổi phiên hoặc tải lại trang không làm mất dòng đã thêm và chữ đã gõ, kể cả file lớn.
- Đóng tab hoặc tải lại trang sẽ ghi nốt các sửa đang chờ.

File từ **4 MB** trở lên (`.xlsx` và `.csv`) được chia trang: các dòng nằm trong bộ nhớ riêng của trình duyệt, bảng chỉ đọc từng trang. File nhỏ hơn giữ toàn bộ dòng trong bộ nhớ.

Giới hạn thực tế của trình duyệt:

- File nén ZIP lớn hơn khoảng 4 GB không mở được.
- Bảng chuỗi dùng chung của Excel nếu lớn hơn bộ nhớ máy thì không nạp hết.

## Bố cục màn sửa

- Màn hình chia hai: form bên trái, bảng bên phải. Kéo thanh giữa để đổi độ rộng. Khi cả hai đang hiện, form chiếm tối đa một nửa màn hình.
- **Ẩn form** / **Hiện form** và **Ẩn bảng** / **Hiện bảng** để chỉ còn một phía.
- Form tự xếp **4 cột** khi đủ rộng (từ khoảng 750 px), **2 cột** khi vừa (từ khoảng 370 px), **1 cột** khi hẹp. Số cột không bị giảm thêm chỉ vì số ô không chia hết.
- Ô chữ nhiều dòng cao lên theo nội dung rồi thấp lại khi xóa bớt. Ô số và ô ngày luôn một dòng.
- Ô ngày có bộ chọn ngày. Định dạng ngày chọn trên từng cột.

## Sửa một dòng

Bấm một hàng trên bảng để đổ dữ liệu vào form. Gõ trên form thì bảng cập nhật theo.

Khi đang sửa một dòng đã có:

- **Sao chép xuống dưới** tạo ngay bên dưới một dòng mang cùng dữ liệu và chọn dòng mới.
- **Thêm dòng mới sau #N** chèn một dòng trống ngay bên dưới dòng đang chọn.
- **Thêm ở cuối** thêm một dòng trống ở cuối bảng.

Khi chưa chọn dòng (đang tạo mới):

- **Tạo dòng mới** hoặc Enter ghi dữ liệu form vào cuối bảng.
- **Làm mới** xóa trắng form.

Sao chép và chèn được ghi ngay. Không cần bấm lưu riêng cho dòng vừa thêm.

**Xóa dòng** xóa dòng đang chọn sau khi xác nhận.

File lớn: form đợi dòng tải xong rồi mới cho sửa. Gõ khi dòng chưa tải không xóa các cột còn lại.

### Phím trên form

| Phím | Tác dụng |
| --- | --- |
| Mũi tên trái / phải | Sang ô cùng hàng. Trong ô chữ, mũi tên đi trong nội dung trước, chạm mép rồi mới sang ô kế |
| Mũi tên lên / xuống | Sang ô hàng trên hoặc hàng dưới |
| Ctrl + Enter | Sang dòng kế tiếp |
| Ctrl + Shift + Enter | Về dòng trước |
| Enter | Ở ô số dòng: nhảy tới dòng đó. Khi đang tạo dòng mới: tạo dòng |

## Sửa ngay trong ô

Bấm đúp một ô trên bảng để sửa tại chỗ. Enter lưu, Esc hủy.

## Sắp xếp hàng và cột

- Mỗi hàng có tay nắm ở cột chọn bên trái. Kéo tay nắm thả lên hàng khác để đổi thứ tự. Với file lớn, chỉ thứ tự hiển thị đổi, không ghi lại từng dòng.
- Kéo tiêu đề cột để đổi thứ tự cột. Bấm tên cột để đổi tên. Menu cột còn thêm cột hai bên, chuyển cột, sắp xếp tăng/giảm, đổi kiểu dữ liệu và xóa cột.

## Lọc, tìm và thanh tổng

- Ô tìm trên bảng lọc theo một cột hoặc tất cả cột. **Nâng cao** ghép nhiều điều kiện bằng Và / Hoặc: chứa, bằng, bắt đầu, kết thúc, giống mẫu (`%` là nhiều ký tự, `_` là một ký tự), trống, lớn hơn, nhỏ hơn.
- File lớn vẫn lọc được; bảng báo đang quét khi danh sách dài.
- Dưới bảng có thanh tổng của các cột số, khớp độ rộng cột phía trên.
- Phân trang bảng: 20, 50, 100, 200, 500 hoặc 1000 dòng mỗi trang.

## Chọn nhiều dòng

Tick ô ở đầu hàng. Khi có dòng được chọn, thanh dưới hiện các thao tác trên đúng các dòng đó:

- Gán một giá trị vào một cột
- Thêm tiền tố hoặc hậu tố
- Viết hoa, viết thường, viết hoa đầu từ, xóa khoảng trắng thừa, xóa nội dung ô
- Xóa các dòng đã chọn

## Tìm và thay thế

Ctrl+H hoặc Ctrl+F mở cửa sổ tìm. Có thể giới hạn cột, phân biệt hoa thường, khớp đúng cả ô, dùng biểu thức chính quy, hoặc tìm ô trống.

Lọc trước để xem các dòng khớp, rồi mới thay thế tất cả. Thay thế chạy nền, có phần trăm, tạm dừng và chạy tiếp.

## Tiến trình

Mở file lớn và thay thế hàng loạt vào danh sách tiến trình. Popup kéo được. Tắt popup không xóa tiến trình. Mở lại từ nút **Tiến trình** trên thanh trên.

## Sheet

Thanh dưới cùng quản lý các sheet: thêm, đổi tên, nhân bản, xóa, kéo đổi thứ tự. Mỗi sheet chọn được dòng nào là tiêu đề cột.

## Xem trước Excel

**Xem trước Excel** hiện lưới gần với file gốc, gồm cả các dòng phía trên tiêu đề. Chữ xuống dòng đủ nội dung, font Calibri 11. Sheet dài được chia trang. File lớn chỉ xem trước một đoạn đầu.

## Hoàn tác

| Phím | Tác dụng |
| --- | --- |
| Ctrl + Z | Hoàn tác |
| Ctrl + Y hoặc Ctrl + Shift + Z | Làm lại |
| F1 | Mở hoặc đóng hướng dẫn trong app |

Hoàn tác được sửa ô, thêm hoặc xóa dòng, sắp xếp hàng, sửa cột và các thao tác sheet đã ghi vào lịch sử.

## Xuất file

**Xuất** ghi lại sổ đang sửa ra `.xlsx` hoặc `.csv`. File lớn được đọc theo thứ tự đang hiển thị, gồm dòng đã chèn và hàng đã kéo thả.

## Giao diện

Menu giao diện đổi theme DaisyUI và tăng hoặc giảm cỡ chữ.

## Phím tắt chung

| Phím | Tác dụng |
| --- | --- |
| F1 | Hướng dẫn |
| Ctrl + Z | Hoàn tác |
| Ctrl + Y hoặc Ctrl + Shift + Z | Làm lại |
| Ctrl + H hoặc Ctrl + F | Tìm và thay thế |
| Esc | Hủy sửa ô, hủy đổi tên, hoặc xóa ô tìm trên bảng |
| Enter | Lưu ô đang sửa, tên cột, tên sheet, hoặc áp dụng bộ lọc nâng cao |
