# Gợi Ý Đặt Hàng — Tài liệu dự án cho Claude

## Tổng quan

Desktop app Python/tkinter giúp team mua hàng **tính toán số lượng cần đặt** dựa trên lịch sử bán hàng và tồn kho hiện tại. App dùng nội bộ tại công ty Bluecircle.

## Cấu trúc file

```
main.py                          # Toàn bộ logic: UI + xử lý + xuất Excel (file duy nhất)
requirements.txt                 # pandas, openpyxl, Pillow
build_windows.bat                # Build file .exe bằng PyInstaller cho Windows
File_template/
  Bluecircle.png                 # Logo công ty (hiện trên header app)
  data.warehouse (42).xlsx       # File mẫu bán hàng (để test)
  stock.quant (13).xlsx          # File mẫu tồn kho (để test)
```

## Input: 2 file Excel

### File Bán Hàng (`data.warehouse` export từ Odoo)
Các cột bắt buộc:
- `SKU`, `Product Item`, `Brand`, `Category`, `Date`, `Quantity`
- `Sale Team`, `Revenue`
- Cột tuỳ chọn (nếu có): `Model`, `Color`, `Frame Size`, `Sub Category`

Lọc tự động: bỏ dòng có SKU chứa `COUPON`/`DISCOUNT` và Category = `SERVICE`.

### File Tồn Kho (`stock.quant` export từ Odoo)
Các cột bắt buộc:
- `Sản phẩm/Mã nội bộ` (SKU), `Số lượng`
- Tuỳ chọn: `Số lượng bảo lưu`, `Sản phẩm/Brand/Display Name`, `Sản phẩm/Nhóm Điểm bán lẻ/Tên hiển thị`

Nhiều dòng/địa điểm lưu kho sẽ được **cộng dồn** theo SKU.

## Logic tính toán chính (hàm `_process`)

1. Đọc & làm sạch file bán hàng
2. Lọc Brand/Category nếu người dùng chọn
3. Pivot theo tháng → tính `TB Tháng (6T Gần Nhất)`
4. Tính thống kê theo năm: `Tổng YYYY`, `TB YYYY`
5. Đọc tồn kho → `Tồn Khả Dụng = Tồn Kho − Tồn Bảo Lưu`
6. **Công thức gợi ý:**
   ```
   Gợi Ý = TB_tháng(6T GN) × (tồn_tối_thiểu + leadtime) − Tồn_Khả_Dụng
   Gợi Ý = max(0, Gợi Ý)
   ```
7. Nhận Xét: `Cần Đặt` / `Đủ Hàng` / `Chưa Có Lịch Sử Bán`
8. SKU chỉ có trong tồn kho (không có lịch sử bán) vẫn được giữ lại

## Output: file Excel

- Lưu vào **Desktop** với tên `GoiYDatHang_{months}T_{brand/cat tag}_{timestamp}.xlsx`
- Sắp xếp theo Doanh Thu Tổng giảm dần
- **Row 1**: SUBTOTAL (tính động khi filter Excel)
- **Row 2**: Header với màu theo nhóm cột
- Màu nhóm cột:
  - Xanh dương đậm: cột cố định (SKU, tên, brand, category)
  - Xanh dương: dữ liệu từng tháng
  - Tím: Tổng/TB theo năm
  - Xanh lá: Gợi Ý Đặt Hàng
  - Vàng/cam: Tồn kho, Doanh thu, Đặt Thực
  - Xám: Nhận Xét
- Cột `Đặt Thực`: để trống để người dùng tự nhập

## Cấu trúc code trong `main.py`

| Phần | Mô tả |
|------|-------|
| `SALES`, `INV` dict | Mapping tên cột file input |
| `strip_sku_prefix()` | Bỏ prefix `[xxx]` trong tên sản phẩm |
| `make_card()`, `make_btn()` | Widget tái sử dụng theo Apple style |
| `LoadingPopup` | Popup spinner khi xử lý |
| `FilterPanel` | Checkbox panel có search, scroll, tick tất cả |
| `App._build_ui()` | Xây dựng toàn bộ giao diện |
| `App._read_files()` | Đọc file & populate bộ lọc Brand/Category |
| `App.run()` | Entry point khi bấm nút "Tạo Gợi Ý" |
| `App._process()` | Toàn bộ logic tính toán |
| `App._export()` | Xuất Excel có format màu sắc |

## UI Design

- **Phong cách**: Apple / iOS (màu `#007AFF`, nền `#F2F2F7`, card trắng)
- Font: `.AppleSystemUIFont` (macOS) / `Helvetica Neue` (Windows)
- Layout: header tối + body scrollable 2 cột (controls trái, log phải)
- Log console: nền tối `#1C1C1E`, chữ xanh lá `#32D74B`

## Build & Deploy

**Chạy local (macOS/Windows):**
```bash
pip install -r requirements.txt
python main.py
```

**Build exe cho Windows:**
```bat
build_windows.bat
```
PyInstaller cần các hidden-import: `PIL`, `PIL.Image`, `PIL.ImageTk`, `PIL._imagingtk`.
File exe cần có `File_template/Bluecircle.png` đi kèm (app dùng `sys._MEIPASS` để tìm).

## Lưu ý khi chỉnh sửa

- Tên cột file input được định nghĩa tập trung ở dict `SALES` và `INV` (dòng ~56–76) — nếu Odoo thay đổi tên cột, chỉ cần sửa ở đây.
- Công thức gợi ý nằm ở `_process()` dòng ~921–926 và được viết lại dưới dạng Excel formula ở `_export()` hàm `excel_formula()`.
- Khi thêm cột mới vào output, cần cập nhật: thứ tự cột `ordered` list (~942), dict `GRP_STYLE` và hàm `grp()` để có màu đúng, và `NOT_SUM` nếu không muốn SUBTOTAL.
