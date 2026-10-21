# Tra cứu thông số: hướng dẫn

Công cụ chạy hoàn toàn trên trình duyệt của máy người dùng. File Excel **không gửi lên mạng**.

## Chạy trên máy (Win 10 / Win 11)

1. Chép cả thư mục `WebExcel` vào máy, ví dụ `D:\TraCuuThongSo`.
2. Bấm đúp `Mo-cong-cu.bat`. Công cụ mở bằng Microsoft Edge, có sẵn trên Windows.
   Có thể mở `index.html` bằng Chrome. **Không dùng Firefox**, vì Firefox không lưu đè được vào file gốc.
3. Không cần cài gì thêm và không cần mạng.

## Đưa lên web (nếu muốn)

Thư mục này là web tĩnh, chỉ cần chép lên bất kỳ nơi nào phục vụ file tĩnh:
VPS (nginx), GitHub Pages, Cloudflare Pages, Netlify... **Bắt buộc dùng https.**
Server không phải xử lý gì, nên nhiều người dùng cùng lúc cũng không lag.

## Chức năng (bản 1)

| Việc | Cách làm |
|---|---|
| Mở file | Nút **Mở file Excel** / **Mở thư mục**, hoặc kéo thả. Lần sau dùng **Mở lại lần trước**. |
| Tìm kiếm | Gõ tên, mã, model, thông số, không cần dấu. Tìm trên mọi file đang mở. |
| Sửa | Bấm vào dòng, sửa ở ngăn bên phải, rồi **Lưu vào file** (Ctrl+S). Ô đã sửa tô vàng, có nút hoàn tác. |
| Sao lưu | Trước lần lưu đầu, bản gốc được tự cất trong mục **Bản sao lưu** (5 bản/file). |
| Lịch sử | **Lịch sử sửa** ghi lại ai sửa ô nào, giá trị cũ và mới. |
| Xuất | Tích chọn SP, bấm **Xuất file Excel**. Giữ định dạng gốc, đánh lại STT, tự chỉnh công thức tổng. Chọn từ nhiều file thì ghép theo file mẫu. |
| Nhận sai cột | Bấm **Cột** cạnh tên sheet để chỉnh. Cấu hình được nhớ cho các file cùng kiểu. |

## Giới hạn cần biết

- Chỉ đọc được `.xlsx`. File `.xls` (Excel 2003) hoặc `.xlsm`: mở bằng Excel, chọn *Lưu thành*, rồi chọn `.xlsx`.
- **Biểu đồ, Pivot, macro, khung chữ (shape)** có thể mất khi lưu đè hoặc xuất. Công cụ sẽ cảnh báo khi file có những thứ này. Định dạng thường (màu, viền, ô gộp, công thức, logo/ảnh, độ rộng cột) được giữ.
- Ô công thức (ví dụ Thành tiền) không sửa trực tiếp. Excel tự tính lại khi mở file.
- Hãy **đóng file trong Excel trước khi bấm Lưu**, nếu không Windows sẽ khoá file.
- Bản sao lưu và lịch sử nằm trong trình duyệt của từng máy. Xoá dữ liệu duyệt web thì chúng cũng mất theo.

## Cấu trúc thư mục

```
index.html          giao diện
Mo-cong-cu.bat      mở bằng Edge
css/app.css         kiểu dáng
js/core.js          lõi: nhận diện bảng, sửa ô, xuất file giữ định dạng
js/store.js         lưu sao lưu / lịch sử / file gần đây (IndexedDB)
js/app.js           điều khiển giao diện
vendor/             ExcelJS 4.4.0, JSZip 3.10.1 (chạy offline)
file-mau/           3 file Excel mẫu để thử
```
