# HỆ THỐNG TỰ ĐỘNG TẢI & TỔNG HỢP HÓA ĐƠN ĐIỆN TỬ PETROLIMEX
> **Công ty Cổ phần Thép POVINA**

Phần mềm tự động hóa toàn diện quy trình xử lý hóa đơn xăng dầu Petrolimex:
1. **Kết nối trực tiếp Microsoft Outlook**: Tự động nhận diện và quét các email thông báo phát hành hóa đơn từ Petrolimex (`Petrolimex@Petrolimex.com.vn`) trong các hòm thư (`vantai povina`, `Phòng kế toán`...).
2. **Tự động truy cập cổng tra cứu**: Truy cập `https://hoadon.petrolimex.com.vn`, tự động vượt mã xác thực Captcha bằng trí tuệ nhân tạo (AI OCR) không cần nhập tay.
3. **Tải file gốc đầy đủ**: Tải file XML hóa đơn gốc chuẩn Thông tư 78 và kết xuất file PDF sắc nét lưu trữ trên máy tính.
4. **Trích xuất thông tin chi tiết**: Trích xuất chính xác 100% biển số xe lấy xăng/dầu, ngày hóa đơn, mẫu số, ký hiệu, số hóa đơn, tên mặt hàng (Dầu DO, Xăng RON 95...), số lượng lít, đơn giá, thành tiền, thuế VAT, tổng tiền, địa chỉ cửa hàng xuất xăng dầu...
5. **Tổng hợp vào Excel chuyên nghiệp**: Ghi tự động vào file `data/Bao_Cao_Hoa_Don_Petrolimex.xlsx` có định dạng đẹp mắt, tự động chống trùng lặp dữ liệu, tích hợp nút liên kết mở trực tiếp file PDF và XML.
6. **Giao diện Web Dashboard**: Quản lý trực quan, hiển thị thống kê tổng tiền, tổng lít xăng dầu, tìm kiếm hóa đơn theo biển số xe hoặc số HĐ, xem trực tiếp PDF ngay trên trình duyệt.

---

## CẤU TRÚC THƯ MỤC DỰ ÁN

```
Petrolimex_Invoice_Tool/
│
├── invoices_pdf/              # Chứa toàn bộ file PDF hóa đơn chính thức tải về
├── invoices_xml/              # Chứa toàn bộ file XML gốc của hóa đơn (chuẩn TT78)
├── data/
│   └── Bao_Cao_Hoa_Don_Petrolimex.xlsx  # File Excel tổng hợp dữ liệu
│
├── core/
│   ├── outlook_reader.py      # Module quét và trích xuất thông tin từ Outlook
│   ├── portal_downloader.py   # Module tự động vượt captcha và tải dữ liệu từ cổng Petrolimex
│   ├── invoice_parser.py      # Module phân tích cú pháp XML/PDF chuẩn TT78
│   ├── excel_manager.py       # Module ghi và định dạng Excel chuyên nghiệp
│   └── manager.py             # Điều phối toàn bộ luồng hoạt động
│
├── web/
│   ├── templates/index.html   # Giao diện Web Dashboard hiện đại
│   └── static/css, js         # Mã nguồn giao diện và xử lý tương tác
│
├── app.py                     # Khởi động máy chủ Web Dashboard
├── cli.py                     # Chạy đồng bộ dạng dòng lệnh (hỗ trợ hẹn giờ Windows Task Scheduler)
├── Chay_Giao_Dien_Web.bat     # [1-CLICK] Nhấp đúp chuột để mở giao diện Web Dashboard
├── Chay_Tu_Dong_Cap_Nhat.bat  # [1-CLICK] Nhấp đúp chuột để tự quét và mở ngay file Excel
└── requirements.txt           # Danh sách thư viện Python
```

---

## CÁCH SỬ DỤNG NHANH

### Cách 1: Sử dụng Giao diện Web Dashboard (Khuyên dùng)
1. Nhấp đúp chuột vào file **`Chay_Giao_Dien_Web.bat`**.
2. Trình duyệt web sẽ tự động mở trang quản lý tại địa chỉ: `http://localhost:5005`.
3. Chọn hòm thư Outlook cần quét (ví dụ: `vantai povina`), chọn số lượng email (ví dụ: 50 email gần nhất).
4. Bấm nút **"Bắt Đầu Quét & Tải Hóa Đơn"**. Hệ thống sẽ tự động giải mã captcha, tải PDF/XML và cập nhật bảng dữ liệu theo thời gian thực.
5. Xem file PDF trực tiếp bằng nút **"Xem PDF"** trên từng dòng hoặc bấm **"Mở File Excel"** để xem báo cáo.

### Cách 2: Chạy Tự Động Nhanh & Mở Ngay Excel
1. Nhấp đúp chuột vào file **`Chay_Tu_Dong_Cap_Nhat.bat`**.
2. Phần mềm sẽ tự động quét Outlook, tải các hóa đơn mới chưa có trong hệ thống và tự động mở file Excel cho bạn.

---

## CÁC TRƯỜNG DỮ LIỆU ĐƯỢC XUẤT VÀO EXCEL

| Cột | Tên Cột | Mô Tả |
| :--- | :--- | :--- |
| **A** | STT | Số thứ tự tăng dần |
| **B** | Ngày nhận mail | Thời gian nhận email thông báo trong Outlook |
| **C** | Ngày hóa đơn | Ngày xuất hóa đơn chính thức |
| **D** | Mẫu số | Mẫu số hóa đơn (ví dụ: `1/002`) |
| **E** | Ký hiệu | Ký hiệu hóa đơn (ví dụ: `K26TAN`) |
| **F** | Số hóa đơn | Số thứ tự hóa đơn điện tử |
| **G** | **Biển số xe** | Biển số xe vào đổ nhiên liệu (ví dụ: `51L-638.55`) |
| **H** | Tên hàng hóa | Loại nhiên liệu (Dầu Điêzen 0,001S, Xăng RON 95...) |
| **I** | ĐVT | Lít |
| **J** | **Số lượng** | Số lít xăng/dầu đã đổ (định dạng `#,##0.000`) |
| **K** | Đơn giá | Đơn giá xăng dầu (định dạng `#,##0`) |
| **L** | Thành tiền | Tiền trước thuế |
| **M** | Tiền thuế VAT | Thuế GTGT |
| **N** | **Tổng thanh toán** | Tổng tiền thanh toán (VNĐ) |
| **O** | Bên bán / Cửa hàng | Đơn vị phát hành (Petrolimex Sài Gòn, Đồng Nai...) |
| **P** | Địa chỉ cửa hàng | Cửa hàng xăng dầu nơi xe vào đổ |
| **Q** | Bên mua | CÔNG TY CỔ PHẦN THÉP POVINA |
| **R** | Mã tra cứu (Fkey) | Mã tra cứu hóa đơn trực tuyến |
| **S** | **File PDF** | **Liên kết bấm trực tiếp để mở file PDF** |
| **T** | **File XML** | **Liên kết bấm trực tiếp để mở file XML gốc** |
