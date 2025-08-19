# 🖼️ Image Crawler - Cào Ảnh Tự Động

Ứng dụng desktop để cào ảnh từ các link sản phẩm, xử lý ảnh và chuyển đổi sang định dạng WebP với nền trắng.

## ✨ Tính Năng Chính

### 🎯 **Cào Ảnh Thông Minh**
- **Link ảnh trực tiếp**: Download ngay lập tức
- **Link trang web**: Tự động crawl và tìm ảnh
- **Đa luồng**: Xử lý song song để tăng tốc độ
- **Nhận đầy đủ**: Không bỏ sót link nào từ Excel

### 🖼️ **Xử Lý Ảnh Sản Phẩm**
- **Nền trắng**: Tự động thêm nền trắng cho ảnh sản phẩm
- **Giữ nguyên kích thước**: Không làm vỡ ảnh gốc
- **Format WebP**: Chuyển đổi sang WebP với chất lượng 85%
- **Tối ưu hóa**: Giảm dung lượng file

### 📝 **Đặt Tên File Thông Minh**
- **Logic JavaScript**: Theo chuẩn đặt tên từ code JavaScript
- **Xử lý mã sản phẩm**: Loại bỏ ghi chú, coating, add-on kit
- **Chuẩn hóa tên**: Tự động làm sạch ký tự không hợp lệ
- **Hỗ trợ add-on**: Tự động thêm suffix "-adk" cho add-on kit

### 📊 **Import Excel**
- **Cấu trúc đơn giản**: Cột A = Mã sản phẩm, Cột B = Link ảnh
- **Xử lý linh hoạt**: Tự động tạo mã nếu trống
- **Debug chi tiết**: Nút debug để xem thông tin chi tiết
- **Log đầy đủ**: Hiển thị quá trình xử lý từng dòng

## 🚀 Cài Đặt

### Yêu Cầu Hệ Thống
- Python 3.7+
- Chrome Browser (cho Selenium)

### Cài Đặt Dependencies
```bash
pip install -r requirements.txt
```

### Chạy Ứng Dụng
```bash
python main.py
```

## 📋 Cách Sử Dụng

### 1. **Chuẩn Bị File Excel**
```
| Cột A (Mã Sản Phẩm) | Cột B (Link Ảnh)                    |
|----------------------|--------------------------------------|
| FR-1H-220V           | https://example.com/image1.jpg      |
| TC-NT-20R            | https://example.com/image2.png      |
| ETC-48 Add-on Kit    | https://example.com/image3.jpg      |
```

### 2. **Import File Excel**
1. Chọn "File Excel (Khuyến nghị)"
2. Click "Chọn File Excel"
3. Chọn file Excel của bạn
4. Xem log để kiểm tra dữ liệu

### 3. **Cấu Hình Xử Lý**
- **Số luồng**: 1-10 (khuyến nghị 5)
- **Thư mục lưu**: Chọn nơi lưu ảnh
- **Xử lý ảnh**: "Ảnh sản phẩm (có nền trắng)"
- **Chế độ crawl**: "Link ảnh trực tiếp"

### 4. **Bắt Đầu Crawl**
1. Click "🚀 Bắt Đầu Crawl"
2. Theo dõi tiến trình trong log
3. Xem thống kê kết quả

## 🧪 Test Tính Năng

### Test Logic Đặt Tên
```bash
python test_naming_demo.py
```

### Test App Đầy Đủ
```bash
python test_full_excel.py
```

## 📝 Logic Đặt Tên File

### Ví Dụ Chuyển Đổi
```
Mã gốc: "FR-1H-220V (with special coating)"
↓
Mã sạch: "fr-1h-220v"
↓
Tên ảnh: "FR-1H-220V"
↓
Filename: "FR-1H-220V.webp"
```

### Xử Lý Đặc Biệt
- **Add-on Kit**: Tự động thêm "-adk"
- **Special Coating**: Loại bỏ ghi chú
- **Ký tự đặc biệt**: Chuyển thành gạch ngang
- **Chuẩn hóa**: Chỉ giữ a-z, A-Z, 0-9, -, _

## 🔧 Tính Năng Debug

### Nút Debug Excel
- Hiển thị thông tin chi tiết file Excel
- Xem dữ liệu từng dòng
- Kiểm tra mapping mã sản phẩm - link

### Log Chi Tiết
- Thông tin file Excel
- Quá trình xử lý từng dòng
- Kết quả mapping
- Thống kê thành công/thất bại

## 📁 Cấu Trúc File

```
crawlerP/
├── main.py                    # App chính
├── image_naming_processor.py  # Module xử lý đặt tên
├── test_naming_demo.py        # Demo logic đặt tên
├── test_full_excel.py         # Test app đầy đủ
├── requirements.txt           # Dependencies
├── README.md                  # Hướng dẫn
└── TROUBLESHOOTING.md         # Sửa lỗi
```

## 🎯 Kết Quả

### File Ảnh Được Tạo
- **Format**: WebP
- **Chất lượng**: 85%
- **Nền**: Trắng (cho ảnh sản phẩm)
- **Tên**: Theo logic JavaScript

### Thống Kê
- Tổng link đã xử lý
- Số ảnh thành công
- Số ảnh thất bại
- Thời gian xử lý

## 🆘 Hỗ Trợ

### Gặp Vấn Đề?
1. Xem file `TROUBLESHOOTING.md`
2. Sử dụng nút "🔍 Debug Excel"
3. Kiểm tra log chi tiết
4. Chia sẻ thông tin debug

### Lỗi Thường Gặp
- **App chỉ nhận 7/12 link**: Xem troubleshooting
- **Lỗi download ảnh**: Kiểm tra link có hợp lệ không
- **Lỗi Selenium**: Cài đặt Chrome browser

## 📞 Liên Hệ

Nếu cần hỗ trợ thêm, hãy chia sẻ:
- Thông tin debug từ app
- File Excel mẫu (nếu có thể)
- Log lỗi chi tiết

**Happy Crawling! 🚀**
