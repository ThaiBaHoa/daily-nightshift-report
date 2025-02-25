# Daily Nightshift Report

Ứng dụng web để tạo và quản lý báo cáo ca đêm, cho phép người dùng nhập dữ liệu, theo dõi và xuất báo cáo ra file Excel.

## Truy cập ứng dụng

Truy cập ứng dụng tại: [https://thaibahoa.github.io/daily-nightshift-report](https://thaibahoa.github.io/daily-nightshift-report)

## Hướng dẫn sử dụng

### 1. Chọn INSPECTOR
- Chọn tên INSPECTOR từ danh sách có sẵn trong dropdown menu
- Tên INSPECTOR sẽ được tự động điền vào báo cáo

### 2. Chọn ngày
- Click vào ô ngày để mở DatePicker
- Chọn ngày cần báo cáo
- Ngày được chọn sẽ tự động cập nhật trong báo cáo

### 3. Nhập dữ liệu báo cáo
- **Target**: Nhập mục tiêu kiểm tra
- **Note**: Ghi chú về tình trạng kiểm tra
- **Corrective action**: Nhập hành động khắc phục (nếu có)
- **Status**: Chọn một trong các trạng thái:
  - Checked: Đã kiểm tra, không có vấn đề
  - Not Check: Chưa kiểm tra
  - Finding: Đã kiểm tra, phát hiện vấn đề

### 4. Xem lại dữ liệu
- Dữ liệu đã nhập sẽ được hiển thị trong bảng phía dưới
- Kiểm tra lại thông tin trước khi xuất file

### 5. Xuất báo cáo
- Click nút "Export to Excel" để xuất báo cáo
- File Excel sẽ được tải về với tên "Daily Nightshift report_YYYYMMDD.xlsx"
- File Excel bao gồm các thông tin:
  - STT (Số thứ tự)
  - Date (Ngày báo cáo)
  - INSPECTOR (Người kiểm tra)
  - Target
  - Note
  - Corrective action
  - Status

### 6. Lưu trữ tạm thời
- Dữ liệu sẽ được tự động lưu trong trình duyệt
- Khi tải lại trang, dữ liệu sẽ được khôi phục
- Dữ liệu sẽ bị xóa khi xuất file Excel

## Yêu cầu hệ thống

- Trình duyệt web hiện đại (Chrome, Firefox, Edge, Safari)
- Kết nối internet để truy cập ứng dụng
- JavaScript được bật trong trình duyệt

## Xử lý sự cố

1. **Không thể chọn ngày**:
   - Kiểm tra xem trình duyệt có hỗ trợ DatePicker không
   - Thử tải lại trang

2. **Không xuất được file Excel**:
   - Đảm bảo đã chọn INSPECTOR
   - Kiểm tra quyền tải xuống file trong trình duyệt

3. **Dữ liệu không được lưu**:
   - Kiểm tra xem trình duyệt có bật JavaScript không
   - Xóa cache và tải lại trang

## Hỗ trợ

Nếu bạn gặp vấn đề hoặc cần hỗ trợ, vui lòng tạo issue tại [GitHub repository](https://github.com/ThaiBaHoa/daily-nightshift-report/issues).
