# Changelog

All notable changes to this project will be documented in this file.

## [1.3.0] - 2025-02-28

### Added
- Thêm tính năng đính kèm hình ảnh vào báo cáo
- Thêm logo Vietjet Air vào đầu ứng dụng
- Thêm khả năng xem trước và xóa hình ảnh đính kèm

### Changed
- Cải thiện tính năng xuất Excel với hỗ trợ hình ảnh
- Tăng kích thước ảnh trong file Excel thêm 15% để dễ nhìn
- Loại bỏ text trong ô chứa hình ảnh trong file Excel
- Điều chỉnh giao diện người dùng để hiển thị hình ảnh đính kèm

### Technical
- Thêm thư viện ExcelJS để hỗ trợ chèn hình ảnh vào file Excel
- Thêm thư viện file-saver để hỗ trợ tải file Excel
- Thêm tính năng resize hình ảnh trước khi lưu (tối đa 800x600px)

## [1.2.0] - 2025-02-25

### Added
- Thêm tính năng chọn STT từ dropdown menu (1-20)
- Cải thiện trải nghiệm người dùng khi nhập dữ liệu

## [1.1.0] - 2025-02-25

### Changed
- Cải thiện xử lý ngày tháng trong ứng dụng
- Loại bỏ cột DATE trùng lặp trong file Excel xuất ra
- Đảm bảo ngày được chọn từ DatePicker được hiển thị nhất quán trong toàn bộ ứng dụng

### Fixed
- Sửa lỗi hiển thị ngày trong phần review
- Sửa lỗi ngày không được cập nhật đúng trong template và dữ liệu

## [1.0.0] - Initial Release

### Added
- Tạo mới báo cáo ca đêm với template có sẵn
- Chọn INSPECTOR từ danh sách có sẵn
- Chọn ngày từ DatePicker
- Nhập dữ liệu cho các trường: Target, Note, Corrective action
- Tùy chọn trạng thái: Checked, Not Check, Finding
- Xuất báo cáo ra file Excel
- Lưu dữ liệu tạm thời trong local storage
- Giao diện thân thiện với người dùng
- Tích hợp với GitHub Pages để triển khai ứng dụng
