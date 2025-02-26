# DAILY NIGHTSHIFT REPORT - HƯỚNG DẪN SỬ DỤNG CHI TIẾT

## Giới thiệu

Daily Nightshift Report là ứng dụng web được thiết kế để giúp người dùng nhập liệu và quản lý báo cáo ca đêm một cách hiệu quả. Ứng dụng cho phép người dùng nhập dữ liệu, lưu trữ tạm thời và xuất báo cáo ra file Excel.

## Truy cập ứng dụng

Bạn có thể truy cập ứng dụng trực tuyến tại: [https://thaibahoa.github.io/daily-nightshift-report/](https://thaibahoa.github.io/daily-nightshift-report/)

## Hướng dẫn sử dụng chi tiết

### 1. Giao diện chính

Giao diện chính của ứng dụng bao gồm các thành phần sau:

- **Phần đầu**: Tiêu đề "Nhập dữ liệu Excel"
- **Phần nhập liệu**: Các trường dữ liệu cần điền
- **Phần nút chức năng**: Các nút để thực hiện các hành động
- **Phần xem trước**: Bảng hiển thị dữ liệu đã nhập

### 2. Các trường dữ liệu

#### 2.1. STT (Số thứ tự)
- Chọn số thứ tự của mục cần nhập liệu từ danh sách dropdown
- Khi chọn STT, các thông tin liên quan sẽ được hiển thị trong các trường tương ứng

#### 2.2. Inspector (Người kiểm tra)
- Chọn tên người kiểm tra từ danh sách có sẵn
- Đây là trường bắt buộc phải chọn trước khi xuất file

#### 2.3. Date (Ngày)
- Chọn ngày thực hiện kiểm tra từ lịch
- Định dạng ngày mặc định là DD/MM/YYYY
- Ứng dụng hỗ trợ nhiều định dạng ngày khác nhau

#### 2.4. Target (Mục tiêu)
- Nhập mục tiêu cần đạt được
- Trường này có thể để trống

#### 2.5. Note (Ghi chú)
- Nhập ghi chú về tình trạng kiểm tra
- Trường này có thể để trống

#### 2.6. Corrective action (Hành động khắc phục)
- Nhập hành động khắc phục nếu có
- Trường này có thể để trống

#### 2.7. Các trường chỉ đọc
- Hiển thị thông tin từ template Excel
- Không thể chỉnh sửa các trường này

#### 2.8. Status (Trạng thái)
- Chọn trạng thái kiểm tra từ danh sách:
  - **Checked**: Đã kiểm tra, không có vấn đề
  - **Not Check**: Chưa kiểm tra
  - **Finding**: Đã kiểm tra, phát hiện vấn đề

### 3. Các nút chức năng

#### 3.1. Cập nhật dữ liệu
- Nhấn nút này để lưu thông tin đã nhập cho STT hiện tại
- Sau khi cập nhật, ứng dụng sẽ tự động chuyển sang STT tiếp theo
- Dữ liệu sẽ được hiển thị trong bảng xem trước

#### 3.2. Xuất File
- Nhấn nút này để xuất dữ liệu ra file Excel
- File Excel sẽ được tải về thiết bị của bạn
- Tên file: "Daily Nightshift report_NGAYTHANGNAM_INSPECTOR.xlsx"
- Lưu ý: Phải chọn Inspector trước khi xuất file

#### 3.3. Save Temp
- Lưu dữ liệu hiện tại vào bộ nhớ tạm của trình duyệt
- Dữ liệu sẽ được khôi phục khi tải lại trang

#### 3.4. Delete Temp
- Xóa dữ liệu tạm đã lưu trong trình duyệt
- Sử dụng khi muốn bắt đầu nhập liệu mới hoàn toàn

### 4. Quy trình làm việc

#### 4.1. Nhập liệu mới
1. Chọn STT cần nhập liệu
2. Chọn Inspector từ danh sách
3. Chọn Date (ngày kiểm tra)
4. Nhập Target, Note, Corrective action (nếu cần)
5. Chọn Status phù hợp
6. Nhấn "Cập nhật dữ liệu" để lưu thông tin
7. Lặp lại các bước trên cho các STT khác nếu cần
8. Nhấn "Xuất File" khi hoàn tất để tạo báo cáo Excel

#### 4.2. Chỉnh sửa dữ liệu đã nhập
1. Chọn STT cần chỉnh sửa
2. Thông tin hiện tại sẽ được hiển thị trong các trường
3. Chỉnh sửa thông tin cần thiết
4. Nhấn "Cập nhật dữ liệu" để lưu thay đổi

#### 4.3. Lưu dữ liệu tạm thời
1. Nhấn "Save Temp" để lưu dữ liệu hiện tại
2. Dữ liệu sẽ được khôi phục khi tải lại trang

#### 4.4. Xuất báo cáo Excel
1. Đảm bảo đã nhập đầy đủ thông tin cần thiết
2. Nhấn "Xuất File" để tạo file Excel
3. File sẽ được tải về thiết bị của bạn

### 5. Tính năng đặc biệt

#### 5.1. Tự động lưu
- Dữ liệu được tự động lưu khi thay đổi Inspector hoặc Date
- Khi thay đổi trường dữ liệu, ứng dụng sẽ tự động lưu vào bộ nhớ tạm

#### 5.2. Thông báo
- Ứng dụng hiển thị thông báo khi:
  - Tải template thành công hoặc thất bại
  - Cập nhật dữ liệu thành công
  - Xuất file thành công hoặc thất bại
  - Xảy ra lỗi khi xử lý dữ liệu

#### 5.3. Xem trước dữ liệu
- Bảng xem trước ở cuối trang hiển thị tất cả dữ liệu đã nhập
- Giúp kiểm tra tổng quan trước khi xuất file

## Xử lý lỗi thường gặp

### 1. Không thể xuất file
- **Nguyên nhân**: Chưa chọn Inspector
- **Giải pháp**: Chọn Inspector từ danh sách trước khi xuất file

### 2. Dữ liệu không được lưu
- **Nguyên nhân**: Trình duyệt không hỗ trợ localStorage hoặc đã đầy
- **Giải pháp**: 
  - Kiểm tra xem trình duyệt có hỗ trợ localStorage không
  - Xóa bớt dữ liệu trong localStorage nếu đã đầy
  - Sử dụng trình duyệt khác

### 3. Lỗi định dạng ngày
- **Nguyên nhân**: Định dạng ngày không hợp lệ
- **Giải pháp**: Sử dụng định dạng DD/MM/YYYY hoặc chọn ngày từ lịch

### 4. Không tải được template
- **Nguyên nhân**: Kết nối internet không ổn định hoặc file template không tồn tại
- **Giải pháp**: 
  - Kiểm tra kết nối internet
  - Tải lại trang
  - Liên hệ quản trị viên nếu vẫn gặp lỗi

## Mẹo sử dụng hiệu quả

1. **Lưu dữ liệu thường xuyên**: Nhấn "Save Temp" định kỳ để tránh mất dữ liệu
2. **Kiểm tra xem trước**: Luôn kiểm tra bảng xem trước trước khi xuất file
3. **Sử dụng bàn phím**: Sử dụng phím Tab để di chuyển giữa các trường nhanh chóng
4. **Xuất file sau khi hoàn tất**: Chỉ xuất file khi đã nhập đầy đủ thông tin cần thiết

## Yêu cầu hệ thống

- Trình duyệt web hiện đại (Chrome, Firefox, Safari, Edge)
- Kết nối internet để tải template và xuất file
- JavaScript được bật trong trình duyệt
- Đủ dung lượng lưu trữ cho localStorage (thường là 5-10MB)

## Liên hệ hỗ trợ

Nếu bạn gặp vấn đề khi sử dụng ứng dụng hoặc có đề xuất cải tiến, vui lòng liên hệ:

- **Email**: thaibahoa.dev@gmail.com
- **GitHub**: [https://github.com/ThaiBaHoa/daily-nightshift-report](https://github.com/ThaiBaHoa/daily-nightshift-report)
- **Tạo issue**: [https://github.com/ThaiBaHoa/daily-nightshift-report/issues](https://github.com/ThaiBaHoa/daily-nightshift-report/issues)
