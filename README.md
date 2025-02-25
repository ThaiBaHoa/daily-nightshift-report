# Ứng dụng Nhập Dữ liệu Excel

Ứng dụng web đáp ứng cho phép người dùng:
- Tải lên file Excel
- Xem và nhập dữ liệu mới
- Tự động cập nhật và tải xuống file Excel với dữ liệu mới

## Cài đặt

1. Cài đặt Node.js (nếu chưa có)
2. Chạy lệnh sau trong thư mục dự án:
```bash
npm install
```

## Chạy ứng dụng

```bash
npm start
```

Ứng dụng sẽ chạy ở địa chỉ [http://localhost:3000](http://localhost:3000)

## Tính năng

- Giao diện thân thiện với thiết bị di động
- Hỗ trợ tải lên file Excel (.xlsx, .xls)
- Nhập dữ liệu với form động dựa trên cấu trúc file Excel
- Xem trước dữ liệu dạng bảng
- Tự động tải xuống file Excel đã cập nhật

## Công nghệ sử dụng

- React
- TypeScript
- Material-UI
- XLSX library
