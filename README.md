# HTM to Excel Converter

Công cụ Python để trích xuất dữ liệu từ các file HTM Strategy Tester Optimization của MetaTrader và tổng hợp vào file Excel.

## Yêu cầu

- Python 3.7 trở lên
- Các thư viện: beautifulsoup4, openpyxl

## Cài đặt

1. Cài đặt các thư viện cần thiết:

```bash
pip install -r requirements.txt
```

Hoặc cài đặt từng thư viện:

```bash
pip install beautifulsoup4 openpyxl
```

## Cách sử dụng

### Cú pháp cơ bản:

```bash
python htm_to_excel.py --input <folder_chứa_file_htm> --output <folder_output> 
```

### Ví dụ:

```bash
python htm_to_excel.py --input D:\OptimizeLot --output D:\Output
```

hoặc sử dụng tham số ngắn gọn:

```bash
python htm_to_excel.py -i D:\OptimizeLot -o D:\Output
```

## Mô tả

Script sẽ:
1. Quét tất cả file `.htm` và `.html` trong folder input
2. Trích xuất dữ liệu từ bảng thứ 2 (bảng kết quả optimization) của mỗi file
3. **Tạo một file Excel riêng cho mỗi file HTM** (ví dụ: `audcad-144-sma.htm` → `audcad-144-sma.xlsx`)

## Dữ liệu được trích xuất

Mỗi file Excel sẽ chứa các cột sau (tất cả là kiểu số, trừ Detail):

- **Pass** (Number): Số thứ tự pass
- **Profit** (Number): Lợi nhuận
- **Total trades** (Number): Tổng số giao dịch
- **Profit factor** (Number): Hệ số lợi nhuận
- **Expected Payoff** (Number): Lợi nhuận kỳ vọng
- **Drawdown $** (Number): Drawdown theo đô la
- **Drawdown %** (Number): Drawdown theo phần trăm
- **Profit/Drawdown$** (Number): Tỷ lệ Profit/Drawdown (được tính tự động)
- **Detail** (Text): Chi tiết tham số (từ title của cột Pass)

## Lưu ý

- Folder output sẽ được tạo tự động nếu chưa tồn tại
- Mỗi file HTM sẽ tạo một file Excel riêng với cùng tên
- File Excel có định dạng header màu xanh và chữ in đậm
- Các cột số sẽ được chuyển đổi sang định dạng Number để dễ tính toán
- Độ rộng cột tự động điều chỉnh (cột Detail tối đa 80 ký tự)

## Xử lý lỗi

Nếu gặp lỗi, script sẽ:
- Hiển thị cảnh báo nếu file HTM không có đủ bảng dữ liệu
- Bỏ qua các dòng dữ liệu bị lỗi và tiếp tục xử lý
- Hiển thị thông báo lỗi chi tiết để dễ dàng debug

## Hỗ trợ

Nếu gặp vấn đề, vui lòng kiểm tra:
1. Python đã được cài đặt đúng version
2. Các thư viện đã được cài đặt
3. Đường dẫn folder input và output có hợp lệ
4. File HTM có đúng định dạng từ MetaTrader Strategy Tester
