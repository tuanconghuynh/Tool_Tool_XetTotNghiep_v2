# Tool_Tool_XetTotNghiep_v2

Phần mềm xét tốt nghiệp từ Excel cho **Viện Đại học Mở HUTECH**.

- Nhận diện bảng theo tên cột và nhóm tiêu đề, hỗ trợ đổi vị trí cột, dòng và sheet.
- Xét bốn nhóm: học tập, hồ sơ, tiếng Anh và học phí; báo rõ dữ liệu trống hoặc trạng thái chưa nhận diện.
- Xuất kết quả xét, báo cáo, file trả kết quả và file DATA 30 cột cho tool email.
- Nhập thông tin đợt tốt nghiệp trước khi xét; tách CCNN khỏi hồ sơ thiếu và ghi riêng thông tin xác minh văn bằng.

## Chạy trên Windows

Cài Python 3.11 trở lên, chạy `CAI_THU_VIEN.bat`, sau đó chạy `CHAY_UNG_DUNG.bat`.

Đọc [HUONG_DAN.md](HUONG_DAN.md) để xem quy tắc xét, thông tin cần nhập và cách xử lý cảnh báo.

## Kiểm thử và đóng gói

```powershell
py -m unittest test_engine.py -v
```

Chạy `BUILD_EXE.bat` để tạo `dist/XetTotNghiep_v2.exe` và `dist/cau_hinh.json`.

Mở `XetTotNghiep_Setup_v2.iss` bằng Inno Setup 6.5 trở lên và bấm Compile để tạo `installer/XetTotNghiep_Setup_v2.0.exe`.

EXE, bộ cài, bản sao lưu và dữ liệu Excel của học viên không được đưa vào Git. Khi phân phối EXE, giữ `cau_hinh.json` cạnh EXE.