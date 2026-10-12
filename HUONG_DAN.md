# Xét tốt nghiệp v2.0

## Sử dụng trên Windows

1. Cài Python 3.11 hoặc mới hơn, có Python Launcher (`py`).
2. Giải nén toàn bộ gói vào cùng một thư mục.
3. Chạy `CAI_THU_VIEN.bat` một lần.
4. Chạy `CHAY_UNG_DUNG.bat`, chọn Excel và thư mục xuất kết quả.
5. Nếu cần ứng dụng độc lập: chạy `BUILD_EXE.bat` trên Windows. Dùng `dist/XetTotNghiep_v2.exe` và giữ `dist/cau_hinh.json` cạnh nó.

Repository chứa mã nguồn và lệnh đóng gói; không lưu EXE, bộ cài hay dữ liệu học viên. Bản v2 đã được build và kiểm tra trên Windows: 11 kiểm thử đạt, EXE tạo đủ 4 file, form nhập thông tin hoạt động. Chạy BUILD_EXE.bat để build lại, sau đó mở XetTotNghiep_Setup_v2.iss bằng Inno Setup để tạo bộ cài.

## Điều kiện xét

Phải đạt cả bốn nhóm:

| Nhóm | Đạt | Không đạt |
|---|---|---|
| Học tập | DatTN = Đạt | Không đạt / Chưa đạt |
| Hồ sơ | Đủ hồ sơ | Thiếu hồ sơ / Không có hồ sơ / trạng thái chờ xác minh bảng điểm CĐ trong file mẫu |
| Tiếng Anh | Đạt / Đủ | Không đạt / Chưa đạt / Thiếu |
| Học phí | Không nợ HP | Nợ HP |

Không tự suy điều kiện từ tín chỉ, điểm trung bình, dấu x đối tượng, số tiền TongHS1 hoặc kết quả tổng của Trung tâm. Cần quy định nghiệp vụ riêng nếu muốn tự tính các điều kiện này.

Mã 0 = đạt, 1 = không đạt, 2 = trống/chưa nhận diện. Có một mã 1 thì không đủ; không có mã 1 nhưng có mã 2 thì CẦN KIỂM TRA; tất cả 0 mới đủ. File trả kết quả có thể chứa `Can kiem tra`, nên phải xác minh các dòng này trước khi nhập hệ thống chỉ cho phép hai trạng thái.

## Linh hoạt với cấu trúc file

- Tìm bảng theo tiêu đề trên mọi sheet và mọi dòng, không giới hạn 20 dòng đầu.
- Cho phép dời dòng tiêu đề, đổi vị trí sheet và cột. Với các cột trùng tên “Tình trạng”, tiêu đề nhóm phía trên phải nằm đúng cột hoặc đúng vùng gộp. Cột mang tên riêng có thể di chuyển độc lập.
- Bỏ dòng không có mã sinh viên và điều kiện; báo lỗi nếu có điều kiện nhưng thiếu mã, mã trùng, thiếu nhóm hoặc nhiều bảng không xác định được.
- Không suy đoán nhóm dựa trên thứ tự. Khi tên trường đổi, thêm tên mới vào `aliases` trong `cau_hinh.json`; khi trạng thái đổi, cập nhật `pass` / `fail` theo từng nhóm. Giữ các khóa nhóm và mã trường có sẵn.
- Nếu có nhiều sheet xét, chạy dòng lệnh với `--sheet "Tên sheet"`. Có nhiều bảng trên cùng sheet thì tách thành sheet riêng.
- Đọc giá trị đã tính được lưu trong Excel. Với công thức chưa có kết quả lưu, mở Excel, tính lại và lưu trước khi chạy. Không tự tính công thức Excel.
- Hỗ trợ `.xlsx`, `.xlsm`, chưa hỗ trợ `.xls` hoặc CSV. VBA được giữ khi xử lý `.xlsm`; các thành phần Excel đặc biệt không được openpyxl hỗ trợ có thể không giữ được hoàn toàn.

## Các file đầu ra

Trong file email, `TRẠNG THÁI HẠ BẬC TN` ghi **Học viên đã được xét hạ bậc tốt nghiệp theo quy định** khi `HaBacTN` có hạ bậc. Giá trị trống/0/Không sẽ để trống cả trạng thái và lý do. `LÝ DO HẠ BẬC` lấy nguyên nội dung thực tế từ `LyDoHBTN`, ví dụ **% học phần thi lại vượt mức cho phép - Học lại 9 TC**; không tự điền ví dụ khi thiếu lý do. Trạng thái chưa rõ sẽ ghi `Cần kiểm tra hạ bậc`. Sheet `Note` có hướng dẫn cho hai cột này.

Khi bấm **BẮT ĐẦU XÉT TỐT NGHIỆP**, ứng dụng mở form nhập thông tin trả email:

- Bắt buộc: Đợt xét TN, Link tra học phí (điền sẵn `https://hocvien.hutech.edu.vn`), Đợt xét TN tiếp theo, Link thông báo đợt TN tiếp theo, Thời gian nhận bằng tạm, Thời gian nhận bằng chính thức.
- Không bắt buộc: Hạn đóng học phí, Hạn bổ sung hồ sơ, Hạn bổ sung CCNN, Hạn phản hồi tổng.
- Thời gian có thể nhập ngày hoặc nội dung thông báo. Link phải bắt đầu bằng `http://` hoặc `https://`.
- Bấm **Tiếp tục xét**, chọn thư mục lưu; các thông tin nhập sẽ được ghi vào tất cả dòng học viên trong `_TOOL_EMAIL.xlsx`. Hạn để trống sẽ để trống cả trong file xuất. Bấm **Hủy** không chạy xét và không tạo file.
- Ứng dụng giữ nội dung nhập cho lần xét tiếp theo trong cùng phiên; đóng ứng dụng sẽ xóa nội dung nhớ. Lệnh không giao diện vẫn có thể chạy theo cách cũ.

1. `_KET_QUA`: giữ các sheet nguồn và thêm `KQ_TOOL_V2`, không chèn cột vào bảng nguồn. Sheet này dành riêng cho công cụ và được thay khi chạy lại.
2. `_BAO_CAO.xlsx`: tổng quan, chi tiết và cảnh báo chưa nhận diện.
3. `_TRA_KET_QUA.xlsx`: 14 cột của bản cũ và thêm cột `Xác minh văn bằng`. Nợ học phí và chứng chỉ tiếng Anh lấy từ đúng nhóm; không theo thứ tự cột. CCNN được tách khỏi `Hồ sơ thiếu`; XMVB/đang chờ giữ riêng kèm tên cột nguồn để kiểm tra. Trường nguồn tùy chọn thiếu thì để trống.
4. `_TOOL_EMAIL.xlsx`: sheet `DATA` đúng 30 cột mẫu ToolEmail, có `XÁC MINH VĂN BẰNG`. Sau khi bấm xét, cuộn xuống `FILE ĐÃ XUẤT` và bấm `Mở file` tại dòng `File trả email`. Đường dẫn cũng hiển thị trong thông báo hoàn tất. Kiểm tra và bổ sung email, đợt xét, các thời hạn và đường dẫn còn thiếu trước khi gửi. `STATUS`, `SENT_TIME`, `NOTE` để trống cho lần gửi mới; không dùng lại lịch sử gửi từ file nguồn. Các dòng `Can kiem tra` phải được xác minh trước khi gửi. Công cụ chỉ tạo Excel, không gửi email.

File đầu ra cùng tên trong thư mục đã chọn sẽ được ghi đè khi chạy lại. File đầu vào không bị ghi đè.

## Kiểm chứng trên file Trung tâm ngày 07.10.2026

- Sheet `TN_T10.26`, dòng tiêu đề 2.
- 250 học viên; bỏ 8 dòng trống cuối file.
- 144 đủ điều kiện; 106 không đủ; 0 cần kiểm tra.
- Kết quả khớp 250/250 dòng với cột kết quả Trung tâm cung cấp.
- Kiểm thử tự động đạt: đảo cột, dời tiêu đề 30 dòng, sheet ghi chú ở đầu, tiêu đề gộp, thiếu nhóm, phủ định, trạng thái lạ và ô trống.

Lệnh chạy không giao diện:

```bat
py xet_tot_nghiep.py "du_lieu.xlsx" --output-dir "ket_qua" --sheet "TN_T10.26"
```

Lệnh kiểm thử: `py -m unittest test_engine.py -v`.
