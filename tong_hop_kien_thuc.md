# Tổng hợp kiến thức
## 1. Các lệnh terminal cơ bản và nâng cao
|Công dụng|Windows (Powershell)|
|---|---|
|Xem vị trí hiện tại | `pwd` |
| Liệt kê nội dung | `ls [-Force]` |
| Liệt kê file và thư mục con, kể cả mục ẩn | `Get-ChildItem -Recurse -Force` |
| Di chuyển giữa các thư mục | `cd folder` |
| Tạo thư mục | `mkdir folder` |
| Di chuyển file | `Move-Item file folder/` |
| Sao chép file | `Copy-Item folder/file other/` |
| Xóa file | `Remove-Item file` |
| Đọc file văn bản | `Get-Content folder/file` |
| Lọc dòng chứa từ khóa qua pipeline | `... \| Select-String "WARNING"` |
| Đếm số dòng nhận qua pipeline | `... \| Measure-Object -Line` |
| Ghi đè hoặc tạo file | `... > file` |
| Ghi nối vào cuối file | `... >> file` |
| Chạy script tại thư mục hiện tại | `.\file.ps1` |
| Tạo file| `ni file`|
---

| Ký hiệu | Công dụng |
|---|---|
| `|` | Chuyển kết quả của lệnh trước sang lệnh sau | 
| `>` | Ghi kết quả ra tệp, thay thế nội dung cũ | 
| `>>` | Ghi thêm kết quả vào cuối tệp | 
| `;` | Chạy các lệnh lần lượt, dù lệnh trước có lỗi |
| `.`| Thư mục hiện tại|
| `..`| Thư mục cha|
---

## 2. Các định dạng trong markdown
| Mục đích | Cú pháp Markdown |
|---|---|
| Tiêu đề cấp 1 | `# Tiêu đề` |
| Tiêu đề cấp 2 | `## Tiêu đề` |
| Tiêu đề cấp 3 | `### Tiêu đề` |
| In đậm | `**Nội dung**` |
| In nghiêng | `*Nội dung*` |
| Đậm và nghiêng | `***Nội dung***` |
| Gạch ngang chữ | `~~Nội dung~~` |
| Chữ dạng mã | `` `code` `` |
| Đường kẻ ngang | `---` |
| Trích dẫn | `> Nội dung trích dẫn` |
| Tạo liên kết| `[Tên liên kết](URL)`|
| Chèn hình ảnh|`![Mô tả hình ảnh](đường_dẫn_ảnh)`
| Tạo task list công việc đã hoàn thành|- [x] |
|Tạo task list công việc chưa hoàn thành| - [ ]|
---


## 3.Microsoft Word đầy đủ kiến thức nhập môn

### 3.1. Các thao tác soạn thảo cơ bản

| Thao tác | Cách thực hiện |
|---|---|
| Tạo tài liệu | File → New |
| Mở tài liệu | File → Open |
| Lưu tài liệu | Ctrl + S |
| Lưu thành tệp khác | File → Save As |
| In tài liệu | Ctrl + P |
| Sao chép, cắt, dán | Ctrl + C, Ctrl + X, Ctrl + V |
| Hoàn tác | Ctrl + Z |
| Tìm từ | Ctrl + F |
| Tìm và thay thế | Ctrl + H |
| Chèn ngắt trang | Ctrl + Enter |

### 3.2. Định dạng chữ và đoạn văn

Các thao tác cần thành thạo:
- Chọn font chữ và cỡ chữ
- In đậm, in nghiêng, gạch chân
- Căn trái, căn giữa, căn phải, căn đều hai bên
- Điều chỉnh giãn dòng và khoảng cách giữa các đoạn
- Thụt đầu dòng
- Tạo danh sách đánh số hoặc dấu đầu dòng

**Phím tắt căn lề thường dùng:**

| Phím tắt | Công dụng |
|---|---|
| Ctrl + L | Căn trái |
| Ctrl + E | Căn giữa |
| Ctrl + R | Căn phải |
| Ctrl + J | Căn đều hai bên |

###  3.3. Heading và mục lục tự động

| Cấp | Sử dụng |
|---|---|
| Heading 1 | Chương hoặc phần lớn |
| Heading 2 | Mục trong chương |
| Heading 3 | Mục nhỏ hơn |

### 3.4. Đánh số trang, đầu trang và chân trang

- Đánh số trang: Insert → Page Number
- Đầu trang: Insert → Header
- Chân trang: Insert → Footer
- Ngày tháng: Insert → Date & Time

### 3.5. Tạo danh sách và bảng

**Danh sách:**

- Bullets: Danh sách có dấu đầu dòng
- Numbering: Danh sách có số thứ tự
- Multilevel List: Danh sách nhiều cấp

**Bảng:**

1. Chọn Insert → Table
2. Chọn số hàng và cột
3. Nhập dữ liệu
4. Dùng Table Design để định dạng
5. Dùng Layout để thêm, xóa hàng hoặc cột và căn chỉnh nội dung

### 3.6. Chèn hình ảnh, liên kết và chú thích

- Chèn ảnh: Insert → Pictures
- Chèn liên kết: Ctrl + K
- Chèn ký hiệu: Insert → Symbol
- Chèn chú thích hình/bảng: References → Insert Caption
- Kiểm tra chính tả: Review → Spelling & Grammar, nếu phiên bản Word hỗ trợ ngôn ngữ đang dùng

### 3.7. Thiết lập trang và xuất tệp

| Tính năng | Vị trí |
|---|---|
| Khổ giấy | Layout → Size |
| Lề trang | Layout → Margins |
| Hướng giấy | Layout → Orientation |
| Chia cột | Layout → Columns |
| Ngắt trang | Insert → Page Break hoặc Ctrl + Enter |
| Xuất PDF | File → Save As hoặc Export → PDF |
---


## 4. Microsoft Excel đầy đủ kiến thức nhập môn

### 4.1. Các khái niệm cơ bản

- **Workbook:** Tệp Excel, thường có đuôi `.xlsx`
- **Worksheet:** Trang tính bên trong workbook
- **Cell:** Ô, ví dụ `A1`
- **Range:** Vùng ô, ví dụ `A1:C5`
- **Formula:** Công thức bắt đầu bằng dấu `=`
- **Function:** Hàm có sẵn như `SUM`, `IF`, `VLOOKUP`


### 4.2. Các phép tính và hàm thông dụng

Giả sử điểm nằm ở các ô B2:B6.

| Hàm | Công dụng | Ví dụ |
|---|---|---|
| SUM | Tính tổng | `=SUM(B2:B6)` |
| AVERAGE | Tính trung bình | `=AVERAGE(B2:B6)` |
| MAX | Giá trị lớn nhất | `=MAX(B2:B6)` |
| MIN | Giá trị nhỏ nhất | `=MIN(B2:B6)` |
| COUNT | Đếm ô chứa số | `=COUNT(B2:B6)` |
| COUNTA | Đếm ô không trống | `=COUNTA(A2:A6)` |
| IF | Kiểm tra điều kiện | `=IF(B2>=5,"Đạt","Trượt")` |
| COUNTIF | Đếm theo điều kiện | `=COUNTIF(B2:B6,">=5")` |
| SUMIF | Tính tổng theo điều kiện | `=SUMIF(A2:A6,"IE02",B2:B6)` |
| IFERROR | Xử lý kết quả lỗi | `=IFERROR(A2/B2,0)` |

### 4.3. Tham chiếu ô tương đối và tuyệt đối

| Loại | Ví dụ | Ý nghĩa |
|---|---|---|
| Tương đối | `A1` | Có thể thay đổi khi sao chép công thức |
| Tuyệt đối | `$A$1` | Cố định cả cột và hàng |
| Hỗn hợp | `$A1` | Cố định cột A |
| Hỗn hợp | `A$1` | Cố định hàng 1 |


### 4.4. VLOOKUP từ cơ bản đến nâng cao

**Cú pháp:**

`=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

| Thành phần | Ý nghĩa |
|---|---|
| lookup_value | Giá trị cần tìm |
| table_array | Vùng dữ liệu tra cứu |
| col_index_num | Số thứ tự cột lấy kết quả |
| range_lookup | FALSE tìm chính xác; TRUE tìm gần đúng |


### 4.5. Sắp xếp, lọc và định dạng dữ liệu

| Tính năng | Công dụng |
|---|---|
| Sort A to Z | Sắp xếp tăng dần |
| Sort Z to A | Sắp xếp giảm dần |
| Filter | Lọc dữ liệu theo điều kiện |
| Format as Table | Định dạng dữ liệu thành bảng |
| Conditional Formatting | Tô màu theo điều kiện |
| Freeze Panes | Cố định hàng hoặc cột khi cuộn |
| Remove Duplicates | Loại bỏ dữ liệu trùng lặp |
| Data Validation | Giới hạn dữ liệu được nhập |


### 4.6. Các lỗi Excel phổ biến

| Lỗi | Nguyên nhân thường gặp |
|---|---|
| `#N/A` | Không tìm thấy giá trị cần tra |
| `#VALUE!` | Dữ liệu không phù hợp với phép tính |
| `#REF!` | Tham chiếu ô không hợp lệ |
| `#DIV/0!` | Chia cho 0 hoặc ô trống |
| `#####` | Cột quá hẹp để hiển thị giá trị |
| Công thức hiện nguyên văn | Ô có thể đang ở định dạng Text hoặc công thức chưa được nhập đúng |
---


## 5. Microsoft PowerPoint 

### 5.1. Các thao tác cơ bản

| Thao tác | Cách thực hiện |
|---|---|
| Tạo bài mới | File → New |
| Thêm slide | Home → New Slide hoặc Ctrl + M |
| Nhân đôi slide/đối tượng | Chọn đối tượng rồi nhấn Ctrl + D |
| Chèn hình ảnh | Insert → Pictures |
| Chèn bảng | Insert → Table |
| Chèn biểu đồ | Insert → Chart |
| Chèn hộp văn bản | Insert → Text Box |
| Chọn giao diện | Design → Themes |
| Chọn hiệu ứng chuyển slide | Transitions |
| Tạo hiệu ứng cho đối tượng | Animations |
| Trình chiếu từ đầu | F5 |
| Trình chiếu từ slide hiện tại | Shift + F5 |
| Kết thúc trình chiếu | Esc |

### 5.2. Phân biệt các tính năng

- **Design:** Thiết kế giao diện, màu sắc và chủ đề.
- **Transitions:** Hiệu ứng chuyển từ slide này sang slide khác.
- **Animations:** Hiệu ứng xuất hiện hoặc di chuyển của chữ, hình ảnh và đối tượng.
- **Slide Show:** Trình chiếu bài thuyết trình.
- **Slide Master:** Thiết lập mẫu chung cho nhiều slide, phù hợp khi muốn thống nhất font, logo và bố cục.

### 5.3. Quy chuẩn thiết kế slide

#### 1. Trang bìa

Có tên đề tài, tên môn học, giảng viên và thành viên nhóm nếu được yêu cầu.

#### 2. Slide nội dung

Mỗi slide tập trung vào một ý chính. Dùng gạch đầu dòng thay vì chép nguyên đoạn văn dài.

#### 3. Slide số liệu

Dùng bảng hoặc biểu đồ phù hợp, có tên biểu đồ và ghi nguồn dữ liệu khi cần.

**Một số quy tắc thực tế:**

- Dùng khoảng 2–3 màu chủ đạo.
- Dùng font dễ đọc, thống nhất toàn bài.
- Tiêu đề khoảng 28–36 pt, nội dung khoảng 18–24 pt là mức tham khảo.
- Không kéo giãn hình ảnh làm biến dạng.
- Căn chỉnh các đối tượng bằng Align và Distribute.
- Không lạm dụng hiệu ứng, âm thanh hoặc hình động.
- Kiểm tra lỗi chính tả và nguồn tham khảo trước khi nộp.


