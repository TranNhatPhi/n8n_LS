# Hanger Automation — upload trực tiếp vào n8n

Dự án chạy trên Docker Desktop cho MacBook. Không cần kết nối SMB và không cần
chép file vào thư mục input.

Người dùng mở form n8n rồi tải lên:

1. Một file đơn hàng `.xls`, `.xlsx`, `.docx` hoặc `.pdf`.
2. Một hoặc nhiều file SOF tương ứng `.xls`, `.xlsx`, `.docx` hoặc `.pdf` (tối đa 20 file).

n8n đóng gói các SOF rồi chuyển file đơn hàng và gói SOF qua mạng Docker nội bộ
cho `hanger-worker`. Worker kiểm tra và đọc gói SOF an toàn, xử lý trong thư mục
tạm, kiểm tra source sheet/cells cho Excel hoặc section/lines cho Word/PDF,
tạo bản sao kết quả rồi tự xóa file upload tạm.
File kết quả nằm trong `data/output` trên MacBook.

## Kiến trúc và an toàn

- `hanger-n8n`: image n8n chính thức, UI tại `http://localhost:5678`.
- `hanger-worker`: Python, OpenPyXL, LibreOffice, python-docx và PyMuPDF.
- Worker không publish cổng ra MacBook; chỉ n8n gọi được.
- `Execute Command` và `Read/Write Files from Disk` bị tắt trong n8n.
- Binary upload được lưu bằng filesystem mode trong Docker volume thay vì RAM.
- File đơn hàng gốc không bao giờ bị sửa.
- Đơn hàng `.docx`/`.pdf` được chuyển thành một bảng Excel chuẩn trước khi kiểm
  tra: chỉ ghi giá trị file có ghi, cột nào file không có thì để trống và dòng đó
  vào `REVIEW`, không suy đoán.
- Chỉ dòng `MATCHED` được ghi hanger; `MISMATCH` và `REVIEW` đi vào sheet
  `HANGER REVIEW`.
- Rule đã duyệt luôn chạy trước. DeepSeek chỉ xử lý các dòng chưa có rule và kết
  quả LLM vẫn phải vượt qua kiểm tra category, Hang/Flat, độ tin cậy, sheet/range
  nguồn và nội dung trích dẫn trong SOF.
- SOF `.docx` và PDF có lớp chữ được đọc trực tiếp, không OCR và không gửi ảnh
  lên DeepSeek. PDF scan/trang ảnh không có chữ bị từ chối rõ ràng. SOF Word/PDF
  chỉ được dùng để điền khi DeepSeek bật và trích dẫn chính xác đoạn chữ; nếu
  không đủ chứng cứ, kết quả vẫn là `REVIEW`.
- Không có SMB username/password trong repository, workflow hoặc container.

## Khởi động trên MacBook

Docker Desktop phải đang mở. Chạy:

```bash
cd /Users/trannhatphi/Desktop/n8n_LS
./scripts/start-macos.sh
```

Lần đầu, script tự tạo `.env` và sinh khóa mã hóa riêng cho n8n. Không còn bước
kiểm tra `/Volumes/LeadingstarSMB`.

Sau khi container chạy:

1. Mở [http://localhost:5678](http://localhost:5678).
2. Tạo tài khoản owner nếu n8n yêu cầu.
3. Import `n8n/hanger-automation.workflow.json`.
4. Mở node **Upload Order và SOF**.
5. Chọn **Execute Workflow** để lấy Test URL, hoặc Publish workflow để dùng
   Production URL
   `http://localhost:5678/form/f63ddfc0-8c14-4f22-9791-d13d5f6bc379`.
6. Trên form, chọn file đơn hàng; tại ô **Các file SOF tương ứng**, giữ
   `Command` để chọn nhiều file (hoặc chọn một nhóm file), rồi bấm xử lý.

Nếu workflow đã được import từ trước, file JSON sửa trong repo **không tự cập
nhật** workflow đang lưu trong n8n. Để giữ các chỉnh sửa hiện có, mở workflow
đang chạy và sửa riêng node **Upload Order và SOF**: trường SOF bật chọn nhiều
file và `acceptFileTypes` thành `.xls,.xlsx,.docx,.pdf`; trường **File đơn
hàng** cũng đổi `acceptFileTypes` thành `.xls,.xlsx,.docx,.pdf`; trong node
**Validate Upload**, dùng phần kiểm tra định dạng từ bản JSON mới. Sau đó lưu và Publish
lại workflow. Không import đè toàn bộ workflow nếu đã chỉnh các node khác.

## Đơn hàng Word/PDF

Ô **File đơn hàng** nhận `.xls`, `.xlsx`, `.docx` và `.pdf`. Với Word/PDF, worker
đọc bảng dòng hàng theo tiêu đề cột (`Qty Ordered`, `Description`,
`Style Number`, `Account`, `Hang/Flat`…), ghép phần mô tả bị xuống dòng vào đúng
dòng hàng phía trên, bỏ dòng tổng cộng, và đọc các giá trị khai báo một lần cho
cả đơn (`PO number: PO07564`) rồi áp cho mọi dòng. Cột không có tên tương ứng —
ví dụ `Unit Price`, `Amount` — bị bỏ chứ không gán bừa vào cột khác. Một trường
được ghi hai lần với hai giá trị khác nhau sẽ bị bỏ trống kèm ghi chú.

Purchase Order in thường không có cột `Account` và `Hang/Flat`. Khi đó mọi dòng
vào `REVIEW` với ghi chú liệt kê đúng những cột file không có, thay vì bị bỏ qua
im lặng. Muốn chạy tự động trọn vẹn thì đơn hàng phải có `Account` (theo cột
hoặc theo dòng `Account: H040M` trong văn bản) và `Hang/Flat`.

PDF scan không có lớp chữ bị từ chối, không OCR. `audit.json` có khối
`order_source` ghi định dạng, số dòng đọc được, cột đã có, cột còn thiếu và các
giá trị khai báo ở cấp đơn hàng.

Mỗi tên file SOF vẫn phải chứa Account tương ứng, ví dụ Account `H040M` cần file như:

```text
H040M Stock Replenishment SO Form 8.31.26.xlsx
```

Điều này giữ nguyên nguyên tắc không đoán file SOF. Nếu tên SOF không khớp
Account, dòng sẽ bị bỏ qua vì không có SOF được chọn cho account đó.

Chỉ Word `.docx` được hỗ trợ, chưa hỗ trợ Word `.doc`. PDF phải có lớp chữ có thể
trích xuất. Với Word chứa nội dung quan trọng dưới dạng ảnh, phần chữ trong ảnh
sẽ không được đọc. Không có OCR hoặc mô hình nhìn ảnh trong luồng này.

## Bật DeepSeek

Không gửi API key qua chat và không đặt API key trong workflow JSON. Mở `.env`
trên MacBook rồi thêm hoặc chỉnh các dòng:

```dotenv
HANGER_LLM_ENABLED=true
DEEPSEEK_API_KEY=sk-your-key-here
DEEPSEEK_MODEL=deepseek-flash
```

Sau đó build lại worker:

```bash
docker compose up -d --build hanger-worker
docker compose exec hanger-worker python -c "import urllib.request; print(urllib.request.urlopen('http://127.0.0.1:8080/health').read().decode())"
```

DeepSeek nhận các nhóm dòng đơn hàng đã bỏ PO và số lượng, cùng phần văn bản có
địa chỉ ô từ những sheet SOF Excel liên quan; với Word/PDF, DeepSeek chỉ nhận
chữ được trích xuất theo section và số dòng. Các dòng trùng điều kiện chỉ tạo một yêu
cầu phân loại. Mặc định mỗi lần chạy tối đa 20 API calls, batch 12 nhóm/call và
chỉ nhận kết quả có confidence từ 0.65. Đây là ngưỡng đã hiệu chuẩn trên dữ liệu
thực tế; thấp hơn 0.60 bắt đầu tăng lỗi rõ rệt. Có thể chỉnh các biến tương ứng trong
`.env`; API key chỉ được truyền vào container `hanger-worker`.

Ngoài các SOF có sheet theo nhóm sản phẩm như `TOPS`, `BOTTOMS`, `SETS`, worker
cũng hiểu bảng theo Label trong sheet `PACKAGING` của SOF HBE. Một quyết định từ
`PACKAGING` chỉ được nhận khi vùng ô trích dẫn chứa đúng Label của dòng đơn hàng.
`FLATPACKED` có thể xác nhận Flat/không hanger. Nếu SOF ghi `Follow US Manual`
nhưng không chứa mã hanger cụ thể, dòng vẫn ở `REVIEW` để tránh tự đoán mã.

### Dùng OpenAI thay cho DeepSeek

Engine dùng cùng giao thức Chat Completions và đọc key từ biến môi trường hoặc
Streamlit Secrets; không ghi API key vào mã nguồn. Cấu hình:

```toml
HANGER_LLM_ENABLED = "true"
HANGER_LLM_PROVIDER = "openai"
OPENAI_API_KEY = "sk-..."
OPENAI_MODEL = "gpt-4.1-mini"
OPENAI_BASE_URL = "https://api.openai.com/v1"
```

Các rule cố định vẫn chạy trước; OpenAI chỉ xử lý các dòng còn `REVIEW` và kết
quả vẫn phải vượt qua kiểm tra bằng chứng SOF.

## Kết quả

Sau khi chạy xong, trình duyệt tự tải file `_KETQUA.xlsx`: node **Tải File Kết
Quả** lấy file từ worker, rồi node **Tải File Về Máy** (`n8n Form`, Page Type =
*Form Ending*, Respond With = *Return Binary File*, field `data`) vừa hiện màn
hình tổng kết vừa đẩy file xuống máy.

Đừng dùng node `Respond to Webhook` ở đây: Form Trigger từ typeVersion 2.2 trở
lên từ chối node đó (*"The 'Respond to Webhook' node is not supported in
workflows initiated by the 'n8n Form Trigger'"*), và trang form submit bằng
`fetch()` rồi đọc `response.text()` nên dữ liệu nhị phân sẽ bị đổ thành chuỗi rác
ra trang thay vì tải xuống. Vì vậy `Respond When` phải giữ nguyên *Workflow
Finishes* — n8n tự chuyển sang chế độ response node khi thấy có node Form nối phía sau.

Bản đầy đủ vẫn nằm trên máy:

```text
data/output/<ten-file>_checked_YYYYMMDD_HHMMSS.xlsx
data/output/<ten-file>_checked_YYYYMMDD_HHMMSS.audit.json
```

Audit JSON chứa kết quả cho từng dòng, SOF source, sheet/cells hoặc
section/lines/excerpt, match method,
confidence, trạng thái và validation note.

## Lệnh vận hành

```bash
# Trạng thái
docker compose ps

# Xem log xử lý
docker compose logs -f hanger-worker

# Xem log n8n
docker compose logs -f n8n

# Dừng nhưng giữ workflow và cấu hình n8n
docker compose down

# Chạy lại
docker compose up -d

# Build lại sau khi sửa Python hoặc rules
docker compose up -d --build hanger-worker
```

Không dùng `docker compose down -v` trừ khi muốn xóa database, workflows và
credential encryption state của n8n.

## Giới hạn upload

Giới hạn hiện tại là 100 MiB cho toàn bộ request. Có thể chỉnh đồng thời:

- `N8N_FORMDATA_FILE_SIZE_MAX` trong `compose.yaml`.
- `HANGER_MAX_UPLOAD_MB` trong `compose.yaml`.

Số SOF tối đa cho mỗi lần gửi mặc định là 20, cấu hình bằng
`HANGER_MAX_SOF_FILES` trong `.env`. Form và worker đều chặn ở mức 20; nếu đổi
giới hạn này, cần đổi cùng giá trị trong node **Validate Upload** của workflow.

Không nên tăng giới hạn nếu Docker Desktop chưa được cấp đủ RAM và dung lượng đĩa.

## Rule và SOF

- `rules/hanger_rules.json`: rules đã duyệt; hiện chứa các case H040M đã cung cấp.
- `config.docker.json`: output path và rules path trong worker.
- `src/hanger_automation.py`: rule engine và Excel writer.
- `src/worker_api.py`: nhận file multipart từ n8n và quản lý thư mục tạm.

Mỗi lần sửa rule hoặc code Python, chạy lại:

```bash
docker compose up -d --build hanger-worker
```

## Kiểm thử

```bash
python3 -m pip install -r requirements.txt
python3 -m pytest -q
```

Test bao phủ tám dòng H040M mẫu, hai chiều mismatch, Hang/Flat trống, PO có số 0
đầu, SOF mơ hồ, source không hợp lệ, upload một/nhiều SOF, kiểm tra gói ZIP và
bảo toàn dữ liệu hanger nhập thủ công.
