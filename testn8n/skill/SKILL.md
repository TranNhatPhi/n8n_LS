
---

name: kiem-hanger
description: "Tra mã hanger cho đơn hàng may mặc bằng cách đọc file SOF (Ship Order Form). Dùng skill này bất cứ khi nào người dùng đưa lên file SOF, file đơn hàng may mặc, hoặc nhắc tới: hanger, mã hanger, hanger code, SOF, SO Form, Ship Order Form, sizer, size clip, FLATPACKED, Hang/Flat, treo hay gấp, hanger color, sticker hanger, account kiểu K413M / H040M / N748M, hoặc hỏi món hàng này dùng hanger gì. Cũng dùng khi người dùng muốn kiểm lại một bảng hanger đã tra, hoặc hỏi vì sao một dòng phải để REVIEW. Kích hoạt cả khi họ không nói chữ 'hanger' — ví dụ 'cái SO Form này quy định gì cho áo thun size 4-6x', 'tra giúp mấy dòng này', 'file này nói size S dùng loại nào'."
---

# Tra mã hanger từ file SOF

Bạn đọc yêu cầu hanger từ file SOF và báo cáo theo từng dòng đơn hàng. Bạn
**trích xuất**, không phán đoán. Không bao giờ quyết định một món hàng "đáng lẽ"
cần hanger gì.

Lý do nguyên tắc này nghiêm ngặt: kết quả `MATCHED` được ghi thẳng vào đơn hàng
và **không ai kiểm lại**. Một dòng đoán sai đi tới tận xưởng. Còn một dòng để
`REVIEW` chỉ tốn vài phút của người kiểm. Hai cái giá đó chênh nhau rất xa, nên
khi phân vân, chọn `REVIEW`.

## Cần những gì

| File | Vai trò |
| -------------------------- | ------------------------------------------------------------------- |
| File SOF của account | Nguồn sự thật duy nhất về mã hanger |
| Dòng đơn hàng cần tra | Có Ref#, Style, mô tả hàng, size, account |
| `Rule_cho_Hanger.docx` | Định nghĩa nhóm sản phẩm của khách (nếu người dùng có) |

Thiếu file SOF thì dừng lại và hỏi — không có nó thì mọi câu trả lời đều là đoán.

Người dùng không đưa `Rule_cho_Hanger.docx` thì dùng
[references/nhan-dien-nhom.md](references/nhan-dien-nhom.md), là bản sao định
nghĩa nhóm của khách. Đừng phân loại theo hiểu biết chung về ngành may: khách có
quy ước riêng, ví dụ romper 1 món thuộc TOPS chứ không phải COVERALLS.

## Đọc file SOF bằng script, đừng nhìn bằng mắt

Nhìn bảng rồi gõ lại địa chỉ ô là nguồn sai lầm lớn nhất: giá trị đúng nhưng ô
sai, và người kiểm mở đúng ô ấy thì không thấy gì.

```bash
python3 scripts/doc_sof.py "SOF.xlsx"                          # liệt kê sheet
python3 scripts/doc_sof.py "SOF.xlsx" --sheet SETS             # cả sheet
python3 scripts/doc_sof.py "SOF.xlsx" --sheet SETS --range B8:G20
```

Script in ra dạng `B11=497B/6008B`, nên địa chỉ ô chỉ việc chép lại.

Không chạy được script (môi trường không có Python hoặc openpyxl) thì vẫn làm
tiếp được, nhưng phải nói rõ với người dùng rằng địa chỉ ô là ước lượng và cần
họ mở file kiểm lại.

## Quy trình, làm đúng thứ tự, cho từng dòng

1. **Phân loại nhóm sản phẩm** từ mô tả hàng, theo định nghĩa của khách.
2. **Chọn sheet** trong SOF phụ trách nhóm đó.
3. **Tìm dòng khớp size** của đơn hàng, và khớp Label nếu sheet chia theo Label.
4. **Trích dẫn 1–3 vùng ô** trên đúng sheet đó, chứa mọi giá trị sắp báo cáo.
5. **Đọc từng giá trị nguyên văn** từ các ô vừa trích dẫn.
6. **Đặt trạng thái sau cùng**, dựa trên những gì thật sự tìm được.

Đặt trạng thái cuối cùng là có chủ đích. Quyết `MATCHED` trước rồi mới đi tìm
bằng chứng là cách người ta tự thuyết phục mình rằng một ô mơ hồ là đủ rõ.

## Trích dẫn

SOF thường để mã hanger trong bảng, còn màu hanger, sizer và quy định sticker
nằm ở câu chữ bên dưới bảng. **Trích dẫn cả hai vùng.** Đừng bỏ một giá trị chỉ
vì bảng không chứa nó — hãy mở rộng vùng trích dẫn.

Đọc màu ngay từ ô chứa mã khi ô đó có nêu màu. Ô ghi
`485B/6010B Black Crown Sizer w/ White Lettering` là hanger **BLACK**. Chỉ dùng
câu mô tả chung khi ô chứa mã và cả dòng đó không nói gì về màu.

Một câu chung chung như "dùng hanger nhựa trắng cho pack gộp size" **không** đè
lên cột size mà đơn hàng thật sự rơi vào. Câu chung mô tả trường hợp khác.

Chỉ trích dẫn vùng thật sự đã dùng, và chỉ trên đúng một sheet đã nêu.

## Cột size

Bảng mỗi nhóm chạy từ size nhỏ đến lớn theo các cột, ví dụ
`Newborn | Infant | Toddler | 4-6x and 4/7 | 7-16 and 8-20`.

Một size không in trong tiêu đề vẫn thuộc cột có khoảng bao phủ nó.

Size chữ người lớn **S, M, L, XL, XXL lớn hơn mọi khoảng size trẻ em**, nên
thuộc cột lớn nhất, cùng cột với `7-16 and 8-20`. Đọc mã và màu từ cột đó.
Đừng trả `REVIEW` chỉ vì size chữ không được ghi trong tiêu đề — cột đó vẫn là
cột điều chỉnh. Cột này thường thuộc dòng hanger "B" màu đen; lấy màu từ chính ô
đó, không lấy từ câu nói về pack gộp size.

## Ba trạng thái của một giá trị

Báo cáo mọi trường **đúng nguyên văn** như trong ô đã trích dẫn — cùng chữ, cùng
số, cùng dấu câu. Không chuẩn hoá, không dịch, không làm gọn. Người kiểm cần
đối chiếu được từng ký tự với file gốc.

Phân biệt ba trường hợp, đừng lẫn:

| | Nghĩa |
| --------------- | ------------------------------- |
| Giá trị thật | File có nêu rõ |
| `NO` | File nói rõ là không cần |
| `KHÔNG_NÊU` | File im lặng, không nhắc gì |

Với `color_sizer`, `sticker_hanger`, `size_sticker_hanger`: SOF chỉ liệt kê phụ
kiện ở chỗ nào cần, nên im lặng nghĩa là không cần — trả `NO`.

Với `hanger_code` và `hanger_color`: im lặng nghĩa là **chưa tìm ra quy định** —
trả `KHÔNG_NÊU`, dòng đó chuyển cho người kiểm.

Khác biệt này quan trọng vì hai loại im lặng mang ý nghĩa trái ngược nhau.

## Quy tắc hàng treo

Hàng treo thì luôn treo trên một chiếc hanger có thật, và chiếc hanger đó có
màu. Nên khi `Hang/Flat` là `Hang`, thì `hanger_code` và `hanger_color` **bắt
buộc** là giá trị thật đọc từ file.

`NO` hoặc `KHÔNG_NÊU` ở một trong hai trường nghĩa là chưa tìm đúng dòng: mở
rộng vùng trích dẫn, hoặc trả `REVIEW`. Không bao giờ báo `Hang` mà thiếu mã
hoặc thiếu màu — đó là mâu thuẫn tự thân.

## Hàng gấp

`FLATPACKED` trong cột `FLATPACKED/HANGER` nghĩa là `Flat`, và cả năm trường
hanger đều là `NO`.

Một dòng có điều kiện chỉ hỗ trợ `Hang` khi điều kiện về sản phẩm hoặc size của
nó đúng với dòng đơn hàng đang xét.

## Trạng thái dòng

**`MATCHED`** — mọi giá trị báo cáo đều đọc được từ vùng ô đã trích dẫn.

**`REVIEW`** — có bất kỳ điểm nào mơ hồ hoặc thiếu, hoặc quy định trỏ tới một
tài liệu không được cung cấp (ví dụ "Follow US Manual" mà không kèm manual),
hoặc muốn điền thì phải đoán. Khi `REVIEW`, để trống các trường kết quả và nói
rõ lý do.

`REVIEW` là câu trả lời đúng và được mong đợi, không phải thất bại.

## Kiểm lại trước khi giao

Chạy script này cho từng dòng `MATCHED`:

```bash
python3 scripts/kiem_trich_dan.py "SOF.xlsx" --sheet SETS --range B10:G12 \
    --gia-tri "497B/6008B" --gia-tri "Newborn"
```

Nó mở đúng vùng ô và xác nhận từng giá trị có thật ở đó. Bắt được cả ba lỗi hay
gặp: sheet không tồn tại, vùng ô ngoài phạm vi dữ liệu, và giá trị không nằm
trong vùng đã trích dẫn.

Script báo thiếu thì có hai khả năng, và phải chọn đúng: hoặc vùng trích dẫn
quá hẹp (mở rộng), hoặc giá trị không thật sự có trong file (sửa lại giá trị,
hoặc chuyển dòng đó sang `REVIEW`). Đừng nới vùng ô một cách máy móc cho tới khi
script hết báo lỗi — như vậy là trích dẫn nửa sheet để hợp thức hoá một giá trị
không đọc được ở đâu cả.

## Định dạng trả về

Trả một bảng đúng 12 cột theo thứ tự này, mỗi dòng đơn hàng một dòng:

```
Account | Ref number | Style | Product category | Hang/Flat | Hanger code |
Hanger color | Color sizer | Sticker hanger | Size sticker hanger | Sheet | Cells
```

Quy ước điền:

- `Account` viết hoa
- `Hang/Flat` chỉ ghi `Hang` hoặc `Flat`
- `Sheet` là tên sheet đã dùng
- `Cells` là vùng đã trích dẫn, ví dụ `B10:G12`; nhiều vùng cách nhau dấu phẩy,
  tối đa 3
- Dòng `REVIEW`: để trống từ `Hanger code` đến `Cells`

Thứ tự 12 cột này khớp với bảng rule mà hệ thống tự động đọc được, nên người
dùng dán thẳng kết quả vào file rule của họ là lần sau khỏi phải tra lại.

Sau bảng, thêm mục **Lý do** liệt kê từng dòng `REVIEW` kèm lý do ngắn gọn và
nêu rõ cần thêm thông tin gì để quyết được. Người dùng cầm phần này đi hỏi khách
hoặc mở đúng chỗ trong SOF, nên hãy cụ thể: "sheet SETS không nêu màu cho cột
7-16" hữu ích hơn nhiều so với "không đủ thông tin".

## Ví dụ

**Đơn hàng:** account `K413M`, Ref `NKG-KPN743`, Style `06N743`,
mô tả `2PC SHORT SET`, size `Newborn`.

Phân loại `SETS` theo định nghĩa của khách (mô tả kết thúc bằng SET). Mở sheet
`SETS`, cột `Newborn`, dòng "Hanger for 2 pc short sets" cho `497B/6008B`.

```
K413M | NKG-KPN743 | 06N743 | SETS | Hang | 497B/6008B | WHITE |
White size clip / black lettering | NO | NO | SETS | B10:G12,B16:C16
```

Hai vùng ô vì mã nằm trong bảng ở `B10:G12`, còn màu và sizer nằm ở dòng mô tả
`B16:C16` bên dưới bảng.

sau đó gủi lại tôi file đơn hàng đầy đủ 
