# M-Docs — Add-in Word định dạng văn bản hành chính tiếng Việt

Add-in Word dạng template macro-enabled (`M-Docs.dotm`), cài vào thư mục **STARTUP** của Word để tự nạp mỗi lần mở Word. Add-in tạo một thẻ ribbon riêng tên **M-Docs** với 4 nhóm nút theo bố cục kiểu V-FORMAT.

- Chỉ dùng **Word trên Windows, Word 2016 trở lên** (kể cả bản không phải Microsoft 365).
- **Không phụ thuộc thư viện ngoài, không cần mạng.** Mã QR được mã hóa hoàn toàn bằng VBA thuần.
- Mọi lệnh đều có thể hoàn tác bằng **một lần Ctrl+Z** (mỗi lệnh chạy trong một Custom Undo Record riêng).

## Chức năng

| Nhóm | Nút | Tác dụng |
|---|---|---|
| Bố cục trang | Căn lề Trang A4 | Mọi section: khổ A4, lề trên 20mm, dưới 20mm, trái 30mm, phải 15mm (Nghị định 30/2020/NĐ-CP). Section ngang giữ nguyên hướng giấy, bộ lề được xoay theo. Không đổi khoảng cách header/footer. |
| Bố cục trang | Chèn trang ngang / Chèn trang dọc | Chèn 1 trang A4 mới tại con trỏ bằng 2 ngắt section "next page". Nội dung sau giữ nguyên hướng giấy và lề cũ. Báo lỗi rõ ràng nếu con trỏ đang trong bảng. |
| Bố cục trang | Xóa trang trắng | Xóa các trang trắng thật (chỉ ký tự trống, không ảnh/bảng/field/content control/chú thích). Trang trắng do ngắt section thì giữ nguyên và báo cáo. Trang cuối không xóa được đoạn cuối thì thu nhỏ đoạn đó xuống 1pt. |
| Công cụ | Xóa khoảng trắng lỗi | Gộp nhiều dấu cách thành 1, xóa dấu cách thừa ở đầu/cuối mỗi đoạn. Giữ nguyên định dạng. |
| Công cụ → Công cụ Văn bản | Xóa khoảng trắng lỗi & dấu cách bất hợp lý | Thêm: bỏ dấu cách trước `, ; : ! ? ) ] } %`, bỏ dấu cách sau `( [ {`, thêm dấu cách thiếu sau `, ; : ! ?`, thêm dấu cách sau dấu chấm chỉ khi chữ sau là CHỮ HOA và từ trước dấu chấm dài ≥ 3 ký tự. **Không** bỏ dấu cách trước dấu chấm (giữ dòng chấm `........`). Không đụng số thập phân `3,5`, giờ `12:30`, URL, email, đường dẫn, viết tắt `TP.HCM`, `v.v.` |
| Công cụ → Công cụ Văn bản | Chuyển sang tiếng Việt không dấu | Chỉ vùng bôi đen (không bôi đen sẽ báo). Xử lý cả đ/Đ. Ví dụ `Nguyễn Văn Đạt` → `Nguyen Van Dat`. |
| Công cụ → Công cụ Văn bản | Sửa dấu thanh toàn văn bản | Kiểu cũ → kiểu mới khi cặp nguyên âm ở **cuối từ**: `hoà→hòa`, `khoẻ→khỏe`, `thuý→thúy`, `Thuỷ→Thủy`, `HOÀ→HÒA`. Giữ nguyên: `hoàng, toán, loài, ngoài, huỳnh, quý, quỳnh, quỷ`. Phạm vi: vùng bôi đen nếu có, không thì cả tài liệu. |
| Chèn nhanh | Tạo mã QR | Từ vùng bôi đen. Byte mode, **UTF-8**, mức sửa lỗi **M**, tự chọn version 1–10 (tối đa 213 byte UTF-8), Reed-Solomon đầy đủ, chọn mask theo điểm phạt thấp nhất, vùng yên tĩnh 4 module. Ảnh 4cm × 4cm, Alt Text = nội dung gốc. Mặc định thay thế vùng bôi đen; nếu vùng chọn kết thúc bằng dấu đoạn thì chèn ở cuối đoạn (trước dấu đoạn). |
| Thông tin | Giới thiệu | Hiện tên, phiên bản, danh sách chức năng. |

Các lệnh ở nhóm Công cụ: có vùng bôi đen thì chạy trên vùng bôi đen, không thì chạy trên cả tài liệu, và báo rõ phạm vi đã áp dụng. Tài liệu đang bị bảo vệ hoặc chỉ đọc sẽ được báo rõ, không crash. Track Changes đang bật thì lệnh vẫn chạy (thay đổi được ghi nhận) và có nhắc trong báo cáo.

## Cấu trúc thư mục

```
M-Docs/
  M-Docs.dotm           FILE ADD-IN ĐÃ DỰNG SẴN — chỉ cần cài là dùng
  Install-M-Docs.bat    Cài one-click: chép .dotm vào STARTUP + gỡ Mark-of-the-Web
  Uninstall-M-Docs.bat  Gỡ one-click: xóa .dotm khỏi STARTUP
  Build-M-Docs.ps1      Script build lại .dotm từ mã nguồn (khi tự sửa mã)
  Install-M-Docs.bat    Bản PowerShell của script cài (tùy chọn thay cho .bat)
  Uninstall-M-Docs.bat  Bản PowerShell của script gỡ (tùy chọn thay cho .bat)
  README.md                File này
  TEST-CHECKLIST.md        Danh sách kiểm thử thủ công từng nút + ca biên
  src/
    modConfig.bas          Hằng số cấu hình (lề, cỡ QR, QR_MODE, SHOW_REPORT...)
    modUtil.bas            VN() giải mã \uXXXX, khung lệnh, kiểm thử tự động
    modRibbon.bas          Callback của ribbon
    modLayout.bas          Căn lề A4, chèn trang ngang/dọc
    modBlankPages.bas      Xóa trang trắng
    modWhitespace.bas      Xóa khoảng trắng lỗi + dấu cách quanh dấu câu
    modVietnamese.bas      Bỏ dấu tiếng Việt, sửa dấu thanh
    modQR.bas              Bộ mã hóa QR thuần VBA + ghi BMP + chèn ảnh
    customUI14.xml         Định nghĩa thẻ ribbon "M-Docs"
  dist/                    Chứa M-Docs.dotm sau khi build bằng script PowerShell
```

## Cách build (chỉ khi muốn dựng lại từ mã nguồn)

File `M-Docs.dotm` trong thư mục này **đã được dựng sẵn** — muốn dùng ngay thì bỏ qua mục này, xem thẳng mục "Cách cài / gỡ". Chỉ khi bạn tự sửa mã trong `src\` và muốn dựng lại file `.dotm` mới cần build.

Điều kiện một lần duy nhất: bật **Trust access to the VBA project object model** (script build có kiểm tra và sẽ dừng kèm hướng dẫn nếu chưa bật):

1. Mở Word → **File → Options → Trust Center → Trust Center Settings…**
2. Mục **Macro Settings** → tích **Trust access to the VBA project object model** → OK.

Sau đó, trong thư mục `M-Docs`:

```powershell
powershell -ExecutionPolicy Bypass -File .\Build-M-Docs.ps1
```

Script sẽ: mở Word qua COM → import 8 module `.bas` vào VBProject của một template mới → lưu thành `dist\M-Docs.dotm` (định dạng 15 = XML Template Macro-Enabled) → mở gói `.dotm` bằng `System.IO.Compression` để thêm part `customUI/customUI14.xml`, quan hệ `http://schemas.microsoft.com/office/2007/relationships/ui/extensibility` vào `_rels/.rels` và Override content-type tương ứng.

## Cách cài / gỡ

**Cài one-click (khuyên dùng):** nhấp đúp **`Install-M-Docs.bat`** (để cùng thư mục với `M-Docs.dotm`). Script chép file vào `%APPDATA%\Microsoft\Word\STARTUP` và gỡ cờ Mark-of-the-Web. **Khởi động lại Word** — thẻ **M-Docs** sẽ xuất hiện trên ribbon. Nếu Word đang mở, script sẽ nhắc tắt và mở lại.

**Gỡ one-click:** nhấp đúp **`Uninstall-M-Docs.bat`** — file trong STARTUP bị xóa, mở lại Word là mất thẻ.

Cách khác (PowerShell, kết quả tương đương):

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-M-Docs.bat
powershell -ExecutionPolicy Bypass -File .\Uninstall-M-Docs.bat
```

## Cấu hình (modConfig.bas)

| Hằng số | Ý nghĩa |
|---|---|
| `MARGIN_TOP_MM` … `MARGIN_RIGHT_MM` | Bộ lề chuẩn Nghị định 30 (20/20/30/15 mm) |
| `QR_SIZE_CM` | Kích thước ảnh QR khi chèn (mặc định 4 cm) |
| `QR_MODE` | `"REPLACE"` = thay thế vùng bôi đen (mặc định); `"AFTER"` = chèn ngay sau vùng bôi đen |
| `QR_MAX_VERSION` | Version QR tối đa (10) |
| `QR_MODULE_PIXELS`, `QR_QUIET_MODULES` | Độ phân giải BMP và vùng yên tĩnh |
| `SHOW_REPORT` | `True`/`False`: bật/tắt hộp thoại báo kết quả sau mỗi lệnh |

Sửa hằng số xong thì build và cài lại.

## Kiểm thử tự động

Trong Word, nhấn **Alt+F8** và chạy:

- **`MDocs_RunAllTests`** — kiểm tra hàm `VN()`, bảng bỏ dấu (ví dụ `Nguyễn Văn Đạt → Nguyen Van Dat`), bảng sửa dấu thanh (cả các ví dụ phải **giữ nguyên** như `hoàng`, `quý`), đồng thời tạo file QR mẫu. Chi tiết từng ca nằm ở cửa sổ Immediate (**Ctrl+G** trong trình soạn VBA), tổng kết hiện bằng MsgBox.
- **`MDocs_CreateTestQRFile`** — tạo file `%TEMP%\MDocs_QR_Test.bmp` chứa mã QR của câu **"Xin chào Việt Nam"** để quét thử bằng điện thoại.

Kiểm thử thủ công đầy đủ trên tài liệu thật: xem **TEST-CHECKLIST.md**.

> Ghi chú kiểm chứng thuật toán QR: phần mã hóa QR (UTF-8, chọn version, Reed-Solomon, ghép khối, đặt dữ liệu, mask, format/version info) đã được đối chiếu ngoại tuyến với một bộ mã hóa QR tham chiếu độc lập — ma trận module trùng khớp từng ô cho cả 8 mask ở các version 2, 7, 8 và 10 (kể cả chuỗi kiểm thử "Xin chào Việt Nam").
>
> Ghi chú về file `.dotm` dựng sẵn: file được lắp ráp theo đúng định dạng gói của Word (vbaProject.bin với đủ 8 module + ThisDocument, ribbon customUI14) và đã kiểm tra bằng hai bộ đọc độc lập (oletools/olevba và compoundfiles): cả 8 module trích xuất ra khớp từng ký tự với mã nguồn `src\*.bas`. Phần chưa kiểm chứng được là hành vi chạy thật trong Word (thẻ ribbon hiện, lệnh chạy trên tài liệu), vì khâu đó cần Word trên Windows — hãy chạy `MDocs_RunAllTests` sau khi cài để xác nhận trên máy của bạn.

## Cách cập nhật phiên bản

1. Sửa mã trong `src\*.bas` (hoặc `customUI14.xml`).
2. Tăng `APP_VERSION` trong `src\modConfig.bas`.
3. Chạy lại `Build-M-Docs.ps1` để dựng bản `.dotm` mới, rồi chạy `Install-M-Docs.bat` (ghi đè bản cũ trong STARTUP).
4. Khởi động lại Word. Nếu ribbon chưa đổi, đóng hẳn Word (kiểm tra không còn tiến trình `WINWORD.EXE`) rồi mở lại.

## Ký số macro (tùy chọn)

1. Tạo chứng thư tự ký: chạy `SelfCert.exe` của Office (thường ở `C:\Program Files\Microsoft Office\root\Office16\SELFCERT.EXE`), đặt tên chứng thư.
2. Mở `M-Docs.dotm` → **Alt+F11** → chọn project của template → **Tools → Digital Signature…** → chọn chứng thư → OK → lưu lại.
3. Trên máy khác, chứng thư tự ký chỉ có tác dụng khi được thêm vào kho **Trusted Publishers** của máy đó. Ký số phù hợp nhất khi dùng chứng thư của tổ chức (CA nội bộ hoặc thương mại).

## Xử lý sự cố thường gặp

**Macro bị chặn / không chạy được**
- File vừa tải về thường dính Mark-of-the-Web: script cài đặt đã `Unblock-File`; nếu chép tay, nhấp phải file → **Properties** → tích **Unblock** → OK.
- Thêm thư mục STARTUP thành **Trusted Location**: File → Options → Trust Center → Trust Center Settings → Trusted Locations → Add new location → chọn `%APPDATA%\Microsoft\Word\STARTUP`. (Mặc định thư mục STARTUP của Word vốn đã được tin cậy cho global template, nhưng chính sách của một số máy có thể tắt.)
- Macro Settings không được để ở mức "Disable all macros without notification".

**Thẻ "M-Docs" không hiện**
- Kiểm tra file có đúng ở `%APPDATA%\Microsoft\Word\STARTUP\M-Docs.dotm` không và đã khởi động lại Word chưa.
- Mở file `.dotm` bằng trình giải nén (đổi đuôi thành `.zip`) để chắc chắn có thư mục `customUI\customUI14.xml` — nếu thiếu, chạy lại script build (bước inject nằm cuối script).
- Nếu Word báo lỗi Custom UI khi mở, XML có thể sai: mở `src\customUI14.xml` bằng trình duyệt hoặc VS Code để kiểm tra well-formed, build lại.
- Kiểm tra File → Options → Customize Ribbon: thẻ tùy chỉnh có bị tắt không.

**Ribbon không cập nhật sau khi cài bản mới**
- Đóng hẳn Word (kể cả cửa sổ ẩn khi mở tài liệu từ Outlook/Explorer), mở lại.
- Chắc chắn không còn bản `M-Docs.dotm` cũ ở một thư mục STARTUP khác (ví dụ thư mục STARTUP tùy chỉnh trong Word Options → Advanced → File Locations).

**Script build báo lỗi AccessVBOM**
- Xem mục "Cách build": bật Trust access to the VBA project object model rồi chạy lại. Đây là yêu cầu của chính Word cho mọi thao tác import module qua COM, không phải của riêng bộ công cụ này.

**Lỗi biên dịch liên quan tham chiếu (References)**
- Mã nguồn được viết theo kiểu late binding cho phần ribbon (`control As Object`) và tránh hằng của thư viện Office, nên project **không cần thêm tham chiếu nào** ngoài các tham chiếu mặc định có sẵn trong template Word (VBA, Word, stdole, Office). Nếu Word báo "Can't find project or library", mở Alt+F11 → Tools → References và kiểm tra có dòng nào bị đánh dấu **MISSING** không; bỏ tích dòng đó (hoặc trỏ lại đúng file thư viện) rồi chạy lại.

**Một nút mất icon nhưng vẫn bấm được**
- Tên `imageMso` là icon có sẵn của Office; một vài tên icon có thể khác nhau giữa các bản Office. Nút vẫn hoạt động bình thường, chỉ thiếu hình. Muốn đổi icon: sửa thuộc tính `imageMso` trong `src\customUI14.xml` rồi build lại.

## Giới hạn đã biết

- **QR tối đa version 10, mức M = 213 byte UTF-8.** Tiếng Việt có dấu tốn 2–3 byte/ký tự nên nội dung thực tế khoảng 70–110 ký tự. Dài hơn sẽ báo lỗi rõ ràng, không treo Word.
- **Sửa dấu thanh** dùng Find với `MatchCase = True` rồi kiểm tra "cặp nguyên âm ở cuối từ" bằng mã, thay vì `MatchWholeWord` — vì một cặp nguyên âm nằm *trong* một từ không bao giờ là "whole word" đối với Find, dùng MatchWholeWord sẽ không bao giờ khớp các ví dụ như `khoẻ`, `thuý`. Kết quả cuối cùng đúng đặc tả (đổi khi ở cuối từ, giữ nguyên `hoàng/quý/...`).
- **Xóa trang trắng** xác định số trang bằng `ComputeStatistics` và vị trí trang bằng `GoTo` — cần Word đã phân trang xong (chế độ Print Layout cho kết quả ổn định nhất).
- **Phương án field DISPLAYBARCODE không được dùng làm đường chính**: field này phụ thuộc bản Word và không đảm bảo đúng UTF-8 với tiếng Việt có dấu. Nếu bắt buộc phải dùng cho nội dung ASCII ngắn, chèn thủ công field `DISPLAYBARCODE "noi dung" QR \s 100` (Ctrl+F9) — chỉ nên dùng khi nội dung không có dấu tiếng Việt.
- Các file `.bas` **phải giữ thuần ASCII**: mọi chuỗi tiếng Việt trong VBA đều đi qua hàm `VN("...\uXXXX...")`. Nếu bạn tự sửa mã và gõ tiếng Việt có dấu trực tiếp vào file `.bas`, chuỗi sẽ bị hỏng khi import (VBA lưu mã nguồn dạng ANSI). Nhãn trên ribbon thì ngược lại: viết tiếng Việt trực tiếp trong `customUI14.xml` (file UTF-8).
