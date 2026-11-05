# Hướng dẫn triển khai app Quản lý Dự án Stellar

Bộ file gồm 3 phần:
- `QLTK_DuLieu.xlsx`: dữ liệu gom từ 3 file LHS, RMH, UNI (25 dự án, 1.636 dòng nhóm/hạng mục, 1.149 ghi chú tuần).
- `Code.gs`: phần máy chủ, chạy trong Apps Script.
- `index.html`: trang web.

Làm lần lượt từ bước 1 đến bước 8, mất khoảng 20 phút.

## 1. Đưa dữ liệu lên Google Drive
1. Mở Google Drive bằng tài khoản của anh, vào thư mục muốn để app (ví dụ `QLTK_App`).
2. Kéo thả file `QLTK_DuLieu.xlsx` vào thư mục đó.
3. Chuột phải vào file → **Mở bằng → Google Trang tính**.
4. Trong trang vừa mở: **Tệp → Lưu dưới dạng Google Trang tính**. Từ giờ chỉ dùng bản Google Trang tính này; file .xlsx gốc có thể xóa.

## 2. Dán code vào Apps Script
1. Trong Google Sheet: **Tiện ích mở rộng → Apps Script**.
2. Ở tệp `Mã.gs` (hoặc `Code.gs`) có sẵn: xóa hết, dán toàn bộ nội dung file `Code.gs`.
3. Bấm dấu **+** cạnh "Tệp" → **HTML** → đặt tên đúng là `index` (không gõ đuôi .html) → xóa nội dung mẫu, dán toàn bộ file `index.html`.
4. Bấm biểu tượng **Lưu**.

## 3. Cấp quyền lần đầu
1. Trên thanh công cụ Apps Script, ô chọn hàm → chọn `onOpen` → **Chạy**.
2. Google hỏi quyền → chọn tài khoản của anh → **Nâng cao → Đi tới … (không an toàn)** → **Cho phép**. Đây là script của chính anh nên Google cảnh báo như vậy là bình thường.

## 4. Tạo mã đăng nhập Authenticator
1. Quay lại Google Sheet, tải lại trang. Sẽ có menu **QLTK** trên thanh menu.
2. **QLTK → Tạo mã đăng nhập cho người chưa có**. Script tạo mã bí mật 16 ký tự cho từng người trong tab NhanSu (cột `ma_bi_mat`).
3. **QLTK → Xem mã đăng nhập**. Lấy dòng của anh (Quang).
4. Trên điện thoại, mở Google Authenticator → **+** → **Nhập khóa thiết lập** → Tên tài khoản: `QLTK`; Khóa: chuỗi 16 ký tự; Loại khóa: **Dựa trên thời gian** → Thêm.

## 5. Triển khai thành trang web
1. Trong Apps Script: **Triển khai → Tùy chọn triển khai mới**.
2. Bánh răng cạnh "Chọn loại" → **Ứng dụng web**.
3. Thực thi với tư cách: **Tôi**. Ai có quyền truy cập: **Bất kỳ ai**.
4. **Triển khai** → sao chép đường link kết thúc bằng `/exec`. Đây là link app.

"Bất kỳ ai" chỉ để mở được trang đăng nhập. Muốn xem hay ghi dữ liệu vẫn phải chọn tên và nhập đúng mã 6 số.

## 6. Dùng thử
1. Mở link `/exec` (trên máy tính hoặc điện thoại) → chọn **Quang · Trưởng** → nhập mã 6 số trong Authenticator.
2. Muốn thử nhanh không cần mã: ở tab **CauHinh** đổi `che_do_thu` thành `Có`. Khi đó ai có link cũng vào được, nên thử xong phải đổi lại `Không`.

## 7. Thêm người dùng
1. Tab **NhanSu**: tên người phải viết đúng như cột `nguoi_pt` của tab CongViec (Bá Bộ, Hùng, Tuyển…), cột `kich_hoat` = `Có`.
2. 3 dòng "(Người duyệt 1–3)": sửa thành tên thật, `vai_tro` = `Duyệt`, `kich_hoat` = `Có`.
3. Chạy lại **QLTK → Tạo mã đăng nhập**, rồi gửi riêng mã bí mật cho từng người. Không gửi vào nhóm chat chung.
4. Vai trò: **Trưởng** sửa được mọi thứ; **Duyệt** chốt RAG, đánh dấu đã rà, sửa mọi hạng mục; **Nhập liệu** chỉ ghi chú và sửa hạng mục mình phụ trách.

## 8. Sao lưu
1. **QLTK → Cài sao lưu tự động mỗi thứ Hai**: khoảng 7h sáng thứ Hai, script xuất một file Excel vào cùng thư mục Drive với Sheet (hoặc thư mục ghi ở `CauHinh → thu_muc_sao_luu`).
2. Apps Script không ghi thẳng được vào ổ Z. Muốn có bản trên Z: cài Google Drive cho máy tính rồi đồng bộ thư mục sao lưu, hoặc chép tay file Excel mỗi tuần.

## Khi sửa code về sau
- Sửa `Code.gs` hoặc `index` → Lưu → **Triển khai → Quản lý triển khai** → bút chì → Phiên bản: **Phiên bản mới** → Triển khai. Làm vậy thì link `/exec` giữ nguyên.
- Không đổi tên cột ở dòng 1 các tab, không chèn cột vào giữa bảng. Cột mới thì thêm ở cuối.

## Dữ liệu cần rà sau khi chuyển
- Tab CongViec, cột `can_ra`: 94 dòng có ngày gõ dạng chữ hoặc tình trạng ghi không chuẩn. Giá trị gốc giữ ở cột `tinh_trang_goc`.
- Tab BaoCaoTuan, cột `tinh_trang` = `Chờ rà`: 144 ghi chú. Gồm 115 ghi chú nằm ở cột tháng 3, 4 (chưa rõ tuần) và 29 ghi chú nằm ở tuần tương lai (phần lớn chép thừa sang tháng 12). Rà trên app ở màn **Tra cứu → Chờ rà**, hoặc ở màn **Duyệt tuần**.
- Ghi chú chuyển từ Excel được gán người nhập = người phụ trách nhóm ở cột A của file gốc.
- Loại vướng mắc (Thanh toán, Thẩm tra/thẩm định, CĐT chưa chốt) được gắn tự động theo từ khóa, nên có thể có chỗ gắn sai.

## Cập nhật bản Báo cáo nhanh (02/11/2026)
1. Dán đè `Code.gs` và `index` bản mới → Lưu → Triển khai → Quản lý triển khai → Phiên bản mới.
2. Tab **CauHinh**:
   - Sửa `phut_phien` thành `10080` để đăng nhập nhớ 7 ngày.
   - Thêm 1 dòng `link_app` (cột A), cột B dán link app. Link này được gửi kèm trong email nhắc. Để trống thì email dùng link `/exec`.
3. Tab **NhanSu**: điền cột `email` cho từng người. Ai không có email thì không nhận được nhắc.
4. Tải lại Sheet → menu **QLTK → Cài nhắc tự động (15h thứ Sáu, 9h thứ Hai)**. Lần đầu Google sẽ hỏi thêm quyền gửi email: bấm Cho phép.
5. Muốn thử ngay: **QLTK → Gửi email nhắc báo cáo ngay**.

Cách nhắc hoạt động:
- 15h thứ Sáu: gửi cho từng người danh sách hạng mục của họ chưa báo cáo trong tuần.
- 9h thứ Hai: gửi lại cho người còn thiếu của tuần trước, đồng thời gửi bản tổng hợp cho Trưởng.

Một hạng mục được tính là "cần báo cáo" khi thuộc dự án Đang chạy và có ít nhất một trong các điều kiện sau:
- trạng thái đang làm hoặc đang chờ;
- tiến độ dở dang;
- có ghi chú trong 4 tuần gần đây;
- deadline CĐT rơi vào 4 tuần tới.

Người có vai trò Nhập liệu khi đăng nhập sẽ vào thẳng màn Báo cáo nhanh.

## Cập nhật: màu theo luồng, phân quyền, tạo dự án mới (02/11/2026)
1. Dán đè `Code.gs` và `index` bằng bản mới.
2. Thêm tệp script mới: bấm **+ → Tập lệnh**, đặt tên `MauHangMuc`, dán nội dung file `MauHangMuc.gs`.
3. Lưu → Triển khai → Quản lý triển khai → **Phiên bản mới**.
4. Tải lại Sheet → menu **QLTK → Tạo tab MauHangMuc**. Tab này chứa 3 bộ hạng mục mẫu (Cao tầng, NOXH, Thấp tầng) lấy từ sheet TỔNG HỢP của 3 file gốc, kèm người phụ trách mặc định cho từng nhóm. Muốn đổi mẫu thì sửa thẳng trên tab này.

Phân quyền ghi:
- Mỗi người chỉ ghi chú và sửa được hạng mục mình phụ trách.
- Trưởng và người duyệt sửa được tất cả.

Dự án mới (chỉ Trưởng và người duyệt):
- Vào **Dự án → + Dự án mới**: nhập tên ngắn, chọn luồng và mẫu, điền mốc CĐT nếu đã có, bỏ chọn những nhóm không làm, chọn người phụ trách từng nhóm → **Tạo dự án**.
- Trong trang dự án, mỗi nhóm có:
  - **Giao cả nhóm cho…**: đổi người phụ trách cho toàn bộ hạng mục trong nhóm.
  - **+ Thêm hạng mục**: thêm hạng mục ngoài mẫu.

## Cập nhật: sửa, bớt hạng mục; ghi chú theo ngày; nút hồ sơ liên quan (05/11/2026)
1. Dán đè `Code.gs` → Lưu → **Triển khai → Quản lý triển khai → Phiên bản mới**.
2. Tải file `index.html` mới lên repo `qlda-stellar`.
3. Tab **LienKet** trong Sheet sẽ tự tạo khi có người thêm nút hồ sơ đầu tiên. Không cần tạo tay.

Quyền:
- **Sửa tên hoặc bớt hạng mục:** Trưởng và người duyệt. Bớt nghĩa là ẩn khỏi app (cột `cap` đổi thành `xoa`), lịch sử ghi chú vẫn giữ. Muốn khôi phục thì sửa cột `cap` về `hangmuc`.
- **Ghi chú:** khi thêm, chọn "Thông tin của ngày". Người nhập hoặc người duyệt được sửa ngày, sửa nội dung hoặc xóa ghi chú.
- **Nút hồ sơ:** người tham gia dự án tạo được. Người tạo, người duyệt và Trưởng được sửa link hoặc xóa. Chỉ người duyệt và Trưởng được sắp xếp thứ tự.
