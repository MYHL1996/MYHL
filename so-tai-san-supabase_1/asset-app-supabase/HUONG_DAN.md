# Hướng dẫn triển khai Sổ Tài Sản với backend Supabase

Làm theo đúng thứ tự bên dưới. Mỗi bước đều ghi rõ làm ở đâu (Supabase Dashboard / máy tính / trình duyệt).

---

## Phần 1 — Tạo backend trên Supabase (10 phút)

1. Vào **supabase.com** → **Start your project** → đăng nhập bằng GitHub hoặc email.
2. Bấm **New project**:
   - Đặt tên project (VD: `so-tai-san`)
   - Đặt **Database Password** — lưu lại mật khẩu này ở nơi an toàn.
   - Chọn Region gần Việt Nam nhất (Singapore).
   - Bấm **Create new project**, đợi khoảng 1–2 phút để khởi tạo.
3. Vào menu bên trái → **SQL Editor** → **New query**.
4. Mở file `supabase-schema.sql` (đi kèm trong gói này), copy toàn bộ nội dung, dán vào ô query, bấm **Run**.
   - Nếu chạy thành công sẽ thấy "Success. No rows returned".
5. Vào menu **Authentication** → **Users** → **Add user** → **Create new user**:
   - Điền email + mật khẩu cho tài khoản quản trị đầu tiên của bạn.
   - Bấm "Auto Confirm User" nếu có, để khỏi cần xác nhận email.
6. Vào menu **Table Editor** → chọn bảng `profiles` → tìm dòng vừa được tự tạo (khớp email bạn vừa nhập) → sửa cột `role` từ `user` thành `admin`, và cột `name` thành tên hiển thị của bạn → **Save**.
7. Vào menu **Settings** (biểu tượng bánh răng) → **API**:
   - Copy **Project URL**
   - Copy khoá **anon public**
   - Hai giá trị này sẽ dùng ở Phần 2.

Vậy là backend đã sẵn sàng — có database thật, xác thực thật, tự sao lưu (Supabase tự backup theo lịch của họ), độc lập hoàn toàn với Claude.

---

## Phần 2 — Chạy ứng dụng trên máy bạn (5 phút)

Cần cài sẵn **Node.js** (bản 18 trở lên) — tải tại nodejs.org nếu chưa có.

1. Mở Terminal (macOS) hoặc PowerShell/CMD (Windows), vào thư mục dự án này:
   ```
   cd duong-dan-toi-thu-muc-nay
   ```
2. Cài các thư viện cần thiết:
   ```
   npm install
   ```
3. Tạo file cấu hình kết nối:
   - Đổi tên file `.env.example` thành `.env`
   - Mở file `.env`, dán **Project URL** và **anon public key** đã copy ở Phần 1 bước 7 vào đúng chỗ.
4. Chạy thử ứng dụng:
   ```
   npm run dev
   ```
5. Mở trình duyệt vào địa chỉ hiện ra (thường là `http://localhost:5173`), đăng nhập bằng email/mật khẩu quản trị đã tạo ở Phần 1 bước 5.

Nếu đăng nhập được và thấy giao diện Sổ Tài Sản — mọi thứ đã hoạt động đúng.

---

## Phần 3 — Đưa lên internet để cả công ty dùng (10 phút)

Cách nhanh nhất: dùng **Vercel** (miễn phí cho quy mô nhỏ).

1. Đưa code lên GitHub:
   - Tạo repository mới trên github.com, đẩy toàn bộ thư mục dự án này lên (không đẩy file `.env` — đã bị `.gitignore` chặn sẵn).
2. Vào **vercel.com** → đăng nhập bằng GitHub → **Add New Project** → chọn repository vừa tạo.
3. Ở bước cấu hình, mở **Environment Variables**, thêm 2 biến giống hệt file `.env`:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
4. Bấm **Deploy**, đợi khoảng 1 phút.
5. Vercel sẽ cấp cho bạn một địa chỉ dạng `https://so-tai-san.vercel.app` — gửi link này cho cả team dùng. Có thể gắn tên miền riêng của công ty ở mục **Domains** sau này.

---

## Sau khi đã dùng thật

- **Thêm người dùng mới**: họ tự vào link app → tab "Tạo tài khoản" → tự đăng ký. Bạn (admin) vào **Cài đặt → Tài khoản đăng nhập** để nâng quyền "Toàn quyền" nếu cần.
- **Quên mật khẩu**: admin bấm biểu tượng khoá 🔒 cạnh tên người đó trong **Cài đặt → Tài khoản đăng nhập** để gửi email đặt lại mật khẩu.
- **Xoá hẳn một tài khoản**: vào Supabase Dashboard → Authentication → Users → xoá (chưa hỗ trợ xoá ngay trong app vì lý do bảo mật — cần quyền admin API).
- **Sao lưu**: app đã có tính năng sao lưu/khôi phục riêng trong **Cài đặt**. Ngoài ra Supabase cũng tự sao lưu hạ tầng của họ theo gói bạn dùng.
- **Chi phí**: gói miễn phí của Supabase đủ dùng cho công ty vừa và nhỏ (500MB database, 50,000 lượt xác thực/tháng). Vercel miễn phí cho 1 dự án cá nhân/nhỏ.

---

## Xử lý sự cố thường gặp

**Mở app lên chỉ thấy "Đang tải sổ tài sản..." mãi không xong, hoặc màn hình trắng:**
Thường do file `.env` chưa đúng, hoặc bạn tạo/sửa file `.env` SAU KHI đã chạy `npm run dev`. Vite chỉ đọc file `.env` lúc khởi động server. Cách sửa:
1. Vào cửa sổ PowerShell/CMD đang chạy `npm run dev`, bấm `Ctrl + C` để dừng.
2. Kiểm tra lại file `.env` — đảm bảo đúng 2 dòng `VITE_SUPABASE_URL=...` và `VITE_SUPABASE_ANON_KEY=...`, không thừa dấu ngoặc kép, không thiếu dấu `=`.
3. Chạy lại `npm run dev`, mở lại trình duyệt (bấm F5 để tải mới hoàn toàn).

**Đăng nhập báo "Sai email hoặc mật khẩu" dù đã tạo đúng ở Supabase Dashboard:**
Kiểm tra tài khoản đó đã được xác nhận email chưa — vào Supabase Dashboard → Authentication → Users, xem cột "Email Confirmed". Nếu chưa, bấm vào tài khoản đó → "Confirm email" thủ công (vì app chạy trên `localhost` không tự gửi được email xác nhận đúng cách).

**Đăng nhập được nhưng vẫn báo "Không tải được dữ liệu":**
Đọc kỹ dòng chữ đỏ hiện ra trong app — nó sẽ nói rõ nguyên nhân (ví dụ: chưa chạy `supabase-schema.sql`, hoặc sai khoá API). Bấm nút **Thử lại**. Nếu vẫn lỗi, vào Supabase Dashboard → SQL Editor, chạy lại toàn bộ file `supabase-schema.sql` để chắc chắn đã tạo đủ bảng.

Nếu vẫn không tự sửa được, chụp lại nguyên màn hình lỗi (kèm dòng chữ đỏ nếu có) gửi cho mình.

