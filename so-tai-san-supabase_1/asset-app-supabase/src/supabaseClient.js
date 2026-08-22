import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // eslint-disable-next-line no-console
  console.error(
    "Thiếu VITE_SUPABASE_URL hoặc VITE_SUPABASE_ANON_KEY. Tạo file .env ở gốc dự án — xem HUONG_DAN.md."
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey, {
  auth: {
    persistSession: true, // giữ đăng nhập qua các lần mở lại — an toàn vì đây là web app độc lập, không phải artifact
    autoRefreshToken: true,
  },
});
