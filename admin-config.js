// ====================================================================
//  KONFIGURASI PANEL ADMIN — BPBD Provinsi Kalimantan Selatan
// ====================================================================
// PENTING: sejak backend (Code.gs) punya action "adminLogin", username
// & password admin TIDAK lagi disimpan di file ini. Kredensial sekarang
// dikelola langsung di sheet "AdminUsers" pada spreadsheet, dan
// diverifikasi oleh server — bukan oleh browser. File ini sekarang
// hanya untuk pengaturan sesi login.
//
// CARA MENAMBAH / MENGGANTI ADMIN:
//  1. Buka spreadsheet -> sheet "AdminUsers".
//  2. Untuk admin baru: tambah baris, isi kolom "Username".
//  3. Untuk kolom "PasswordHash": buka admin.html, tekan F12 -> Console,
//     lalu jalankan (ganti PASSWORD_BARU_ANDA dengan password yang mau dipakai):
//
//     crypto.subtle.digest('SHA-256', new TextEncoder().encode('PASSWORD_BARU_ANDA'))
//       .then(buf => console.log([...new Uint8Array(buf)]
//       .map(b => b.toString(16).padStart(2,'0')).join('')));
//
//     Salin hasilnya ke kolom "PasswordHash", lalu simpan sheet-nya.
//     Tidak perlu deploy ulang Web App atau mengubah file apa pun.
//  4. Untuk mencabut akses seorang admin, hapus saja barisnya di sheet ini.
//
// Akun bawaan (dibuat otomatis pertama kali): username "admin",
// password "GantiSekarang123" — GANTI SEGERA lewat langkah di atas.
// ====================================================================

// Lama sesi login bertahan (dalam jam) sebelum admin harus login ulang.
// Sesi otomatis berakhir juga saat tab/browser ditutup.
const ADMIN_SESSION_HOURS = 8;
