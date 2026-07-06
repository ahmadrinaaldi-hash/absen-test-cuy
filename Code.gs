/**
 * ABSENSI BPBD PROVINSI KALIMANTAN SELATAN
 * Backend Google Apps Script — menghubungkan Web App ke Google Sheets + Drive
 *
 * CARA PAKAI: lihat README.md di root project untuk langkah deploy lengkap.
 *
 * Struktur Spreadsheet (dibuat otomatis oleh setupSheets() jika belum ada):
 *  - Sheet "Pegawai"     : ID | Nama | Jabatan | Bidang | FotoURL
 *  - Sheet "LogAbsensi"  : Timestamp | ID | Nama | Status | Keterangan | Ketepatan | FotoAbsen | FotoBukti | LokasiLat | LokasiLng | LokasiAkurasi
 *                          (kolom "Ketepatan" berisi "Tepat Waktu"/"Telat", hanya diisi untuk Status MASUK)
 *  - Sheet "OpsiIzin"    : No | Teks
 *  - Sheet "AdminUsers"  : Username | PasswordHash (untuk login Panel Admin — lihat adminLogin())
 *  - Sheet "Pengaturan"  : Key | Value (dipakai fitur pembatasan radius lokasi absen / geofencing,
 *                          diatur lewat tab "📍 Lokasi Absen" — lihat getPengaturan()/simpanPengaturan() —
 *                          DAN batas jam masuk Tepat Waktu/Telat, diatur lewat tab "⏰ Jam Masuk",
 *                          lihat getPengaturanWaktu()/simpanPengaturanWaktu())
 *
 * Penyimpanan foto di Google Drive:
 *  - Foto selfie absen (MASUK/KELUAR/IZIN) -> folder "Foto Absensi BPBD" (rata, semua karyawan)
 *  - Foto bukti izin (opsional, misal surat sakit) -> folder "Foto Bukti Izin BPBD" / {Nama Karyawan} / ...
 *    supaya gampang dicari per orang.
 *
 * Rekap bulanan (halaman rekap.html):
 *  - Tidak pakai sheet baru — dihitung langsung dari sheet "LogAbsensi" tiap kali diminta
 *    lewat action "getRekapBulanan" (lihat fungsi getRekapBulanan() di bawah).
 *  - "Alpha" = hari kerja (Senin-Jumat) dalam bulan tsb dikurangi hari yang sudah tercatat
 *    MASUK atau IZIN. Ganti array HARI_KERJA di getRekapBulanan() kalau Sabtu juga hari kerja.
 */

const SHEET_PEGAWAI = 'Pegawai';
const SHEET_LOG = 'LogAbsensi';
const SHEET_IZIN = 'OpsiIzin';
const SHEET_ADMIN = 'AdminUsers';
const SHEET_SETTING = 'Pengaturan';
const DRIVE_FOLDER_NAME = 'Foto Absensi BPBD'; // folder Drive utk simpan foto absen

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function setupSheets() {
  const ss = getSS();
  // Kolom "Piket" (Tidak Piket / Piket Pagi / Piket Malam) menentukan batas jam
  // Tepat Waktu/Telat mana yang berlaku untuk pegawai tsb — lihat getRekapBulanan(),
  // catatAbsen(), dan tab "⏰ JAM MASUK" di Panel Admin.
  const PEGAWAI_HEADERS = ['ID', 'Nama', 'Jabatan', 'Bidang', 'FotoURL', 'Piket'];
  if (!ss.getSheetByName(SHEET_PEGAWAI)) {
    const sh = ss.insertSheet(SHEET_PEGAWAI);
    sh.appendRow(PEGAWAI_HEADERS);
  } else {
    // migrasi: sheet Pegawai lama mungkin belum punya kolom "Piket" — tambahkan
    // otomatis di ujung kanan tanpa mengubah/menghapus data yang sudah ada.
    // Pegawai lama otomatis dianggap "Tidak Piket" (lihat normalisasiPiket()).
    const sh = ss.getSheetByName(SHEET_PEGAWAI);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    PEGAWAI_HEADERS.forEach(h => {
      if (headers.indexOf(h) === -1) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      }
    });
  }
  // Header lengkap yang seharusnya ada di sheet LogAbsensi. Kolom Lokasi*
  // ditambahkan untuk menyimpan titik GPS saat foto absen diambil. Kolom
  // "Ketepatan" ditambahkan untuk menyimpan status "Tepat Waktu" / "Telat"
  // (hanya diisi untuk status MASUK — lihat hitungKetepatanWaktu()).
  const LOG_HEADERS = ['Timestamp', 'ID', 'Nama', 'Status', 'Keterangan', 'Ketepatan', 'FotoAbsen', 'FotoBukti', 'LokasiLat', 'LokasiLng', 'LokasiAkurasi'];
  if (!ss.getSheetByName(SHEET_LOG)) {
    const sh = ss.insertSheet(SHEET_LOG);
    sh.appendRow(LOG_HEADERS);
  } else {
    // migrasi: sheet LogAbsensi lama mungkin belum punya sebagian kolom di atas
    // (mis. FotoBukti atau Lokasi*) — tambahkan otomatis di ujung kanan,
    // tanpa mengubah/menghapus data yang sudah ada.
    const sh = ss.getSheetByName(SHEET_LOG);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    LOG_HEADERS.forEach(h => {
      if (headers.indexOf(h) === -1) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      }
    });
  }
  if (!ss.getSheetByName(SHEET_IZIN)) {
    const sh = ss.insertSheet(SHEET_IZIN);
    sh.appendRow(['No', 'Teks']);
    sh.appendRow([1, 'Izin Sakit']);
    sh.appendRow([2, 'Izin Pribadi']);
    sh.appendRow([3, 'Izin Lapangan']);
  }
  if (!ss.getSheetByName(SHEET_ADMIN)) {
    const sh = ss.insertSheet(SHEET_ADMIN);
    sh.appendRow(['Username', 'PasswordHash']);
    // Akun bawaan: username "admin", password "GantiSekarang123".
    // GANTI SEGERA lewat sheet ini — lihat penjelasan cara pakai adminLogin di bawah.
    sh.appendRow(['admin', sha256Hex('GantiSekarang123')]);
  }
  // Sheet "Pengaturan" (key-value) — dipakai untuk fitur pembatasan radius lokasi absen
  // DAN untuk batas jam masuk (Tepat Waktu / Telat) — kini terpisah 3: pegawai "Tidak
  // Piket" (reguler), "Piket Pagi", dan "Piket Malam". Diatur lewat tab "📍 Lokasi Absen"
  // dan tab "⏰ Jam Masuk" di Panel Admin, tidak perlu diedit manual di sini.
  if (!ss.getSheetByName(SHEET_SETTING)) {
    const sh = ss.insertSheet(SHEET_SETTING);
    sh.appendRow(['Key', 'Value']);
    sh.appendRow(['GeofenceAktif', false]);
    sh.appendRow(['LokasiLat', '']);
    sh.appendRow(['LokasiLng', '']);
    sh.appendRow(['RadiusMeter', 100]);
    sh.appendRow(['JamBatasTelat', '08:00:01']);
    sh.appendRow(['JamBatasTelatPiketPagi', '08:00:01']);
    sh.appendRow(['JamBatasTelatPiketMalam', '19:00:01']);
  } else {
    // migrasi: tambahkan key yang belum ada kalau sheet Pengaturan sudah ada
    // sebelumnya (dibuat sebelum fitur/kolom ini ada).
    const sh = ss.getSheetByName(SHEET_SETTING);
    const keys = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues().map(r => r[0]);
    if (keys.indexOf('JamBatasTelat') === -1) sh.appendRow(['JamBatasTelat', '08:00:01']);
    if (keys.indexOf('JamBatasTelatPiketPagi') === -1) sh.appendRow(['JamBatasTelatPiketPagi', '08:00:01']);
    if (keys.indexOf('JamBatasTelatPiketMalam') === -1) sh.appendRow(['JamBatasTelatPiketMalam', '19:00:01']);
  }
}

// Nilai "Piket" yang sah untuk pegawai. Nilai lain (kosong / typo / data lama
// sebelum kolom ini ada) otomatis dianggap "Tidak Piket" supaya tidak error.
const PIKET_VALID = ['Tidak Piket', 'Piket Pagi', 'Piket Malam'];
function normalisasiPiket(v) {
  const s = String(v || 'Tidak Piket').trim();
  return PIKET_VALID.indexOf(s) !== -1 ? s : 'Tidak Piket';
}


function sheetToObjects(sheetName) {
  const sh = getSS().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return values.map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => (obj[h] = row[j]));
    return obj;
  });
}

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Menambah baris baru berdasarkan nama header sheet (bukan urutan tetap),
// supaya aman walau urutan/kolom di sheet berubah atau ada kolom tambahan.
function appendRowByHeaders(sheet, dataObj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => (dataObj[h] !== undefined ? dataObj[h] : ''));
  sheet.appendRow(row);
}

// Hash SHA-256 dalam bentuk hex — hasilnya SAMA PERSIS dengan hasil
// crypto.subtle.digest('SHA-256', ...) di browser, jadi hash yang dibuat
// lewat Console browser (lihat penjelasan di adminLogin()) bisa langsung
// ditempel ke sheet "AdminUsers" di sini.
function sha256Hex(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

// ---------- PENGATURAN (geofencing radius absen) ----------
// Sheet "Pengaturan" berbentuk Key/Value supaya gampang ditambah setting lain di masa depan.
function getPengaturan() {
  const rows = sheetToObjects(SHEET_SETTING); // [{Key:'GeofenceAktif', Value:false}, ...]
  const map = {};
  rows.forEach(r => { map[r.Key] = r.Value; });
  return {
    geofenceAktif: map.GeofenceAktif === true || String(map.GeofenceAktif).toUpperCase() === 'TRUE',
    lokasiLat: map.LokasiLat !== '' && map.LokasiLat !== undefined ? Number(map.LokasiLat) : null,
    lokasiLng: map.LokasiLng !== '' && map.LokasiLng !== undefined ? Number(map.LokasiLng) : null,
    radiusMeter: map.RadiusMeter !== '' && map.RadiusMeter !== undefined ? Number(map.RadiusMeter) : 100,
    jamBatasTelat: map.JamBatasTelat !== '' && map.JamBatasTelat !== undefined ? String(map.JamBatasTelat) : '08:00:01'
  };
}

// ---------- PENGATURAN (batas jam masuk — Tepat Waktu / Telat, per Jenis Piket) ----------
// Dipisah dari simpanPengaturan() (geofencing) karena validasinya berbeda,
// tapi tetap disimpan di sheet "Pengaturan" yang sama. Ada 3 batas jam terpisah
// supaya pegawai "Piket Pagi" / "Piket Malam" tidak salah dianggap Telat memakai
// jam pegawai reguler — lihat pilihJamBatas() dan catatAbsen().
function getPengaturanWaktu() {
  const p = getPengaturan();
  const rows = sheetToObjects(SHEET_SETTING);
  const map = {};
  rows.forEach(r => { map[r.Key] = r.Value; });
  return {
    jamBatasTelat: p.jamBatasTelat,
    jamBatasTelatPiketPagi: map.JamBatasTelatPiketPagi !== '' && map.JamBatasTelatPiketPagi !== undefined ? String(map.JamBatasTelatPiketPagi) : '08:00:01',
    jamBatasTelatPiketMalam: map.JamBatasTelatPiketMalam !== '' && map.JamBatasTelatPiketMalam !== undefined ? String(map.JamBatasTelatPiketMalam) : '19:00:01'
  };
}

// Memilih batas jam yang sesuai berdasarkan Jenis Piket pegawai.
function pilihJamBatas(jenisPiket, waktuSetting) {
  if (jenisPiket === 'Piket Pagi') return waktuSetting.jamBatasTelatPiketPagi;
  if (jenisPiket === 'Piket Malam') return waktuSetting.jamBatasTelatPiketMalam;
  return waktuSetting.jamBatasTelat; // "Tidak Piket" / default
}

function simpanPengaturanWaktu(body) {
  const polaJam = /^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$/;
  const jam = String(body.jamBatasTelat || '').trim();
  const jamPagi = String(body.jamBatasTelatPiketPagi || '').trim();
  const jamMalam = String(body.jamBatasTelatPiketMalam || '').trim();
  if (!polaJam.test(jam) || !polaJam.test(jamPagi) || !polaJam.test(jamMalam)) {
    return { ok: false, error: 'Format jam tidak valid pada salah satu kolom. Gunakan format JJ:MM:DD, contoh 08:00:01' };
  }
  const sh = getSS().getSheetByName(SHEET_SETTING);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const setValue = (key, value) => {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === key) { sh.getRange(i + 2, 2).setValue(value); return; }
    }
    sh.appendRow([key, value]); // jaga-jaga kalau key belum ada
  };
  setValue('JamBatasTelat', jam);
  setValue('JamBatasTelatPiketPagi', jamPagi);
  setValue('JamBatasTelatPiketMalam', jamMalam);
  return { ok: true };
}

// Menentukan "Tepat Waktu" atau "Telat" berdasarkan jam saat absen MASUK
// dibandingkan dengan batas yang diatur di Panel Admin (default 08:00:01).
// Aturan: jika jam absen LEBIH BESAR dari batas -> "Telat".
//         jika jam absen SAMA DENGAN atau SEBELUM batas -> "Tepat Waktu".
function hitungKetepatanWaktu(tanggal, jamBatasStr) {
  const cocok = /^(\d{1,2}):(\d{1,2}):(\d{1,2})$/.exec(String(jamBatasStr || '08:00:01').trim());
  const batasDetik = cocok
    ? (Number(cocok[1]) * 3600 + Number(cocok[2]) * 60 + Number(cocok[3]))
    : (8 * 3600 + 0 * 60 + 1); // fallback 08:00:01 kalau format tersimpan rusak
  const detikSekarang = tanggal.getHours() * 3600 + tanggal.getMinutes() * 60 + tanggal.getSeconds();
  return detikSekarang > batasDetik ? 'Telat' : 'Tepat Waktu';
}

function simpanPengaturan(body) {
  const aktif = !!body.aktif;
  if (aktif) {
    // Kalau geofencing mau diaktifkan, titik lokasi & radius wajib valid.
    const lat = Number(body.lat), lng = Number(body.lng), radius = Number(body.radius);
    if (!isFinite(lat) || !isFinite(lng) || lat === 0 && lng === 0) {
      return { ok: false, error: 'Titik lokasi kantor belum diisi / tidak valid' };
    }
    if (!isFinite(radius) || radius <= 0) {
      return { ok: false, error: 'Radius harus lebih besar dari 0 meter' };
    }
  }
  const sh = getSS().getSheetByName(SHEET_SETTING);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const setValue = (key, value) => {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === key) { sh.getRange(i + 2, 2).setValue(value); return; }
    }
    sh.appendRow([key, value]); // jaga-jaga kalau key belum ada
  };
  setValue('GeofenceAktif', aktif);
  setValue('LokasiLat', body.lat !== undefined ? Number(body.lat) : '');
  setValue('LokasiLng', body.lng !== undefined ? Number(body.lng) : '');
  setValue('RadiusMeter', body.radius !== undefined ? Number(body.radius) : 100);
  return { ok: true };
}

// Jarak antar 2 koordinat GPS dalam meter (rumus Haversine)
function hitungJarakMeter(lat1, lng1, lat2, lng2) {
  const R = 6371000; // radius bumi dalam meter
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// ---------- REKAP BULANAN (per pegawai: Masuk / Tepat Waktu / Telat / Izin / Alpha) ----------
// Dihitung langsung dari sheet "LogAbsensi" setiap kali diminta (tidak perlu sheet baru),
// supaya datanya selalu sesuai dengan log absensi yang sebenarnya.
//
// Definisi "Alpha": hari kerja dalam bulan tsb DIKURANGI hari yang sudah ada
// catatan MASUK atau IZIN. Kalau bulan yang diminta adalah bulan BERJALAN,
// hari kerja hanya dihitung sampai HARI INI (bukan sampai akhir bulan),
// supaya sisa hari yang belum terjadi tidak salah dianggap "Alpha". Kalau
// bulan yang diminta belum mulai (di masa depan), semua nilai dikembalikan 0.
//
// Hari kerja yang dihitung: Senin s/d Jumat (getDay(): 1=Senin ... 5=Jumat).
// Kalau kantor Anda juga masuk hari Sabtu, tambahkan angka 6 ke array HARI_KERJA di bawah.
function getRekapBulanan(bulanParam, tahunParam) {
  const now = new Date();
  const bulan = bulanParam ? Number(bulanParam) : (now.getMonth() + 1); // 1-12
  const tahun = tahunParam ? Number(tahunParam) : now.getFullYear();

  const HARI_KERJA = [1, 2, 3, 4, 5]; // Senin-Jumat. Tambahkan 6 kalau Sabtu juga hari kerja.

  const akanDatang = (tahun > now.getFullYear()) || (tahun === now.getFullYear() && bulan > now.getMonth() + 1);
  const bulanBerjalan = (tahun === now.getFullYear() && bulan === now.getMonth() + 1);
  const tanggalAkhir = bulanBerjalan ? now.getDate() : new Date(tahun, bulan, 0).getDate();

  let totalHariKerja = 0;
  if (!akanDatang) {
    for (let d = 1; d <= tanggalAkhir; d++) {
      if (HARI_KERJA.indexOf(new Date(tahun, bulan - 1, d).getDay()) !== -1) totalHariKerja++;
    }
  }

  const pegawaiList = sheetToObjects(SHEET_PEGAWAI);
  const logList = sheetToObjects(SHEET_LOG);

  const hasil = pegawaiList.map(p => {
    const idPegawai = String(p.ID).trim();
    const hariMasuk = new Set();
    const hariTelat = new Set();
    const hariIzin = new Set();
    // Union hari yang sudah "tertangani" (MASUK ATAU IZIN) — dipakai KHUSUS
    // untuk menghitung Alpha. Sengaja dipisah dari hariMasuk/hariIzin: kalau
    // 1 hari yang sama punya KEDUA log (mis. sudah MASUK lalu menyusul
    // mengajukan IZIN di hari yang sama), memakai (totalMasuk - totalIzin)
    // akan mengurangi hari kerja 2x untuk 1 hari yang sama sehingga Alpha
    // jadi lebih kecil dari yang seharusnya. Union memastikan 1 hari hanya
    // dihitung 1x, sesuai definisi di komentar atas fungsi ini.
    const hariTertangani = new Set();

    logList.forEach(log => {
      if (String(log.ID).trim() !== idPegawai) return;
      const ts = log.Timestamp instanceof Date ? log.Timestamp : new Date(log.Timestamp);
      if (!ts || isNaN(ts.getTime())) return;
      if (ts.getFullYear() !== tahun || (ts.getMonth() + 1) !== bulan) return;
      const tglKey = ts.getFullYear() + '-' + (ts.getMonth() + 1) + '-' + ts.getDate();
      if (log.Status === 'MASUK') {
        hariMasuk.add(tglKey);
        hariTertangani.add(tglKey);
        if (log.Ketepatan === 'Telat') hariTelat.add(tglKey);
      } else if (log.Status === 'IZIN') {
        hariIzin.add(tglKey);
        hariTertangani.add(tglKey);
      }
    });

    const totalMasuk = hariMasuk.size;
    const totalTelat = hariTelat.size;
    const totalIzin = hariIzin.size;

    return {
      ID: p.ID,
      Nama: p.Nama,
      Jabatan: p.Jabatan || '',
      Bidang: p.Bidang || '',
      Piket: normalisasiPiket(p.Piket),
      TotalMasuk: totalMasuk,
      TotalTepatWaktu: totalMasuk - totalTelat,
      TotalTelat: totalTelat,
      TotalIzin: totalIzin,
      TotalAlpha: Math.max(0, totalHariKerja - hariTertangani.size)
    };
  });

  return { bulan: bulan, tahun: tahun, totalHariKerja: totalHariKerja, data: hasil };
}

function doGet(e) {
  setupSheets();
  const action = e.parameter.action;
  try {
    if (action === 'getPegawai') return jsonOut({ ok: true, data: sheetToObjects(SHEET_PEGAWAI) });
    if (action === 'getLog') return jsonOut({ ok: true, data: sheetToObjects(SHEET_LOG).reverse() });
    if (action === 'getOpsiIzin') return jsonOut({ ok: true, data: sheetToObjects(SHEET_IZIN) });
    if (action === 'getPengaturanLokasi') return jsonOut({ ok: true, data: getPengaturan() });
    if (action === 'getPengaturanWaktu') return jsonOut({ ok: true, data: getPengaturanWaktu() });
    if (action === 'getRekapBulanan') {
      const bulan = e.parameter.bulan ? Number(e.parameter.bulan) : '';
      const tahun = e.parameter.tahun ? Number(e.parameter.tahun) : '';
      return jsonOut({ ok: true, data: getRekapBulanan(bulan, tahun) });
    }
    if (action === 'cariPegawai') {
      const id = String(e.parameter.id || '').trim();
      const list = sheetToObjects(SHEET_PEGAWAI);
      const found = list.find(p => String(p.ID).trim() === id);
      return jsonOut({ ok: true, data: found || null });
    }
    return jsonOut({ ok: false, error: 'Aksi tidak dikenali' });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function doPost(e) {
  setupSheets();
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'Body tidak valid' });
  }
  const action = body.action;
  try {
    switch (action) {
      case 'addPegawai': return jsonOut(addPegawai(body));
      case 'updatePegawai': return jsonOut(updatePegawai(body));
      case 'deletePegawai': return jsonOut(deleteRow(SHEET_PEGAWAI, body.row));
      case 'updateLog': return jsonOut(updateLog(body));
      case 'addOpsiIzin': return jsonOut(addOpsiIzin(body));
      case 'updateOpsiIzin': return jsonOut(updateOpsiIzin(body));
      case 'deleteOpsiIzin': return jsonOut(deleteRow(SHEET_IZIN, body.row));
      case 'absen': return jsonOut(catatAbsen(body));
      case 'adminLogin': return jsonOut(adminLogin(body));
      case 'simpanPengaturanLokasi': return jsonOut(simpanPengaturan(body));
      case 'simpanPengaturanWaktu': return jsonOut(simpanPengaturanWaktu(body));
      default: return jsonOut({ ok: false, error: 'Aksi tidak dikenali' });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ---------- EDIT LOG ABSENSI (dari popup "Edit Log Absensi Lengkap" di Panel Admin) ----------
/**
 * body: { row, timestamp, status, keterangan }
 *  - row       : nomor baris di sheet LogAbsensi (properti "_row" dari getLog)
 *  - timestamp : string tanggal+jam (ISO 8601, dikirim admin.html via toISOString())
 *  - status    : 'MASUK' | 'KELUAR' | 'IZIN'
 *  - keterangan: teks bebas (opsional)
 *
 * "Ketepatan" (Tepat Waktu/Telat) DIHITUNG ULANG di sini memakai jam & Jenis
 * Piket pegawai yang sama seperti saat absen pertama kali dicatat (lihat
 * catatAbsen()) — supaya kalau admin mengoreksi jam MASUK, kolom Ketepatan
 * ikut ter-update secara konsisten, bukan sekadar disalin apa adanya.
 * Untuk status KELUAR/IZIN, "Ketepatan" sengaja dikosongkan (tidak relevan),
 * sama seperti perilaku catatAbsen().
 */
function updateLog(body) {
  const row = Number(body.row);
  if (!row || row < 2) {
    return { ok: false, error: 'Baris data tidak valid' };
  }
  const status = String(body.status || '').trim().toUpperCase();
  if (['MASUK', 'KELUAR', 'IZIN'].indexOf(status) === -1) {
    return { ok: false, error: 'Status absensi tidak valid' };
  }
  const timestamp = body.timestamp ? new Date(body.timestamp) : null;
  if (!timestamp || isNaN(timestamp.getTime())) {
    return { ok: false, error: 'Tanggal/jam absen tidak valid' };
  }

  const sh = getSS().getSheetByName(SHEET_LOG);
  if (row > sh.getLastRow()) {
    return { ok: false, error: 'Baris log tidak ditemukan (mungkin sudah dihapus/berubah, muat ulang halaman)' };
  }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf('ID') + 1;
  const idPegawai = idCol ? String(sh.getRange(row, idCol).getValue()).trim() : '';

  let ketepatan = '';
  if (status === 'MASUK') {
    const pegawai = sheetToObjects(SHEET_PEGAWAI).find(p => String(p.ID).trim() === idPegawai);
    const jenisPiket = normalisasiPiket(pegawai ? pegawai.Piket : 'Tidak Piket');
    const jamBatas = pilihJamBatas(jenisPiket, getPengaturanWaktu());
    ketepatan = hitungKetepatanWaktu(timestamp, jamBatas);
  }

  // Tulis berdasarkan NAMA kolom (bukan nomor tetap) supaya aman walau urutan
  // kolom di sheet LogAbsensi berbeda — sama seperti pola appendRowByHeaders().
  const setCell = (namaKolom, nilai) => {
    const col = headers.indexOf(namaKolom) + 1;
    if (col) sh.getRange(row, col).setValue(nilai);
  };
  setCell('Timestamp', timestamp);
  setCell('Status', status);
  setCell('Keterangan', String(body.keterangan || '').trim());
  setCell('Ketepatan', ketepatan);

  return { ok: true, ketepatan: ketepatan };
}

// ---------- PEGAWAI ----------
function addPegawai(body) {
  const id = String(body.id || '').trim();
  const nama = String(body.nama || '').trim();
  if (!id || !nama) {
    return { ok: false, error: 'ID dan Nama wajib diisi' };
  }
  // Cegah ID ganda — kalau tidak dicek, cariPegawai() di mesin absen akan
  // selalu mengambil data yang PALING ATAS saja dan membingungkan admin.
  const bentrok = sheetToObjects(SHEET_PEGAWAI).find(p => String(p.ID).trim() === id);
  if (bentrok) {
    return { ok: false, error: `ID "${id}" sudah dipakai oleh ${bentrok.Nama}. Gunakan ID lain atau edit data yang sudah ada.` };
  }
  const sh = getSS().getSheetByName(SHEET_PEGAWAI);
  sh.appendRow([id, nama, body.jabatan || '', body.bidang || '', body.fotoUrl || '', normalisasiPiket(body.piket)]);
  return { ok: true };
}

function updatePegawai(body) {
  const id = String(body.id || '').trim();
  const nama = String(body.nama || '').trim();
  if (!body.row) {
    return { ok: false, error: 'Baris data tidak valid' };
  }
  if (!id || !nama) {
    return { ok: false, error: 'ID dan Nama wajib diisi' };
  }
  // Cegah ID ganda dengan pegawai LAIN (baris yang sedang diedit sendiri dikecualikan)
  const bentrok = sheetToObjects(SHEET_PEGAWAI).find(p => p._row !== Number(body.row) && String(p.ID).trim() === id);
  if (bentrok) {
    return { ok: false, error: `ID "${id}" sudah dipakai oleh ${bentrok.Nama}. Gunakan ID lain.` };
  }
  const sh = getSS().getSheetByName(SHEET_PEGAWAI);
  sh.getRange(body.row, 1, 1, 6).setValues([[id, nama, body.jabatan || '', body.bidang || '', body.fotoUrl || '', normalisasiPiket(body.piket)]]);
  return { ok: true };
}

// ---------- OPSI IZIN ----------
function addOpsiIzin(body) {
  const teks = String(body.teks || '').trim();
  if (!teks) return { ok: false, error: 'Teks izin wajib diisi' };
  const sh = getSS().getSheetByName(SHEET_IZIN);
  const data = sheetToObjects(SHEET_IZIN);
  // Nomor berikutnya dihitung dari NILAI "No" TERBESAR yang ada, bukan dari
  // jumlah baris fisik — supaya tidak terjadi nomor ganda kalau sebelumnya
  // ada opsi izin di tengah yang pernah dihapus.
  const nextNo = data.reduce((max, o) => Math.max(max, Number(o.No) || 0), 0) + 1;
  sh.appendRow([nextNo, teks]);
  return { ok: true };
}

function updateOpsiIzin(body) {
  const teks = String(body.teks || '').trim();
  if (!body.row) return { ok: false, error: 'Baris data tidak valid' };
  if (!teks) return { ok: false, error: 'Teks izin wajib diisi' };
  const sh = getSS().getSheetByName(SHEET_IZIN);
  sh.getRange(body.row, 2).setValue(teks);
  return { ok: true };
}

// ---------- GENERIC DELETE ----------
function deleteRow(sheetName, row) {
  if (!row || Number(row) < 2) {
    return { ok: false, error: 'Baris data tidak valid' };
  }
  getSS().getSheetByName(sheetName).deleteRow(Number(row));
  return { ok: true };
}

// ---------- LOGIN ADMIN ----------
/**
 * CARA PAKAI action "adminLogin":
 *
 * 1) Dari halaman admin.html, kirim POST ke API_URL dengan body:
 *      { "action": "adminLogin", "username": "...", "password": "..." }
 *    (password dikirim APA ADANYA / plaintext — aman karena Web App Apps
 *    Script selalu diakses lewat HTTPS, jadi terenkripsi saat pengiriman.
 *    Fungsi inilah yang akan meng-hash lalu mencocokkannya, bukan browser.)
 *
 * 2) Balasannya:
 *      { "ok": true }                          -> login berhasil
 *      { "ok": false, "error": "pesan..." }     -> login gagal
 *
 * 3) Kredensial admin disimpan di sheet "AdminUsers" (dibuat otomatis oleh
 *    setupSheets() dengan 1 akun bawaan: username "admin",
 *    password "GantiSekarang123"). Kolom "PasswordHash" berisi hash
 *    SHA-256 dari password, BUKAN password aslinya.
 *
 * UNTUK MENAMBAH / MENGGANTI ADMIN (tidak perlu ubah kode ini sama sekali):
 *    a. Buka spreadsheet -> sheet "AdminUsers".
 *    b. Untuk admin baru: tambah baris baru, isi kolom "Username".
 *    c. Untuk kolom "PasswordHash", buat hash-nya dulu:
 *       - Buka admin.html di Chrome/Edge/Firefox, tekan F12 -> tab Console.
 *       - Jalankan (ganti PASSWORD_BARU_ANDA dengan password yang diinginkan):
 *           crypto.subtle.digest('SHA-256', new TextEncoder().encode('PASSWORD_BARU_ANDA'))
 *             .then(buf => console.log([...new Uint8Array(buf)]
 *             .map(b => b.toString(16).padStart(2,'0')).join('')));
 *       - Salin hasilnya (deretan huruf/angka) ke kolom "PasswordHash".
 *    d. Simpan sheet-nya. Tidak perlu deploy ulang Web App.
 *    e. Untuk menghapus akses seorang admin, cukup hapus barisnya di sheet ini.
 */
function adminLogin(body) {
  const username = String(body.username || '').trim();
  const password = String(body.password || '');
  if (!username || !password) {
    return { ok: false, error: 'Username dan password wajib diisi' };
  }
  const sh = getSS().getSheetByName(SHEET_ADMIN);
  if (!sh || sh.getLastRow() < 2) {
    return { ok: false, error: 'Belum ada akun admin terdaftar di sheet AdminUsers' };
  }
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const hashInput = sha256Hex(password);
  const cocok = rows.some(r => String(r[0]).trim() === username && String(r[1]).trim() === hashInput);
  if (!cocok) {
    return { ok: false, error: 'Username atau password salah' };
  }
  return { ok: true };
}

// ---------- ABSEN (dari mesin absensi) ----------
function catatAbsen(body) {
  // body: { id, nama, status: 'MASUK'|'KELUAR'|'IZIN', keterangan, fotoBase64, fotoBukti,
  //         lokasiLat, lokasiLng, lokasiAkurasi }

  // Validasi radius (geofencing) — dicek di SERVER, bukan cuma di HP pegawai,
  // supaya tidak bisa diakali dengan GPS palsu / mengubah kode di browser.
  const pengaturan = getPengaturan();
  if (pengaturan.geofenceAktif) {
    const lat = Number(body.lokasiLat), lng = Number(body.lokasiLng);
    if (!body.lokasiLat || !body.lokasiLng || !isFinite(lat) || !isFinite(lng)) {
      return { ok: false, error: 'Lokasi GPS wajib aktif untuk absen. Aktifkan izin lokasi lalu ambil foto ulang.' };
    }
    // Jaga-jaga: kalau geofencing AKTIF tapi titik lokasi kantor belum/tidak
    // valid (mis. terhapus manual di sheet "Pengaturan"), hitungJarakMeter()
    // akan menghasilkan NaN dan perbandingan "jarak > radius" DIAM-DIAM
    // selalu false — akibatnya absen malah lolos tanpa pengecekan radius
    // sama sekali, padahal admin mengira pembatasan lokasi sedang aktif.
    // Lebih aman menolak absen dengan pesan jelas daripada meloloskannya.
    if (pengaturan.lokasiLat === null || pengaturan.lokasiLng === null ||
        !isFinite(pengaturan.lokasiLat) || !isFinite(pengaturan.lokasiLng)) {
      return { ok: false, error: 'Pembatasan radius lokasi sedang aktif tetapi titik lokasi kantor belum diatur. Hubungi admin untuk mengatur ulang di Panel Admin > Lokasi Absen.' };
    }
    const jarak = hitungJarakMeter(lat, lng, pengaturan.lokasiLat, pengaturan.lokasiLng);
    if (jarak > pengaturan.radiusMeter) {
      return {
        ok: false,
        error: `Anda berada di luar radius absen yang diizinkan (jarak ${Math.round(jarak)} m, maksimal ${pengaturan.radiusMeter} m dari lokasi kantor).`
      };
    }
  }

  let fotoUrl = '';
  if (body.fotoBase64) {
    fotoUrl = simpanFotoKeDrive(body.fotoBase64, body.id, body.status);
  }
  let fotoBuktiUrl = '';
  if (body.fotoBukti) {
    fotoBuktiUrl = simpanFotoBuktiKeDrive(body.fotoBukti, body.id, body.nama);
  }
  const waktuAbsen = new Date();
  // "Ketepatan" (Tepat Waktu / Telat) hanya relevan untuk jam MASUK (check-in) —
  // status KELUAR dan IZIN sengaja dikosongkan karena konsep telat/tepat waktu
  // tidak berlaku untuk keduanya. Batas jam yang dipakai disesuaikan dengan
  // Jenis Piket pegawai (Tidak Piket / Piket Pagi / Piket Malam) supaya pegawai
  // piket malam tidak salah dianggap Telat memakai jam pegawai reguler.
  let ketepatan = '';
  if (body.status === 'MASUK') {
    const pegawai = sheetToObjects(SHEET_PEGAWAI).find(p => String(p.ID).trim() === String(body.id || '').trim());
    const jenisPiket = normalisasiPiket(pegawai ? pegawai.Piket : 'Tidak Piket');
    const jamBatas = pilihJamBatas(jenisPiket, getPengaturanWaktu());
    ketepatan = hitungKetepatanWaktu(waktuAbsen, jamBatas);
  }
  const sh = getSS().getSheetByName(SHEET_LOG);
  appendRowByHeaders(sh, {
    Timestamp: waktuAbsen,
    ID: body.id,
    Nama: body.nama,
    Status: body.status,
    Keterangan: body.keterangan || '',
    Ketepatan: ketepatan,
    FotoAbsen: fotoUrl,
    FotoBukti: fotoBuktiUrl,
    LokasiLat: body.lokasiLat || '',
    LokasiLng: body.lokasiLng || '',
    LokasiAkurasi: body.lokasiAkurasi || ''
  });
  return { ok: true, fotoUrl: fotoUrl, fotoBuktiUrl: fotoBuktiUrl, ketepatan: ketepatan };
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function simpanFotoKeDrive(base64Data, id, status) {
  const folder = getOrCreateFolder();
  const cleaned = base64Data.replace(/^data:image\/\w+;base64,/, '');
  const bytes = Utilities.base64Decode(cleaned);
  const blob = Utilities.newBlob(bytes, 'image/jpeg',
    `absen_${id}_${status}_${new Date().getTime()}.jpg`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?id=${file.getId()}`;
}

// ---------- FOTO BUKTI IZIN (disimpan per nama karyawan, biar gampang dicari) ----------
const DRIVE_FOLDER_BUKTI_IZIN = 'Foto Bukti Izin BPBD';

function sanitizeFolderName(name) {
  // Rapikan nama supaya aman dipakai sebagai nama folder Drive
  const bersih = String(name || 'Tanpa Nama').trim().replace(/[\/\\]/g, '-').replace(/\s+/g, ' ');
  return bersih || 'Tanpa Nama';
}

function getOrCreateSubfolder(parentFolder, subfolderName) {
  const folders = parentFolder.getFoldersByName(subfolderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(subfolderName);
}

function getOrCreateBuktiIzinFolderForPegawai(nama) {
  const rootFolders = DriveApp.getFoldersByName(DRIVE_FOLDER_BUKTI_IZIN);
  const root = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(DRIVE_FOLDER_BUKTI_IZIN);
  return getOrCreateSubfolder(root, sanitizeFolderName(nama));
}

function simpanFotoBuktiKeDrive(base64Data, id, nama) {
  // Struktur Drive: "Foto Bukti Izin BPBD" / {Nama Karyawan} / bukti_{id}_{timestamp}.jpg
  const folder = getOrCreateBuktiIzinFolderForPegawai(nama);
  const cleaned = base64Data.replace(/^data:image\/\w+;base64,/, '');
  const bytes = Utilities.base64Decode(cleaned);
  const blob = Utilities.newBlob(bytes, 'image/jpeg',
    `bukti_${id}_${new Date().getTime()}.jpg`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?id=${file.getId()}`;
}
