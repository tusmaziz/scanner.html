/**
 * Fungsi utama untuk menampilkan halaman web HTML ketika URL Web App diakses.
 */
function doGet(e) {
  // Kalau dibuka dari scanner.html lewat link (?kode=...&ctx=...), teruskan ke halaman.
  // Nilai dibersihkan ketat supaya aman disisipkan ke dalam JavaScript halaman.
  const q = (e && e.parameter) || {};
  const kode = String(q.kode || '').replace(/[^A-Za-z0-9._\- ]/g, '').trim().slice(0, 100);
  const allowedCtx = ['lookup', 'pinjam', 'stok-masuk'];
  const ctx = allowedCtx.indexOf(q.ctx) !== -1 ? q.ctx : 'lookup';
  const scanParams = kode ? { kode: kode, ctx: ctx } : null;

  const html = HtmlService.createHtmlOutputFromFile('index').getContent()
    .replace('/*__SCAN_PARAMS__*/null', function () { return JSON.stringify(scanParams); });

  return HtmlService.createHtmlOutput(html)
    .setTitle('Sistem Inventaris Sekolah')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Memeriksa dan membuat tabel/sheet awal secara otomatis di Google Sheets jika belum ada.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet Users
  // Cek jumlah baris (data kosong), bukan cuma apakah sheet ada -
  // supaya kalau baris akun sempat terhapus, akun default dibuat ulang
  // otomatis saat setupDatabase() jalan lagi (misalnya lewat loginUser).
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
  }
  if (userSheet.getDataRange().getNumRows() <= 1) {
    if (userSheet.getLastRow() === 0) {
      userSheet.appendRow(['Username', 'Password', 'Nama', 'Role']);
    }
    userSheet.appendRow(['admin', hashPassword('admin123'), 'Administrator Sarpras', 'Admin']);
    userSheet.appendRow(['guru', hashPassword('guru123'), 'Bapak/Ibu Guru', 'Guru']);
  }

  // 2. Sheet Barang
  let barangSheet = ss.getSheetByName('Barang');
  if (!barangSheet) {
    barangSheet = ss.insertSheet('Barang');
    barangSheet.appendRow(['id', 'kode', 'nama', 'kategori', 'jumlah', 'satuan', 'ruangan', 'kondisi']);
    barangSheet.appendRow(['1', 'INV-ELEK-001', 'Proyektor Epson EB-E500', 'Elektronik', 8, 'Unit', 'Lab Komputer', 'Baik']);
    barangSheet.appendRow(['2', 'INV-ELEK-002', 'Laptop Chromebook Acer', 'Elektronik', 25, 'Unit', 'Ruang Guru', 'Baik']);
    barangSheet.appendRow(['3', 'INV-ATK-001', 'Spidol Whiteboard Snowman', 'ATK', 4, 'Dus', 'Gudang Utama', 'Baik']);
    barangSheet.appendRow(['4', 'INV-ATK-002', 'Kertas HVS A4 80gr PaperOne', 'ATK', 2, 'Rim', 'Gudang Utama', 'Baik']);
    barangSheet.appendRow(['5', 'INV-MEB-001', 'Meja Belajar Siswa Kayu', 'Mebel', 120, 'Buah', 'Kelas 7A', 'Baik']);
    barangSheet.appendRow(['6', 'INV-MEB-002', 'Kursi Lipat Chitose', 'Mebel', 45, 'Buah', 'Aula Sekolah', 'Rusak Ringan']);
  }

  // 3. Sheet Ruangan
  let ruangSheet = ss.getSheetByName('Ruangan');
  if (!ruangSheet) {
    ruangSheet = ss.insertSheet('Ruangan');
    ruangSheet.appendRow(['id', 'nama', 'pj']);
    ruangSheet.appendRow(['1', 'Lab Komputer', 'Pak Hendra, S.Kom']);
    ruangSheet.appendRow(['2', 'Ruang Guru', 'Bu Rina, M.Pd']);
    ruangSheet.appendRow(['3', 'Gudang Utama', 'Pak Joko (Sarpras)']);
    ruangSheet.appendRow(['4', 'Kelas 7A', 'Wali Kelas 7A']);
  }

  // 4. Sheet Pinjaman
  let pinjamSheet = ss.getSheetByName('Pinjaman');
  if (!pinjamSheet) {
    pinjamSheet = ss.insertSheet('Pinjaman');
    pinjamSheet.appendRow(['id', 'peminjam', 'role', 'barangNama', 'jumlah', 'tanggal', 'status']);
    pinjamSheet.appendRow(['1', 'Pak Budi, S.Pd', 'Guru / Staf', 'Laptop Chromebook Acer', 1, '2026-08-10', 'DIPINJAM']);
  }

  // 5. Sheet StokLog
  let logSheet = ss.getSheetByName('StokLog');
  if (!logSheet) {
    logSheet = ss.insertSheet('StokLog');
    logSheet.appendRow(['id', 'tanggal', 'jenis', 'nama', 'jumlah', 'ket']);
    logSheet.appendRow(['1', '2026-08-01', 'MASUK', 'Proyektor Epson EB-E500', 3, 'Pengadaan BOS 2026']);
  }

  // 6. Sheet Pengaturan (Kop Sekolah / identitas untuk laporan cetak)
  let settingSheet = ss.getSheetByName('Pengaturan');
  if (!settingSheet) {
    settingSheet = ss.insertSheet('Pengaturan');
  }
  if (settingSheet.getDataRange().getNumRows() <= 1) {
    if (settingSheet.getLastRow() === 0) {
      settingSheet.appendRow(['key', 'value']);
    }
    settingSheet.appendRow(['namaSekolah', 'SMA Al Aziz Community School']);
    settingSheet.appendRow(['alamatSekolah', '']);
    settingSheet.appendRow(['namaKepsek', '']);
  }

  return "Database Google Sheet Siap!";
}

/* =========================================================
 * OTENTIKASI & OTORISASI
 * Semua fungsi yang mengubah/menghapus data WAJIB menerima
 * `token` dan divalidasi lewat requireRole() di server -
 * jangan pernah percaya status login yang dikirim dari client.
 * ========================================================= */

/**
 * Hash password dengan SHA-256 (bukan enkripsi dua arah - cukup untuk
 * memverifikasi kecocokan tanpa menyimpan teks asli password).
 */
function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password));
  return digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

/**
 * Membuat token sesi acak dan menyimpan role pemiliknya di server-side
 * cache (bukan di client), berlaku maksimal 6 jam.
 */
function createSession(user) {
  const token = Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  cache.put('session_' + token, JSON.stringify(user), 21600); // 6 jam, batas maksimum CacheService
  return token;
}

/**
 * Mengambil data sesi dari token. Mengembalikan null kalau token
 * tidak ada / sudah kedaluwarsa.
 */
function getSession(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('session_' + token);
  return raw ? JSON.parse(raw) : null;
}

/**
 * Wajib dipanggil di awal setiap fungsi yang mengubah data.
 * allowedRoles contoh: ['Admin'] atau ['Admin', 'Guru'].
 * Melempar error kalau token tidak valid / role tidak diizinkan -
 * error ini akan diteruskan ke withFailureHandler() di client.
 */
function requireRole(token, allowedRoles) {
  const session = getSession(token);
  if (!session) {
    throw new Error('Sesi tidak valid atau sudah habis, silakan login ulang.');
  }
  if (allowedRoles && allowedRoles.indexOf(session.role) === -1) {
    throw new Error('Anda tidak memiliki izin untuk melakukan aksi ini.');
  }
  return session;
}

/**
 * Otentikasi Login Pengguna (Membedakan Role Admin vs Guru)
 */
function loginUser(username, password) {
  setupDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const hashedInput = hashPassword(password);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && String(data[i][1]) === hashedInput) {
      const user = {
        username: data[i][0],
        nama: data[i][2],
        role: data[i][3] // Admin atau Guru
      };
      const token = createSession(user);
      return { success: true, token: token, user: user };
    }
  }
  return { success: false, message: 'Username atau Password salah!' };
}

/**
 * Mengambil Seluruh Data dari Google Sheets (wajib sudah login, role apa pun)
 */
function getInitialData(token) {
  requireRole(token, ['Admin', 'Guru']);
  setupDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return {
    barang: sheetToObjects(ss.getSheetByName('Barang')),
    ruangan: sheetToObjects(ss.getSheetByName('Ruangan')),
    pinjam: sheetToObjects(ss.getSheetByName('Pinjaman')),
    stokLog: sheetToObjects(ss.getSheetByName('StokLog')),
    settings: getSettingsMap()
  };
}

/**
 * Helper internal: baca sheet Pengaturan (key-value) jadi object biasa.
 */
function getSettingsMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pengaturan');
  const map = {};
  if (!sheet) return map;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    map[rows[i][0]] = rows[i][1];
  }
  return map;
}

/**
 * Menyimpan Kop Sekolah / Nama Kepala Sekolah untuk laporan cetak (Khusus Admin)
 */
function saveSettingsGS(token, settings) {
  requireRole(token, ['Admin']);
  setupDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pengaturan');
  const data = sheet.getDataRange().getValues();

  Object.keys(settings).forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(settings[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, settings[key]]);
    }
  });

  return { success: true };
}

/**
 * Helper mengubah baris sheet menjadi objek JSON
 */
function sheetToObjects(sheet) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  const result = [];

  for (let i = 1; i < rows.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = rows[i][j];
    }
    result.push(obj);
  }
  return result;
}

/**
 * Menambah / Mengubah Data Barang (Khusus Admin)
 */
function saveBarangGS(token, item) {
  requireRole(token, ['Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Barang');
  const data = sheet.getDataRange().getValues();

  const id = item.id || Utilities.getUuid();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      foundRow = i + 1;
      break;
    }
  }

  const jumlah = Number(item.jumlah);
  const jumlahAman = isNaN(jumlah) || jumlah < 0 ? 0 : jumlah;

  const row = [id, item.kode, item.nama, item.kategori, jumlahAman, item.satuan, item.ruangan, item.kondisi];

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, 8).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true, id: id };
}

/**
 * Menghapus Data Barang (Khusus Admin)
 */
function deleteBarangGS(token, id) {
  requireRole(token, ['Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Barang');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

/**
 * Catat Peminjaman Baru (Bisa diakses Guru & Admin)
 */
function savePinjamGS(token, item) {
  const session = requireRole(token, ['Admin', 'Guru']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pinjaman');
  const id = item.id || Utilities.getUuid();
  // Peminjam & role diambil dari sesi server, bukan dari input client,
  // supaya orang tidak bisa mencatut nama pengguna lain.
  sheet.appendRow([id, session.nama, session.role, item.barangNama, item.jumlah, item.tanggal, 'DIPINJAM']);
  return { success: true, id: id };
}

/**
 * Kembalikan Barang Dipinjam (Khusus Admin)
 */
function returnPinjamGS(token, id) {
  requireRole(token, ['Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pinjaman');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 7).setValue('DIKEMBALIKAN');
      return { success: true };
    }
  }
  return { success: false };
}

/**
 * Tambah Log Stok Masuk / Keluar (Khusus Admin)
 */
function addStokLogGS(token, log) {
  requireRole(token, ['Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StokLog');
  const id = log.id || Utilities.getUuid();
  sheet.appendRow([id, log.tanggal, log.jenis, log.nama, log.jumlah, log.ket]);
  return { success: true, id: id };
}

/**
 * Menambah / Mengubah Data Ruangan (Khusus Admin)
 */
function saveRuanganGS(token, ruang) {
  requireRole(token, ['Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Ruangan');
  const data = sheet.getDataRange().getValues();

  const id = ruang.id || Utilities.getUuid();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, 3).setValues([[id, ruang.nama, ruang.pj]]);
  } else {
    sheet.appendRow([id, ruang.nama, ruang.pj]);
  }
  return { success: true, id: id };
}

/**
 * Menghapus Data Ruangan (Khusus Admin)
 */
function deleteRuanganGS(token, id) {
  requireRole(token, ['Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Ruangan');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}
