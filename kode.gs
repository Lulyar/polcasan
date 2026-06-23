/* ============================================================
   KODE.GS — Backend Google Apps Script
   Sistem Informasi Klinik Polcasan
   Database: Google Spreadsheet
   ============================================================ */

// ─── KONFIGURASI ────────────────────────────────────────────
const SPREADSHEET_ID = '1WDCD_c4lrUeOL5MA3AYyzkH2W_sWG9EuUXv97Ao5mvg';

const SHEET_NAMES = {
  PATIENTS: 'Patients',
  APPOINTMENTS: 'Appointments',
  QUEUE: 'Queue',
  PAYMENTS: 'Payments',
  THERAPISTS: 'Therapists',
  ADMIN: 'Admin',
  RECORDS: 'MedicalRecords'
};

// ─── SERVE HTML & API ─────────────────────────────────────────
function doGet(e) {
  // Handle API GET requests from external apps (CORS simple request backup)
  if (e && e.parameter && e.parameter.action) {
    try {
      const action = e.parameter.action;
      const args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
      if (typeof this[action] === 'function') {
        const result = this[action].apply(null, args);
        return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Fungsi tidak ditemukan: ' + action }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Gagal memproses: ' + err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Serve static HTML page
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Klinik Polcasan — Sistem Informasi Fisioterapi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://img.icons8.com/ios-filled/50/D4577A/plus-math.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const args = postData.args || [];
    
    if (typeof this[action] === 'function') {
      const result = this[action].apply(null, args);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Fungsi tidak ditemukan: ' + action }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Gagal memproses: ' + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ─── INISIALISASI SPREADSHEET ───────────────────────────────
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // Helper helper to clear or create sheet
  function setupSheet(name, headers) {
    let s = ss.getSheetByName(name);
    if (s) {
      s.clear();
      s.appendRow(headers);
    } else {
      s = ss.insertSheet(name);
      s.appendRow(headers);
    }
    s.setFrozenRows(1);
    return s;
  }

  // 1. Sheet: Patients
  const patientHeaders = ['PatientID', 'Nama', 'NIK', 'NoTelepon', 'Email', 'TanggalLahir', 'JenisKelamin', 'Alamat', 'Password', 'TanggalDaftar', 'Role'];
  const patientSheet = setupSheet(SHEET_NAMES.PATIENTS, patientHeaders);
  const samplePatients = [
    ['P001', 'Siti Aminah', '3326012405950001', '081234567890', 'siti@mail.com', '1995-05-24', 'Perempuan', 'Jl. Mawar No. 12, Pekalongan', '123', now, 'pasien'],
    ['P002', 'Budi Utomo', '3326012405950002', '081234567891', 'budi@mail.com', '1988-11-12', 'Laki-laki', 'Jl. Melati No. 45, Semarang', '123', now, 'pasien'],
    ['P003', 'Ahmad Fauzi', '3326012405950003', '081234567892', 'fauzi@mail.com', '1990-02-15', 'Laki-laki', 'Jl. Dahlia No. 7, Bandung', '123', now, 'pasien'],
    ['P004', 'Rini Lestari', '3326012405950004', '081234567893', 'rini@mail.com', '1993-07-22', 'Perempuan', 'Jl. Kenanga No. 89, Surabaya', '123', now, 'pasien'],
    ['P005', 'Hendra Wijaya', '3326012405950005', '081234567894', 'hendra@mail.com', '1985-04-30', 'Laki-laki', 'Jl. Anggrek No. 101, Yogyakarta', '123', now, 'pasien'],
    ['P006', 'Dewi Sartika', '3326012405950006', '081234567895', 'dewi@mail.com', '1997-09-08', 'Perempuan', 'Jl. Flamboyan No. 34, Solo', '123', now, 'pasien'],
    ['P007', 'Bambang Sugeng', '3326012405950007', '081234567896', 'bambang@mail.com', '1978-01-25', 'Laki-laki', 'Jl. Kamboja No. 56, Malang', '123', now, 'pasien'],
    ['P008', 'Lilis Suryani', '3326012405950008', '081234567897', 'lilis@mail.com', '1982-06-18', 'Perempuan', 'Jl. Tulip No. 23, Bogor', '123', now, 'pasien'],
    ['P009', 'Joko Widodo', '3326012405950009', '081234567898', 'joko@mail.com', '1961-06-21', 'Laki-laki', 'Jl. Solo No. 10, Depok', '123', now, 'pasien'],
    ['P010', 'Megawati Putri', '3326012405950010', '081234567899', 'mega@mail.com', '1972-01-23', 'Perempuan', 'Jl. Teuku Umar No. 27, Tangerang', '123', now, 'pasien'],
    ['P011', 'Prabowo Subianto', '3326012405950011', '081234567811', 'prabowo@mail.com', '1951-10-17', 'Laki-laki', 'Jl. Kertanegara No. 4, Jakarta', '123', now, 'pasien'],
    ['P012', 'Ganjar Pranowo', '3326012405950012', '081234567812', 'ganjar@mail.com', '1968-10-28', 'Laki-laki', 'Jl. Gajah Mada No. 8, Semarang', '123', now, 'pasien'],
    ['P013', 'Anies Baswedan', '3326012405950013', '081234567813', 'anies@mail.com', '1969-05-07', 'Laki-laki', 'Jl. Lebak Bulus No. 2, Jakarta', '123', now, 'pasien'],
    ['P014', 'Susi Pudjiastuti', '3326012405950014', '081234567814', 'susi@mail.com', '1965-01-15', 'Perempuan', 'Jl. Pananjung No. 13, Pangandaran', '123', now, 'pasien'],
    ['P015', 'Sri Mulyani', '3326012405950015', '081234567815', 'sri@mail.com', '1962-08-26', 'Perempuan', 'Jl. Sudirman No. 1, Jakarta', '123', now, 'pasien'],
    ['P016', 'Basuki Hadimuljono', '3326012405950016', '081234567816', 'basuki@mail.com', '1954-11-05', 'Laki-laki', 'Jl. Pattimura No. 20, Jakarta', '123', now, 'pasien'],
    ['P017', 'Retno Marsudi', '3326012405950017', '081234567817', 'retno@mail.com', '1962-11-27', 'Perempuan', 'Jl. Pejambon No. 6, Jakarta', '123', now, 'pasien'],
    ['P018', 'Erick Thohir', '3326012405950018', '081234567818', 'erick@mail.com', '1970-05-30', 'Laki-laki', 'Jl. Merdeka Barat No. 8, Jakarta', '123', now, 'pasien'],
    ['P019', 'Sandiaga Uno', '3326012405950019', '081234567819', 'sandi@mail.com', '1969-06-28', 'Laki-laki', 'Jl. Senopati No. 12, Jakarta', '123', now, 'pasien'],
    ['P020', 'Gibran Rakabuming', '3326012405950020', '081234567820', 'gibran@mail.com', '1987-10-01', 'Laki-laki', 'Jl. Solo No. 1, Surakarta', '123', now, 'pasien']
  ];
  samplePatients.forEach(row => patientSheet.appendRow(row));

  // 2. Sheet: Appointments
  const aptHeaders = ['AppointmentID', 'PatientID', 'NamaPasien', 'Spesialisasi', 'Tanggal', 'Jam', 'TherapistID', 'NamaTherapist', 'Status', 'NomorAntrean', 'TanggalDibuat'];
  const aptSheet = setupSheet(SHEET_NAMES.APPOINTMENTS, aptHeaders);
  const sampleApt = [
    ['APT001', 'P001', 'Siti Aminah', 'Fisioterapi Muskuloskeletal', today, '08:00', 'T001', 'Khoirunnisa, S.Ftr.Ftr., M.fis', 'Selesai', 'A001', now],
    ['APT002', 'P002', 'Budi Utomo', 'Fisioterapi Neuromuskular', today, '09:00', 'T002', 'dr. Sari Dewi, Sp.FT', 'Selesai', 'A002', now],
    ['APT003', 'P003', 'Ahmad Fauzi', 'Fisioterapi Kardiovaskulopulmonal', today, '10:00', 'T003', 'dr. Budi Santoso, Sp.FT', 'Selesai', 'A003', now],
    ['APT004', 'P004', 'Rini Lestari', 'Fisioterapi Pediatri', today, '11:00', 'T004', 'dr. Maya Putri, Sp.FT', 'Check-in', 'A004', now],
    ['APT005', 'P005', 'Hendra Wijaya', 'Fisioterapi Geriatri', today, '13:00', 'T005', 'dr. Hendra Wijaya, Sp.FT', 'Check-in', 'A005', now],
    ['APT006', 'P006', 'Dewi Sartika', 'Fisioterapi Kesehatan Wanita', today, '14:00', 'T006', 'dr. Linda Kusuma, Sp.FT', 'Terjadwal', 'A006', now],
    ['APT007', 'P007', 'Bambang Sugeng', 'Fisioterapi Sport', today, '15:00', 'T007', 'dr. Riko Firmansyah, Sp.FT', 'Terjadwal', 'A007', now],
    ['APT008', 'P008', 'Lilis Suryani', 'Fisioterapi Muskuloskeletal', today, '16:00', 'T001', 'Khoirunnisa, S.Ftr.Ftr., M.fis', 'Terjadwal', 'A008', now],
    ['APT009', 'P009', 'Joko Widodo', 'Fisioterapi Neuromuskular', today, '17:00', 'T002', 'dr. Sari Dewi, Sp.FT', 'Dibatalkan', 'A009', now],
    ['APT010', 'P010', 'Megawati Putri', 'Fisioterapi Kardiovaskulopulmonal', today, '17:00', 'T003', 'dr. Budi Santoso, Sp.FT', 'Terjadwal', 'A010', now]
  ];
  sampleApt.forEach(row => aptSheet.appendRow(row));

  // 3. Sheet: Queue
  const queueHeaders = ['QueueID', 'AppointmentID', 'NomorAntrean', 'NamaPasien', 'Spesialisasi', 'Status', 'Tanggal', 'WaktuPanggil'];
  const queueSheet = setupSheet(SHEET_NAMES.QUEUE, queueHeaders);
  const sampleQueue = [
    ['Q001', 'APT001', 'A001', 'Siti Aminah', 'Fisioterapi Muskuloskeletal', 'Selesai', today, '08:05:00'],
    ['Q002', 'APT002', 'A002', 'Budi Utomo', 'Fisioterapi Neuromuskular', 'Selesai', today, '09:04:00'],
    ['Q003', 'APT003', 'A003', 'Ahmad Fauzi', 'Fisioterapi Kardiovaskulopulmonal', 'Selesai', today, '10:03:00'],
    ['Q004', 'APT004', 'A004', 'Rini Lestari', 'Fisioterapi Pediatri', 'Dipanggil', today, '11:02:00'],
    ['Q005', 'APT005', 'A005', 'Hendra Wijaya', 'Fisioterapi Geriatri', 'Menunggu', today, ''],
    ['Q006', 'APT006', 'A006', 'Dewi Sartika', 'Fisioterapi Kesehatan Wanita', 'Menunggu', today, ''],
    ['Q007', 'APT007', 'A007', 'Bambang Sugeng', 'Fisioterapi Sport', 'Menunggu', today, ''],
    ['Q008', 'APT008', 'A008', 'Lilis Suryani', 'Fisioterapi Muskuloskeletal', 'Menunggu', today, ''],
    ['Q010', 'APT010', 'A010', 'Megawati Putri', 'Fisioterapi Kardiovaskulopulmonal', 'Menunggu', today, '']
  ];
  sampleQueue.forEach(row => queueSheet.appendRow(row));

  // 4. Sheet: Payments
  const payHeaders = ['PaymentID', 'AppointmentID', 'PatientID', 'NamaPasien', 'Layanan', 'JumlahTagihan', 'MetodePembayaran', 'StatusAsuransi', 'NamaAsuransi', 'StatusBayar', 'TanggalBayar', 'NomorKuitansi'];
  const paySheet = setupSheet(SHEET_NAMES.PAYMENTS, payHeaders);
  const samplePayments = [
    ['PAY001', 'APT001', 'P001', 'Siti Aminah', 'Fisioterapi Muskuloskeletal', '250000', 'QRIS', 'Tidak', '-', 'Lunas', now, 'KWT-20260623-001'],
    ['PAY002', 'APT002', 'P002', 'Budi Utomo', 'Fisioterapi Neuromuskular', '300000', 'Tunai', 'Ya', 'BPJS Kesehatan', 'Lunas', now, 'KWT-20260623-002'],
    ['PAY003', 'APT003', 'P003', 'Ahmad Fauzi', 'Fisioterapi Kardiovaskulopulmonal', '350000', 'Transfer Bank', 'Ya', 'Mandiri Inhealth', 'Lunas', now, 'KWT-20260623-003'],
    ['PAY004', 'APT004', 'P004', 'Rini Lestari', 'Fisioterapi Pediatri', '250000', 'Kartu Debit', 'Tidak', '-', 'Lunas', now, 'KWT-20260623-004'],
    ['PAY005', 'APT005', 'P005', 'Hendra Wijaya', 'Fisioterapi Geriatri', '280000', 'QRIS', 'Tidak', '-', 'Lunas', now, 'KWT-20260623-005']
  ];
  samplePayments.forEach(row => paySheet.appendRow(row));

  // 5. Sheet: Therapists
  const therapistHeaders = ['TherapistID', 'Nama', 'Spesialisasi', 'JadwalHari', 'JamMulai', 'JamSelesai', 'Status'];
  const therapistSheet = setupSheet(SHEET_NAMES.THERAPISTS, therapistHeaders);
  const sampleTherapists = [
    ['T001', 'Khoirunnisa, S.Ftr.Ftr., M.fis', 'Fisioterapi Muskuloskeletal', 'Senin,Selasa,Rabu', '08:00', '15:00', 'Aktif'],
    ['T002', 'dr. Sari Dewi, Sp.FT', 'Fisioterapi Neuromuskular', 'Senin,Rabu,Jumat', '09:00', '16:00', 'Aktif'],
    ['T003', 'dr. Budi Santoso, Sp.FT', 'Fisioterapi Kardiovaskulopulmonal', 'Selasa,Kamis,Sabtu', '08:00', '14:00', 'Aktif'],
    ['T004', 'dr. Maya Putri, Sp.FT', 'Fisioterapi Pediatri', 'Senin,Selasa,Kamis', '10:00', '17:00', 'Aktif'],
    ['T005', 'dr. Hendra Wijaya, Sp.FT', 'Fisioterapi Geriatri', 'Rabu,Kamis,Jumat', '08:00', '15:00', 'Aktif'],
    ['T006', 'dr. Linda Kusuma, Sp.FT', 'Fisioterapi Kesehatan Wanita', 'Senin,Rabu,Sabtu', '09:00', '16:00', 'Aktif'],
    ['T007', 'dr. Riko Firmansyah, Sp.FT', 'Fisioterapi Sport', 'Selasa,Kamis,Sabtu', '07:00', '14:00', 'Aktif']
  ];
  sampleTherapists.forEach(row => therapistSheet.appendRow(row));

  // 6. Sheet: Admin
  const adminHeaders = ['AdminID', 'Username', 'Password', 'Nama', 'Role'];
  const adminSheet = setupSheet(SHEET_NAMES.ADMIN, adminHeaders);
  const sampleAdmin = [
    ['A001', 'admin', '123', 'Administrator', 'admin'],
    ['A002', 'kasir', '123', 'Kasir Klinik', 'kasir'],
    ['T001', 'terapis', '123', 'Khoirunnisa, S.Ftr.Ftr., M.fis', 'terapis']
  ];
  sampleAdmin.forEach(row => adminSheet.appendRow(row));

  // 7. Sheet: MedicalRecords
  const recordHeaders = ['RecordID', 'PatientID', 'NamaPasien', 'Diagnosa', 'Spesialisasi', 'Catatan', 'TherapistID', 'NamaTherapist', 'Tanggal'];
  const recordSheet = setupSheet(SHEET_NAMES.RECORDS, recordHeaders);
  const sampleRecords = [
    ['REC001', 'P001', 'Siti Aminah', 'Low Back Pain', 'Fisioterapi Muskuloskeletal', 'Mengalami nyeri punggung bawah kronis. Dilakukan terapi manual, pemanasan infra-merah, dan edukasi posisi duduk tegak.', 'T001', 'Khoirunnisa, S.Ftr.Ftr., M.fis', today],
    ['REC002', 'P002', 'Budi Utomo', 'Hemiparese Dextra post Stroke', 'Fisioterapi Neuromuskular', 'Terapi latihan gerak pasif dan penguatan otot ekstremitas kanan atas dan bawah.', 'T002', 'dr. Sari Dewi, Sp.FT', today],
    ['REC003', 'P004', 'Rini Lestari', 'Keterlambatan Motorik Kasar', 'Fisioterapi Pediatri', 'Terapi stimulasi merangkak dan berdiri tegak dengan penyangga.', 'T004', 'dr. Maya Putri, Sp.FT', today],
    ['REC004', 'P003', 'Ahmad Fauzi', 'Asma Bronkial', 'Fisioterapi Kardiovaskulopulmonal', 'Terapi inhalasi (nebulisasi) diikuti dengan latihan teknik batuk efektif dan postural drainage. Kondisi dada mulai membaik.', 'T003', 'dr. Budi Santoso, Sp.FT', today],
    ['REC005', 'P005', 'Hendra Wijaya', 'Osteoarthritis Genu', 'Fisioterapi Geriatri', 'Nyeri lutut kiri kronis karena faktor usia. Diberikan terapi TENS dan latihan penguatan otot kuadrisep.', 'T005', 'dr. Hendra Wijaya, Sp.FT', today],
    ['REC006', 'P006', 'Dewi Sartika', 'Frozen Shoulder', 'Fisioterapi Kesehatan Wanita', 'Keterbatasan gerak sendi bahu kanan. Dilakukan mobilisasi sendi dan terapi ultrasound.', 'T006', 'dr. Linda Kusuma, Sp.FT', today],
    ['REC007', 'P007', 'Bambang Sugeng', 'Ankle Sprain Grade II', 'Fisioterapi Sport', 'Cedera pergelangan kaki kanan saat bermain sepak bola. Dilakukan terapi kompres dingin, taping, dan latihan proprioseptif.', 'T007', 'dr. Riko Firmansyah, Sp.FT', today],
    ['REC008', 'P008', 'Lilis Suryani', 'Cervical Syndrome', 'Fisioterapi Muskuloskeletal', 'Nyeri leher menjalar ke lengan kanan akibat spasme otot leher. Diberikan traksi leher and infra-merah.', 'T001', 'Khoirunnisa, S.Ftr.Ftr., M.fis', today],
    ['REC009', 'P009', 'Joko Widodo', 'Hernia Nucleus Pulposus (HNP)', 'Fisioterapi Neuromuskular', 'Saraf terjepit di pinggang. Diberikan terapi traksi lumbal dan edukasi core stability exercises.', 'T002', 'dr. Sari Dewi, Sp.FT', today],
    ['REC010', 'P010', 'Megawati Putri', 'Post-Operative Pulmonary Complications', 'Fisioterapi Kardiovaskulopulmonal', 'Latihan pernapasan dalam (deep breathing exercises) dan ambulasi dini pasca operasi perut.', 'T003', 'dr. Budi Santoso, Sp.FT', today]
  ];
  sampleRecords.forEach(row => recordSheet.appendRow(row));

  return 'Inisialisasi berhasil dengan 10 data dummy!';
}

// ─── UTILITAS ───────────────────────────────────────────────
function generateID(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5);
}

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function getTodayString() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// ─── REGISTRASI PASIEN ──────────────────────────────────────
function registerPatient(data) {
  try {
    const sheet = getSheet(SHEET_NAMES.PATIENTS);
    if (!sheet) return { success: false, message: 'Sheet Patients tidak ditemukan. Jalankan initializeSheets() terlebih dahulu.' };

    // Cek duplikat NIK atau NoTelepon
    const existing = getSheetData(SHEET_NAMES.PATIENTS);
    const dupNIK = existing.find(p => p.NIK === data.nik);
    if (dupNIK) return { success: false, message: 'NIK sudah terdaftar.' };

    const dupPhone = existing.find(p => p.NoTelepon === data.noTelepon);
    if (dupPhone) return { success: false, message: 'Nomor telepon sudah terdaftar.' };

    const patientID = generateID('P');
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      patientID, data.nama, data.nik, data.noTelepon, data.email || '',
      data.tanggalLahir || '', data.jenisKelamin || '', data.alamat || '',
      data.password, now, 'pasien'
    ]);

    return {
      success: true,
      message: 'Registrasi berhasil!',
      patientID: patientID
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── LOGIN PASIEN ───────────────────────────────────────────
function loginPatient(identifier, password) {
  try {
    const patients = getSheetData(SHEET_NAMES.PATIENTS);
    const patient = patients.find(p =>
      (p.NIK === identifier || p.NoTelepon === identifier) && p.Password === password
    );

    if (!patient) return { success: false, message: 'NIK/No. Telepon atau password salah.' };

    return {
      success: true,
      message: 'Login berhasil!',
      patient: {
        id: patient.PatientID,
        nama: patient.Nama,
        nik: patient.NIK,
        noTelepon: patient.NoTelepon,
        email: patient.Email,
        role: patient.Role
      }
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── LOGIN ADMIN ────────────────────────────────────────────
function loginAdmin(username, password) {
  try {
    const admins = getSheetData(SHEET_NAMES.ADMIN);
    const admin = admins.find(a => a.Username === username && a.Password === password);

    if (!admin) return { success: false, message: 'Username atau password salah.' };

    return {
      success: true,
      message: 'Login admin berhasil!',
      admin: {
        id: admin.AdminID,
        nama: admin.Nama,
        role: admin.Role
      }
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── DAFTAR FISIOTERAPIS ────────────────────────────────────
function getTherapists() {
  try {
    const therapists = getSheetData(SHEET_NAMES.THERAPISTS);
    return { success: true, data: therapists };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── JADWAL FISIOTERAPIS ────────────────────────────────────
function getTherapistSchedule(spesialisasi) {
  try {
    const therapists = getSheetData(SHEET_NAMES.THERAPISTS);
    const filtered = therapists.filter(t => t.Spesialisasi === spesialisasi && t.Status === 'Aktif');
    return { success: true, data: filtered };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── SLOT TERSEDIA ──────────────────────────────────────────
function getAvailableSlots(tanggal, spesialisasi) {
  try {
    const therapists = getSheetData(SHEET_NAMES.THERAPISTS);
    const appointments = getSheetData(SHEET_NAMES.APPOINTMENTS);

    const dayNames = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const dateObj = new Date(tanggal);
    const dayName = dayNames[dateObj.getDay()];

    const availableTherapists = therapists.filter(t =>
      t.Spesialisasi === spesialisasi &&
      t.Status === 'Aktif' &&
      t.JadwalHari.split(',').map(d => d.trim()).includes(dayName)
    );

    if (availableTherapists.length === 0) {
      return { success: true, data: [], message: 'Tidak ada fisioterapis yang tersedia pada hari tersebut.' };
    }

    const slots = [];
    availableTherapists.forEach(therapist => {
      const startHour = parseInt(therapist.JamMulai.split(':')[0]);
      const endHour = parseInt(therapist.JamSelesai.split(':')[0]);

      for (let h = startHour; h < endHour; h++) {
        const timeSlot = (h < 10 ? '0' + h : h) + ':00';
        const booked = appointments.find(a =>
          a.Tanggal === tanggal &&
          a.Jam === timeSlot &&
          a.TherapistID === therapist.TherapistID &&
          a.Status !== 'Dibatalkan'
        );

        if (!booked) {
          slots.push({
            therapistID: therapist.TherapistID,
            namaTherapist: therapist.Nama,
            spesialisasi: therapist.Spesialisasi,
            jam: timeSlot,
            tanggal: tanggal
          });
        }
      }
    });

    return { success: true, data: slots };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── BOOKING JANJI TEMU ─────────────────────────────────────
function bookAppointment(patientID, namaPasien, spesialisasi, tanggal, jam, therapistID, namaTherapist) {
  try {
    const sheet = getSheet(SHEET_NAMES.APPOINTMENTS);
    const queueSheet = getSheet(SHEET_NAMES.QUEUE);

    // Generate nomor antrean hari ini
    const todayAppointments = getSheetData(SHEET_NAMES.APPOINTMENTS)
      .filter(a => a.Tanggal === tanggal && a.Status !== 'Dibatalkan');
    const nomorAntrean = 'A' + String(todayAppointments.length + 1).padStart(3, '0');

    const appointmentID = generateID('APT');
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      appointmentID, patientID, namaPasien, spesialisasi,
      tanggal, jam, therapistID, namaTherapist,
      'Terjadwal', nomorAntrean, now
    ]);

    // Tambah ke antrean
    const queueID = generateID('Q');
    queueSheet.appendRow([
      queueID, appointmentID, nomorAntrean, namaPasien,
      spesialisasi, 'Menunggu', tanggal, ''
    ]);

    return {
      success: true,
      message: 'Booking berhasil!',
      data: {
        appointmentID: appointmentID,
        nomorAntrean: nomorAntrean,
        tanggal: tanggal,
        jam: jam,
        spesialisasi: spesialisasi,
        namaTherapist: namaTherapist
      }
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── RIWAYAT JANJI PASIEN ───────────────────────────────────
function getPatientAppointments(patientID) {
  try {
    const appointments = getSheetData(SHEET_NAMES.APPOINTMENTS);
    const filtered = appointments.filter(a => a.PatientID === patientID);
    return { success: true, data: filtered };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── ANTREAN HARI INI ───────────────────────────────────────
function getTodayQueue() {
  try {
    const today = getTodayString();
    const queue = getSheetData(SHEET_NAMES.QUEUE).filter(q => q.Tanggal === today);
    return { success: true, data: queue };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── PANGGIL ANTREAN BERIKUTNYA ─────────────────────────────
function callNextQueue() {
  try {
    const today = getTodayString();
    const sheet = getSheet(SHEET_NAMES.QUEUE);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const tanggalIdx = headers.indexOf('Tanggal');
    const statusIdx = headers.indexOf('Status');
    const waktuPanggilIdx = headers.indexOf('WaktuPanggil');

    // Selesaikan yang sedang dipanggil
    for (let i = 1; i < data.length; i++) {
      if (data[i][tanggalIdx] === today && data[i][statusIdx] === 'Dipanggil') {
        sheet.getRange(i + 1, statusIdx + 1).setValue('Selesai');
      }
    }

    // Panggil berikutnya
    for (let i = 1; i < data.length; i++) {
      if (data[i][tanggalIdx] === today && data[i][statusIdx] === 'Menunggu') {
        sheet.getRange(i + 1, statusIdx + 1).setValue('Dipanggil');
        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
        sheet.getRange(i + 1, waktuPanggilIdx + 1).setValue(now);

        return {
          success: true,
          data: {
            nomorAntrean: data[i][headers.indexOf('NomorAntrean')],
            namaPasien: data[i][headers.indexOf('NamaPasien')],
            spesialisasi: data[i][headers.indexOf('Spesialisasi')]
          }
        };
      }
    }

    return { success: true, data: null, message: 'Tidak ada antrean yang menunggu.' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── ANTREAN SAAT INI (UNTUK DISPLAY) ───────────────────────
function getCurrentServing() {
  try {
    const today = getTodayString();
    const queue = getSheetData(SHEET_NAMES.QUEUE);
    const current = queue.find(q => q.Tanggal === today && q.Status === 'Dipanggil');
    const waiting = queue.filter(q => q.Tanggal === today && q.Status === 'Menunggu');

    return {
      success: true,
      data: {
        current: current || null,
        waitingCount: waiting.length,
        waitingList: waiting
      }
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── ADMIN: SEMUA PASIEN ────────────────────────────────────
function getAllPatients() {
  try {
    const patients = getSheetData(SHEET_NAMES.PATIENTS);
    const sanitized = patients.map(p => ({
      id: p.PatientID,
      nama: p.Nama,
      nik: p.NIK,
      noTelepon: p.NoTelepon,
      email: p.Email,
      tanggalDaftar: p.TanggalDaftar
    }));
    return { success: true, data: sanitized };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── ADMIN: SEMUA JANJI HARI INI ────────────────────────────
function getTodayAppointments() {
  try {
    const today = getTodayString();
    const appointments = getSheetData(SHEET_NAMES.APPOINTMENTS)
      .filter(a => a.Tanggal === today);
    return { success: true, data: appointments };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── ADMIN: CHECK-IN PASIEN ─────────────────────────────────
function checkInPatient(appointmentID) {
  try {
    const sheet = getSheet(SHEET_NAMES.APPOINTMENTS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('AppointmentID');
    const statusIdx = headers.indexOf('Status');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === appointmentID) {
        sheet.getRange(i + 1, statusIdx + 1).setValue('Check-in');
        return { success: true, message: 'Pasien berhasil check-in.' };
      }
    }

    return { success: false, message: 'Appointment tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── ADMIN: TAMBAH PASIEN BARU (WALK-IN) ────────────────────
function addWalkInPatient(data) {
  try {
    // Registrasi pasien baru
    const regResult = registerPatient({
      nama: data.nama,
      nik: data.nik,
      noTelepon: data.noTelepon,
      email: data.email || '',
      password: data.nik, // default password = NIK
      tanggalLahir: data.tanggalLahir || '',
      jenisKelamin: data.jenisKelamin || '',
      alamat: data.alamat || ''
    });

    if (!regResult.success) return regResult;

    // Langsung booking jika ada spesialisasi
    if (data.spesialisasi) {
      const today = getTodayString();
      const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');

      const bookResult = bookAppointment(
        regResult.patientID, data.nama, data.spesialisasi,
        today, now, 'WALK-IN', 'Walk-in'
      );

      return {
        success: true,
        message: 'Pasien baru ditambahkan dan langsung terdaftar.',
        data: bookResult.data
      };
    }

    return { success: true, message: 'Pasien baru berhasil didaftarkan.', patientID: regResult.patientID };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── PEMBAYARAN ─────────────────────────────────────────────
function processPayment(data) {
  try {
    const sheet = getSheet(SHEET_NAMES.PAYMENTS);
    const paymentID = generateID('PAY');
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const nomorKuitansi = 'KWT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + Math.random().toString(36).substr(2, 5).toUpperCase();

    sheet.appendRow([
      paymentID, data.appointmentID, data.patientID, data.namaPasien,
      data.layanan, data.jumlahTagihan, data.metodePembayaran,
      data.statusAsuransi || 'Tidak', data.namaAsuransi || '-',
      'Lunas', now, nomorKuitansi
    ]);

    // Update status appointment
    const aptSheet = getSheet(SHEET_NAMES.APPOINTMENTS);
    const aptData = aptSheet.getDataRange().getValues();
    const headers = aptData[0];
    const idIdx = headers.indexOf('AppointmentID');
    const statusIdx = headers.indexOf('Status');

    for (let i = 1; i < aptData.length; i++) {
      if (aptData[i][idIdx] === data.appointmentID) {
        aptSheet.getRange(i + 1, statusIdx + 1).setValue('Selesai');
        break;
      }
    }

    return {
      success: true,
      message: 'Pembayaran berhasil diproses.',
      data: {
        paymentID: paymentID,
        nomorKuitansi: nomorKuitansi,
        tanggalBayar: now
      }
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── RIWAYAT PEMBAYARAN ────────────────────────────────────
function getPaymentHistory() {
  try {
    const payments = getSheetData(SHEET_NAMES.PAYMENTS);
    return { success: true, data: payments };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── STATISTIK DASHBOARD ────────────────────────────────────
function getDashboardStats() {
  try {
    const today = getTodayString();
    const patients = getSheetData(SHEET_NAMES.PATIENTS);
    const appointments = getSheetData(SHEET_NAMES.APPOINTMENTS);
    const todayApts = appointments.filter(a => a.Tanggal === today);
    const queue = getSheetData(SHEET_NAMES.QUEUE).filter(q => q.Tanggal === today);
    const payments = getSheetData(SHEET_NAMES.PAYMENTS);
    const todayPayments = payments.filter(p => p.TanggalBayar && p.TanggalBayar.toString().startsWith(today));

    let todayRevenue = 0;
    todayPayments.forEach(p => {
      todayRevenue += parseFloat(p.JumlahTagihan) || 0;
    });

    return {
      success: true,
      data: {
        totalPasien: patients.length,
        janjiHariIni: todayApts.length,
        antreanMenunggu: queue.filter(q => q.Status === 'Menunggu').length,
        antreanSelesai: queue.filter(q => q.Status === 'Selesai').length,
        pendapatanHariIni: todayRevenue,
        totalPembayaran: todayPayments.length
      }
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ─── REKAM MEDIS (MEDICAL RECORDS) ──────────────────────────
function getMedicalRecords() {
  try {
    const data = getSheetData(SHEET_NAMES.RECORDS);
    return { success: true, data: data };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function addMedicalRecord(patientID, namaPasien, diagnosa, spesialisasi, catatan, therapistID, namaTherapist, tanggal) {
  try {
    const sheet = getSheet(SHEET_NAMES.RECORDS);
    if (!sheet) return { success: false, message: 'Sheet MedicalRecords tidak ditemukan.' };
    
    const recordID = 'REC' + Date.now();
    sheet.appendRow([recordID, patientID, namaPasien, diagnosa, spesialisasi, catatan, therapistID, namaTherapist, tanggal]);
    
    return { success: true, message: 'Rekam medis berhasil ditambahkan!' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}
function getQueueStatusDirect(identifier) {
  try {
    const queue = getSheetData(SHEET_NAMES.QUEUE);
    const appointments = getSheetData(SHEET_NAMES.APPOINTMENTS);
    
    // Cari di queue hari ini berdasarkan NomorAntrean (misal A001) atau QueueID
    let matchedQueue = queue.find(q => q.NomorAntrean === identifier || q.QueueID === identifier);
    
    if (!matchedQueue) {
      // Cari lewat appointments (bisa berdasarkan NomorAntrean atau AppointmentID)
      const matchedApt = appointments.find(a => 
        a.NomorAntrean === identifier || 
        a.AppointmentID === identifier
      );
      if (matchedApt) {
        matchedQueue = queue.find(q => q.AppointmentID === matchedApt.AppointmentID);
      }
    }
    
    if (!matchedQueue) {
      return { success: false, message: 'Antrean atau ID tidak ditemukan untuk hari ini.' };
    }
    
    // Dapatkan data janji temu untuk info detail
    const apt = appointments.find(a => a.AppointmentID === matchedQueue.AppointmentID);
    
    return {
      success: true,
      data: {
        nomorAntrean: matchedQueue.NomorAntrean,
        namaPasien: matchedQueue.NamaPasien,
        spesialisasi: matchedQueue.Spesialisasi,
        status: matchedQueue.Status,
        waktuPanggil: matchedQueue.WaktuPanggil || '-',
        tanggal: matchedQueue.Tanggal,
        therapist: apt ? apt.NamaTherapist : '-',
        jam: apt ? apt.Jam : '-'
      }
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

