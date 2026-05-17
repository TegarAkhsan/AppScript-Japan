function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SIPADJA - Desa Japanan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Fungsi khusus untuk memancing otorisasi Google Drive (Write Permission)
function authorizeDrive() {
  const file = DriveApp.createFile("temp_auth_trigger.txt", "trigger");
  file.setTrashed(true); // Hapus langsung agar tidak nyampah
  Logger.log("Selesai memanggil DriveApp (Write Permission)");
}

// Fungsi untuk mengetes akses folder spesifik
function testFolderAccess() {
  try {
    Logger.log("Mencoba mengambil folder ID: " + FOLDER_ID);
    const folder = DriveApp.getFolderById(FOLDER_ID);
    Logger.log("Folder berhasil diambil: " + folder.getName());
    
    Logger.log("Mencoba membuat file dummy di folder tersebut...");
    const file = folder.createFile("test_akses.txt", "Tes Akses");
    Logger.log("File dummy berhasil dibuat! URL: " + file.getUrl());
    
    Logger.log("Menghapus file dummy...");
    file.setTrashed(true);
    Logger.log("Semua tes berhasil! Akun kamu punya akses penuh ke folder ini.");
  } catch (e) {
    Logger.log("TES GAGAL! Error detail: " + e.toString());
  }
}

// Configuration
const SPREADSHEET_ID = '1DLif1iFWVWCxQDX69epfcxA-NST7n291lmFO5hFq8Qc'; 
const FOLDER_ID = '1gzZnph9C_06ipxeDfkIPqIkl7PC6GpWs';

function checkLogin(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users') || createUsersSheet(ss);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === username && data[i][1].toString() === password) {
      return { success: true, user: username };
    }
  }
  return { success: false };
}

function saveArchive(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Archives') || createArchivesSheet(ss);
    
    // Pastikan header benar
    fixHeaders(sheet);
    
    let fileUrl = '';
    if (formData.fileData && formData.fileName) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const contentType = formData.fileData.substring(formData.fileData.indexOf(":") + 1, formData.fileData.indexOf(";"));
      const bytes = Utilities.base64Decode(formData.fileData.split(",")[1]);
      const blob = Utilities.newBlob(bytes, contentType, formData.fileName);
      const file = folder.createFile(blob);
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (sharingError) {
        Logger.log("Gagal mengatur hak akses publik (kemungkinan diblokir kebijakan organisasi): " + sharingError.toString());
      }
      fileUrl = file.getUrl();
    }

    sheet.appendRow([
      new Date(), // Timestamp
      formData.nomorSurat || '',
      formData.judulArsip || '',
      formData.jenisSurat || '',
      formData.kategoriArsip || '',
      formData.tanggalArsip || '',
      formData.statusArsip || '',
      formData.masaRetensi || '',
      formData.keterangan || '',
      fileUrl
    ]);
    
    return { success: true, url: fileUrl };
  } catch (e) {
    throw new Error('Gagal menyimpan data: ' + e.message);
  }
}

function getArchives() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Archives');
  
  if (!sheet) {
    throw new Error('Sheet "Archives" tidak ditemukan!');
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return JSON.stringify([]);
  
  const headers = data[0].map(h => h.toString().toLowerCase().replace(/ /g, ''));
  const rows = data.slice(1);
  
  const result = rows.map((row, i) => {
    let obj = { rowNumber: i + 2 }; // Simpan nomor baris asli (header + 1)
    headers.forEach((header, index) => {
      if (header) {
        let val = row[index];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd");
        }
        obj[header] = (val !== null && val !== undefined) ? val.toString() : '';
      }
    });
    return obj;
  });
  
  return JSON.stringify(result);
}

function deleteArchive(rowNumber) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Archives');
  sheet.deleteRow(rowNumber);
  return { success: true };
}

function updateArchive(rowNumber, formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Archives');
  
  // Ambil baris yang ada untuk mempertahankan file link jika tidak diubah
  const currentRow = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
  
  const updatedData = [
    currentRow[0], // Tetap gunakan timestamp asli
    formData.nomorSurat,
    formData.judulArsip,
    formData.jenisSurat,
    formData.kategoriArsip,
    formData.tanggalArsip,
    formData.statusArsip,
    formData.masaRetensi,
    formData.keterangan,
    currentRow[9] // Tetap gunakan link file yang lama (untuk saat ini edit file belum didukung)
  ];
  
  sheet.getRange(rowNumber, 1, 1, 10).setValues([updatedData]);
  return { success: true };
}

function fixHeaders(sheet) {
  const targetHeaders = ['Timestamp', 'Nomor Surat', 'Judul', 'Jenis Surat', 'Kategori', 'Tanggal', 'Status', 'Masa Retensi', 'Keterangan', 'File Link'];
  sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
}

function createArchivesSheet(ss) {
  let sheet = ss.getSheetByName('Archives');
  if (!sheet) sheet = ss.insertSheet('Archives');
  const headers = ['Timestamp', 'Nomor Surat', 'Judul', 'Jenis Surat', 'Kategori', 'Tanggal', 'Status', 'Masa Retensi', 'Keterangan', 'File Link'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}
function generateDummyData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Archives') || createArchivesSheet(ss);
  
  const jenisSurat = ['Surat Keputusan', 'Surat Keterangan', 'Surat Undangan', 'Berita Acara', 'Surat Pengantar'];
  const kategori = ['Jual Beli Warisan', 'Surat Pernyataan Warisan', 'Sertifikat Tanah', 'Foto Kegiatan', 'Peraturan Desa', 'Lainnya'];
  const status = ['Aktif', 'Inaktif'];
  
  const dummyData = [];
  const today = new Date();
  
  for (let i = 1; i <= 20; i++) {
    const randomDate = new Date();
    randomDate.setFullYear(today.getFullYear() - Math.floor(Math.random() * 8)); // Random year in past 8 years
    randomDate.setMonth(Math.floor(Math.random() * 12));
    randomDate.setDate(Math.floor(Math.random() * 28));
    
    dummyData.push([
      new Date(), // Timestamp
      `SKD/0${i}/2024`, // Nomor Surat
      `Dokumen Arsip Dummy ${i}`, // Judul
      jenisSurat[Math.floor(Math.random() * jenisSurat.length)], // Jenis
      kategori[Math.floor(Math.random() * kategori.length)], // Kategori
      Utilities.formatDate(randomDate, "GMT+7", "yyyy-MM-dd"), // Tanggal
      status[Math.floor(Math.random() * status.length)], // Status
      Math.floor(Math.random() * 10) + 1, // Masa Retensi (1-10 thn)
      `Keterangan otomatis untuk data dummy nomor ${i}`, // Keterangan
      'https://drive.google.com/file/d/1gzZnph9C_06ipxeDfkIPqIkl7PC6GpWs/view' // Dummy Link
    ]);
  }
  
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, dummyData.length, dummyData[0].length).setValues(dummyData);
  return { success: true, count: dummyData.length };
}
