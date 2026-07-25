// Tempel isi file ini ke Extensions > Apps Script pada Google Sheet Anda,
// lalu Deploy > Manage deployments > edit deployment yang sudah ada > New version
// (supaya URL /exec tetap sama dan tidak perlu diganti di index.html/admin.html).
//
// PENTING: script ini memakai skema kolom BARU (Tahap/Jilid/Juz/Halaman) supaya
// bisa dirangking. Ini beda dari skema lama (Capaian/Detail bebas teks).
// Disarankan pakai TAB SHEET BARU (nama di bawah, default "Data") supaya data
// lama Anda tidak tertimpa. Data lama tetap aman di tab lamanya, hanya tidak
// otomatis ikut muncul di leaderboard versi baru.

const SHEET_NAME = 'Data'; // ganti sesuai nama tab sheet Anda

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Nama', 'Tahap', 'Jilid', 'Juz', 'Halaman', 'Catatan']);
  }
  return sheet;
}

function doGet(e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const rows = values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return {
        nama: obj.nama || '',
        tahap: obj.tahap || '',
        jilid: obj.jilid || '',
        juz: obj.juz || '',
        halaman: obj.halaman || '',
        catatan: obj.catatan || '',
        tanggal: obj.timestamp instanceof Date ? obj.timestamp.toISOString() : String(obj.timestamp || '')
      };
    });

  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = getSheet_();
  const data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date(),
    data.nama || '',
    data.tahap || '',
    data.jilid || '',
    data.juz || '',
    data.halaman || '',
    data.catatan || ''
  ]);

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
