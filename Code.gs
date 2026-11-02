
/***************************************
 * APLIKASI BOOKING WISATA GUNUNG BROMO
 * Database: Google Spreadsheet
 ***************************************/

const CONFIG = {
  SHEET_GURU: 'DATA_GURU',
  SHEET_BOOKING: 'BOOKING',
  MAX_PESERTA: 5
};

/**
 * Menampilkan halaman aplikasi
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('BOOKING WISATA KE GUNUNG BROMO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Mengambil daftar guru dan karyawan
 */
function getDataGuru() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_GURU);

  if (!sheet) {
    throw new Error('Sheet DATA_GURU tidak ditemukan.');
  }

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const header = data[0].map(v => String(v).trim());

  const idxID = header.indexOf('ID_GURU');
  const idxNama = header.indexOf('NAMA_GURU');

  if (idxID === -1 || idxNama === -1) {
    throw new Error('Header DATA_GURU tidak sesuai.');
  }

  return data.slice(1)
    .filter(row => row[idxID] && row[idxNama])
    .map(row => ({
      id: String(row[idxID]).trim(),
      nama: String(row[idxNama]).trim()
    }));
}

/**
 * Menyimpan booking
 */
function simpanBooking(idGuru, tambahan) {
  const lock = LockService.getScriptLock();

  try {
    // Validasi input
    idGuru = String(idGuru || '').trim();
    tambahan = Number(tambahan);

    if (!idGuru) {
      throw new Error('Silakan pilih nama guru/karyawan.');
    }

    if (!Number.isInteger(tambahan) ||
        tambahan < 0 ||
        tambahan > CONFIG.MAX_PESERTA) {
      throw new Error('Jumlah peserta tambahan tidak valid.');
    }

    // Mencegah transaksi bersamaan
    lock.waitLock(15000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetGuru = ss.getSheetByName(CONFIG.SHEET_GURU);
    const sheetBooking = ss.getSheetByName(CONFIG.SHEET_BOOKING);

    if (!sheetGuru || !sheetBooking) {
      throw new Error('Sheet database tidak ditemukan.');
    }

    // Verifikasi identitas berdasarkan database
    const guru = getDataGuru().find(item => item.id === idGuru);

    if (!guru) {
      throw new Error('Data guru/karyawan tidak ditemukan.');
    }

    // Validasi header booking
    const headers = sheetBooking
      .getRange(1, 1, 1, 7)
      .getValues()[0]
      .map(v => String(v).trim());

    const expected = [
      'ID_BOOKING',
      'TANGGAL_BOOKING',
      'ID_GURU',
      'NAMA_GURU',
      'TAMBAHAN',
      'TOTAL_KURSI',
      'STATUS'
    ];

    if (expected.some((h, i) => headers[i] !== h)) {
      throw new Error('Header sheet BOOKING tidak sesuai.');
    }

    // Cek booking sebelumnya yang masih aktif
    const lastRow = sheetBooking.getLastRow();

    if (lastRow >= 2) {
      const records = sheetBooking
        .getRange(2, 1, lastRow - 1, 7)
        .getValues();

      const sudahBooking = records.some(row =>
        String(row[2]).trim() === idGuru &&
        String(row[6]).trim().toUpperCase() === 'AKTIF'
      );

      if (sudahBooking) {
        throw new Error(
          'Anda sudah melakukan booking sebelumnya. ' +
          'Setiap guru/karyawan hanya diperbolehkan satu kali booking.'
        );
      }
    }

    // Hitung kursi
    const totalKursi = 1 + tambahan;

    // Buat ID booking unik
    const idBooking =
      'BKM-' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyyMMdd-HHmmss'
      ) +
      '-' +
      Utilities.getUuid().substring(0, 6).toUpperCase();

    const tanggal = new Date();

    // Simpan transaksi
    sheetBooking.appendRow([
      idBooking,
      tanggal,
      guru.id,
      guru.nama,
      tambahan,
      totalKursi,
      'AKTIF'
    ]);

    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Booking berhasil disimpan!',
      idBooking: idBooking,
      nama: guru.nama,
      tambahan: tambahan,
      totalKursi: totalKursi
    };

  } catch (error) {
    throw new Error(error.message || 'Terjadi kesalahan saat menyimpan booking.');

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
