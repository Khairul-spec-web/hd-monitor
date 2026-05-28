/**
 * ═══════════════════════════════════════════════════════════════════
 *  HD MONITOR — Google Apps Script Backend
 *  Hospital Tambunan, KKM
 *  Versi: 1.0 | Tarikh: 2026
 * ═══════════════════════════════════════════════════════════════════
 *
 *  CARA PASANG:
 *  1. Buka Google Sheets → Extensions → Apps Script
 *  2. Tampal keseluruhan fail ini, gantikan kod sedia ada
 *  3. Klik ikon 💾 Save
 *  4. Klik Deploy → New deployment → Web App
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. Klik Deploy → Authorize → salin URL Web App
 *  6. Tampal URL tersebut dalam aplikasi HD Monitor (tab Tetapan)
 * ═══════════════════════════════════════════════════════════════════
 */

// ── KONFIGURASI ──────────────────────────────────────────────────
const CONFIG = {
  SHEET_REKOD:     'Rekod HD',        // Nama tab untuk rekod utama
  SHEET_INTRA:     'Data Intradialitik', // Nama tab untuk data 2-jam
  SHEET_LOG:       'Log Sistem',      // Nama tab untuk log ralat/aktiviti
  FREEZE_ROWS:     1,                 // Beku baris header
  DATE_FORMAT:     'dd/MM/yyyy',
  TIME_ZONE:       'Asia/Kuching',    // Malaysia (UTC+8)
};

// ── HEADER LAJUR REKOD UTAMA ──────────────────────────────────────
const HEADERS_REKOD = [
  'No.',
  'Timestamp',
  'Hospital',
  'Unit',
  'Tarikh',
  'Masa Mula',
  'Masa Tamat',
  'Nama Pesakit',
  'No. Pesakit',
  'Syif',
  'Staff Mula',
  'Staff Tamat',
  'Berat Kering (kg)',
  'Dialyser',
  'Preskripsi',
  'Penampilan Am',
  'Fistula',
  'Penilaian Pra-HD',
  // PRE HD
  'BP Pre (mmHg)',
  'Nadi Pre (/min)',
  'Berat Pre (kg)',
  'Suhu Pre (°C)',
  'SpO2 Pre (%)',
  'P/S Pre (/10)',
  'IDW (kg)',
  'BP Pre Berdiri',
  // POST HD
  'BP Post (mmHg)',
  'Nadi Post (/min)',
  'Berat Post (kg)',
  'Suhu Post (°C)',
  'SpO2 Post (%)',
  'P/S Post (/10)',
  'TFR Sebenar (kg)',
  'IDW Baki (kg)',
  // PENGIRAAN AUTO
  'Perbezaan BP Sistolik',
  'Pencapaian UF (%)',
  'Status TFR',
];

// ── HEADER LAJUR DATA INTRADIALITIK ──────────────────────────────
const HEADERS_INTRA = [
  'No.',
  'Timestamp Rekod',
  'Nama Pesakit',
  'No. Pesakit',
  'Tarikh',
  'Bacaan Ke-',
  'Masa',
  'BP (mmHg)',
  'Nadi (/min)',
  'Suhu (°C)',
  'Qb (ml/min)',
  'Heparin (unit)',
  'VP (mmHg)',
  'TFR (ml)',
  'TMP (mmHg)',
];

// ─────────────────────────────────────────────────────────────────
//  doPost — Titik masuk utama. Dipanggil oleh aplikasi HD Monitor.
// ─────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    // Parse data JSON dari aplikasi
    const data = JSON.parse(e.postData.contents);

    // Abaikan request ujian (test ping)
    if (data.test === true) {
      return buildResponse('OK — HD Monitor Apps Script berjalan dengan baik.');
    }

    // Pastikan spreadsheet dan semua tab wujud
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRekod = getOrCreateSheet(ss, CONFIG.SHEET_REKOD, HEADERS_REKOD);
    const sheetIntra = getOrCreateSheet(ss, CONFIG.SHEET_INTRA, HEADERS_INTRA);
    const sheetLog   = getOrCreateSheet(ss, CONFIG.SHEET_LOG,   ['Timestamp','Jenis','Mesej','IP']);

    // Timestamp semasa (Malaysia UTC+8)
    const now = new Date();
    const ts = Utilities.formatDate(now, CONFIG.TIME_ZONE, "dd/MM/yyyy HH:mm:ss");

    // ── PENGIRAAN AUTO ───────────────────────────────────────────
    let diffBP = '';
    try {
      const sysPre  = parseInt(String(data.bp_pre  || '').split('/')[0]);
      const sysPost = parseInt(String(data.bp_post || '').split('/')[0]);
      if (!isNaN(sysPre) && !isNaN(sysPost)) {
        diffBP = (sysPost - sysPre).toString();
      }
    } catch (_) {}

    let pencapaianUF = '';
    try {
      const ufGoal = parseFloat(String(data.preskripsi || '').match(/UF:([\d.]+)/)?.[1] || '');
      const tfrKg  = parseFloat(String(data.tfr || '').replace(' kg',''));
      if (!isNaN(ufGoal) && ufGoal > 0 && !isNaN(tfrKg)) {
        pencapaianUF = ((tfrKg / ufGoal) * 100).toFixed(1) + '%';
      }
    } catch (_) {}

    let statusTFR = '';
    try {
      const tfrKg = parseFloat(String(data.tfr || '').replace(' kg',''));
      const bkKg  = parseFloat(data.berat_kering || '');
      const wtPost = parseFloat(data.wt_post || '');
      if (!isNaN(tfrKg) && !isNaN(bkKg) && !isNaN(wtPost)) {
        const ib = wtPost - bkKg;
        if (Math.abs(ib) <= 0.5)     statusTFR = 'Mencapai berat kering';
        else if (ib > 0.5)           statusTFR = 'Masih ' + ib.toFixed(1) + ' kg atas berat kering';
        else                         statusTFR = 'Di bawah berat kering ' + Math.abs(ib).toFixed(1) + ' kg';
      }
    } catch (_) {}

    // ── REKOD UTAMA ───────────────────────────────────────────────
    const nextNo = sheetRekod.getLastRow(); // Nombor baris (termasuk header)
    const row = [
      nextNo,                   // No.
      ts,                       // Timestamp
      data.hospital || 'Hospital Tambunan',
      data.unit     || 'Unit Hemodialisis',
      data.tarikh       || '',
      data.masa_mula    || '',
      data.masa_tamat   || '',
      data.nama         || '',
      data.no_pesakit   || '',
      data.syif         || '',
      data.staff_mula   || '',
      data.staff_tamat  || '',
      data.berat_kering || '',
      data.dialyser     || '',
      data.preskripsi   || '',
      data.penampilan   || '',
      data.fistula      || '',
      data.penilaian    || '',
      data.bp_pre       || '',
      data.hr_pre       || '',
      data.wt_pre       || '',
      data.tmp_pre      || '',
      data.spo2_pre     || '',
      data.ps_pre       || '',
      data.idw          || '',
      data.bp_pre_berdiri || '',
      data.bp_post      || '',
      data.hr_post      || '',
      data.wt_post      || '',
      data.tmp_post     || '',
      data.spo2_post    || '',
      data.ps_post      || '',
      data.tfr          || '',
      data.idw_baki     || '',
      diffBP,
      pencapaianUF,
      statusTFR,
    ];
    sheetRekod.appendRow(row);

    // Warna baris alternatif untuk kebolehbacaan
    applyRowFormatting(sheetRekod);

    // ── DATA INTRADIALITIK ─────────────────────────────────────────
    if (data.data_intra) {
      try {
        const intraList = JSON.parse(data.data_intra);
        intraList.forEach((item, idx) => {
          const intraRow = [
            sheetIntra.getLastRow(),
            ts,
            data.nama || '',
            data.no_pesakit || '',
            data.tarikh || '',
            idx + 1,
            item.masa  || '',
            item.bp    || '',
            item.hr    || '',
            item.tmp   || '',
            item.qb    || '',
            item.hep   || '',
            item.vp    || '',
            item.tfr   || '',
            item.tmp2  || '',
          ];
          sheetIntra.appendRow(intraRow);
        });
      } catch (_) {}
    }

    // ── LOG AKTIVITI ──────────────────────────────────────────────
    sheetLog.appendRow([ts, 'DATA_MASUK', 'Rekod diterima: ' + (data.nama || 'Tiada nama'), '']);

    // Semak amaran kritikal dan tandakan dalam log
    checkCriticalAlerts(data, sheetLog, ts);

    return buildResponse('OK');

  } catch (err) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetLog = getOrCreateSheet(ss, CONFIG.SHEET_LOG, ['Timestamp','Jenis','Mesej','IP']);
      sheetLog.appendRow([new Date().toISOString(), 'RALAT', err.message, '']);
    } catch (_) {}
    return buildResponse('RALAT: ' + err.message);
  }
}

// ─────────────────────────────────────────────────────────────────
//  doGet — Untuk ujian sambungan dari pelayar (browser)
// ─────────────────────────────────────────────────────────────────
function doGet(e) {
  return buildResponse('HD Monitor Apps Script — Aktif ✓ | Hospital Tambunan KKM');
}

// ─────────────────────────────────────────────────────────────────
//  Semak amaran kritikal dan rekod dalam log
// ─────────────────────────────────────────────────────────────────
function checkCriticalAlerts(data, sheetLog, ts) {
  const alerts = [];

  // Semak BP
  try {
    const sys = parseInt(String(data.bp_pre || '').split('/')[0]);
    const dia = parseInt(String(data.bp_pre || '').split('/')[1]);
    if (sys >= 180 || dia >= 110) alerts.push('⚠ HIPERTENSI TERUK Pre: ' + data.bp_pre);
    if (sys < 90  || dia < 60)   alerts.push('⚠ HIPOTENSI Pre: ' + data.bp_pre);
  } catch (_) {}

  // Semak SpO2
  try {
    const spo2 = parseInt(data.spo2_pre || '');
    if (spo2 < 94) alerts.push('⚠ SpO2 RENDAH Pre: ' + spo2 + '%');
  } catch (_) {}

  // Semak Post BP
  try {
    const sys = parseInt(String(data.bp_post || '').split('/')[0]);
    const dia = parseInt(String(data.bp_post || '').split('/')[1]);
    if (sys < 90 || dia < 60) alerts.push('⚠ HIPOTENSI Post: ' + data.bp_post);
  } catch (_) {}

  alerts.forEach(msg => {
    sheetLog.appendRow([ts, 'AMARAN_KLINIKAL', (data.nama || '') + ' — ' + msg, '']);
  });
}

// ─────────────────────────────────────────────────────────────────
//  getOrCreateSheet — Dapatkan atau cipta tab dengan header
// ─────────────────────────────────────────────────────────────────
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Tulis header
    sheet.appendRow(headers);
    // Format header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#185FA5');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontSize(11);
    headerRange.setWrap(false);
    // Beku baris header
    sheet.setFrozenRows(CONFIG.FREEZE_ROWS);
    // Auto-resize lajur
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

// ─────────────────────────────────────────────────────────────────
//  applyRowFormatting — Warna selang-seli untuk kebolehbacaan
// ─────────────────────────────────────────────────────────────────
function applyRowFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(lastRow, 1, 1, lastCol);
  range.setFontSize(11);
  if (lastRow % 2 === 0) {
    range.setBackground('#EAF3DE'); // Hijau muda untuk baris genap
  } else {
    range.setBackground('#ffffff'); // Putih untuk baris ganjil
  }
  // Sempadan bawah untuk keterbacaan
  range.setBorder(false, false, true, false, false, false, '#B4B2A9', SpreadsheetApp.BorderStyle.SOLID);
}

// ─────────────────────────────────────────────────────────────────
//  buildResponse — Bina respons HTTP yang betul untuk CORS
// ─────────────────────────────────────────────────────────────────
function buildResponse(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: msg, time: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────────
//  setupSpreadsheet — Jalankan sekali untuk persediaan awal.
//  Pergi ke Apps Script Editor → pilih fungsi ini → klik Run.
// ─────────────────────────────────────────────────────────────────
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Namakan semula Sheet1 sebagai tab rekod utama
  const firstSheet = ss.getSheets()[0];
  if (firstSheet.getName() === 'Sheet1') {
    firstSheet.setName(CONFIG.SHEET_REKOD);
  }

  // Cipta semua tab
  getOrCreateSheet(ss, CONFIG.SHEET_REKOD, HEADERS_REKOD);
  getOrCreateSheet(ss, CONFIG.SHEET_INTRA, HEADERS_INTRA);
  getOrCreateSheet(ss, CONFIG.SHEET_LOG,   ['Timestamp','Jenis','Mesej','IP']);

  // Cipta tab Ringkasan
  let sheetSum = ss.getSheetByName('Ringkasan');
  if (!sheetSum) {
    sheetSum = ss.insertSheet('Ringkasan');
    sheetSum.getRange('A1').setValue('HD Monitor — Ringkasan');
    sheetSum.getRange('A1').setFontSize(16).setFontWeight('bold');
    sheetSum.getRange('A2').setValue('Hospital Tambunan, Sabah | KKM');
    sheetSum.getRange('A3').setValue('Dikemas kini secara automatik apabila data baru diterima.');
    sheetSum.getRange('A3').setFontColor('#5F5E5A');
  }

  SpreadsheetApp.getUi().alert(
    '✅ Persediaan selesai!\n\n' +
    'Tab berikut telah dicipta:\n' +
    '• Rekod HD\n' +
    '• Data Intradialitik\n' +
    '• Log Sistem\n' +
    '• Ringkasan\n\n' +
    'Seterusnya, deploy sebagai Web App dan salin URL ke dalam aplikasi HD Monitor.'
  );
}
