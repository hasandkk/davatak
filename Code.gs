// =====================================================
// Teminat Group - Dava Takip Sistemi
// Google Apps Script Backend
// =====================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Teminat Group - Dava Takip Sistemi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─── Sheet Names ───────────────────────────────────
const SHEET_NAMES = {
  OTO_HASAR:       'OTO HASAR',
  OTO_DISI:        'OTO DIŞI HASAR',
  DEGER_KAYBI:     'DEĞER KAYBI',
  HAK_MAHRUMIYETI: 'HAK MAHRUMİYETİ',
  HASAR_FARKI:     'HASAR FARKI',
  KAZANC_KAYBI:    'KAZANÇ KAYBI',
  DASK:            'DASK',
  AYIPLI_MAL:      'AYIPLI MAL'
};

const ALL_SHEETS = Object.values(SHEET_NAMES);

// Header config for each sheet (used for initializeSheets and addRow)
const SHEET_HEADERS = {
  'OTO HASAR': [
    'SAYI','PLAKA','DOSYA AÇILIŞ','KAZA TARİHİ','ARAÇ SAHİBİ',
    'T.C/V.N','SİGORTA ŞİRKETİ','TRAFİK/KASKO','POLİÇE NO',
    'DOSYA NUMARASI','EKSPERİ','DURUM','ALINAN ÖDEME'
  ],
  'OTO DIŞI HASAR': [
    'SAYI','ADRES','DOSYA AÇILIŞ','HASAR TARİHİ','SİGORTALI',
    'T.C/V.N','SİGORTA ŞİRKETİ','DASK/KONUT','POLİÇE NO',
    'DOSYA NUMARASI','EKSPERİ','DURUM','ALINAN ÖDEME'
  ],
  'DEĞER KAYBI': [
    'S NO','PLAKA','TC','DAVACI','DAVALI','DAVALI SİGORTA','DOSYA NO','KONU DOSYA',
    'SİGORTA MÜRACAAT TARİHİ','DOSYA BEDELİ','SİGORTA DÖNÜŞ TARİHİ','GELEN BEDEL',
    'TAHKİM MÜRACAAT TARİHİ','TAHKİM DOSYA NO','TAHKİM ÜCRETİ','BİLİRKİŞİ ÜCRETİ',
    'DOSYA DURUMU','KARAR TARİHİ','İCRA DOSYA NUM.','İCRA TARİHİ','GELEN BEDEL2'
  ],
  'HAK MAHRUMİYETİ': [
    'S NO','PLAKA','TC','DAVACI','DAVALI','DAVALI SİGORTA','DOSYA NO','KONU DOSYA',
    'SİGORTA MÜRACAAT TARİHİ','DOSYA BEDELİ','SİGORTA DÖNÜŞ TARİHİ','GELEN BEDEL',
    'TAHKİM MÜRACAAT TARİHİ','TAHKİM DOSYA NUMARASI','TAHKİM ÜCRETİ','BİLİRKİŞİ ÜCRETİ',
    'DOSYA DURUMU','KARAR TARİHİ','İCRA DOSYA NUM.','İCRA TARİHİ','GELEN BEDEL2'
  ],
  'HASAR FARKI': [
    'S NO','PLAKA','TC','DAVACI','DAVALI','DAVALI SİGORTA','DOSYA NO','KONU DOSYA',
    'SİGORTA MÜRACAAT TARİHİ','DOSYA BEDELİ','SİGORTA DÖNÜŞ TARİHİ','GELEN BEDEL',
    'TAHKİM MÜRACAAT TARİHİ','TAHKİM DOSYA NUMARASI','TAHKİM ÜCRETİ','BİLİRKİŞİ ÜCRETİ',
    'DOSYA DURUMU','KARAR TARİHİ','İCRA TARİHİ','GELEN BEDEL2'
  ],
  'KAZANÇ KAYBI': [
    'S NO','PLAKA','TC','DAVACI','DAVALI','DAVALI SİGORTA','DOSYA NO','KONU DOSYA',
    'SİGORTA MÜRACAAT TARİHİ','DOSYA BEDELİ','SİGORTA DÖNÜŞ TARİHİ','GELEN BEDEL',
    'TAHKİM MÜRACAAT TARİHİ','TAHKİM DOSYA NUMARASI','TAHKİM ÜCRETİ','BİLİRKİŞİ ÜCRETİ',
    'DOSYA DURUMU','KARAR TARİHİ','İCRA TARİHİ','GELEN BEDEL2'
  ],
  'DASK': [
    'S NO','DAVACI','DAVALI SİGORTA','DOSYA NO','KONU DOSYA',
    'SİGORTA MÜRACAAT TARİHİ','DOSYA BEDELİ','SİGORTA DÖNÜŞ TARİHİ','GELEN BEDEL',
    'TAHKİM MÜRACAAT TARİHİ','TAHKİM DOSYA NUMARASI','TAHKİM ÜCRETİ','BİLİRKİŞİ ÜCRETİ',
    'DOSYA DURUMU','KARAR TARİHİ','İCRA TARİHİ','GELEN BEDEL2'
  ],
  'AYIPLI MAL': [
    'S NO','PLAKA','TC-VN','DAVACI','DAVALI','DAVALI SİGORTA','DOSYA NO','KONU DOSYA',
    'DOSYA BEDELİ','SİGORTA DÖNÜŞ TARİHİ','GELEN BEDEL',
    'TAHKİM MÜRACAAT TARİHİ','TAHKİM DOSYA NUMARASI','TAHKİM ÜCRETİ','BİLİRKİŞİ ÜCRETİ',
    'DOSYA DURUMU','KARAR TARİHİ','İCRA DOSYA NUM.','İCRA TARİHİ','GELEN BEDEL2'
  ]
};

// ─── Helpers ───────────────────────────────────────

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function findHeaderRow(values) {
  for (var i = 0; i < values.length; i++) {
    if (values[i].some(function(c) { return c !== null && c !== ''; })) {
      return i;
    }
  }
  return 0;
}

function formatCellValue(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  }
  return val === null ? '' : val;
}

// ─── CRUD Operations ───────────────────────────────

function getSheetData(sheetName) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sayfa bulunamadı: ' + sheetName };

    var range = sheet.getDataRange();
    if (range.getNumRows() < 1) return { success: true, headers: [], data: [] };

    var values = range.getValues();
    var headerRowIdx = findHeaderRow(values);
    var headers = values[headerRowIdx].map(function(h) { return String(h || ''); });

    var validHeaders = headers.filter(function(h) {
      return h && !h.startsWith('Sütun');
    });

    var rows = [];
    for (var i = headerRowIdx + 1; i < values.length; i++) {
      var row = values[i];
      if (!row.some(function(c) { return c !== null && c !== ''; })) continue;

      var obj = { _rowIndex: i + 1 };
      headers.forEach(function(header, j) {
        if (header && !header.startsWith('Sütun')) {
          obj[header] = formatCellValue(row[j]);
        }
      });
      rows.push(obj);
    }

    return { success: true, headers: validHeaders, data: rows };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function addRow(sheetName, rowData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sayfa bulunamadı: ' + sheetName };

    var values = sheet.getDataRange().getValues();
    var headerRowIdx = findHeaderRow(values);
    var headers = values[headerRowIdx];

    // Count existing data rows for auto-numbering
    var dataRows = values.slice(headerRowIdx + 1).filter(function(r) {
      return r.some(function(c) { return c !== null && c !== ''; });
    });
    var nextNo = dataRows.length + 1;

    var newRow = headers.map(function(header, idx) {
      if (idx === 0 && (header === 'SAYI' || header === 'S NO')) return nextNo;
      return rowData[header] !== undefined ? rowData[header] : '';
    });

    sheet.appendRow(newRow);

    // Basic formatting for new row
    var lastRow = sheet.getLastRow();
    var bg = lastRow % 2 === 0 ? '#f8f9fa' : '#ffffff';
    sheet.getRange(lastRow, 1, 1, headers.length).setBackground(bg);

    return { success: true, message: 'Kayıt başarıyla eklendi (#' + nextNo + ')' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function updateRow(sheetName, rowIndex, rowData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sayfa bulunamadı: ' + sheetName };

    var values = sheet.getDataRange().getValues();
    var headerRowIdx = findHeaderRow(values);
    var headers = values[headerRowIdx];
    var existingRow = values[rowIndex - 1];

    var updatedRow = headers.map(function(header, j) {
      if (!header || header.startsWith('Sütun')) return existingRow[j];
      if (rowData[header] !== undefined && rowData[header] !== null) return rowData[header];
      return existingRow[j];
    });

    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
    return { success: true, message: 'Kayıt güncellendi' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteRow(sheetName, rowIndex) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sayfa bulunamadı: ' + sheetName };

    sheet.deleteRow(rowIndex);
    return { success: true, message: 'Kayıt silindi' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── Dashboard Stats ────────────────────────────────

function getDashboardStats() {
  try {
    var ss = getSpreadsheet();
    var totalCases = 0;
    var activeCount = 0;
    var completedCount = 0;
    var sheetCounts = {};
    var recentActivity = [];

    ALL_SHEETS.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (!sheet) { sheetCounts[name] = 0; return; }

      var values = sheet.getDataRange().getValues();
      var headerRowIdx = findHeaderRow(values);
      var headers = values[headerRowIdx];

      var dataRows = values.slice(headerRowIdx + 1).filter(function(r) {
        return r.some(function(c) { return c !== null && c !== ''; });
      });

      sheetCounts[name] = dataRows.length;
      totalCases += dataRows.length;

      var statusCol = -1;
      headers.forEach(function(h, i) {
        if (h === 'DURUM' || h === 'DOSYA DURUMU') statusCol = i;
      });

      if (statusCol >= 0) {
        dataRows.forEach(function(row) {
          var s = String(row[statusCol] || '').toUpperCase();
          if (!s) return;
          if (s.includes('İCRADA') || s.includes('TAMAMLAND') || s.includes('ALINDI') ||
              s.includes('ÖDENDİ') || s.includes('KAPANDI')) {
            completedCount++;
          } else {
            activeCount++;
          }
        });
      }
    });

    return {
      success: true,
      totalCases: totalCases,
      activeCount: activeCount,
      completedCount: completedCount,
      sheetCounts: sheetCounts
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── Global Search ─────────────────────────────────

function searchAll(query) {
  try {
    var ss = getSpreadsheet();
    var results = [];
    var q = query.toLowerCase().trim();
    if (!q) return { success: true, results: [] };

    ALL_SHEETS.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      var values = sheet.getDataRange().getValues();
      var headerRowIdx = findHeaderRow(values);
      var headers = values[headerRowIdx];

      for (var i = headerRowIdx + 1; i < values.length; i++) {
        var row = values[i];
        var matches = row.some(function(cell) {
          return cell !== null && cell !== '' &&
                 String(cell).toLowerCase().includes(q);
        });

        if (matches) {
          var obj = { _sheet: sheetName, _rowIndex: i + 1 };
          headers.forEach(function(h, j) {
            if (h && !h.startsWith('Sütun')) {
              obj[h] = formatCellValue(row[j]);
            }
          });
          results.push(obj);
        }
      }
    });

    return { success: true, results: results };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── Sheet Initializer ─────────────────────────────

function initializeSheets() {
  try {
    var ss = getSpreadsheet();
    var created = [];
    var existing = [];

    Object.keys(SHEET_HEADERS).forEach(function(name) {
      var headers = SHEET_HEADERS[name];
      var sheet = ss.getSheetByName(name);

      if (!sheet) {
        sheet = ss.insertSheet(name);
        created.push(name);
      } else {
        existing.push(name);
      }

      // Only write headers if first row is empty
      var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      if (firstRow.every(function(c) { return c === '' || c === null; })) {
        var headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setValues([headers]);
        headerRange.setBackground('#1a3c5e');
        headerRange.setFontColor('#ffffff');
        headerRange.setFontWeight('bold');
        headerRange.setFontSize(10);
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 50);

        // Auto-resize columns 2+
        for (var c = 2; c <= headers.length; c++) {
          sheet.setColumnWidth(c, 140);
        }
      }
    });

    var msg = '';
    if (created.length) msg += 'Oluşturuldu: ' + created.join(', ') + '. ';
    if (existing.length) msg += 'Mevcut: ' + existing.join(', ') + '.';
    return { success: true, message: msg || 'Tüm sayfalar hazır.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── Bulk Import from another sheet (optional helper) ─

function getSheetNames() {
  try {
    var ss = getSpreadsheet();
    var names = ss.getSheets().map(function(s) { return s.getName(); });
    return { success: true, names: names };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
