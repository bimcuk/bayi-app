function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Catatan Si Kecil 🍼')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================
   HELPER
========================================================= */

function getOrCreateSheet(sheetName, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }

  return sheet;
}

/* =========================================================
   1. PROFIL BAYI
========================================================= */

function simpanProfilData(profilObj) {
  var sheet = getOrCreateSheet(
    'Profil Bayi',
    ['Nama Bayi', 'Tanggal Lahir', 'Avatar Data']
  );

  var lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  sheet.getRange(2, 1, 1, 3).setValues([[
    profilObj.babyName || 'Adek Bayi',
    profilObj.birthDate || '',
    profilObj.avatarUrl || ''
  ]]);

  return "OK";
}

function ambilProfilData() {
  var sheet = getOrCreateSheet(
    'Profil Bayi',
    ['Nama Bayi', 'Tanggal Lahir', 'Avatar Data']
  );

  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return {
      babyName: "Adek Bayi",
      birthDate: "",
      avatarUrl: ""
    };
  }

  var row = sheet.getRange(2, 1, 1, Math.max(sheet.getLastColumn(), 3)).getValues()[0];
  
  var babyNameVal = row[0];
  var birthVal = row[1];
  var avatarVal = row[2];

  if (row.length >= 4 && (row[0] === 'Catatan Si Kecil' || row[0] === 'App Title' || row[0].toString().length > 30)) {
    babyNameVal = row[1] || 'Adek Bayi';
    birthVal = row[2];
    avatarVal = row[3];
  }

  if (birthVal instanceof Date) {
    birthVal = Utilities.formatDate(
      birthVal,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
  }

  return {
    babyName: babyNameVal || "Adek Bayi",
    birthDate: birthVal || "",
    avatarUrl: avatarVal || ""
  };
}

/* =========================================================
   2. ASIP + SUFOR (Riwayat Minum)
========================================================= */

function simpanSemuaData(dataArray) {
  var sheet = getOrCreateSheet(
    'Riwayat Minum',
    [
      'Waktu',
      'Jumlah (ml)',
      'Jenis Susu',
      'Catatan',
      'Jam Selesai',
      'Posisi Payudara',
      'Durasi (Menit)'
    ]
  );

  var lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }

  var filteredData = (dataArray || []).filter(function(item) {
    return (
      item &&
      item.jenis &&
      item.jenis !== 'Direct Breastfeeding (DBF)'
    );
  });

  if (filteredData.length > 0) {
    filteredData.sort(function(a, b) {
      var timeA = new Date(String(a.waktu).replace(' ', 'T')).getTime();
      var timeB = new Date(String(b.waktu).replace(' ', 'T')).getTime();
      return timeB - timeA;
    });

    var rows = filteredData.map(function(item) {
      return [
        item.waktu || '',
        item.jumlah || '',
        item.jenis || '',
        item.catatan || '',
        '',
        '',
        ''
      ];
    });

    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }

  return "OK";
}

function ambilData() {
  var sheet = getOrCreateSheet(
    'Riwayat Minum',
    [
      'Waktu',
      'Jumlah (ml)',
      'Jenis Susu',
      'Catatan',
      'Jam Selesai',
      'Posisi Payudara',
      'Durasi (Menit)'
    ]
  );

  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var result = [];

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][2] || '') === 'Direct Breastfeeding (DBF)') {
      continue;
    }

    var waktuVal = values[i][0];
    if (waktuVal instanceof Date) {
      waktuVal = Utilities.formatDate(
        waktuVal,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm"
      );
    }

    result.push({
      waktu: waktuVal,
      jumlah: values[i][1],
      jenis: values[i][2],
      catatan: values[i][3],
      jamSelesai: values[i][4] || '',
      payudara: values[i][5] || '',
      durasiMinut: values[i][6] || ''
    });
  }

  result.sort(function(a, b) {
    var timeA = new Date(String(a.waktu).replace(' ', 'T')).getTime();
    var timeB = new Date(String(b.waktu).replace(' ', 'T')).getTime();
    return timeB - timeA;
  });

  return result;
}

/* =========================================================
   3. DBF (Riwayat DBF)
========================================================= */

function getOrCreateDBFSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Riwayat DBF');

  if (!sheet) {
    sheet = ss.insertSheet('Riwayat DBF');
    sheet.appendRow([
      'Tanggal',
      'Jam Mulai',
      'Jam Selesai',
      'Durasi (Menit)',
      'Posisi Payudara',
      'Catatan'
    ]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  }

  return sheet;
}

function simpanDataDBF(dataArray) {
  var sheet = getOrCreateDBFSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }

  if (!dataArray || dataArray.length === 0) {
    return "OK";
  }

  var cleanData = dataArray.filter(function(item) {
    return item && item.tanggal;
  });

  cleanData.sort(function(a, b) {
    var timeA = new Date(String(a.tanggal) + 'T' + String(a.jamMulai || '00:00')).getTime();
    var timeB = new Date(String(b.tanggal) + 'T' + String(b.jamMulai || '00:00')).getTime();
    return timeB - timeA;
  });

  var rows = cleanData.map(function(item) {
    return [
      item.tanggal || '',
      item.jamMulai || '',
      item.jamSelesai || '',
      item.durasiMinut || '',
      item.payudara || '',
      item.catatan || ''
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  return "OK";
}

function ambilDataDBF() {
  var sheet = getOrCreateDBFSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var result = [];

  for (var i = 0; i < values.length; i++) {
    var tanggal = values[i][0];
    if (tanggal instanceof Date) {
      tanggal = Utilities.formatDate(tanggal, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    tanggal = String(tanggal || '');

    var jamMulai = values[i][1];
    if (jamMulai instanceof Date) {
      jamMulai = Utilities.formatDate(jamMulai, Session.getScriptTimeZone(), "HH:mm");
    }
    jamMulai = String(jamMulai || '');

    var jamSelesai = values[i][2];
    if (jamSelesai instanceof Date) {
      jamSelesai = Utilities.formatDate(jamSelesai, Session.getScriptTimeZone(), "HH:mm");
    }
    jamSelesai = String(jamSelesai || '');

    result.push({
      tanggal: tanggal,
      jamMulai: jamMulai,
      jamSelesai: jamSelesai,
      durasiMinut: values[i][3] || '',
      payudara: values[i][4] || '',
      catatan: values[i][5] || '',
      waktu: tanggal + (jamMulai ? ' ' + jamMulai : '')
    });
  }

  result.sort(function(a, b) {
    var timeA = new Date(String(a.waktu).replace(' ', 'T')).getTime();
    var timeB = new Date(String(b.waktu).replace(' ', 'T')).getTime();
    return timeB - timeA;
  });

  return result;
}

/* =========================================================
   4. PERTUMBUHAN FISIK (BB, PB, Lingkar Kepala)
========================================================= */

function getOrCreateFisikSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pertumbuhan Fisik');

  if (!sheet) {
    sheet = ss.insertSheet('Pertumbuhan Fisik');
    sheet.appendRow([
      'Tanggal',
      'Berat Badan (kg)',
      'Panjang Badan (cm)',
      'Lingkar Kepala (cm)'
    ]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    return sheet;
  }

  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 3)).getValues()[0];
  var headerLK = 'Lingkar Kepala (cm)';

  if (headers.length < 4 || String(headers[3] || '') !== headerLK) {
    sheet.getRange(1, 4).setValue(headerLK);
  }

  return sheet;
}

function simpanDataFisik(dataArray) {
  var sheet = getOrCreateFisikSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  if (dataArray && dataArray.length > 0) {
    dataArray.sort(function(a, b) {
      return new Date(b.tgl).getTime() - new Date(a.tgl).getTime();
    });

    var rows = dataArray.map(function(item) {
      return [
        item.tgl || '',
        item.berat !== undefined ? item.berat : '',
        item.panjang !== undefined ? item.panjang : '',
        item.lingkarKepala !== undefined ? item.lingkarKepala : ''
      ];
    });

    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  return "OK";
}

function ambilDataFisik() {
  var sheet = getOrCreateFisikSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var result = [];

  for (var i = 0; i < values.length; i++) {
    var tglVal = values[i][0];
    if (tglVal instanceof Date) {
      tglVal = Utilities.formatDate(tglVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    result.push({
      tgl: tglVal,
      berat: values[i][1],
      panjang: values[i][2],
      lingkarKepala: values[i][3]
    });
  }

  result.sort(function(a, b) {
    return new Date(b.tgl).getTime() - new Date(a.tgl).getTime();
  });

  return result;
}

/* =========================================================
   5. POPOK
========================================================= */

function simpanDataPopok(dataArray) {
  var sheet = getOrCreateSheet(
    'Riwayat Popok',
    [
      'Waktu',
      'Jenis',
      'Urutan Popok',
      'Keterangan Pup/Pipis',
      'Catatan'
    ]
  );

  var lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }

  if (dataArray && dataArray.length > 0) {
    dataArray.sort(function(a, b) {
      var timeA = new Date(String(a.waktu).replace(' ', 'T')).getTime();
      var timeB = new Date(String(b.waktu).replace(' ', 'T')).getTime();
      return timeB - timeA;
    });

    var rows = dataArray.map(function(item) {
      return [
        item.waktu,
        item.jenis,
        item.urutan || 1,
        item.detail || '',
        item.catatan || ''
      ];
    });

    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  return "OK";
}

function ambilDataPopok() {
  var sheet = getOrCreateSheet(
    'Riwayat Popok',
    [
      'Waktu',
      'Jenis',
      'Urutan Popok',
      'Keterangan Pup/Pipis',
      'Catatan'
    ]
  );

  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var result = [];

  for (var i = 0; i < values.length; i++) {
    var waktuVal = values[i][0];
    if (waktuVal instanceof Date) {
      waktuVal = Utilities.formatDate(waktuVal, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    }

    result.push({
      waktu: waktuVal,
      jenis: values[i][1],
      urutan: values[i][2] || 1,
      detail: values[i][3] || '',
      catatan: values[i][4] || ''
    });
  }

  result.sort(function(a, b) {
    var timeA = new Date(String(a.waktu).replace(' ', 'T')).getTime();
    var timeB = new Date(String(b.waktu).replace(' ', 'T')).getTime();
    return timeB - timeA;
  });

  return result;
}