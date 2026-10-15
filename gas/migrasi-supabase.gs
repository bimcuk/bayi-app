/*******************************************************
 * MIGRASI GOOGLE SHEETS → SUPABASE
 * Tumbangkan Adek
 *
 * PENTING:
 * - Tidak menghapus data Spreadsheet
 * - Service Role Key disimpan di Script Properties
 * - Migrasi ditujukan ke 1 user Supabase
 *******************************************************/


/* =====================================================
   KONFIGURASI
===================================================== */

const SUPABASE_URL =
  'https://evvqxxsojeiisrsuygky.supabase.co';

const TARGET_USER_ID =
  'b5309355-c88c-4600-b9b6-8e0c622b56dc';


/* =====================================================
   AMBIL SERVICE ROLE / SECRET KEY
===================================================== */

function getSupabaseKey() {

  const key =
    PropertiesService
      .getScriptProperties()
      .getProperty('SUPABASE_SERVICE_ROLE_KEY');

  if (!key) {
    throw new Error(
      'SUPABASE_SERVICE_ROLE_KEY belum disimpan di Script Properties.'
    );
  }

  return key;
}


/* =====================================================
   HEADER SUPABASE REQUEST
===================================================== */

function supabaseHeaders() {

  const key = getSupabaseKey();

  return {
    'apikey': key,
    'Authorization': 'Bearer ' + key,
    'Content-Type': 'application/json',
    'Prefer': 'return=representation'
  };
}


/* =====================================================
   REQUEST SUPABASE
===================================================== */

function supabaseRequest(method, table, payload, query) {

  let url =
    SUPABASE_URL +
    '/rest/v1/' +
    table;

  if (query) {
    url += '?' + query;
  }

  const options = {
    method: method,
    headers: supabaseHeaders(),
    muteHttpExceptions: true
  };

  if (payload !== undefined && payload !== null) {
    options.payload =
      JSON.stringify(payload);
  }

  const response =
    UrlFetchApp.fetch(url, options);

  const code =
    response.getResponseCode();

  const text =
    response.getContentText();

  if (code < 200 || code >= 300) {

    throw new Error(
      'Supabase error ' +
      code +
      ':\n' +
      text
    );
  }

  if (!text) {
    return null;
  }

  return JSON.parse(text);
}


/* =====================================================
   UTILITAS
===================================================== */

function cleanValue(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return null;
  }

  if (
    typeof value === 'string' &&
    value.trim() === ''
  ) {
    return null;
  }

  return value;
}


function formatDate(value) {

  if (!value) {
    return null;
  }

  if (value instanceof Date) {

    return Utilities.formatDate(
      value,
      'Asia/Jakarta',
      'yyyy-MM-dd'
    );
  }

  return String(value).trim();
}


function formatDateTime(value) {

  if (!value) {
    return null;
  }

  if (value instanceof Date) {

    return Utilities.formatDate(
      value,
      'Asia/Jakarta',
      "yyyy-MM-dd'T'HH:mm:ssXXX"
    );
  }

  let text =
    String(value).trim();

  text =
    text.replace(' ', 'T');

  return text;
}


function formatTime(value) {

  if (!value) {
    return null;
  }

  if (value instanceof Date) {

    return Utilities.formatDate(
      value,
      'Asia/Jakarta',
      'HH:mm:ss'
    );
  }

  let text =
    String(value).trim();

  if (/^\d{1,2}:\d{2}$/.test(text)) {
    return text + ':00';
  }

  return text;
}


function numberValue(value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return null;
  }

  if (typeof value === 'number') {
    return value;
  }

  let text =
    String(value)
      .trim()
      .replace(',', '.');

  const number =
    Number(text);

  return isNaN(number)
    ? null
    : number;
}


/* =====================================================
   CARI HEADER
===================================================== */

function getHeaderMap(sheet) {

  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {
    throw new Error(
      'Sheet "' +
      sheet.getName() +
      '" tidak memiliki kolom.'
    );
  }

  const headers =
    sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0];

  const map = {};

  headers.forEach(function(header, index) {

    const name =
      String(header || '')
        .trim();

    if (name) {
      map[name] = index;
    }

  });

  return map;
}


function getCell(row, headerMap, headerName) {

  if (
    headerMap[headerName] === undefined
  ) {
    return '';
  }

  return row[
    headerMap[headerName]
  ];
}


/* =====================================================
   HITUNG DATA SHEET
===================================================== */

function getDataRows(sheet) {

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {
    return [];
  }

  return sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      lastColumn
    )
    .getValues();
}


/* =====================================================
   HAPUS DATA TEST SUPABASE USER INI
 *
 * HANYA DATA DI SUPABASE.
 * TIDAK MENYENTUH GOOGLE SHEETS.
===================================================== */

function resetSupabaseUserData() {

  const tables = [
    'milk_records',
    'dbf_records',
    'diaper_records',
    'growth_records',
    'profiles'
  ];

  tables.forEach(function(table) {

    supabaseRequest(
      'delete',
      table,
      null,
      'user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID)
    );

  });

  Logger.log(
    'Data Supabase untuk user berhasil dibersihkan.'
  );
}


/* =====================================================
   MIGRASI PROFIL
===================================================== */

function migrasiProfil() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Profil Bayi');

  if (!sheet) {
    throw new Error(
      'Sheet "Profil Bayi" tidak ditemukan.'
    );
  }

  const headers =
    getHeaderMap(sheet);

  const rows =
    getDataRows(sheet);

  if (rows.length === 0) {

    Logger.log(
      'Profil Bayi kosong.'
    );

    return 0;
  }

  const row =
    rows[0];

  const babyName =
    cleanValue(
      getCell(
        row,
        headers,
        'Nama Bayi'
      )
    ) || 'Adek Bayi';

  const birthDate =
    formatDate(
      getCell(
        row,
        headers,
        'Tanggal Lahir'
      )
    ); 

  const payload = {

    user_id:
      TARGET_USER_ID,

    baby_name:
      babyName,

    birth_date:
      birthDate,

    avatar_url:
      null

  };

  const result =
    supabaseRequest(
      'post',
      'profiles',
      payload,
      ''
    );

  Logger.log(
    'Profil berhasil dimigrasikan.'
  );

  return result
    ? result.length
    : 1;
}


/* =====================================================
   MIGRASI RIWAYAT MINUM
===================================================== */

function migrasiMinum() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat Minum');

  if (!sheet) {
    throw new Error(
      'Sheet "Riwayat Minum" tidak ditemukan.'
    );
  }

  const headers =
    getHeaderMap(sheet);

  const rows =
    getDataRows(sheet);

  if (rows.length === 0) {
    return 0;
  }

  const records = [];

  rows.forEach(function(row) {

    const waktu =
      formatDateTime(
        getCell(
          row,
          headers,
          'Waktu'
        )
      );

    if (!waktu) {
      return;
    }

    records.push({

      user_id:
        TARGET_USER_ID,

      waktu:
        waktu,

      jumlah:
        numberValue(
          getCell(
            row,
            headers,
            'Jumlah (ml)'
          )
        ),

      jenis:
        cleanValue(
          getCell(
            row,
            headers,
            'Jenis Susu'
          )
        ),

      catatan:
        cleanValue(
          getCell(
            row,
            headers,
            'Catatan'
          )
        )

    });

  });

  if (records.length === 0) {
    return 0;
  }

  supabaseRequest(
    'post',
    'milk_records',
    records,
    ''
  );

  Logger.log(
    'Riwayat Minum: ' +
    records.length +
    ' data.'
  );

  return records.length;
}


/* =====================================================
   MIGRASI DBF
===================================================== */

function migrasiDBF() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat DBF');

  if (!sheet) {
    throw new Error(
      'Sheet "Riwayat DBF" tidak ditemukan.'
    );
  }

  const headers =
    getHeaderMap(sheet);

  const rows =
    getDataRows(sheet);

  if (rows.length === 0) {
    return 0;
  }

  const records = [];

  rows.forEach(function(row) {

    const tanggal =
      formatDate(
        getCell(
          row,
          headers,
          'Tanggal'
        )
      );

    if (!tanggal) {
      return;
    }

    records.push({

      user_id:
        TARGET_USER_ID,

      tanggal:
        tanggal,

      jam_mulai:
        formatTime(
          getCell(
            row,
            headers,
            'Jam Mulai'
          )
        ),

      jam_selesai:
        formatTime(
          getCell(
            row,
            headers,
            'Jam Selesai'
          )
        ),

      durasi_menit:
        numberValue(
          getCell(
            row,
            headers,
            'Durasi (Menit)'
          )
        ),

      payudara:
        cleanValue(
          getCell(
            row,
            headers,
            'Posisi Payudara'
          )
        ),

      catatan:
        cleanValue(
          getCell(
            row,
            headers,
            'Catatan'
          )
        )

    });

  });

  if (records.length === 0) {
    return 0;
  }

  supabaseRequest(
    'post',
    'dbf_records',
    records,
    ''
  );

  Logger.log(
    'Riwayat DBF: ' +
    records.length +
    ' data.'
  );

  return records.length;
}


/* =====================================================
   MIGRASI POPOK
===================================================== */

function migrasiPopok() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat Popok');

  if (!sheet) {
    throw new Error(
      'Sheet "Riwayat Popok" tidak ditemukan.'
    );
  }

  const headers =
    getHeaderMap(sheet);

  const rows =
    getDataRows(sheet);

  if (rows.length === 0) {
    return 0;
  }

  const records = [];

  rows.forEach(function(row) {

    const waktu =
      formatDateTime(
        getCell(
          row,
          headers,
          'Waktu'
        )
      );

    if (!waktu) {
      return;
    }

    let urutan =
      numberValue(
        getCell(
          row,
          headers,
          'Urutan Popok'
        )
      );

    if (urutan === null) {
      urutan = 1;
    }

    records.push({

      user_id:
        TARGET_USER_ID,

      waktu:
        waktu,

      jenis:
        cleanValue(
          getCell(
            row,
            headers,
            'Jenis'
          )
        ),

      urutan:
        urutan,

      detail:
        cleanValue(
          getCell(
            row,
            headers,
            'Keterangan Pup/Pipis'
          )
        ),

      catatan:
        cleanValue(
          getCell(
            row,
            headers,
            'Catatan'
          )
        )

    });

  });

  if (records.length === 0) {
    return 0;
  }

  supabaseRequest(
    'post',
    'diaper_records',
    records,
    ''
  );

  Logger.log(
    'Riwayat Popok: ' +
    records.length +
    ' data.'
  );

  return records.length;
}


/* =====================================================
   MIGRASI PERTUMBUHAN FISIK
===================================================== */

function migrasiFisik() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Pertumbuhan Fisik');

  if (!sheet) {
    throw new Error(
      'Sheet "Pertumbuhan Fisik" tidak ditemukan.'
    );
  }

  const headers =
    getHeaderMap(sheet);

  const rows =
    getDataRows(sheet);

  if (rows.length === 0) {
    return 0;
  }

  const records = [];

  rows.forEach(function(row) {

    const tgl =
      formatDate(
        getCell(
          row,
          headers,
          'Tanggal'
        )
      );

    if (!tgl) {
      return;
    }

    records.push({

      user_id:
        TARGET_USER_ID,

      tgl:
        tgl,

      berat:
        numberValue(
          getCell(
            row,
            headers,
            'Berat Badan (kg)'
          )
        ),

      panjang:
        numberValue(
          getCell(
            row,
            headers,
            'Panjang Badan (cm)'
          )
        ),

      lingkar_kepala:
        numberValue(
          getCell(
            row,
            headers,
            'Lingkar Kepala (cm)'
          )
        )

    });

  });

  if (records.length === 0) {
    return 0;
  }

  supabaseRequest(
    'post',
    'growth_records',
    records,
    ''
  );

  Logger.log(
    'Pertumbuhan Fisik: ' +
    records.length +
    ' data.'
  );

  return records.length;
}


/* =====================================================
   MIGRASI SEMUA
===================================================== */

function migrasiSemuaDataKeSupabase() {

  Logger.log(
    '========================================'
  );

  Logger.log(
    'MULAI MIGRASI DATA'
  );

  Logger.log(
    '========================================'
  );


  // Bersihkan data percobaan / hasil migrasi
  // untuk user target di SUPABASE.
  //
  // GOOGLE SHEETS TIDAK DISENTUH.

  resetSupabaseUserData();


  const hasil = {};


  hasil.profil =
    migrasiProfil();

  hasil.minum =
    migrasiMinum();

  hasil.dbf =
    migrasiDBF();

  hasil.popok =
    migrasiPopok();

  hasil.fisik =
    migrasiFisik();


  Logger.log(
    '========================================'
  );

  Logger.log(
    'MIGRASI SELESAI'
  );

  Logger.log(
    JSON.stringify(
      hasil,
      null,
      2
    )
  );

  Logger.log(
    '========================================'
  );


  return hasil;
}

function tesKoneksiSupabase() {

  const result =
    supabaseRequest(
      'get',
      'profiles',
      null,
      'select=id,user_id,baby_name,birth_date&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID)
    );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}

function validasiMigrasiSupabase() {

  Logger.log('========================================');
  Logger.log('VALIDASI MIGRASI');
  Logger.log('========================================');


  /* =====================================================
     PROFIL
  ===================================================== */

  const profilSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Profil Bayi');

  const profilRows =
    getDataRows(profilSheet);

  const profilHeaders =
    getHeaderMap(profilSheet);

  const profilLocal =
    profilRows.length > 0
      ? {
          nama: getCell(
            profilRows[0],
            profilHeaders,
            'Nama Bayi'
          ),
          lahir: formatDate(
            getCell(
              profilRows[0],
              profilHeaders,
              'Tanggal Lahir'
            )
          )
        }
      : null;

  const profilSupabase =
    supabaseRequest(
      'get',
      'profiles',
      null,
      'select=baby_name,birth_date&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID)
    );

  Logger.log('--- PROFIL ---');
  Logger.log(
    'Spreadsheet: ' +
    JSON.stringify(profilLocal)
  );
  Logger.log(
    'Supabase: ' +
    JSON.stringify(profilSupabase)
  );


  /* =====================================================
     RIWAYAT MINUM
  ===================================================== */

  const minumSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat Minum');

  const minumRows =
    getDataRows(minumSheet);

  const minumHeaders =
    getHeaderMap(minumSheet);

  const minumLocal = [];

  [0, minumRows.length - 1].forEach(function(index) {

    if (
      index < 0 ||
      !minumRows[index]
    ) {
      return;
    }

    const row =
      minumRows[index];

    minumLocal.push({

      waktu:
        formatDateTime(
          getCell(
            row,
            minumHeaders,
            'Waktu'
          )
        ),

      jumlah:
        numberValue(
          getCell(
            row,
            minumHeaders,
            'Jumlah (ml)'
          )
        ),

      jenis:
        cleanValue(
          getCell(
            row,
            minumHeaders,
            'Jenis Susu'
          )
        ),

      catatan:
        cleanValue(
          getCell(
            row,
            minumHeaders,
            'Catatan'
          )
        )

    });

  });

  const minumSupabase =
    supabaseRequest(
      'get',
      'milk_records',
      null,
      'select=waktu,jumlah,jenis,catatan&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID) +
      '&order=waktu.desc'
    );

  Logger.log('--- RIWAYAT MINUM ---');
  Logger.log(
    'Spreadsheet sample: ' +
    JSON.stringify(minumLocal)
  );

  Logger.log(
    'Supabase pertama: ' +
    JSON.stringify(
      minumSupabase[0]
    )
  );

  Logger.log(
    'Supabase terakhir: ' +
    JSON.stringify(
      minumSupabase[minumSupabase.length - 1]
    )
  );


  /* =====================================================
     DBF
  ===================================================== */

  const dbfSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat DBF');

  const dbfRows =
    getDataRows(dbfSheet);

  const dbfHeaders =
    getHeaderMap(dbfSheet);

  const dbfLocal = [];

  [0, dbfRows.length - 1].forEach(function(index) {

    if (
      index < 0 ||
      !dbfRows[index]
    ) {
      return;
    }

    const row =
      dbfRows[index];

    dbfLocal.push({

      tanggal:
        formatDate(
          getCell(
            row,
            dbfHeaders,
            'Tanggal'
          )
        ),

      jam_mulai:
        formatTime(
          getCell(
            row,
            dbfHeaders,
            'Jam Mulai'
          )
        ),

      jam_selesai:
        formatTime(
          getCell(
            row,
            dbfHeaders,
            'Jam Selesai'
          )
        ),

      durasi_menit:
        numberValue(
          getCell(
            row,
            dbfHeaders,
            'Durasi (Menit)'
          )
        ),

      payudara:
        cleanValue(
          getCell(
            row,
            dbfHeaders,
            'Posisi Payudara'
          )
        ),

      catatan:
        cleanValue(
          getCell(
            row,
            dbfHeaders,
            'Catatan'
          )
        )

    });

  });

  const dbfSupabase =
    supabaseRequest(
      'get',
      'dbf_records',
      null,
      'select=tanggal,jam_mulai,jam_selesai,durasi_menit,payudara,catatan&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID) +
      '&order=tanggal.desc,jam_mulai.desc'
    );

  Logger.log('--- RIWAYAT DBF ---');

  Logger.log(
    'Spreadsheet sample: ' +
    JSON.stringify(dbfLocal)
  );

  Logger.log(
    'Supabase pertama: ' +
    JSON.stringify(
      dbfSupabase[0]
    )
  );

  Logger.log(
    'Supabase terakhir: ' +
    JSON.stringify(
      dbfSupabase[dbfSupabase.length - 1]
    )
  );


  /* =====================================================
     POPOK
  ===================================================== */

  const popokSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat Popok');

  const popokRows =
    getDataRows(popokSheet);

  const popokHeaders =
    getHeaderMap(popokSheet);

  const popokLocal = [];

  [0, popokRows.length - 1].forEach(function(index) {

    if (
      index < 0 ||
      !popokRows[index]
    ) {
      return;
    }

    const row =
      popokRows[index];

    popokLocal.push({

      waktu:
        formatDateTime(
          getCell(
            row,
            popokHeaders,
            'Waktu'
          )
        ),

      jenis:
        cleanValue(
          getCell(
            row,
            popokHeaders,
            'Jenis'
          )
        ),

      urutan:
        numberValue(
          getCell(
            row,
            popokHeaders,
            'Urutan Popok'
          )
        ) || 1,

      detail:
        cleanValue(
          getCell(
            row,
            popokHeaders,
            'Keterangan Pup/Pipis'
          )
        ),

      catatan:
        cleanValue(
          getCell(
            row,
            popokHeaders,
            'Catatan'
          )
        )

    });

  });

  const popokSupabase =
    supabaseRequest(
      'get',
      'diaper_records',
      null,
      'select=waktu,jenis,urutan,detail,catatan&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID) +
      '&order=waktu.desc'
    );

  Logger.log('--- RIWAYAT POPOK ---');

  Logger.log(
    'Spreadsheet sample: ' +
    JSON.stringify(popokLocal)
  );

  Logger.log(
    'Supabase pertama: ' +
    JSON.stringify(
      popokSupabase[0]
    )
  );

  Logger.log(
    'Supabase terakhir: ' +
    JSON.stringify(
      popokSupabase[popokSupabase.length - 1]
    )
  );


  /* =====================================================
     PERTUMBUHAN FISIK
  ===================================================== */

  const fisikSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Pertumbuhan Fisik');

  const fisikRows =
    getDataRows(fisikSheet);

  const fisikHeaders =
    getHeaderMap(fisikSheet);

  const fisikLocal = [];

  [0, fisikRows.length - 1].forEach(function(index) {

    if (
      index < 0 ||
      !fisikRows[index]
    ) {
      return;
    }

    const row =
      fisikRows[index];

    fisikLocal.push({

      tgl:
        formatDate(
          getCell(
            row,
            fisikHeaders,
            'Tanggal'
          )
        ),

      berat:
        numberValue(
          getCell(
            row,
            fisikHeaders,
            'Berat Badan (kg)'
          )
        ),

      panjang:
        numberValue(
          getCell(
            row,
            fisikHeaders,
            'Panjang Badan (cm)'
          )
        ),

      lingkar_kepala:
        numberValue(
          getCell(
            row,
            fisikHeaders,
            'Lingkar Kepala (cm)'
          )
        )

    });

  });

  const fisikSupabase =
    supabaseRequest(
      'get',
      'growth_records',
      null,
      'select=tgl,berat,panjang,lingkar_kepala&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID) +
      '&order=tgl.desc'
    );

  Logger.log('--- PERTUMBUHAN FISIK ---');

  Logger.log(
    'Spreadsheet sample: ' +
    JSON.stringify(fisikLocal)
  );

  Logger.log(
    'Supabase pertama: ' +
    JSON.stringify(
      fisikSupabase[0]
    )
  );

  Logger.log(
    'Supabase terakhir: ' +
    JSON.stringify(
      fisikSupabase[fisikSupabase.length - 1]
    )
  );


  Logger.log('========================================');
  Logger.log('VALIDASI SELESAI');
  Logger.log('========================================');
}

function validasiExactDBF() {

  Logger.log('========================================');
  Logger.log('VALIDASI EXACT DBF');
  Logger.log('========================================');


  // =====================================================
  // AMBIL DATA DARI SPREADSHEET
  // =====================================================

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat DBF');

  const headers =
    getHeaderMap(sheet);

  const rows =
    getDataRows(sheet);

  const spreadsheetData = [];

  rows.forEach(function(row) {

    const tanggal =
      formatDate(
        getCell(
          row,
          headers,
          'Tanggal'
        )
      );

    if (!tanggal) {
      return;
    }

    spreadsheetData.push({

      tanggal:
        tanggal,

      jam_mulai:
        formatTime(
          getCell(
            row,
            headers,
            'Jam Mulai'
          )
        ),

      jam_selesai:
        formatTime(
          getCell(
            row,
            headers,
            'Jam Selesai'
          )
        ),

      durasi_menit:
        numberValue(
          getCell(
            row,
            headers,
            'Durasi (Menit)'
          )
        ),

      payudara:
        cleanValue(
          getCell(
            row,
            headers,
            'Posisi Payudara'
          )
        ),

      catatan:
        cleanValue(
          getCell(
            row,
            headers,
            'Catatan'
          )
        )

    });

  });


  // =====================================================
  // AMBIL SEMUA DATA DARI SUPABASE
  // =====================================================

  const supabaseData =
    supabaseRequest(
      'get',
      'dbf_records',
      null,
      'select=tanggal,jam_mulai,jam_selesai,durasi_menit,payudara,catatan' +
      '&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID) +
      '&order=tanggal.asc,jam_mulai.asc'
    );


  Logger.log(
    'Jumlah Spreadsheet: ' +
    spreadsheetData.length
  );

  Logger.log(
    'Jumlah Supabase: ' +
    supabaseData.length
  );


  // =====================================================
  // BUAT KEY UNIK SETIAP DATA
  // =====================================================

  function makeKey(item) {

    return [
      item.tanggal || '',
      item.jam_mulai || '',
      item.jam_selesai || '',
      item.durasi_menit === null
        ? ''
        : item.durasi_menit,
      item.payudara || '',
      item.catatan || ''
    ].join('|');

  }


  const spreadsheetKeys = {};

  spreadsheetData.forEach(function(item) {

    const key =
      makeKey(item);

    spreadsheetKeys[key] =
      (spreadsheetKeys[key] || 0) + 1;

  });


  const supabaseKeys = {};

  supabaseData.forEach(function(item) {

    const normalized = {

      tanggal:
        item.tanggal || '',

      jam_mulai:
        item.jam_mulai
          ? String(item.jam_mulai)
          : '',

      jam_selesai:
        item.jam_selesai
          ? String(item.jam_selesai)
          : '',

      durasi_menit:
        item.durasi_menit === null
          ? null
          : Number(item.durasi_menit),

      payudara:
        item.payudara || '',

      catatan:
        item.catatan || ''

    };

    const key =
      makeKey(normalized);

    supabaseKeys[key] =
      (supabaseKeys[key] || 0) + 1;

  });


  // =====================================================
  // CARI DATA YANG HILANG / BERLEBIH
  // =====================================================

  const missingFromSupabase = [];
  const extraInSupabase = [];


  Object.keys(spreadsheetKeys)
    .forEach(function(key) {

      const localCount =
        spreadsheetKeys[key];

      const remoteCount =
        supabaseKeys[key] || 0;

      if (localCount > remoteCount) {

        for (
          let i = 0;
          i < localCount - remoteCount;
          i++
        ) {
          missingFromSupabase.push(key);
        }

      }

    });


  Object.keys(supabaseKeys)
    .forEach(function(key) {

      const remoteCount =
        supabaseKeys[key];

      const localCount =
        spreadsheetKeys[key] || 0;

      if (remoteCount > localCount) {

        for (
          let i = 0;
          i < remoteCount - localCount;
          i++
        ) {
          extraInSupabase.push(key);
        }

      }

    });


  // =====================================================
  // HASIL
  // =====================================================

  Logger.log('----------------------------------------');

  Logger.log(
    'Data hilang di Supabase: ' +
    missingFromSupabase.length
  );

  if (missingFromSupabase.length > 0) {

    Logger.log(
      JSON.stringify(
        missingFromSupabase,
        null,
        2
      )
    );

  }


  Logger.log(
    'Data tambahan di Supabase: ' +
    extraInSupabase.length
  );

  if (extraInSupabase.length > 0) {

    Logger.log(
      JSON.stringify(
        extraInSupabase,
        null,
        2
      )
    );

  }


  Logger.log('----------------------------------------');


  if (
    missingFromSupabase.length === 0 &&
    extraInSupabase.length === 0
  ) {

    Logger.log(
      '✅ DBF 100% IDENTIK'
    );

  } else {

    Logger.log(
      '⚠️ ADA PERBEDAAN DATA DBF'
    );

  }


  Logger.log('========================================');
  Logger.log('VALIDASI EXACT DBF SELESAI');
  Logger.log('========================================');
}

function perbaiki1DBFHilang() {

  const payload = {

    user_id:
      TARGET_USER_ID,

    tanggal:
      '2026-09-24',

    jam_mulai:
      '00:19:08',

    jam_selesai:
      '01:19:08',

    durasi_menit:
      60,

    payudara:
      'Kiri',

    catatan:
      null

  };

  const result =
    supabaseRequest(
      'post',
      'dbf_records',
      payload,
      ''
    );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}

/* =====================================================
   VALIDASI EXACT SEMUA DATA
 *
 * HANYA MEMBACA DATA.
 * TIDAK MENGUBAH SPREADSHEET.
 * TIDAK MENGUBAH SUPABASE.
===================================================== */

function validasiExactSemuaData() {

  Logger.log('========================================');
  Logger.log('VALIDASI EXACT SEMUA DATA');
  Logger.log('========================================');


  const hasil = {};


  /* =====================================================
     HELPER
  ===================================================== */

  function makeKey(values) {

    return values.map(function(value) {

      if (
        value === null ||
        value === undefined
      ) {
        return '';
      }

      return String(value);

    }).join('|');

  }


  function hitungKey(data) {

    const map = {};

    data.forEach(function(item) {

      const key = makeKey(item);

      map[key] =
        (map[key] || 0) + 1;

    });

    return map;

  }


  function bandingkan(localData, remoteData) {

    const localKeys =
      hitungKey(localData);

    const remoteKeys =
      hitungKey(remoteData);

    const missing = [];
    const extra = [];

    Object.keys(localKeys)
      .forEach(function(key) {

        const localCount =
          localKeys[key];

        const remoteCount =
          remoteKeys[key] || 0;

        if (localCount > remoteCount) {

          for (
            let i = 0;
            i < localCount - remoteCount;
            i++
          ) {
            missing.push(key);
          }

        }

      });


    Object.keys(remoteKeys)
      .forEach(function(key) {

        const remoteCount =
          remoteKeys[key];

        const localCount =
          localKeys[key] || 0;

        if (remoteCount > localCount) {

          for (
            let i = 0;
            i < remoteCount - localCount;
            i++
          ) {
            extra.push(key);
          }

        }

      });


    return {
      localCount: localData.length,
      remoteCount: remoteData.length,
      missing: missing,
      extra: extra
    };

  }


  function logHasil(nama, result) {

    Logger.log('');
    Logger.log('----------------------------------------');
    Logger.log(nama);
    Logger.log('----------------------------------------');

    Logger.log(
      'Jumlah Spreadsheet: ' +
      result.localCount
    );

    Logger.log(
      'Jumlah Supabase: ' +
      result.remoteCount
    );

    Logger.log(
      'Data hilang di Supabase: ' +
      result.missing.length
    );

    if (result.missing.length > 0) {

      Logger.log(
        'MISSING:\n' +
        JSON.stringify(
          result.missing,
          null,
          2
        )
      );

    }

    Logger.log(
      'Data tambahan di Supabase: ' +
      result.extra.length
    );

    if (result.extra.length > 0) {

      Logger.log(
        'EXTRA:\n' +
        JSON.stringify(
          result.extra,
          null,
          2
        )
      );

    }

    if (
      result.missing.length === 0 &&
      result.extra.length === 0
    ) {

      Logger.log(
        '✅ ' +
        nama +
        ' 100% IDENTIK'
      );

    } else {

      Logger.log(
        '⚠️ ' +
        nama +
        ' ADA PERBEDAAN'
      );

    }

  }


  /* =====================================================
     1. PROFIL
  ===================================================== */

  const profilSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Profil Bayi');

  const profilHeaders =
    getHeaderMap(profilSheet);

  const profilRows =
    getDataRows(profilSheet);

  const profilLocal = [];

  if (profilRows.length > 0) {

    const row =
      profilRows[0];

    profilLocal.push([
      cleanValue(
        getCell(
          row,
          profilHeaders,
          'Nama Bayi'
        )
      ) || 'Adek Bayi',

      formatDate(
        getCell(
          row,
          profilHeaders,
          'Tanggal Lahir'
        )
      )
    ]);

  }


  const profilSupabase =
    supabaseRequest(
      'get',
      'profiles',
      null,
      'select=baby_name,birth_date' +
      '&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID)
    );


  const profilRemote =
    profilSupabase.map(function(item) {

      return [
        item.baby_name || 'Adek Bayi',
        item.birth_date || ''
      ];

    });


  hasil.profil =
    bandingkan(
      profilLocal,
      profilRemote
    );


  logHasil(
    'PROFIL',
    hasil.profil
  );


  /* =====================================================
     2. RIWAYAT MINUM
  ===================================================== */

  const minumSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat Minum');

  const minumHeaders =
    getHeaderMap(minumSheet);

  const minumRows =
    getDataRows(minumSheet);

  const minumLocal = [];

  minumRows.forEach(function(row) {

    const waktu =
      formatDateTime(
        getCell(
          row,
          minumHeaders,
          'Waktu'
        )
      );

    if (!waktu) {
      return;
    }

    minumLocal.push([

      waktu,

      numberValue(
        getCell(
          row,
          minumHeaders,
          'Jumlah (ml)'
        )
      ),

      cleanValue(
        getCell(
          row,
          minumHeaders,
          'Jenis Susu'
        )
      ) || '',

      cleanValue(
        getCell(
          row,
          minumHeaders,
          'Catatan'
        )
      ) || ''

    ]);

  });


  const minumSupabase =
    supabaseRequest(
      'get',
      'milk_records',
      null,
      'select=waktu,jumlah,jenis,catatan' +
      '&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID) +
      '&order=waktu.asc'
    );


  const minumRemote =
    minumSupabase.map(function(item) {

      let waktu =
        item.waktu || '';

      /*
       * Supabase mengembalikan timestamptz
       * dalam UTC.
       *
       * Kita ubah ke waktu Jakarta
       * supaya bisa dibandingkan dengan
       * Spreadsheet.
       */

      if (waktu) {

        const date =
          new Date(waktu);

        waktu =
          Utilities.formatDate(
            date,
            'Asia/Jakarta',
            "yyyy-MM-dd'T'HH:mm:ssXXX"
          );

      }


      return [

        waktu,

        item.jumlah === null
          ? null
          : Number(item.jumlah),

        item.jenis || '',

        item.catatan || ''

      ];

    });


  hasil.minum =
    bandingkan(
      minumLocal,
      minumRemote
    );


  logHasil(
    'RIWAYAT MINUM',
    hasil.minum
  );


  /* =====================================================
     3. DBF
  ===================================================== */

  const dbfSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat DBF');

  const dbfHeaders =
    getHeaderMap(dbfSheet);

  const dbfRows =
    getDataRows(dbfSheet);

  const dbfLocal = [];

  dbfRows.forEach(function(row) {

    const tanggal =
      formatDate(
        getCell(
          row,
          dbfHeaders,
          'Tanggal'
        )
      );

    if (!tanggal) {
      return;
    }

    dbfLocal.push([

      tanggal,

      formatTime(
        getCell(
          row,
          dbfHeaders,
          'Jam Mulai'
        )
      ) || '',

      formatTime(
        getCell(
          row,
          dbfHeaders,
          'Jam Selesai'
        )
      ) || '',

      numberValue(
        getCell(
          row,
          dbfHeaders,
          'Durasi (Menit)'
        )
      ),

      cleanValue(
        getCell(
          row,
          dbfHeaders,
          'Posisi Payudara'
        )
      ) || '',

      cleanValue(
        getCell(
          row,
          dbfHeaders,
          'Catatan'
        )
      ) || ''

    ]);

  });


  const dbfSupabase =
    supabaseRequest(
      'get',
      'dbf_records',
      null,
      'select=tanggal,jam_mulai,jam_selesai,durasi_menit,payudara,catatan' +
      '&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID) +
      '&order=tanggal.asc,jam_mulai.asc'
    );


  const dbfRemote =
    dbfSupabase.map(function(item) {

      return [

        item.tanggal || '',

        item.jam_mulai
          ? String(item.jam_mulai)
          : '',

        item.jam_selesai
          ? String(item.jam_selesai)
          : '',

        item.durasi_menit === null
          ? null
          : Number(item.durasi_menit),

        item.payudara || '',

        item.catatan || ''

      ];

    });


  hasil.dbf =
    bandingkan(
      dbfLocal,
      dbfRemote
    );


  logHasil(
    'DBF',
    hasil.dbf
  );


  /* =====================================================
     4. POPOK
  ===================================================== */

  const popokSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Riwayat Popok');

  const popokHeaders =
    getHeaderMap(popokSheet);

  const popokRows =
    getDataRows(popokSheet);

  const popokLocal = [];

  popokRows.forEach(function(row) {

    const waktu =
      formatDateTime(
        getCell(
          row,
          popokHeaders,
          'Waktu'
        )
      );

    if (!waktu) {
      return;
    }

    let urutan =
      numberValue(
        getCell(
          row,
          popokHeaders,
          'Urutan Popok'
        )
      );

    if (urutan === null) {
      urutan = 1;
    }


    popokLocal.push([

      waktu,

      cleanValue(
        getCell(
          row,
          popokHeaders,
          'Jenis'
        )
      ) || '',

      urutan,

      cleanValue(
        getCell(
          row,
          popokHeaders,
          'Keterangan Pup/Pipis'
        )
      ) || '',

      cleanValue(
        getCell(
          row,
          popokHeaders,
          'Catatan'
        )
      ) || ''

    ]);

  });


  const popokSupabase =
    supabaseRequest(
      'get',
      'diaper_records',
      null,
      'select=waktu,jenis,urutan,detail,catatan' +
      '&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID) +
      '&order=waktu.asc'
    );


  const popokRemote =
    popokSupabase.map(function(item) {

      let waktu =
        item.waktu || '';

      if (waktu) {

        const date =
          new Date(waktu);

        waktu =
          Utilities.formatDate(
            date,
            'Asia/Jakarta',
            "yyyy-MM-dd'T'HH:mm:ssXXX"
          );

      }


      return [

        waktu,

        item.jenis || '',

        item.urutan === null
          ? 1
          : Number(item.urutan),

        item.detail || '',

        item.catatan || ''

      ];

    });


  hasil.popok =
    bandingkan(
      popokLocal,
      popokRemote
    );


  logHasil(
    'RIWAYAT POPOK',
    hasil.popok
  );


  /* =====================================================
     5. PERTUMBUHAN FISIK
  ===================================================== */

  const fisikSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Pertumbuhan Fisik');

  const fisikHeaders =
    getHeaderMap(fisikSheet);

  const fisikRows =
    getDataRows(fisikSheet);

  const fisikLocal = [];

  fisikRows.forEach(function(row) {

    const tgl =
      formatDate(
        getCell(
          row,
          fisikHeaders,
          'Tanggal'
        )
      );

    if (!tgl) {
      return;
    }

    fisikLocal.push([

      tgl,

      numberValue(
        getCell(
          row,
          fisikHeaders,
          'Berat Badan (kg)'
        )
      ),

      numberValue(
        getCell(
          row,
          fisikHeaders,
          'Panjang Badan (cm)'
        )
      ),

      numberValue(
        getCell(
          row,
          fisikHeaders,
          'Lingkar Kepala (cm)'
        )
      )

    ]);

  });


  const fisikSupabase =
    supabaseRequest(
      'get',
      'growth_records',
      null,
      'select=tgl,berat,panjang,lingkar_kepala' +
      '&user_id=eq.' +
      encodeURIComponent(TARGET_USER_ID) +
      '&order=tgl.asc'
    );


  const fisikRemote =
    fisikSupabase.map(function(item) {

      return [

        item.tgl || '',

        item.berat === null
          ? null
          : Number(item.berat),

        item.panjang === null
          ? null
          : Number(item.panjang),

        item.lingkar_kepala === null
          ? null
          : Number(item.lingkar_kepala)

      ];

    });


  hasil.fisik =
    bandingkan(
      fisikLocal,
      fisikRemote
    );


  logHasil(
    'PERTUMBUHAN FISIK',
    hasil.fisik
  );


  /* =====================================================
     RINGKASAN
  ===================================================== */

  Logger.log('');
  Logger.log('========================================');
  Logger.log('RINGKASAN VALIDASI');
  Logger.log('========================================');


  let semuaAman = true;


  Object.keys(hasil)
    .forEach(function(nama) {

      const result =
        hasil[nama];

      const aman =
        result.missing.length === 0 &&
        result.extra.length === 0;

      if (!aman) {
        semuaAman = false;
      }


      Logger.log(
        (aman ? '✅ ' : '⚠️ ') +
        nama.toUpperCase() +
        ' | Spreadsheet: ' +
        result.localCount +
        ' | Supabase: ' +
        result.remoteCount +
        ' | Hilang: ' +
        result.missing.length +
        ' | Tambahan: ' +
        result.extra.length
      );

    });


  Logger.log('----------------------------------------');


  if (semuaAman) {

    Logger.log(
      '🎉 SEMUA DATA 100% IDENTIK'
    );

  } else {

    Logger.log(
      '⚠️ MASIH ADA PERBEDAAN DATA'
    );

  }


  Logger.log('========================================');
  Logger.log('VALIDASI EXACT SEMUA DATA SELESAI');
  Logger.log('========================================');


  return hasil;
}