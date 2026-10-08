const CLOUD_FILE_ID = '1TDpnpSV1zWnEEuiOBdoN4GmxKqbk_NTd';
const ACCESS_TOKEN = 'UJwLuJ0_CVu23ROijfsUwTaY6LjanmsdlASVF8s1md5VxvE_';
const EXPECTED_HEADERS = ['Ссылка', 'Название хар-ки', 'Значение', 'Ссылка на раздел'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ok:false, error:'Пустой POST-запрос'});
    }

    const body = JSON.parse(e.postData.contents);
    if (String(body.token || '') !== ACCESS_TOKEN) {
      return json_({ok:false, error:'Неверный токен'});
    }

    const requestedId = String(body.file_id || '').trim();
    if (requestedId && requestedId !== CLOUD_FILE_ID) {
      return json_({ok:false, error:'Неверный ID облачного файла'});
    }

    const incomingCsv = String(body.csv || '');
    if (!incomingCsv.trim()) {
      return json_({ok:false, error:'Пустой CSV'});
    }

    const file = DriveApp.getFileById(CLOUD_FILE_ID);
    const currentCsv = file.getBlob().getDataAsString('UTF-8');

    const existingRows = parseRows_(currentCsv);
    const incomingRows = parseRows_(incomingCsv);

    const merged = [];
    const keys = new Set();

    existingRows.forEach(function(row) {
      const key = row.join('\t');
      if (!keys.has(key)) {
        keys.add(key);
        merged.push(row);
      }
    });

    let added = 0;
    incomingRows.forEach(function(row) {
      const key = row.join('\t');
      if (!keys.has(key)) {
        keys.add(key);
        merged.push(row);
        added++;
      }
    });

    file.setContent(toCsv_(merged));

    return json_({
      ok: true,
      added: added,
      total: merged.length,
      file_id: CLOUD_FILE_ID,
      url: 'https://drive.google.com/file/d/' + CLOUD_FILE_ID + '/view'
    });
  } catch (err) {
    return json_({ok:false, error:String(err && err.message ? err.message : err)});
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({
    ok:true,
    service:'FE-RUS SEO Manager cloud CSV endpoint',
    file_id:CLOUD_FILE_ID
  });
}

function parseRows_(csv) {
  const text = String(csv || '').replace(/^\uFEFF/, '').trim();
  if (!text) return [];

  const matrix = Utilities.parseCsv(text, ';');
  if (!matrix.length) return [];

  const rows = [];
  const start = looksLikeHeader_(matrix[0]) ? 1 : 0;
  for (let i = start; i < matrix.length; i++) {
    const row = matrix[i] || [];
    const url = normalizeUrl_(row[0]);
    const characteristic = String(row.length > 1 ? row[1] : '').trim();
    const value = String(row.length > 2 ? row[2] : '').trim();
    // Поддержка старого двухколоночного файла:
    // [Ссылка, Ссылка на раздел] -> [Ссылка, '', '', Ссылка на раздел].
    let parent = '';
    if (row.length >= 4) {
      parent = normalizeUrl_(row[3]);
    } else if (row.length === 2) {
      parent = normalizeUrl_(row[1]);
    }
    if (url && parent) rows.push([url, characteristic, value, parent]);
  }
  return rows;
}

function looksLikeHeader_(row) {
  if (!row || !row.length) return false;
  return String(row[0] || '').trim() === EXPECTED_HEADERS[0];
}

function normalizeUrl_(value) {
  return String(value || '').trim().replace(/\/+$/, '/') || '';
}

function csvEscape_(value) {
  const s = String(value == null ? '' : value);
  if (/[";\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function toCsv_(rows) {
  const out = [EXPECTED_HEADERS.join(';')];
  rows.forEach(function(row) {
    out.push(
      csvEscape_(row[0]) + ';' +
      csvEscape_(row[1]) + ';' +
      csvEscape_(row[2]) + ';' +
      csvEscape_(row[3])
    );
  });
  return out.join('\n') + '\n';
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
