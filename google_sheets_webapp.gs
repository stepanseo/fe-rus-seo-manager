const SPREADSHEET_ID = '1NbEr9ZDR5dmFa_8zayfSQtQ5VL1dLn2jSN0nNT1fkD8';
const SHEET_NAME = ''; // пусто = первый лист таблицы
const ACCESS_TOKEN = 'Bi9OUKpZtyuRj0hY86DAP8w9FcoyZqO1QTzSccmP83A';

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

    const rows = Array.isArray(body.rows) ? body.rows : [];
    if (!rows.length) {
      return json_({ok:false, error:'Нет строк для записи'});
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = SHEET_NAME
      ? ss.getSheetByName(SHEET_NAME)
      : ss.getSheets()[0];
    if (!sh) throw new Error('Лист не найден');

    // Заголовок.
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 2).setValues([['Ссылка', 'Ссылка на раздел']]);
    } else {
      const header = sh.getRange(1, 1, 1, Math.max(2, sh.getLastColumn())).getValues()[0];
      if (String(header[0] || '') !== 'Ссылка' || String(header[1] || '') !== 'Ссылка на раздел') {
        sh.insertRows(1, 1);
        sh.getRange(1, 1, 1, 2).setValues([['Ссылка', 'Ссылка на раздел']]);
      }
    }

    const lastRow = sh.getLastRow();
    const existing = new Set();
    if (lastRow >= 2) {
      const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
      values.forEach(r => {
        const u = String(r[0] || '').trim().replace(/\/+$/, '/') ;
        const p = String(r[1] || '').trim().replace(/\/+$/, '/') ;
        if (u && p) existing.add(u + '\t' + p);
      });
    }

    const toAppend = [];
    rows.forEach(item => {
      const url = String(item.url || '').trim().replace(/\/+$/, '/') ;
      const parent = String(item.parent_url || '').trim().replace(/\/+$/, '/') ;
      if (!url || !parent) return;
      const key = url + '\t' + parent;
      if (!existing.has(key)) {
        existing.add(key);
        toAppend.push([url, parent]);
      }
    });

    if (toAppend.length) {
      sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 2).setValues(toAppend);
    }

    return json_({
      ok: true,
      added: toAppend.length,
      total: Math.max(0, sh.getLastRow() - 1)
    });
  } catch (err) {
    return json_({ok:false, error:String(err && err.message ? err.message : err)});
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ok:true, service:'FE-RUS SEO Manager Google Sheets endpoint'});
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
