const ROOT_FOLDER_ID = '1uyJ9YqbX7v1wYT2FX4hmO3GIWk0u-YYs';
const COMMON_FOLDER_NAME = '00_ОБЩИЕ';
const COMMON_FILE_NAME = 'seo_create_links.csv';
const CLOUD_FILE_ID = '1YcEPhor7Up_5Wt2wCni86irElRqqkc0Y';

const ACCESS_TOKEN = 'UJwLuJ0_CVu23ROijfsUwTaY6LjanmsdlASVF8s1md5VxvE_';
const EXPECTED_HEADERS = ['Ссылка', 'Название хар-ки', 'Значение', 'Ссылка на раздел'];

function doGet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const info = getOrCreateCommonFile_();
    return json_({
      ok:true,
      service:'FE-RUS SEO Manager cloud storage v1.23.0',
      root_folder_id:ROOT_FOLDER_ID,
      root_folder_url:'https://drive.google.com/drive/folders/' + ROOT_FOLDER_ID,
      common_folder_id:info.folderId,
      common_file_id:info.fileId,
      file_id:info.fileId,
      file_url:'https://drive.google.com/file/d/' + info.fileId + '/view'
    });
  } catch (err) {
    return json_({ok:false, error:String(err && err.message ? err.message : err)});
  } finally {
    lock.releaseLock();
  }
}

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

    const action = String(body.action || '').trim();

    if (action === 'ensure_structure') {
      const common = getOrCreateCommonFile_();
      const sectionName = String(body.section_name || '').trim();
      if (!sectionName) {
        return json_({
          ok:true,
          root_folder_id:ROOT_FOLDER_ID,
          common_folder_id:common.folderId,
          common_file_id:common.fileId
        });
      }

      const section = getOrCreateSectionFolder_(sectionName);
      return json_({
        ok:true,
        root_folder_id:ROOT_FOLDER_ID,
        root_folder_url:'https://drive.google.com/drive/folders/' + ROOT_FOLDER_ID,
        common_folder_id:common.folderId,
        common_file_id:common.fileId,
        section_id:section.id,
        section_name:section.name,
        section_url:section.url
      });
    }

    if (action === 'download_file') {
      const sectionName = String(body.section_name || '').trim();
      const fileName = sanitizeFileName_(String(body.file_name || '').trim());
      if (!sectionName || !fileName) {
        return json_({ok:false, error:'Не задан section_name или file_name'});
      }

      const section = getOrCreateSectionFolder_(sectionName);
      const file = findFileInFolder_(section.folder, fileName);
      if (!file) {
        return json_({
          ok:true,
          found:false,
          section_id:section.id,
          section_name:section.name,
          file_name:fileName
        });
      }

      const blob = file.getBlob();
      const zipped = Utilities.gzip(blob, fileName + '.gz');
      return json_({
        ok:true,
        found:true,
        section_id:section.id,
        section_name:section.name,
        file_id:file.getId(),
        file_name:fileName,
        compressed:true,
        content_b64:Utilities.base64Encode(zipped.getBytes()),
        size:blob.getBytes().length,
        modified_at:file.getLastUpdated().toISOString()
      });
    }

    if (action === 'upload_file') {
      const sectionName = String(body.section_name || '').trim();
      const fileName = sanitizeFileName_(String(body.file_name || '').trim());
      const b64 = String(body.content_b64 || '');
      const compressed = Boolean(body.compressed);

      if (!sectionName || !fileName || !b64) {
        return json_({ok:false, error:'Не заданы section_name, file_name или content_b64'});
      }

      const section = getOrCreateSectionFolder_(sectionName);
      let bytes = Utilities.base64Decode(b64);

      if (compressed) {
        const zipped = Utilities.newBlob(bytes, 'application/gzip', fileName + '.gz');
        bytes = Utilities.ungzip(zipped).getBytes();
      }

      const mime = mimeForFile_(fileName);
      const blob = Utilities.newBlob(bytes, mime, fileName);
      let file = findFileInFolder_(section.folder, fileName);

      if (file) {
        file.setContent(blob.getDataAsString('UTF-8'));
      } else {
        file = section.folder.createFile(blob);
      }

      return json_({
        ok:true,
        section_id:section.id,
        section_name:section.name,
        file_id:file.getId(),
        file_name:fileName,
        size:bytes.length,
        modified_at:file.getLastUpdated().toISOString(),
        url:'https://drive.google.com/file/d/' + file.getId() + '/view'
      });
    }

    // Совместимость с накопительным seo_create_links.csv.
    const common = getOrCreateCommonFile_();
    const requestedId = String(body.file_id || '').trim();
    if (requestedId && requestedId !== common.fileId) {
      return json_({ok:false, error:'Неверный ID облачного файла'});
    }

    const incomingCsv = String(body.csv || '');
    if (!incomingCsv.trim()) {
      return json_({ok:false, error:'Пустой CSV'});
    }

    const file = DriveApp.getFileById(common.fileId);
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
      ok:true,
      added:added,
      total:merged.length,
      file_id:common.fileId,
      folder_id:common.folderId,
      url:'https://drive.google.com/file/d/' + common.fileId + '/view'
    });
  } catch (err) {
    return json_({ok:false, error:String(err && err.message ? err.message : err)});
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateCommonFile_() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const commonFolder = getOrCreateFolderByName_(root, COMMON_FOLDER_NAME);

  try {
    const existing = DriveApp.getFileById(CLOUD_FILE_ID);
    return {folderId:commonFolder.getId(), fileId:existing.getId()};
  } catch (e) {}

  const files = commonFolder.getFilesByName(COMMON_FILE_NAME);
  if (files.hasNext()) {
    return {folderId:commonFolder.getId(), fileId:files.next().getId()};
  }

  const content = EXPECTED_HEADERS.join(';') + '\n';
  const file = commonFolder.createFile(COMMON_FILE_NAME, content, MimeType.CSV);
  return {folderId:commonFolder.getId(), fileId:file.getId()};
}

function getOrCreateSectionFolder_(sectionName) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const clean = sanitizeFolderName_(sectionName);
  if (!clean) throw new Error('Пустое название раздела');

  const folders = root.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    const title = f.getName();
    if (title === clean || title.replace(/^\d+_/, '') === clean) {
      return {
        id:f.getId(),
        name:title,
        folder:f,
        url:'https://drive.google.com/drive/folders/' + f.getId()
      };
    }
  }

  let maxNo = 0;
  const all = root.getFolders();
  const re = /^(\d+)_/;
  while (all.hasNext()) {
    const title = all.next().getName();
    const m = title.match(re);
    if (m) maxNo = Math.max(maxNo, Number(m[1]));
  }

  const number = String(maxNo + 1).padStart(2, '0');
  const folder = root.createFolder(number + '_' + clean);

  return {
    id:folder.getId(),
    name:folder.getName(),
    folder:folder,
    url:'https://drive.google.com/drive/folders/' + folder.getId()
  };
}

function getOrCreateFolderByName_(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function findFileInFolder_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

function sanitizeFolderName_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|#%{}~]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 120);
}

function sanitizeFileName_(name) {
  const s = String(name || '').trim();
  if (!/^[A-Za-z0-9._-]+$/.test(s)) {
    throw new Error('Недопустимое имя файла: ' + s);
  }
  return s;
}

function mimeForFile_(name) {
  if (/\.json$/i.test(name)) return 'application/json';
  if (/\.csv$/i.test(name)) return MimeType.CSV;
  if (/\.log$/i.test(name)) return 'text/plain';
  return 'text/plain';
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

    let parent = '';
    if (row.length >= 4) parent = normalizeUrl_(row[3]);
    else if (row.length === 2) parent = normalizeUrl_(row[1]);

    if (url && parent) rows.push([url, characteristic, value, parent]);
  }

  return rows;
}

function looksLikeHeader_(row) {
  return !!(row && row.length && String(row[0] || '').trim() === EXPECTED_HEADERS[0]);
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
