// FE-RUS SEO Manager cloud storage Web App v1.23.10 — добавлена поддержка рабочих файлов TOP-10
const ROOT_FOLDER_ID = '1uyJ9YqbX7v1wYT2FX4hmO3GIWk0u-YYs';
const COMMON_FOLDER_NAME = '00_ОБЩИЕ';
const COMMON_FILE_NAME = 'seo_create_links.csv';

// Уже созданный общий CSV в папке 00_ОБЩИЕ.
// Если файл будет удалён, скрипт найдёт/создаст его заново.
const CLOUD_FILE_ID = '1YcEPhor7Up_5Wt2wCni86irElRqqkc0Y';

const ACCESS_TOKEN = 'UJwLuJ0_CVu23ROijfsUwTaY6LjanmsdlASVF8s1md5VxvE_';
const EXPECTED_HEADERS = ['Ссылка', 'Название хар-ки', 'Значение', 'Ссылка на раздел'];
const ALLOWED_SECTION_FILES = new Set([
  // Основная версия TOP-30
  'seo_top30_result.csv',
  'seo_top30_progress.json',
  'seo_serp_page_classification.csv',
  'seo_url_check_audit.csv',
  'seo_ocfilter_check_audit.json',
  'wordkeeper_error.log',

  // Экспериментальная версия TOP-10
  'seo_top10_result.csv',
  'seo_top10_progress.json',
  'seo_top10_serp_page_classification.csv',
  'seo_top10_url_check_audit.csv',
  'seo_top10_ocfilter_check_audit.json',
  'wordkeeper_top10_error.log'
]);

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

    // Список уже созданных папок-разделов.
    if (String(body.action || '') === 'list_sections') {
      const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
      const sections = [];
      const folders = root.getFolders();
      while (folders.hasNext()) {
        const f = folders.next();
        if (f.getName() === COMMON_FOLDER_NAME) continue;
        sections.push({
          id: f.getId(),
          name: f.getName(),
          url: 'https://drive.google.com/drive/folders/' + f.getId()
        });
      }
      sections.sort(function(a,b){ return a.name.localeCompare(b.name, 'ru'); });
      return json_({ok:true, sections:sections});
    }

    // Служебный запрос: создать/проверить общую структуру.
    if (String(body.action || '') === 'ensure_structure') {
      const common = getOrCreateCommonFile_();
      let section = null;
      const sectionName = String(body.section_name || '').trim();
      if (sectionName) {
        section = getOrCreateSectionFolder_(sectionName);
      }

      return json_({
        ok:true,
        root_folder_id:ROOT_FOLDER_ID,
        root_folder_url:'https://drive.google.com/drive/folders/' + ROOT_FOLDER_ID,
        common_folder_id:common.folderId,
        common_folder_url:'https://drive.google.com/drive/folders/' + common.folderId,
        common_file_id:common.fileId,
        common_file_url:'https://drive.google.com/file/d/' + common.fileId + '/view',
        section:section,
        section_id: section ? section.id : '',
        section_name: section ? section.name : '',
        section_url: section ? section.url : ''
      });
    }

    // Быстрая синхронизация раздела: один запрос возвращает метаданные всех
    // рабочих файлов. Содержимое передаём только вторым запросом для файлов,
    // которые действительно изменились/отсутствуют локально.
    if (String(body.action || '') === 'sync_section_meta') {
      const sectionName = String(body.section_name || '').trim();
      const section = getOrCreateSectionFolder_(sectionName);
      const folder = DriveApp.getFolderById(section.id);
      const fileMap = {};
      const files = folder.getFiles();
      while (files.hasNext()) {
        const f = files.next();
        const name = f.getName();
        if (!ALLOWED_SECTION_FILES.has(name)) continue;
        const bytes = f.getSize();
        fileMap[name] = {
          file_name:name,
          file_id:f.getId(),
          found:bytes > 0,
          bytes:bytes,
          updated_ms:f.getLastUpdated().getTime()
        };
      }

      const resultFiles = [];
      ALLOWED_SECTION_FILES.forEach(function(name) {
        if (fileMap[name]) {
          resultFiles.push(fileMap[name]);
        } else {
          resultFiles.push({file_name:name, found:false, bytes:0, updated_ms:0, file_id:''});
        }
      });

      return json_({
        ok:true,
        section_id:section.id,
        section_name:section.name,
        section_url:section.url,
        files:resultFiles
      });
    }

    // Второй и обычно последний запрос синхронизации. Можно одновременно
    // скачать изменившиеся файлы и загрузить отсутствующие/обновлённые.
    if (String(body.action || '') === 'sync_section_data') {
      const sectionName = String(body.section_name || '').trim();
      const section = getOrCreateSectionFolder_(sectionName);
      const folder = DriveApp.getFolderById(section.id);
      const requestedDownloads = Array.isArray(body.download_files) ? body.download_files : [];
      const uploadFiles = body.upload_files && typeof body.upload_files === 'object' ? body.upload_files : {};

      const fileMap = {};
      const files = folder.getFiles();
      while (files.hasNext()) {
        const f = files.next();
        const name = f.getName();
        if (!ALLOWED_SECTION_FILES.has(name)) continue;
        fileMap[name] = f;
      }

      const uploaded = [];
      Object.keys(uploadFiles).forEach(function(fileName) {
        validateSectionFile_(fileName);
        const item = uploadFiles[fileName] || {};
        const compressed = item.compressed !== false;
        let raw = Utilities.base64Decode(String(item.content_b64 || ''));
        if (compressed) {
          const gzipBlob = Utilities.newBlob(raw, 'application/gzip', fileName + '.gz');
          raw = Utilities.ungzip(gzipBlob).getBytes();
        }
        if (!raw || raw.length === 0) return;

        let file = fileMap[fileName];
        if (!file) {
          const mime = fileName.toLowerCase().endsWith('.json') ? MimeType.PLAIN_TEXT : MimeType.CSV;
          file = folder.createFile(fileName, '', mime);
          fileMap[fileName] = file;
        }
        const mime = fileName.toLowerCase().endsWith('.json') ? 'application/json' : 'text/csv';
        const content = Utilities.newBlob(raw, mime, fileName).getDataAsString('UTF-8');
        file.setContent(content);
        uploaded.push({
          file_name:fileName,
          file_id:file.getId(),
          folder_id:folder.getId(),
          section_name:folder.getName(),
          bytes:raw.length,
          updated_ms:file.getLastUpdated().getTime()
        });
      });

      const downloaded = [];
      requestedDownloads.forEach(function(fileName) {
        validateSectionFile_(fileName);
        const file = fileMap[fileName];
        if (!file) return;
        const raw = file.getBlob().getBytes();
        if (!raw || raw.length === 0) return;
        const zipped = Utilities.gzip(file.getBlob(), fileName + '.gz');
        downloaded.push({
          file_name:fileName,
          file_id:file.getId(),
          folder_id:folder.getId(),
          section_name:folder.getName(),
          bytes:raw.length,
          updated_ms:file.getLastUpdated().getTime(),
          compressed:true,
          content_b64:Utilities.base64Encode(zipped.getBytes())
        });
      });

      return json_({ok:true, section_id:section.id, section_name:section.name, uploaded:uploaded, downloaded:downloaded});
    }

    // Скачать рабочий файл конкретного раздела.
    if (String(body.action || '') === 'download_file') {
      const sectionName = String(body.section_name || '').trim();
      const fileName = String(body.file_name || '').trim();
      validateSectionFile_(fileName);
      const info = getOrCreateSectionFile_(sectionName, fileName);
      if (!info.fileId) return json_({ok:true, found:false, section_id:info.folderId, section_name:info.folderName});
      const file = DriveApp.getFileById(info.fileId);
      const raw = file.getBlob().getBytes();

      // Пустой файл не считаем валидной облачной копией.
      // Это защищает локальный кэш от затирания после неудачного upload.
      if (!raw || raw.length === 0) {
        return json_({
          ok:true,
          found:false,
          empty:true,
          section_id:info.folderId,
          section_name:info.folderName,
          file_name:fileName,
          file_id:info.fileId
        });
      }

      const zipped = Utilities.gzip(file.getBlob(), fileName + '.gz');
      return json_({
        ok:true,
        found:true,
        compressed:true,
        content_b64:Utilities.base64Encode(zipped.getBytes()),
        file_id:info.fileId,
        folder_id:info.folderId,
        section_name:info.folderName,
        file_name:fileName
      });
    }

    // Загрузить/обновить рабочий файл конкретного раздела.
    if (String(body.action || '') === 'upload_file') {
      const sectionName = String(body.section_name || '').trim();
      const fileName = String(body.file_name || '').trim();
      validateSectionFile_(fileName);
      const info = getOrCreateSectionFile_(sectionName, fileName);
      let raw = Utilities.base64Decode(String(body.content_b64 || ''));
      if (body.compressed) {
        const gzipBlob = Utilities.newBlob(
          raw,
          'application/gzip',
          fileName + '.gz'
        );
        raw = Utilities.ungzip(gzipBlob).getBytes();
      }
      const mime = fileName.toLowerCase().endsWith('.json')
        ? 'application/json'
        : 'text/csv';
      const content = Utilities.newBlob(raw, mime, fileName).getDataAsString('UTF-8');
      const file = DriveApp.getFileById(info.fileId);

      // При миграции сохраняем текущую версию перед заменой.
      if (body.backup_existing && file.getSize() > 0) {
        const stamp = Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone() || 'Europe/Moscow',
          'yyyyMMdd_HHmmss'
        );
        const backupName = fileName.replace(/\.[^.]+$/, '') + '.before_migration_' + stamp + '.csv';
        file.makeCopy(backupName, DriveApp.getFolderById(info.folderId));
      }

      file.setContent(content);
      return json_({
        ok:true,
        file_id:info.fileId,
        folder_id:info.folderId,
        section_name:info.folderName,
        file_name:fileName,
        bytes:raw.length
      });
    }

    const fileInfo = getOrCreateCommonFile_();

    const requestedId = String(body.file_id || '').trim();
    if (requestedId && requestedId !== fileInfo.fileId) {
      return json_({ok:false, error:'Неверный ID облачного файла'});
    }

    const incomingCsv = String(body.csv || '');
    if (!incomingCsv.trim()) {
      return json_({ok:false, error:'Пустой CSV'});
    }

    const file = DriveApp.getFileById(fileInfo.fileId);
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
      file_id:fileInfo.fileId,
      folder_id:fileInfo.folderId,
      url:'https://drive.google.com/file/d/' + fileInfo.fileId + '/view'
    });
  } catch (err) {
    return json_({ok:false, error:String(err && err.message ? err.message : err)});
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  // Лёгкая проверка /exec. Никаких обращений к Drive:
  // программа вызывает GET только для проверки, что Web App опубликован.
  return json_({
    ok:true,
    service:'FE-RUS SEO Manager cloud storage v1.23.9',
    root_folder_id:ROOT_FOLDER_ID,
    root_folder_url:'https://drive.google.com/drive/folders/' + ROOT_FOLDER_ID
  });
}

function validateSectionFile_(fileName) {
  if (!ALLOWED_SECTION_FILES.has(fileName)) {
    throw new Error('Недопустимое имя рабочего файла: ' + fileName);
  }
}

function getOrCreateSectionFile_(sectionName, fileName) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const clean = sanitizeFolderName_(sectionName);
  if (!clean) throw new Error('Пустое название раздела');

  let folder = null;
  const folders = root.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    const title = f.getName();
    if (title === clean || title.replace(/^\d+_/, '') === clean) {
      folder = f;
      break;
    }
  }
  if (!folder) { const created = getOrCreateSectionFolder_(clean); folder = DriveApp.getFolderById(created.id); }

  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    return {folderId:folder.getId(), folderName:folder.getName(), fileId:files.next().getId()};
  }

  const mime = fileName.toLowerCase().endsWith('.json') ? MimeType.PLAIN_TEXT : MimeType.CSV;
  const file = folder.createFile(fileName, '', mime);
  return {folderId:folder.getId(), folderName:folder.getName(), fileId:file.getId()};
}

function getOrCreateCommonFile_() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const commonFolder = getOrCreateFolderByName_(root, COMMON_FOLDER_NAME);

  // Сначала пытаемся использовать текущий ID.
  try {
    const existing = DriveApp.getFileById(CLOUD_FILE_ID);
    if (existing) {
      return {folderId:commonFolder.getId(), fileId:existing.getId()};
    }
  } catch (e) {
    // Файл удалён/недоступен — ищем по имени ниже.
  }

  const files = commonFolder.getFilesByName(COMMON_FILE_NAME);
  if (files.hasNext()) {
    return {folderId:commonFolder.getId(), fileId:files.next().getId()};
  }

  const content = EXPECTED_HEADERS.join(';') + '\n';
  const file = commonFolder.createFile(COMMON_FILE_NAME, content, MimeType.CSV);
  return {folderId:commonFolder.getId(), fileId:file.getId()};
}

function getOrCreateFolderByName_(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function getOrCreateSectionFolder_(sectionName) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const clean = sanitizeFolderName_(sectionName);
  if (!clean) throw new Error('Пустое название раздела');

  // Ищем уже существующий раздел независимо от его номера.
  const folders = root.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    const title = f.getName();
    if (title === clean || title.replace(/^\d+_/, '') === clean) {
      return {
        id:f.getId(),
        name:title,
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
    url:'https://drive.google.com/drive/folders/' + folder.getId()
  };
}

function sanitizeFolderName_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|#%{}~]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 120);
}

function parseRows_(csv) {
  // Старые версии могли записать символы \\n буквально вместо перевода строки.
  // Нормализуем такой файл при следующей записи, чтобы существующие данные не потерялись.
  const text = String(csv || '').replace(/^\uFEFF/, '').replace(/\\n/g, '\n').trim();
  if (!text) return [];

  // При первоначальном создании через Google Sheets CSV мог оказаться
  // с разделителем запятая. Поддерживаем оба варианта.
  const delimiter = text.indexOf(';') >= 0 ? ';' : ',';
  const matrix = Utilities.parseCsv(text, delimiter);
  if (!matrix.length) return [];

  const rows = [];
  const start = looksLikeHeader_(matrix[0]) ? 1 : 0;

  for (let i = start; i < matrix.length; i++) {
    const row = matrix[i] || [];
    const url = normalizeUrl_(row[0]);
    const characteristic = String(row.length > 1 ? row[1] : '').trim();
    const value = String(row.length > 2 ? row[2] : '').trim();

    let parent = '';
    if (row.length >= 4) {
      parent = normalizeUrl_(row[3]);
    } else if (row.length === 2) {
      parent = normalizeUrl_(row[1]);
    }

    if (url && parent) {
      rows.push([url, characteristic, value, parent]);
    }
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
  if (/[";\\r\\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
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
