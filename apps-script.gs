/**
 * Google Apps Script — общее хранилище для "Учёт товаров".
 * Хранит товары и настройки в этой Google Таблице.
 *
 * В отличие от предыдущей версии, каждое изменение (добавление/редактирование/
 * удаление товара) отправляется на сервер ОТДЕЛЬНЫМ запросом и применяется
 * точечно — к одной строке по её id, — а не перезаписывает весь лист целиком.
 * Это устраняет гонки, из-за которых при почти одновременных действиях
 * (или при возврате на вкладку сразу после сохранения) часть данных могла
 * "слетать". Дополнительно запись защищена LockService — пока один запрос
 * пишет в таблицу, остальные ждут своей очереди вместо того, чтобы писать
 * одновременно и портить данные друг друга.
 *
 * УСТАНОВКА:
 * 1. Откройте sheets.google.com → создайте новую таблицу.
 * 2. Меню Расширения → Apps Script.
 * 3. Удалите всё содержимое файла Code.gs и вставьте туда этот файл целиком.
 * 4. (Необязательно, но рекомендуется) смените APP_KEY ниже на свою строку —
 *    и точно такую же впишите в index.html в константу SHEET_API_KEY.
 * 5. Сохраните (значок дискеты).
 * 6. Развернуть → Новое развёртывание → тип "Веб-приложение":
 *      Выполнять как: Я
 *      У кого есть доступ: Все
 * 7. Скопируйте "URL веб-приложения" (заканчивается на /exec).
 * 8. Вставьте этот URL в index.html в константу SHEET_API_URL.
 *
 * ВАЖНО: если у вас уже было старое развёртывание — создайте новую версию
 * существующего развёртывания (Управление развёртываниями → редактировать →
 * новая версия), тогда URL менять не придётся.
 */

// Смените на любую свою строку — и здесь, и в index.html в константе
// SHEET_API_KEY (значения должны совпадать буква в букву). Это простая защита
// от посторонних, которые случайно узнают ссылку /exec.
const APP_KEY = 'constructor-secret-2026';

const ITEMS_SHEET = 'items';
const SETTINGS_SHEET = 'settings';
const ITEM_FIELDS = ['id','img','name','item','weight','coeff','sale','buyDate','saleDate','comment','category','qty'];
const SETTINGS_FIELDS = ['coeff','cny','uah'];

function checkKey(e) {
  const key = (e.parameter && e.parameter.key) || '';
  return key === APP_KEY;
}

function doGet(e) {
  if (!checkKey(e)) return jsonOut({ok: false, error: 'bad key'});
  const data = readAll();
  return jsonOut({ok: true, items: data.items, settings: data.settings});
}

function doPost(e) {
  if (!checkKey(e)) return jsonOut({ok: false, error: 'bad key'});
  const lock = LockService.getScriptLock();
  try {
    // ждём до 10 секунд, если таблицу прямо сейчас пишет другой запрос
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut({ok: false, error: 'Таблица занята, попробуйте ещё раз'});
  }
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || '{}');

    if (payload.type === 'settings') {
      writeSettings(payload.data || {});
      return jsonOut({ok: true});
    }

    // payload.type === 'item' (или не указан — для обратной совместимости)
    const sheet = getSheet(ITEMS_SHEET, ITEM_FIELDS);

    if (payload.action === 'add') {
      sheet.appendRow(ITEM_FIELDS.map(f => cellVal(payload.data[f])));
    } else if (payload.action === 'update') {
      const rowIndex = findRowById(sheet, payload.data.id);
      if (rowIndex > -1) {
        sheet.getRange(rowIndex, 1, 1, ITEM_FIELDS.length)
          .setValues([ITEM_FIELDS.map(f => cellVal(payload.data[f]))]);
      } else {
        // такой строки не нашлось (например, рассинхронизация) — добавляем,
        // чтобы правка точно не потерялась
        sheet.appendRow(ITEM_FIELDS.map(f => cellVal(payload.data[f])));
      }
    } else if (payload.action === 'delete') {
      const rowIndex = findRowById(sheet, payload.id);
      if (rowIndex > -1) sheet.deleteRow(rowIndex);
    } else {
      return jsonOut({ok: false, error: 'unknown action: ' + payload.action});
    }

    return jsonOut({ok: true});
  } catch (err) {
    return jsonOut({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

function writeSettings(s) {
  const sh = getSheet(SETTINGS_SHEET, SETTINGS_FIELDS);
  sh.getRange(2, 1, 1, SETTINGS_FIELDS.length)
    .setValues([SETTINGS_FIELDS.map(f => cellVal(s[f]))]);
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { // i=1: пропускаем строку заголовков
    if (String(data[i][0]) === String(id)) {
      return i + 1; // Таблицы считают строки с 1, строка 1 — заголовок
    }
  }
  return -1;
}

function cellVal(v) {
  return (v === undefined || v === null) ? '' : String(v);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name, headerRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headerRow);
    // весь лист как текст — иначе Таблицы сами превращают даты/числа в
    // типизированные значения, и при следующем чтении получаем "мусор"
    sh.getRange(1, 1, sh.getMaxRows(), headerRow.length).setNumberFormat('@');
  }
  return sh;
}

// Значение из ячейки может оказаться реальным объектом Date (если Таблица
// сама распознала текст как дату) — тогда String(value) даёт что-то вроде
// "Wed Jul 01 2026 00:00:00 GMT+0300 (...)" вместо "2026-07-01". Приводим
// такие значения обратно к простой строке "yyyy-MM-dd".
function cellToString(value) {
  if (value === undefined || value === null || value === '') return '';
  if (value instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  return String(value);
}

function readAll() {
  const itemsSh = getSheet(ITEMS_SHEET, ITEM_FIELDS);
  const rows = itemsSh.getDataRange().getValues();
  const items = rows.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => {
      const o = {};
      ITEM_FIELDS.forEach((f, i) => { o[f] = cellToString(r[i]); });
      return o;
    });

  const setSh = getSheet(SETTINGS_SHEET, SETTINGS_FIELDS);
  const setRows = setSh.getDataRange().getValues();
  let settings = {coeff: 6.4, cny: 6.8, uah: 45};
  if (setRows.length > 1) {
    settings = {};
    SETTINGS_FIELDS.forEach((f, i) => { settings[f] = cellToString(setRows[1][i]); });
  }
  return {items: items, settings: settings};
}