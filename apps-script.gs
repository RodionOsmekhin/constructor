/**
 * Google Apps Script — общее хранилище для "Учёт товаров".
 * Хранит товары и настройки в этой Google Таблице и отдаёт/принимает их как JSON,
 * чтобы страница на GitHub Pages могла читать и писать данные, общие для всех.
 *
 * УСТАНОВКА:
 * 1. Откройте sheets.google.com → создайте новую таблицу.
 * 2. Меню Расширения → Apps Script.
 * 3. Удалите всё содержимое файла Code.gs и вставьте туда этот файл целиком.
 * 4. Сохраните (значок дискеты).
 * 5. Развернуть → Новое развёртывание → тип "Веб-приложение":
 *      Выполнять как: Я
 *      У кого есть доступ: Все
 * 6. Скопируйте "URL веб-приложения" (заканчивается на /exec).
 * 7. Вставьте этот URL в index.html в константу SHEET_API_URL.
 *
 * Подробности — в файле GOOGLE_SHEETS_SETUP.md рядом с этим файлом.
 */

const ITEMS_SHEET = 'items';
const SETTINGS_SHEET = 'settings';
const ITEM_FIELDS = ['id','img','name','item','weight','coeff','sale','buyDate','saleDate','comment'];
const SETTINGS_FIELDS = ['coeff','cny','uah'];

function doGet(e) {
  return jsonOut(readAll());
}

function doPost(e) {
  const body = JSON.parse((e.postData && e.postData.contents) || '{}');
  writeAll(body);
  return jsonOut({ok: true});
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
  }
  return sh;
}

function readAll() {
  const itemsSh = getSheet(ITEMS_SHEET, ITEM_FIELDS);
  const rows = itemsSh.getDataRange().getValues();
  const items = rows.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => {
      const o = {};
      ITEM_FIELDS.forEach((f, i) => {
        o[f] = (r[i] === undefined || r[i] === null) ? '' : String(r[i]);
      });
      return o;
    });

  const setSh = getSheet(SETTINGS_SHEET, SETTINGS_FIELDS);
  const setRows = setSh.getDataRange().getValues();
  let settings = {coeff: 6.4, cny: 6.8, uah: 45};
  if (setRows.length > 1) {
    settings = {};
    SETTINGS_FIELDS.forEach((f, i) => { settings[f] = setRows[1][i]; });
  }
  return {items: items, settings: settings};
}

function writeAll(data) {
  const itemsSh = getSheet(ITEMS_SHEET, ITEM_FIELDS);
  itemsSh.clearContents();
  itemsSh.appendRow(ITEM_FIELDS);
  const items = Array.isArray(data.items) ? data.items : [];
  if (items.length) {
    const rows = items.map(p =>
      ITEM_FIELDS.map(f => (p[f] !== undefined && p[f] !== null) ? String(p[f]) : ''));
    itemsSh.getRange(2, 1, rows.length, ITEM_FIELDS.length).setValues(rows);
  }

  const setSh = getSheet(SETTINGS_SHEET, SETTINGS_FIELDS);
  setSh.clearContents();
  setSh.appendRow(SETTINGS_FIELDS);
  const s = data.settings || {};
  setSh.appendRow(SETTINGS_FIELDS.map(f => (s[f] !== undefined ? s[f] : '')));
}
