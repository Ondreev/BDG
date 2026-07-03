/**
 * BDG — облачная синхронизация через Google Таблицу.
 *
 * Как установить (один раз, ~2 минуты):
 * 1. Откройте вашу таблицу BDGApp → меню «Расширения» → «Apps Script».
 * 2. Удалите всё в редакторе и вставьте целиком этот файл. Нажмите «Сохранить» (значок дискеты).
 * 3. Кнопка «Развернуть» (справа сверху) → «Новое развертывание» → тип «Веб-приложение»:
 *      - Выполнять от имени: «От моего имени»
 *      - У кого есть доступ: «Все»
 *    → «Начать развертывание» → разрешите доступ своему аккаунту → скопируйте URL веб-приложения
 *      (вида https://script.google.com/macros/s/XXXX/exec).
 * 4. Вставьте этот URL в приложении: Меню → «Облако — копия данных» → поле «Адрес веб-приложения».
 * 5. ВАЖНО: отмените публикацию таблицы в интернете (Файл → Поделиться → Публикация в интернете →
 *    Отменить публикацию) и закройте доступ по ссылке. Скрипту публичный доступ к таблице не нужен,
 *    а данные пользователей не должны быть видны посторонним.
 *
 * Скрипт сам создаст лист «users»: login | pin | updated | data1..data8.
 * Каждый пользователь видит только свои данные: строка отдаётся только при совпадении логина и PIN.
 */

var SHEET_NAME = 'users';
var CHUNK = 45000;   // лимит ячейки Google Sheets — 50 000 символов
var MAXCH = 8;       // до ~360 КБ данных на пользователя

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var login = String(req.login || '').replace(/\D/g, '');
    var pin = String(req.pin || '');
    if (!/^\d{10,15}$/.test(login)) return out({ ok: false, error: 'bad_login' });
    if (!/^\d{4}$/.test(pin)) return out({ ok: false, error: 'bad_pin' });

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sh = getSheet();
      var row = findRow(sh, login);

      if (req.action === 'save') {
        var json = JSON.stringify(req.payload || {});
        if (json.length > CHUNK * MAXCH) return out({ ok: false, error: 'too_big' });
        var chunks = [];
        for (var i = 0; i < json.length; i += CHUNK) chunks.push(json.slice(i, i + CHUNK));
        while (chunks.length < MAXCH) chunks.push('');
        var now = new Date().toISOString();
        if (row === -1) {
          // первый вход — регистрация: логин занимается, PIN фиксируется
          sh.appendRow([login, "'" + pin, now].concat(chunks));
        } else {
          if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
          sh.getRange(row, 3).setValue(now);
          sh.getRange(row, 4, 1, MAXCH).setValues([chunks]);
        }
        return out({ ok: true, at: now });
      }

      if (req.action === 'load') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var vals = sh.getRange(row, 4, 1, MAXCH).getDisplayValues()[0];
        var data = vals.join('');
        return out({
          ok: true,
          payload: data ? JSON.parse(data) : null,
          at: String(sh.getRange(row, 3).getValue())
        });
      }

      return out({ ok: false, error: 'bad_action' });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out({ ok: false, error: 'server: ' + String(err) });
  }
}

function doGet() {
  return out({ ok: true, service: 'BDG sync', hint: 'use POST' });
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    var head = ['login', 'pin', 'updated'];
    for (var i = 1; i <= MAXCH; i++) head.push('data' + i);
    sh.appendRow(head);
  }
  return sh;
}

function findRow(sh, login) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var logins = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < logins.length; i++) {
    if (String(logins[i][0]).replace(/\D/g, '') === login) return i + 2;
  }
  return -1;
}

function getPin(sh, row) {
  return String(sh.getRange(row, 2).getDisplayValue()).replace(/\D/g, '');
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
