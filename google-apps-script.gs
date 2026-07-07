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
 *
 * Печать с телефона на компьютер использует тот же скрипт: задания складываются во
 * временный лист «print_jobs» и удаляются оттуда сразу после того, как их заберёт
 * компьютер. Если вы уже развернули более раннюю версию скрипта — вставьте этот файл
 * заново и создайте новую версию развёртывания (Развернуть → Управление развертываниями →
 * значок карандаша → Версия: «Новая версия» → Развернуть), иначе печать с телефона будет
 * отвечать ошибкой bad_action.
 */

var SHEET_NAME = 'users';
var CHUNK = 45000;   // лимит ячейки Google Sheets — 50 000 символов
var MAXCH = 8;       // до ~360 КБ данных на пользователя

var PRINT_SHEET_NAME = 'print_jobs';
var PRINT_MAXCH = 12; // до ~540 КБ на одно задание печати (документ + этикетки с QR)

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

      // печать с телефона на компьютер: задание кладётся в очередь и разбирается той же
      // парой логин/PIN на другом устройстве — компьютер сам его печатает и удаляет из очереди
      if (req.action === 'print_push') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var pjson = JSON.stringify(req.payload || {});
        if (pjson.length > CHUNK * PRINT_MAXCH) return out({ ok: false, error: 'too_big' });
        var pchunks = [];
        for (var j = 0; j < pjson.length; j += CHUNK) pchunks.push(pjson.slice(j, j + CHUNK));
        while (pchunks.length < PRINT_MAXCH) pchunks.push('');
        var psh = getPrintSheet();
        psh.appendRow([login, new Date().toISOString()].concat(pchunks));
        return out({ ok: true });
      }

      if (req.action === 'print_poll') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var psh2 = getPrintSheet();
        var prow = findRow(psh2, login);
        if (prow === -1) return out({ ok: true, job: null });
        var pvals = psh2.getRange(prow, 3, 1, PRINT_MAXCH).getDisplayValues()[0];
        var pdata = pvals.join('');
        psh2.deleteRow(prow);
        return out({ ok: true, job: pdata ? JSON.parse(pdata) : null });
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

function getPrintSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(PRINT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PRINT_SHEET_NAME);
    var head = ['login', 'created'];
    for (var i = 1; i <= PRINT_MAXCH; i++) head.push('chunk' + i);
    sh.appendRow(head);
  }
  return sh;
}

function getPin(sh, row) {
  return String(sh.getRange(row, 2).getDisplayValue()).replace(/\D/g, '');
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
