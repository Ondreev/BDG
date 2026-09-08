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
 * компьютер.
 *
 * Обмен карточками между Поставщиком и Покупателем — тоже через этот скрипт: карточки
 * складываются в лист «catalog_shares» по номеру телефона получателя и удаляются оттуда,
 * когда он их примет или отклонит.
 *
 * Заказы поставщику — тоже через этот скрипт: покупатель отправляет магазин со своими
 * количествами на номер поставщика, заказ складывается в лист «orders» и удаляется
 * оттуда, когда поставщик его примет или отклонит. У поставщика принятый заказ
 * становится обычным списком закупок (просто помечен «входящий»).
 *
 * Фото товаров хранятся не в самой таблице (там текстовые ячейки), а в Google Диске —
 * скрипт сам создаст на Диске папку «BDG_photos» с подпапками по логинам. Это требует
 * доступа к Диску: при первом сохранении фото после обновления скрипта Google попросит
 * заново подтвердить разрешения — это нормально, просто разрешите.
 *
 * Если вы уже развернули более раннюю версию скрипта — вставьте этот файл заново и
 * создайте новую версию развёртывания (Развернуть → Управление развертываниями →
 * значок карандаша → Версия: «Новая версия» → Развернуть), иначе новые функции (печать
 * с телефона, обмен карточками, заказы поставщику, фото товаров) будут отвечать ошибкой
 * bad_action.
 */

var SHEET_NAME = 'users';
var CHUNK = 45000;   // лимит ячейки Google Sheets — 50 000 символов
var MAXCH = 8;       // до ~360 КБ данных на пользователя

var PRINT_SHEET_NAME = 'print_jobs';
var PRINT_MAXCH = 12; // до ~540 КБ на одно задание печати (документ + этикетки с QR)

var SHARE_SHEET_NAME = 'catalog_shares';
var SHARE_MAXCH = 12; // до ~540 КБ на одну передачу товаров

var ORDER_SHEET_NAME = 'orders';
var ORDER_MAXCH = 12; // до ~540 КБ на один заказ поставщику

var BACKUP_SHEET_NAME = 'backups';
var BACKUP_WINDOW_MS = 24 * 60 * 60 * 1000; // сутки, как и просили — окно, пока предлагаем восстановить

var HISTORY_SHEET_NAME = 'history';
var HISTORY_MAX_VERSIONS = 30;          // сколько последних точек возврата хранить на аккаунт
var HISTORY_MIN_GAP_MS = 3 * 60 * 1000; // не чаще одной точки в 3 минуты — иначе частые
                                         // автосохранения при активной работе быстро
                                         // замусорили бы историю почти одинаковыми снимками
var HISTORY_NO_CHANGES = 'Без изменений';

var PHOTO_ROOT_FOLDER_NAME = 'BDG_photos';
var PHOTO_MAX_BASE64_LEN = 2000000; // с запасом достаточно для сжатого фото с телефона

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
        var payload = req.payload || {};
        var now = new Date().toISOString();
        if (row === -1) {
          // первый вход — регистрация: логин занимается, PIN фиксируется
          var json0 = JSON.stringify(payload);
          if (json0.length > CHUNK * MAXCH) return out({ ok: false, error: 'too_big' });
          var chunks0 = [];
          for (var i0 = 0; i0 < json0.length; i0 += CHUNK) chunks0.push(json0.slice(i0, i0 + CHUNK));
          while (chunks0.length < MAXCH) chunks0.push('');
          sh.appendRow([login, "'" + pin, now].concat(chunks0));
          appendHistoryEntry(login, payload, json0);
          return out({ ok: true, at: now });
        }
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var existingVals = sh.getRange(row, 4, 1, MAXCH).getDisplayValues()[0];
        var existingJson = existingVals.join('');
        var existingLen = existingJson.length;
        // существующее состояние нужно и для слияния (ниже), и для дневника изменений
        // (appendHistoryEntry) — разбираем один раз, чтобы не парсить JSON дважды
        var existingPayload = null;
        if (existingJson) { try { existingPayload = JSON.parse(existingJson); } catch (eParse) {} }
        // каталог товаров сливаем с уже сохранённым на сервере по каждой карточке
        // отдельно (см. mergeCatalogDBServer), а не заменяем целиком — иначе устройство
        // с чуть более старой локальной копией могло вслепую затереть чужую свежую
        // правку (например, только что добавленное на другом устройстве фото).
        // req.skipMerge — единственное исключение: человек явно нажал "Всё равно
        // заменить" в диалоге предупреждения о потере данных, и этот диалог прямым
        // текстом обещает полную замену, поэтому в этом случае делаем ровно её
        if (!req.skipMerge && existingPayload) {
          try {
            payload.catalogDB = mergeCatalogDBServer(existingPayload.catalogDB, payload.catalogDB);
            // то же самое для списков закупок (см. mergeListsServer) — без этого
            // список с чуть более старой локальной копией на другом устройстве мог
            // вслепую затереть только что добавленный сюда товар
            var mergedLists = mergeListsServer(existingPayload.lists, existingPayload.deletedListIds, payload.lists, payload.deletedListIds);
            payload.lists = mergedLists.lists;
            payload.deletedListIds = mergedLists.deletedListIds;
            // то же самое для долгов и планов — любое несливаемое поле рискует быть
            // тихо затёртым push'ем с другого устройства, даже если тот push был
            // вообще не про долги/планы (каждая отправка несёт полный слепок состояния)
            var mergedDebts = mergePlainRecordsServer(existingPayload.debts, payload.debts, existingPayload.deletedDebtIds, payload.deletedDebtIds);
            payload.debts = mergedDebts.records;
            payload.deletedDebtIds = mergedDebts.deletedIds;
            var mergedPlans = mergePlansServer(existingPayload.plans, existingPayload.deletedPlanIds, payload.plans, payload.deletedPlanIds);
            payload.plans = mergedPlans.plans;
            payload.deletedPlanIds = mergedPlans.deletedPlanIds;
          } catch (eMerge) {}
        }
        var json = JSON.stringify(payload);
        if (json.length > CHUNK * MAXCH) return out({ ok: false, error: 'too_big' });
        var chunks = [];
        for (var i = 0; i < json.length; i += CHUNK) chunks.push(json.slice(i, i + CHUNK));
        while (chunks.length < MAXCH) chunks.push('');
        // защита от случайной перезаписи большого объёма данных почти пустыми.
        // Самый частый сценарий потери: слетел логин на устройстве, его ввели заново
        // и нажали "Выгрузить" раньше, чем успели что-то скачать — без этой проверки
        // такое одним запросом стирает всё, что накопилось в облаке с других устройств.
        var risky = existingLen > 500 && json.length < existingLen * 0.3;
        if (risky && !req.force) {
          return out({ ok: false, error: 'data_loss_risk', existingSize: existingLen, incomingSize: json.length });
        }
        // риск подтверждён явно (или кто-то намеренно решил уничтожить данные) — то, что
        // заменяется, кладём в резервную копию на сутки, чтобы это всегда можно было отменить
        if (risky) saveBackup(login, existingVals);
        sh.getRange(row, 3).setValue(now);
        sh.getRange(row, 4, 1, MAXCH).setValues([chunks]);
        // точка возврата в "Историю изменений" — независимо от бэкапа выше (тот
        // хранит только одну последнюю рискованную замену на сутки), эта пишет
        // регулярные снимки с коротким описанием, что изменилось с прошлого снимка
        appendHistoryEntry(login, payload, json);
        return out({ ok: true, at: now });
      }

      if (req.action === 'load') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var vals = sh.getRange(row, 4, 1, MAXCH).getDisplayValues()[0];
        var data = vals.join('');
        var backupInfo = getBackupInfo(login);
        return out({
          ok: true,
          payload: data ? JSON.parse(data) : null,
          at: String(sh.getRange(row, 3).getValue()),
          hasBackup: backupInfo.has,
          backupAt: backupInfo.at
        });
      }

      // отменить недавнюю рискованную перезапись (см. saveBackup выше) — доступно сутки
      // с момента замены, с любого устройства, вошедшего тем же логином+PIN
      if (req.action === 'restore_backup') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var bsh = getBackupSheet();
        var brow = findRow(bsh, login);
        if (brow === -1) return out({ ok: false, error: 'backup_not_found' });
        var binfo = getBackupInfo(login);
        if (!binfo.has) return out({ ok: false, error: 'backup_expired' });
        var bvals = bsh.getRange(brow, 3, 1, MAXCH).getDisplayValues()[0];
        var bdata = bvals.join('');
        var restoredAt = new Date().toISOString();
        sh.getRange(row, 3).setValue(restoredAt);
        sh.getRange(row, 4, 1, MAXCH).setValues([bvals]);
        bsh.deleteRow(brow); // восстановили — больше не предлагаем повторно
        return out({ ok: true, at: restoredAt, payload: bdata ? JSON.parse(bdata) : null });
      }

      // список точек возврата ("История изменений") — только дата и короткое описание,
      // без самих данных (тяжёлые колонки не читаем, чтобы не гонять зря лишний объём)
      if (req.action === 'history_list') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var hsh2 = getHistorySheet();
        var hlast = hsh2.getLastRow();
        var history = [];
        if (hlast >= 2) {
          var hrows = hsh2.getRange(2, 1, hlast - 1, 3).getDisplayValues();
          for (var hi = 0; hi < hrows.length; hi++) {
            if (String(hrows[hi][0]).replace(/\D/g, '') === login) {
              history.push({ at: hrows[hi][1], summary: hrows[hi][2] || '' });
            }
          }
        }
        history.reverse(); // новые сверху
        return out({ ok: true, history: history });
      }

      // "Использовать как истину" — явный откат к выбранному снимку из истории.
      // Текущее состояние перед заменой само кладётся в историю (как обычный снимок),
      // поэтому сам откат тоже можно отменить, выбрав снимок "до отката"
      if (req.action === 'history_restore') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var targetAt = String((req.payload && req.payload.at) || '');
        var hsh3 = getHistorySheet();
        var hRows = findHistoryRows(hsh3, login);
        var targetRow = -1;
        for (var tr = 0; tr < hRows.length; tr++) {
          if (String(hsh3.getRange(hRows[tr], 2).getDisplayValue()) === targetAt) { targetRow = hRows[tr]; break; }
        }
        if (targetRow === -1) return out({ ok: false, error: 'history_not_found' });
        var targetVals = hsh3.getRange(targetRow, 4, 1, MAXCH).getDisplayValues()[0];
        var targetJson = targetVals.join('');
        var targetPayload = targetJson ? JSON.parse(targetJson) : {};

        var curVals = sh.getRange(row, 4, 1, MAXCH).getDisplayValues()[0];
        var curJson = curVals.join('');
        if (curJson) {
          try { appendHistoryEntry(login, JSON.parse(curJson), curJson); } catch (eSnap) {}
        }

        var restoredAt2 = new Date(Date.now() + 1).toISOString();
        sh.getRange(row, 3).setValue(restoredAt2);
        sh.getRange(row, 4, 1, MAXCH).setValues([targetVals]);
        // отдельная точка "после отката" — записываем её ВСЕГДА, а не через обычный
        // appendHistoryEntry: иначе следующее обычное сохранение искало бы описание
        // изменений относительно последней ЗАПИСАННОЙ точки, которой оказался бы снимок
        // "перед откатом" (то есть состояние ДО отката) — и текст истории стал бы неверным
        var hsh4 = getHistorySheet();
        hsh4.appendRow([login, restoredAt2, 'Откат к более ранней версии'].concat(targetVals));
        pruneHistory(hsh4, login);
        return out({ ok: true, at: restoredAt2, payload: targetPayload });
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

      // передача карточек товаров между разными аккаунтами (Поставщик -> Покупатель).
      // получателя аутентифицировать нечем — отправитель знает только его номер телефона,
      // поэтому это "почтовый ящик": положить может любой, а прочитать/удалить — только
      // сам получатель своей парой логин+PIN. Отправитель аутентифицируется как обычно.
      // toLogin === 'ALL' — публичное предложение "всем": его видят все аккаунты, а
      // строка на сервере не удаляется по одиночному accept/decline — только отправитель
      // мог бы её убрать (такого действия пока нет), каждый получатель просто прячет
      // её у себя локально (state.dismissedShares), поэтому share_remove для него не зовём.
      if (req.action === 'share_push') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var rawTo = String((req.payload && req.payload.toLogin) || '');
        var toLogin = rawTo === 'ALL' ? 'ALL' : rawTo.replace(/\D/g, '');
        if (toLogin !== 'ALL' && !/^\d{10,15}$/.test(toLogin)) return out({ ok: false, error: 'bad_to_login' });
        var title = String((req.payload && req.payload.title) || '').slice(0, 80);
        var summary = String((req.payload && req.payload.summary) || '').slice(0, 200);
        var sjson = JSON.stringify({ items: (req.payload && req.payload.items) || [] });
        if (sjson.length > CHUNK * SHARE_MAXCH) return out({ ok: false, error: 'too_big' });
        var schunks = [];
        for (var k = 0; k < sjson.length; k += CHUNK) schunks.push(sjson.slice(k, k + CHUNK));
        while (schunks.length < SHARE_MAXCH) schunks.push('');
        var ssh = getShareSheet();
        var shareId = Utilities.getUuid();
        var itemCount = (req.payload && req.payload.items && req.payload.items.length) || 0;
        ssh.appendRow([shareId, login, toLogin, new Date().toISOString(), itemCount, title, summary].concat(schunks));
        return out({ ok: true, id: shareId });
      }

      if (req.action === 'share_list') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var ssh2 = getShareSheet();
        var last2 = ssh2.getLastRow();
        var shares = [];
        if (last2 >= 2) {
          var rows2 = ssh2.getRange(2, 1, last2 - 1, 7).getDisplayValues();
          for (var m = 0; m < rows2.length; m++) {
            var rowToLogin = String(rows2[m][2]);
            if (rowToLogin.replace(/\D/g, '') === login || rowToLogin === 'ALL') {
              shares.push({
                id: rows2[m][0], fromLogin: rows2[m][1], toLogin: rowToLogin, created: rows2[m][3],
                itemCount: +rows2[m][4] || 0, title: rows2[m][5] || '', summary: rows2[m][6] || ''
              });
            }
          }
        }
        return out({ ok: true, shares: shares });
      }

      if (req.action === 'share_fetch') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var frow = findShareRowForRecipient(getShareSheet(), String((req.payload && req.payload.id) || ''), login);
        if (frow === -1) return out({ ok: false, error: 'share_not_found' });
        var fvals = getShareSheet().getRange(frow, 8, 1, SHARE_MAXCH).getDisplayValues()[0];
        var fdata = fvals.join('');
        return out({ ok: true, payload: fdata ? JSON.parse(fdata) : { items: [] } });
      }

      if (req.action === 'share_remove') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var rrow = findShareRowForRecipient(getShareSheet(), String((req.payload && req.payload.id) || ''), login);
        if (rrow === -1) return out({ ok: false, error: 'share_not_found' });
        getShareSheet().deleteRow(rrow);
        return out({ ok: true });
      }

      // заказ поставщику: покупатель отправляет магазин со своими количествами (не карточки
      // каталога) на номер поставщика; тот принимает — и у себя получает обычный список
      // (движок редактирования/печати общий, специального кода на приёмной стороне не нужно).
      // Тот же "почтовый ящик", что и catalog_shares, но отдельный лист — другая структура
      // данных (есть qty, нет sku) и другая логика приёма.
      if (req.action === 'order_push') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var oToLogin = String((req.payload && req.payload.toLogin) || '').replace(/\D/g, '');
        if (!/^\d{10,15}$/.test(oToLogin)) return out({ ok: false, error: 'bad_to_login' });
        var storeName = String((req.payload && req.payload.storeName) || '').slice(0, 80);
        var fromPhone = String((req.payload && req.payload.fromPhone) || '').slice(0, 30);
        var ojson = JSON.stringify({ items: (req.payload && req.payload.items) || [] });
        if (ojson.length > CHUNK * ORDER_MAXCH) return out({ ok: false, error: 'too_big' });
        var ochunks = [];
        for (var n = 0; n < ojson.length; n += CHUNK) ochunks.push(ojson.slice(n, n + CHUNK));
        while (ochunks.length < ORDER_MAXCH) ochunks.push('');
        var osh = getOrderSheet();
        var orderId = Utilities.getUuid();
        var oItemCount = (req.payload && req.payload.items && req.payload.items.length) || 0;
        osh.appendRow([orderId, login, oToLogin, new Date().toISOString(), oItemCount, storeName, fromPhone].concat(ochunks));
        return out({ ok: true, id: orderId });
      }

      if (req.action === 'order_list') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var osh2 = getOrderSheet();
        var olast = osh2.getLastRow();
        var orders = [];
        if (olast >= 2) {
          var orows = osh2.getRange(2, 1, olast - 1, 7).getDisplayValues();
          for (var p = 0; p < orows.length; p++) {
            if (String(orows[p][2]).replace(/\D/g, '') === login) {
              orders.push({
                id: orows[p][0], fromLogin: orows[p][1], toLogin: orows[p][2], created: orows[p][3],
                itemCount: +orows[p][4] || 0, storeName: orows[p][5] || '', fromPhone: orows[p][6] || ''
              });
            }
          }
        }
        return out({ ok: true, orders: orders });
      }

      if (req.action === 'order_fetch') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var ofrow = findOrderRowForRecipient(getOrderSheet(), String((req.payload && req.payload.id) || ''), login);
        if (ofrow === -1) return out({ ok: false, error: 'order_not_found' });
        var ovals = getOrderSheet().getRange(ofrow, 8, 1, ORDER_MAXCH).getDisplayValues()[0];
        var odata = ovals.join('');
        return out({ ok: true, payload: odata ? JSON.parse(odata) : { items: [] } });
      }

      if (req.action === 'order_remove') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var orrow = findOrderRowForRecipient(getOrderSheet(), String((req.payload && req.payload.id) || ''), login);
        if (orrow === -1) return out({ ok: false, error: 'order_not_found' });
        getOrderSheet().deleteRow(orrow);
        return out({ ok: true });
      }

      // фото товара: сохраняется в Google Диске, а не в самой таблице (там текстовые
      // ячейки) — в данных товара хранится только id файла. Ссылка на просмотр строится
      // на клиенте (lh3.googleusercontent.com/d/{id}), доступ открыт "всем, у кого есть
      // ссылка" — иначе браузеру нечем было бы её показать в <img>.
      if (req.action === 'photo_upload') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var b64 = String((req.payload && req.payload.data) || '');
        if (!b64) return out({ ok: false, error: 'no_data' });
        if (b64.length > PHOTO_MAX_BASE64_LEN) return out({ ok: false, error: 'too_big' });
        var mime = String((req.payload && req.payload.mime) || 'image/jpeg');
        var bytes = Utilities.base64Decode(b64);
        var blob = Utilities.newBlob(bytes, mime, login + '_' + Date.now() + '.jpg');
        var folder = getUserPhotoFolder(login);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return out({ ok: true, id: file.getId() });
      }

      // удаление фото (заменили на другое или убрали) — чистим Диск, чтобы файлы не
      // копились без дела; проверяем, что файл лежит в папке именно этого логина
      if (req.action === 'photo_delete') {
        if (row === -1) return out({ ok: false, error: 'not_found' });
        if (getPin(sh, row) !== pin) return out({ ok: false, error: 'wrong_pin' });
        var pid = String((req.payload && req.payload.id) || '');
        if (pid) {
          try {
            var ownFolder = getUserPhotoFolder(login);
            var pfile = DriveApp.getFileById(pid);
            var parents = pfile.getParents();
            var owns = false;
            while (parents.hasNext()) { if (parents.next().getId() === ownFolder.getId()) owns = true; }
            if (owns) pfile.setTrashed(true);
          } catch (ignored) {} // файла уже нет/недоступен — удалять нечего
        }
        return out({ ok: true });
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

function getShareSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHARE_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHARE_SHEET_NAME);
    var head = ['id', 'fromLogin', 'toLogin', 'created', 'itemCount', 'title', 'summary'];
    for (var i = 1; i <= SHARE_MAXCH; i++) head.push('chunk' + i);
    sh.appendRow(head);
  }
  return sh;
}

// строка отдаётся тому, кому адресована (toLogin), либо любому, если это публичное
// предложение "всем" (toLogin === 'ALL') — это и есть проверка доступа
function findShareRowForRecipient(sh, id, toLogin) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return -1;
  var vals = sh.getRange(2, 1, last - 1, 3).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    var rowTo = String(vals[i][2]);
    if (String(vals[i][0]) === id && (rowTo.replace(/\D/g, '') === toLogin || rowTo === 'ALL')) return i + 2;
  }
  return -1;
}

function getOrderSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ORDER_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ORDER_SHEET_NAME);
    var head = ['id', 'fromLogin', 'toLogin', 'created', 'itemCount', 'storeName', 'fromPhone'];
    for (var i = 1; i <= ORDER_MAXCH; i++) head.push('chunk' + i);
    sh.appendRow(head);
  }
  return sh;
}

// заказы всегда адресные (нет варианта "всем", как у share) — строку отдаём и удаляем
// только тому, чей логин совпадает с toLogin
function findOrderRowForRecipient(sh, id, toLogin) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return -1;
  var vals = sh.getRange(2, 1, last - 1, 3).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === id && String(vals[i][2]).replace(/\D/g, '') === toLogin) return i + 2;
  }
  return -1;
}

function getPin(sh, row) {
  return String(sh.getRange(row, 2).getDisplayValue()).replace(/\D/g, '');
}

// резервная копия того, что было заменено рискованной выгрузкой — одна на логин
// (новая замена перезаписывает предыдущую резервную копию, а не копится бесконечно)
function getBackupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BACKUP_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(BACKUP_SHEET_NAME);
    var head = ['login', 'createdAt'];
    for (var i = 1; i <= MAXCH; i++) head.push('data' + i);
    sh.appendRow(head);
  }
  return sh;
}
function saveBackup(login, chunksVals) {
  var bsh = getBackupSheet();
  var brow = findRow(bsh, login);
  var now = new Date().toISOString();
  if (brow === -1) {
    bsh.appendRow([login, now].concat(chunksVals));
  } else {
    bsh.getRange(brow, 2).setValue(now);
    bsh.getRange(brow, 3, 1, MAXCH).setValues([chunksVals]);
  }
}
function getBackupInfo(login) {
  var bsh = getBackupSheet();
  var brow = findRow(bsh, login);
  if (brow === -1) return { has: false };
  var at = String(bsh.getRange(brow, 2).getDisplayValue());
  var age = Date.now() - new Date(at).getTime();
  return { has: age >= 0 && age < BACKUP_WINDOW_MS, at: at };
}

/* ================= История изменений (точки возврата) =================
   В отличие от "backups" выше (одна аварийная копия на сутки, только для
   рискованных замен), это регулярный журнал: на каждое сохранение — с
   троттлингом по времени — добавляется снимок ПОСЛЕ этого сохранения плюс
   короткое описание того, что изменилось с предыдущего снимка. Список этих
   точек показывается на клиенте под "стрелкой назад"; выбор точки и
   подтверждение "Использовать как истину" откатывает облако к ней (see
   history_restore в doPost) — сам откат тоже становится снимком, поэтому
   ошибку выбора можно исправить тем же способом. */
function getHistorySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(HISTORY_SHEET_NAME);
    var head = ['login', 'createdAt', 'summary'];
    for (var i = 1; i <= MAXCH; i++) head.push('data' + i);
    sh.appendRow(head);
  }
  return sh;
}
// строки данного логина в порядке добавления (сверху вниз = от старых к новым,
// поскольку новые точки всегда дописываются в конец через appendRow)
function findHistoryRows(hsh, login) {
  var last = hsh.getLastRow();
  if (last < 2) return [];
  var logins = hsh.getRange(2, 1, last - 1, 1).getDisplayValues();
  var rows = [];
  for (var i = 0; i < logins.length; i++) {
    if (String(logins[i][0]).replace(/\D/g, '') === login) rows.push(i + 2);
  }
  return rows;
}
// не даём истории расти бесконечно — оставляем только последние HISTORY_MAX_VERSIONS
// точек на логин, удаляя лишние старые (снизу вверх по номеру строки, чтобы удаление
// одной строки не сбивало номера ещё не обработанных)
function pruneHistory(hsh, login) {
  var rows = findHistoryRows(hsh, login);
  var excess = rows.length - HISTORY_MAX_VERSIONS;
  if (excess <= 0) return;
  var toDelete = rows.slice(0, excess).sort(function (a, b) { return b - a; });
  for (var i = 0; i < toDelete.length; i++) hsh.deleteRow(toDelete[i]);
}
// добавляет точку возврата, если есть что фиксировать. Сравнивает не с "существующими
// перед этим save данными" (existingPayload передавать сюда не нужно), а с ПРЕДЫДУЩЕЙ
// уже записанной точкой истории — так описание точки корректно отражает всё, что
// накопилось с прошлой точки, даже если несколько сохранений подряд попали в окно
// троттлинга и не получили собственных точек
function appendHistoryEntry(login, newPayload, newJson) {
  var hsh = getHistorySheet();
  var rows = findHistoryRows(hsh, login);
  var now = Date.now();
  var basePayload = null;
  if (rows.length) {
    var lastRow = rows[rows.length - 1];
    var lastAt = new Date(hsh.getRange(lastRow, 2).getDisplayValue()).getTime();
    if (!isNaN(lastAt) && now - lastAt < HISTORY_MIN_GAP_MS) return; // слишком рано для новой точки
    try {
      var lastVals = hsh.getRange(lastRow, 4, 1, MAXCH).getDisplayValues()[0];
      var lastJson = lastVals.join('');
      if (lastJson) basePayload = JSON.parse(lastJson);
    } catch (eBase) {}
  }
  var summary = rows.length ? buildHistorySummary(basePayload, newPayload) : 'Начальный снимок';
  if (summary === HISTORY_NO_CHANGES) return; // нечего фиксировать
  var chunks = [];
  for (var i = 0; i < newJson.length; i += CHUNK) chunks.push(newJson.slice(i, i + CHUNK));
  while (chunks.length < MAXCH) chunks.push('');
  hsh.appendRow([login, new Date(now).toISOString(), summary].concat(chunks));
  pruneHistory(hsh, login);
}
// сравнение двух плоских массивов записей по id — сколько добавлено/убрано/изменено.
// "изменено" — грубое сравнение по JSON.stringify всей записи (для короткой сводки в
// истории точности достаточно, здесь не нужна логика полноценного слияния)
function diffRecordsById(prevArr, newArr) {
  prevArr = prevArr || []; newArr = newArr || [];
  var prevMap = {}, i, it;
  for (i = 0; i < prevArr.length; i++) prevMap[prevArr[i].id] = prevArr[i];
  var seen = {}, added = 0, changed = 0;
  for (i = 0; i < newArr.length; i++) {
    it = newArr[i]; seen[it.id] = true;
    var prev = prevMap[it.id];
    if (!prev) added++;
    else if (JSON.stringify(prev) !== JSON.stringify(it)) changed++;
  }
  var removed = 0;
  for (i = 0; i < prevArr.length; i++) { if (!seen[prevArr[i].id]) removed++; }
  return { added: added, removed: removed, changed: changed };
}
function historyPartLabel(name, d) {
  if (!d.added && !d.removed && !d.changed) return null;
  var bits = [];
  if (d.added) bits.push('+' + d.added);
  if (d.removed) bits.push('-' + d.removed);
  if (d.changed) bits.push('изм. ' + d.changed);
  return name + ': ' + bits.join(', ');
}
// короткое, человекочитаемое описание того, что изменилось между двумя полными
// снимками состояния — используется как подпись точки в "Истории изменений"
function buildHistorySummary(prevPayload, newPayload) {
  prevPayload = prevPayload || {};
  newPayload = newPayload || {};
  var parts = [], i, j;

  var dProducts = diffRecordsById(
    (prevPayload.catalogDB && prevPayload.catalogDB.products) || [],
    (newPayload.catalogDB && newPayload.catalogDB.products) || []
  );
  var pLabel = historyPartLabel('Каталог', dProducts);
  if (pLabel) parts.push(pLabel);

  var prevItems = [], newItems = [];
  var pLists = prevPayload.lists || [], nLists = newPayload.lists || [];
  for (i = 0; i < pLists.length; i++) {
    var pStores = (pLists[i].data && pLists[i].data.stores) || [];
    for (j = 0; j < pStores.length; j++) prevItems = prevItems.concat(pStores[j].items || []);
  }
  for (i = 0; i < nLists.length; i++) {
    var nStores = (nLists[i].data && nLists[i].data.stores) || [];
    for (j = 0; j < nStores.length; j++) newItems = newItems.concat(nStores[j].items || []);
  }
  var dItems = diffRecordsById(prevItems, newItems);
  var prevItemMap = {};
  for (i = 0; i < prevItems.length; i++) prevItemMap[prevItems[i].id] = prevItems[i];
  var boughtOn = 0, boughtOff = 0;
  for (i = 0; i < newItems.length; i++) {
    var ni = newItems[i], pi = prevItemMap[ni.id];
    if (pi && !pi.bought && ni.bought) boughtOn++;
    if (pi && pi.bought && !ni.bought) boughtOff++;
  }
  if (dItems.added || dItems.removed || dItems.changed || boughtOn || boughtOff) {
    var ip = [];
    if (dItems.added) ip.push('+' + dItems.added);
    if (dItems.removed) ip.push('-' + dItems.removed);
    if (boughtOn) ip.push('куплено ' + boughtOn);
    if (boughtOff) ip.push('возврат ' + boughtOff);
    var otherChanged = dItems.changed - boughtOn - boughtOff;
    if (otherChanged > 0) ip.push('изм. ' + otherChanged);
    parts.push('Списки: ' + ip.join(', '));
  }

  var dDebts = diffRecordsById(prevPayload.debts || [], newPayload.debts || []);
  var debtLabel = historyPartLabel('Долги', dDebts);
  if (debtLabel) parts.push(debtLabel);

  var prevTasks = [], newTasks = [];
  var pPlans = prevPayload.plans || [], nPlans = newPayload.plans || [];
  for (i = 0; i < pPlans.length; i++) prevTasks = prevTasks.concat(pPlans[i].items || []);
  for (i = 0; i < nPlans.length; i++) newTasks = newTasks.concat(nPlans[i].items || []);
  var dTasks = diffRecordsById(prevTasks, newTasks);
  var tasksLabel = historyPartLabel('Планы', dTasks);
  if (tasksLabel) parts.push(tasksLabel);

  return parts.length ? parts.join('; ') : HISTORY_NO_CHANGES;
}

// у каждого логина своя подпапка внутри общей "BDG_photos" — так фото разных аккаунтов
// на одном скрипте не путаются и их проще найти вручную на Диске при необходимости
function getPhotosRootFolder() {
  var it = DriveApp.getFoldersByName(PHOTO_ROOT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_ROOT_FOLDER_NAME);
}
function getUserPhotoFolder(login) {
  var root = getPhotosRootFolder();
  var it = root.getFoldersByName(login);
  return it.hasNext() ? it.next() : root.createFolder(login);
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// слияние каталога товаров по каждой записи отдельно (та же логика, что и на клиенте
// в mergeCatalogDB) — выполняется здесь, на сервере, внутри той же блокировки, что и
// сама запись, одним запросом. Раньше это делал клиент: сначала отдельным запросом
// скачивал текущие данные, сливал у себя и только потом отправлял — на каждое
// сохранение уходило два обращения к скрипту вместо одного, и при активной работе
// с нескольких устройств запросы начали упираться в тайм-аут общей блокировки
// (LockService — она общая для всех действий и всех пользователей сразу).
// Здесь слияние атомарно: гонка между чтением текущих данных и записью исключена,
// потому что и то, и другое происходит под одной и той же уже захваченной блокировкой.
function mergeCatalogDBServer(existingDB, incomingDB) {
  existingDB = existingDB || {};
  incomingDB = incomingDB || {};
  var tomb = {}, id;
  var exTomb = existingDB.deletedProducts || {};
  var inTomb = incomingDB.deletedProducts || {};
  for (id in exTomb) tomb[id] = Number(exTomb[id]) || 0;
  for (id in inTomb) tomb[id] = Math.max(tomb[id] || 0, Number(inTomb[id]) || 0);

  var prodMap = {}, order = [];
  var exProducts = existingDB.products || [];
  var inProducts = incomingDB.products || [];
  var i, p, cur, curAt, pAt;
  for (i = 0; i < exProducts.length; i++) {
    p = exProducts[i];
    if (!(p.id in prodMap)) order.push(p.id);
    prodMap[p.id] = p;
  }
  for (i = 0; i < inProducts.length; i++) {
    p = inProducts[i];
    cur = prodMap[p.id];
    if (!cur) { prodMap[p.id] = p; order.push(p.id); continue; }
    curAt = Number(cur.updatedAt) || 0;
    pAt = Number(p.updatedAt) || 0;
    if (pAt > curAt || (pAt === curAt && !cur.photoId && p.photoId)) prodMap[p.id] = p;
  }
  var products = [];
  for (i = 0; i < order.length; i++) {
    p = prodMap[order[i]];
    if ((tomb[p.id] || 0) <= (Number(p.updatedAt) || 0)) products.push(p);
  }

  var catMap = {}, catOrder = [], c;
  var exCats = existingDB.categories || [];
  var inCats = incomingDB.categories || [];
  for (i = 0; i < exCats.length; i++) { c = exCats[i]; if (!(c.id in catMap)) catOrder.push(c.id); catMap[c.id] = c; }
  for (i = 0; i < inCats.length; i++) { c = inCats[i]; if (!(c.id in catMap)) catOrder.push(c.id); catMap[c.id] = c; }
  var categories = [];
  for (i = 0; i < catOrder.length; i++) categories.push(catMap[catOrder[i]]);

  var locMap = {}, locOrder = [], l;
  var exLocs = existingDB.locations || [];
  var inLocs = incomingDB.locations || [];
  for (i = 0; i < exLocs.length; i++) { l = exLocs[i]; if (!(l.id in locMap)) locOrder.push(l.id); locMap[l.id] = l; }
  for (i = 0; i < inLocs.length; i++) { l = inLocs[i]; if (!(l.id in locMap)) locOrder.push(l.id); locMap[l.id] = l; }
  var locations = [];
  for (i = 0; i < locOrder.length; i++) locations.push(locMap[locOrder[i]]);

  var merged = {};
  for (id in existingDB) merged[id] = existingDB[id];
  for (id in incomingDB) merged[id] = incomingDB[id];
  merged.categories = categories;
  merged.products = products;
  merged.deletedProducts = tomb;
  merged.locations = locations;
  return merged;
}

// то же слияние по записям, что и mergeCatalogDBServer, но для списков закупок —
// портировано из клиентского mergeListContents/mergeLists (index.html) 1-в-1, чтобы
// поведение при сохранении (здесь) и при скачивании (на клиенте) не расходилось.
// Магазины и позиции объединяются по id ГЛОБАЛЬНО по всему списку (не по каждому
// магазину отдельно) — иначе товар, перемещённый между магазинами на одном из
// устройств, задвоился бы после слияния
function mergeListContentsServer(existingList, incomingList) {
  var exData = existingList.data || {}, inData = incomingList.data || {};
  var delStoreIds = {}, id, i;
  var exDelStores = exData.deletedStoreIds || [], inDelStores = inData.deletedStoreIds || [];
  for (i = 0; i < exDelStores.length; i++) delStoreIds[exDelStores[i]] = true;
  for (i = 0; i < inDelStores.length; i++) delStoreIds[inDelStores[i]] = true;
  var delItemIds = {};
  var exDelItems = exData.deletedItemIds || [], inDelItems = inData.deletedItemIds || [];
  for (i = 0; i < exDelItems.length; i++) delItemIds[exDelItems[i]] = true;
  for (i = 0; i < inDelItems.length; i++) delItemIds[inDelItems[i]] = true;

  // существование магазина/позиции — чистое объединение (никогда не теряем то, что
  // знает только одна сторона), а вот СОДЕРЖИМОЕ и РАСПОЛОЖЕНИЕ для того, что есть
  // на обеих сторонах, берём из ПРИСЫЛАЕМОЙ версии (incoming — то, что это устройство
  // сохраняет прямо сейчас): без пометок времени на каждом поле это единственный
  // способ не потерять, например, переименование магазина или перемещение товара —
  // если бы тут побеждала "уже сохранённая" версия, любая правка существующей записи
  // молча пропадала бы при каждом сохранении
  var storeOrder = [], storeMeta = {};
  function regStore(st, canOverwrite) {
    if (!(st.id in storeMeta)) storeOrder.push(st.id);
    if (canOverwrite || !(st.id in storeMeta)) storeMeta[st.id] = { id: st.id, name: st.name, collapsed: st.collapsed, storeDirId: st.storeDirId };
  }
  var exStores = exData.stores || [], inStores = inData.stores || [];
  for (i = 0; i < exStores.length; i++) regStore(exStores[i], false);
  for (i = 0; i < inStores.length; i++) regStore(inStores[i], true);

  // позиции — глобально по id во всём списке; для общих id содержимое и расположение
  // решаются по updatedAt каждой позиции (чья правка новее), а не "присылаемая всегда
  // побеждает" — присланный снимок может быть устаревшим (например, из-за защиты от
  // дедлока автосинка на клиенте, когда правка отправляется, даже если в облаке уже
  // есть более новая версия), и без метки времени такой снимок мог бы откатить чужую
  // более свежую отметку "куплено"
  var itemLoc = {}, st, j, it;
  for (i = 0; i < exStores.length; i++) {
    st = exStores[i];
    for (j = 0; j < (st.items || []).length; j++) { it = st.items[j]; itemLoc[it.id] = { storeId: st.id, item: it }; }
  }
  for (i = 0; i < inStores.length; i++) {
    st = inStores[i];
    for (j = 0; j < (st.items || []).length; j++) {
      it = st.items[j];
      var curLoc = itemLoc[it.id];
      if (!curLoc || (Number(it.updatedAt) || 0) >= (Number(curLoc.item.updatedAt) || 0)) itemLoc[it.id] = { storeId: st.id, item: it };
    }
  }

  var stores = [];
  for (i = 0; i < storeOrder.length; i++) {
    var sid = storeOrder[i];
    if (delStoreIds[sid]) continue;
    var meta = storeMeta[sid];
    var items = [];
    for (var key in itemLoc) {
      var entry = itemLoc[key];
      if (entry.storeId === sid && !delItemIds[entry.item.id]) items.push(entry.item);
    }
    stores.push({ id: meta.id, name: meta.name, collapsed: meta.collapsed, storeDirId: meta.storeDirId, items: items });
  }

  var merged = {};
  for (id in existingList) merged[id] = existingList[id];
  for (id in incomingList) merged[id] = incomingList[id];
  var mergedData = {};
  for (id in exData) mergedData[id] = exData[id];
  for (id in inData) mergedData[id] = inData[id];
  mergedData.stores = stores;
  var delStoreArr = [], delItemArr = [];
  for (id in delStoreIds) delStoreArr.push(id);
  for (id in delItemIds) delItemArr.push(id);
  mergedData.deletedStoreIds = delStoreArr;
  mergedData.deletedItemIds = delItemArr;
  merged.data = mergedData;
  return merged;
}
// слияние массива списков целиком — портировано из клиентского mergeLists. Раньше
// state.lists при сохранении просто заменялся целиком присланным — если на одном
// устройстве список только что пополнили, а другое устройство (с чуть более старой
// локальной копией того же списка) сохранялось следом, весь список откатывался к
// версии второго устройства и добавленный товар пропадал, хотя первое устройство
// уже успешно его отправило
function mergeListsServer(existingLists, existingDeletedListIds, incomingLists, incomingDeletedListIds) {
  var delListIds = {}, id, i;
  existingDeletedListIds = existingDeletedListIds || [];
  incomingDeletedListIds = incomingDeletedListIds || [];
  for (i = 0; i < existingDeletedListIds.length; i++) delListIds[existingDeletedListIds[i]] = true;
  for (i = 0; i < incomingDeletedListIds.length; i++) delListIds[incomingDeletedListIds[i]] = true;

  var order = [], map = {};
  existingLists = existingLists || []; incomingLists = incomingLists || [];
  for (i = 0; i < existingLists.length; i++) { order.push(existingLists[i].id); map[existingLists[i].id] = existingLists[i]; }
  for (i = 0; i < incomingLists.length; i++) {
    var l = incomingLists[i];
    if (!(l.id in map)) { order.push(l.id); map[l.id] = l; }
    else map[l.id] = mergeListContentsServer(map[l.id], l);
  }
  var lists = [], delListArr = [];
  for (i = 0; i < order.length; i++) { if (!delListIds[order[i]]) lists.push(map[order[i]]); }
  for (id in delListIds) delListArr.push(id);
  return { lists: lists, deletedListIds: delListArr };
}
// универсальное слияние плоского массива записей по id (долги и т.п.) — портировано
// из клиентского mergePlainRecords: существование объединяется, содержимое общей
// записи берётся из ПРИСЫЛАЕМОЙ (incoming) версии — то же самое правило, что и для
// товаров каталога и позиций списков: без этого правки долгов/планов, отправленные
// с ОДНОГО устройства, стирались бы следующим push'ом ЛЮБОГО другого устройства,
// даже если тот push был вообще не про долги/планы — каждая отправка несёт полный
// слепок состояния целиком
function mergePlainRecordsServer(existingRecords, incomingRecords, existingDeletedIds, incomingDeletedIds) {
  var delIds = {}, id, i;
  existingDeletedIds = existingDeletedIds || []; incomingDeletedIds = incomingDeletedIds || [];
  for (i = 0; i < existingDeletedIds.length; i++) delIds[existingDeletedIds[i]] = true;
  for (i = 0; i < incomingDeletedIds.length; i++) delIds[incomingDeletedIds[i]] = true;

  var order = [], map = {};
  existingRecords = existingRecords || []; incomingRecords = incomingRecords || [];
  for (i = 0; i < existingRecords.length; i++) { if (!(existingRecords[i].id in map)) order.push(existingRecords[i].id); map[existingRecords[i].id] = existingRecords[i]; }
  for (i = 0; i < incomingRecords.length; i++) {
    var rec = incomingRecords[i];
    if (!(rec.id in map)) { order.push(rec.id); map[rec.id] = rec; continue; }
    var curRec = map[rec.id];
    if ((Number(rec.updatedAt) || 0) >= (Number(curRec.updatedAt) || 0)) map[rec.id] = rec;
  }
  var records = [], delArr = [];
  for (i = 0; i < order.length; i++) { if (!delIds[order[i]]) records.push(map[order[i]]); }
  for (id in delIds) delArr.push(id);
  return { records: records, deletedIds: delArr };
}
// слияние планов (чек-листов) — та же идея, что mergeListsServer/mergeListContentsServer,
// но на один уровень вложенности мельче: план -> задачи (без промежуточного "магазина")
function mergePlanContentsServer(existingPlan, incomingPlan) {
  var exDel = existingPlan.deletedItemIds || [], inDel = incomingPlan.deletedItemIds || [];
  var delItemIds = {}, i;
  for (i = 0; i < exDel.length; i++) delItemIds[exDel[i]] = true;
  for (i = 0; i < inDel.length; i++) delItemIds[inDel[i]] = true;
  var mergedRecs = mergePlainRecordsServer(existingPlan.items, incomingPlan.items, [], []);
  var items = [];
  for (i = 0; i < mergedRecs.records.length; i++) { if (!delItemIds[mergedRecs.records[i].id]) items.push(mergedRecs.records[i]); }
  var merged = {}, id2;
  for (id2 in existingPlan) merged[id2] = existingPlan[id2];
  for (id2 in incomingPlan) merged[id2] = incomingPlan[id2];
  merged.items = items;
  var delArr = [];
  for (id2 in delItemIds) delArr.push(id2);
  merged.deletedItemIds = delArr;
  return merged;
}
function mergePlansServer(existingPlans, existingDeletedPlanIds, incomingPlans, incomingDeletedPlanIds) {
  var delPlanIds = {}, id, i;
  existingDeletedPlanIds = existingDeletedPlanIds || []; incomingDeletedPlanIds = incomingDeletedPlanIds || [];
  for (i = 0; i < existingDeletedPlanIds.length; i++) delPlanIds[existingDeletedPlanIds[i]] = true;
  for (i = 0; i < incomingDeletedPlanIds.length; i++) delPlanIds[incomingDeletedPlanIds[i]] = true;

  var order = [], map = {};
  existingPlans = existingPlans || []; incomingPlans = incomingPlans || [];
  for (i = 0; i < existingPlans.length; i++) { order.push(existingPlans[i].id); map[existingPlans[i].id] = existingPlans[i]; }
  for (i = 0; i < incomingPlans.length; i++) {
    var p = incomingPlans[i];
    if (!(p.id in map)) { order.push(p.id); map[p.id] = p; }
    else map[p.id] = mergePlanContentsServer(map[p.id], p);
  }
  var plans = [], delArr = [];
  for (i = 0; i < order.length; i++) { if (!delPlanIds[order[i]]) plans.push(map[order[i]]); }
  for (id in delPlanIds) delArr.push(id);
  return { plans: plans, deletedPlanIds: delArr };
}
