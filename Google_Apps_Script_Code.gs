/**
 * BLSTCS.RU - Google Apps Script для синхронизации расчетов баллистики и базы аккаунтов (v3.3)
 * Таблица: "Расчет баллистики" (https://docs.google.com/spreadsheets/d/1QsGYXKAYsQOCguMVwBq-BhCaeRhzL33YVekkr49RYM8/edit)
 * 
 * ВАЖНО ПРИ ОБНОВЛЕНИИ СКРИПТА:
 * Чтобы изменения вступили в силу, в Apps Script нажмите:
 * 1. "Развернуть" (Deploy) -> "Управление развертываниями" (Manage deployments).
 * 2. Нажмите иконку карандаша (Edit).
 * 3. В выпадающем списке "Версия" (Version) выберите: "Новая версия" (New version).
 * 4. Нажмите "Развернуть" (Deploy).
 */

var SHEET_HISTORY = 'История расчетов';
var SHEET_USERS = 'Пользователи';

// ================= GET ЗАПРОСЫ =================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'get_history';
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Проверка связи (ping)
    if (action === 'ping') {
      var histSheet = getOrCreateHistorySheet(ss);
      var userSheet = getOrCreateUsersSheet(ss);
      return jsonResponse({
        status: 'ok',
        message: 'BLSTCS Google Sheets API v3.3 подключено успешно!',
        spreadsheet: ss.getName(),
        history_rows: Math.max(0, histSheet.getLastRow() - 1),
        users_count: Math.max(0, userSheet.getLastRow() - 1),
        time: new Date().toISOString()
      });
    }

    // 2. Получение истории расчетов
    if (action === 'get_history') {
      var sheet = getOrCreateHistorySheet(ss);
      var data = sheet.getDataRange().getValues();

      if (data.length <= 1) {
        return jsonResponse({ status: 'ok', count: 0, data: [] });
      }

      var rows = [];
      for (var i = 1; i < data.length; i++) {
        var r = data[i];
        if (!r[0] && !r[1] && !r[2]) continue;
        rows.push({
          id: r[0],
          created_at: r[1] ? (r[1] instanceof Date ? r[1].toLocaleString("ru-RU") : r[1].toString()) : '',
          product_name: r[2] || '',
          product_code: r[3] || '',
          material: r[4] || '',
          layers: r[5] || '',
          single_area: r[6] || '',
          cover_material: r[7] || '',
          accessories_summary: r[8] || '',
          material_cost: r[9] || 0,
          work_cost: r[10] || 0,
          total_cost: r[11] || 0,
          wholesale_price: r[12] || 0,
          sale_price: r[13] || 0,
          profit: r[14] || 0,
          weight: r[15] || 0,
          author: r[16] || '',
          parts_summary: r[17] || ''
        });
      }
      return jsonResponse({ status: 'ok', count: rows.length, data: rows });
    }

    // 3. Получение списка пользователей
    if (action === 'get_users') {
      var uSheet = getOrCreateUsersSheet(ss);
      var uData = uSheet.getDataRange().getValues();
      var users = [];
      for (var j = 1; j < uData.length; j++) {
        var u = uData[j];
        if (!u[1]) continue;
        users.push({
          id: u[0],
          login: u[1],
          role: u[2],
          name: u[4],
          status: u[5],
          created_at: u[6],
          last_login: u[7]
        });
      }
      return jsonResponse({ status: 'ok', count: users.length, users: users });
    }

    return jsonResponse({ status: 'ok', message: 'BLSTCS API active' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ================= POST ЗАПРОСЫ =================
function doPost(e) {
  try {
    var body = e.postData ? e.postData.contents : '';
    var payload = {};
    if (body) {
      try {
        payload = JSON.parse(body);
      } catch (parseErr) {
        payload = {};
      }
    }
    if ((!payload || Object.keys(payload).length === 0) && e.parameter) {
      payload = e.parameter;
    }

    var action = payload.action || 'save_history';
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Сохранение расчетов с дедупликацией (защита от повторных дубликатов)
    if (action === 'save_history') {
      var histSheet = getOrCreateHistorySheet(ss);
      var records = payload.records || (payload.record ? [payload.record] : []);
      
      // Читаем уже существующие ключи для предотвращения дублирования строк
      var existingData = histSheet.getDataRange().getValues();
      var existingKeys = {};
      for (var k = 1; k < existingData.length; k++) {
        var exRow = existingData[k];
        var idKey = exRow[0] ? exRow[0].toString() : '';
        var matchKey = (exRow[1] || '') + '|' + (exRow[2] || '') + '|' + (exRow[4] || '') + '|' + (exRow[5] || '');
        if (idKey) existingKeys[idKey] = true;
        if (matchKey) existingKeys[matchKey] = true;
      }

      var addedCount = 0;
      for (var i = 0; i < records.length; i++) {
        var r = records[i];
        if (!r) continue;

        var rIdKey = r.id ? r.id.toString() : '';
        var rMatchKey = (r.created_at || '') + '|' + (r.product_name || '') + '|' + (r.material || '') + '|' + (r.layers || '');

        if (existingKeys[rIdKey] || existingKeys[rMatchKey]) {
          continue; // Уже есть в таблице, пропускаем дубликат
        }
        existingKeys[rIdKey] = true;
        existingKeys[rMatchKey] = true;

        histSheet.appendRow([
          r.id || (Date.now() + '_' + i),
          r.created_at || new Date().toLocaleString("ru-RU"),
          r.product_name || 'Баллистический пакет',
          r.product_code || '',
          r.material || '',
          r.layers || '',
          r.single_area ? Number(r.single_area).toFixed(4) : '',
          r.cover_summary || r.cover_material || '',
          r.accessories_summary || '',
          r.material_cost ? Math.round(r.material_cost) : '',
          r.work_cost ? Math.round(r.work_cost) : '',
          r.total_cost ? Math.round(r.total_cost) : '',
          r.wholesale_price ? Math.round(r.wholesale_price) : '',
          r.sale_price ? Math.round(r.sale_price) : '',
          r.profit ? Math.round(r.profit) : '',
          r.weight ? Number(r.weight).toFixed(3) : '',
          r.author || payload.device_name || 'Пользователь',
          r.parts_summary || ''
        ]);
        addedCount++;
      }

      return jsonResponse({
        status: 'ok',
        message: 'Успешно добавлено новых строк: ' + addedCount,
        added_count: addedCount,
        total_in_sheet: histSheet.getLastRow() - 1
      });
    }

    // 2. Проверка пользователя / вход
    if (action === 'verify_user') {
      var userSheet = getOrCreateUsersSheet(ss);
      var allUsers = userSheet.getDataRange().getValues();
      var inputLogin = (payload.login || '').trim().toLowerCase();
      var inputHash = payload.password_hash || '';

      for (var u = 1; u < allUsers.length; u++) {
        var row = allUsers[u];
        var rowLogin = (row[1] || '').toString().trim().toLowerCase();
        var rowRole = (row[2] || 'Сотрудник').toString().trim();
        var rowHash = (row[3] || '').toString().trim();
        var rowStatus = (row[5] || 'Активен').toString().trim();

        if (rowLogin === inputLogin) {
          if (rowStatus !== 'Активен') {
            return jsonResponse({ status: 'blocked', message: 'Учетная запись заблокирована' });
          }
          if (rowHash === inputHash) {
            userSheet.getRange(u + 1, 8).setValue(new Date().toLocaleString("ru-RU"));
            return jsonResponse({
              status: 'ok',
              login: row[1],
              role: rowRole,
              is_admin: (rowRole.toLowerCase().indexOf('админ') !== -1 || rowRole.toLowerCase() === 'admin'),
              name: row[4] || row[1]
            });
          } else {
            return jsonResponse({ status: 'wrong_password', message: 'Неверный пароль' });
          }
        }
      }
      return jsonResponse({ status: 'not_found', message: 'Пользователь не найден' });
    }

    return jsonResponse({ status: 'error', message: 'Неизвестное действие: ' + action });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 1. Создание листа "История расчетов"
function getOrCreateHistorySheet(ss) {
  var sheet = ss.getSheetByName(SHEET_HISTORY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HISTORY);
    var headers = [
      'ID записи',
      'Дата и время',
      'Наименование изделия',
      'Артикул / Код',
      'Материал',
      'Слоёв',
      'Площадь (м²)',
      'Чехол (ткани)',
      'Фурнитура / Опции',
      'Себестоимость матер. (руб)',
      'Работа (руб)',
      'Итого себестоимость (руб)',
      'Оптовая цена (руб)',
      'Реком. розница (руб)',
      'Прибыль розницы (руб)',
      'Вес (кг)',
      'Автор / Устройство',
      'Детали кроя'
    ];
    sheet.appendRow(headers);

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0f172a');
    headerRange.setFontColor('#f8fafc');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    var widths = [110, 140, 200, 120, 100, 70, 100, 160, 160, 130, 100, 140, 130, 130, 130, 80, 130, 220];
    for (var i = 0; i < widths.length; i++) {
      sheet.setColumnWidth(i + 1, widths[i]);
    }
  }
  return sheet;
}

// 2. Создание листа "Пользователи"
function getOrCreateUsersSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    var headers = [
      'ID',
      'Логин',
      'Роль',
      'Хэш пароля (SHA-256)',
      'Имя / Должность',
      'Статус',
      'Дата добавления',
      'Последний вход'
    ];
    sheet.appendRow(headers);

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1e293b');
    headerRange.setFontColor('#f8fafc');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    var widths = [90, 140, 130, 280, 200, 100, 140, 140];
    for (var i = 0; i < widths.length; i++) {
      sheet.setColumnWidth(i + 1, widths[i]);
    }

    sheet.appendRow([
      'USR-001',
      'admin',
      'Администратор',
      '75ad31cc31193c26bff6c1b871819fb5004ef13f0c6e94a3ae8e76743641d2db',
      'Главный администратор',
      'Активен',
      new Date().toLocaleString("ru-RU"),
      ''
    ]);

    sheet.appendRow([
      'USR-002',
      'user',
      'Сотрудник',
      'b7b260d7fe03c0e686972dc206d1270d8330e5fb38b960710bd3e40c3dfba561',
      'Менеджер расчетов',
      'Активен',
      new Date().toLocaleString("ru-RU"),
      ''
    ]);
  }
  return sheet;
}
