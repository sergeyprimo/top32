
// Удаление мусорных дубликатов и конфликтных вкладок
function cleanupGarbageSheets(ss) {
  var deleted = [];
  try {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var s = sheets[i];
      var name = s.getName();
      // Удаляем любые вкладки с _conflict и временные автоматические архивы с таймштампами
      if (name.indexOf('_conflict') !== -1 || 
          name.indexOf('Архив этапа (') !== -1 || 
          name.indexOf('Итоги Финала (') !== -1 ||
          name.indexOf('Итоги Этапа 1 (Архив)') !== -1 ||
          name === '📁 Полные результаты сеток') {
        if (ss.getSheets().length > 1) {
          try {
            ss.deleteSheet(s);
            deleted.push(name);
          } catch(eDel) {}
        }
      }
    }
  } catch(e) {}
  return deleted;
}

// ================================================================
// СКРИПТ АВАРИЙНОГО РЕЗЕРВНОГО КОНТУРА ДЛЯ GOOGLE ТАБЛИЦЫ
// Проект: DRIFT BANDA / ГАЗДЕРЖИМ.РФ
// ВАРИАНТ 1: ЕДИНЫЙ СКВОЗНОЙ ПРОТОКОЛ (ВСЕ СТАДИИ) + АРХИВ ЭТАПОВ
//
// Инструкция по обновлению / установке:
// 1. Откройте вашу Google Таблицу (DRIFT BANDA - РЕЗЕРВНАЯ СЕТКА И КВАЛИФИКАЦИЯ)
// 2. В верхнем меню выберите: Расширения (Extensions) -> Apps Script
// 3. Сотрите весь старый код в редакторе и вставьте этот код целиком
// 4. Справа вверху нажмите: Развернуть (Deploy) -> Управление развертываниями (Manage deployments)
// 5. Нажмите на значок карандаша (Изменить) у текущего активного веб-приложения:
//    - Версия: Новая версия (New version)
//    - Нажмите синюю кнопку "Развернуть" (Deploy) и затем "Готово"
// 6. Всё готово!
//    - Каждые 30 сек в «🏆 Сетка заездов» сквозным списком пишутся ВСЕ стадии (ТОП-32, ТОП-24, ТОП-16, ТОП-8, Полуфиналы, Финал).
//    - По кнопке «📁 Сохранить этап» создаётся отдельная постоянная архивная вкладка со всеми результатами и подиумом призёров!
// ================================================================

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = (e && e.parameter && e.parameter.action) || 'ping';

    // Ручная или автоматическая очистка мусорных вкладок
    if (action === 'cleanup_garbage_sheets' || action === 'cleanup_conflicts') {
      var deletedTabs = cleanupGarbageSheets(ss);
      logAudit(ss, (e && e.parameter && e.parameter.author) || 'Система', 'Очистка мусорных вкладок', 'Удалено дубликатов: ' + deletedTabs.length);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        version: '2.4',
        message: 'Мусорные вкладки успешно удалены!',
        deleted: deletedTabs,
        sheets: ss.getSheets().map(function(s) { return s.getName(); })
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Запрос экстренного восстановления в судейскую панель
    if (action === 'get_backup') {
      var snapshot = PropertiesService.getScriptProperties().getProperty('tourney_snapshot');
      if (!snapshot) {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'empty',
          message: 'Резервная копия в таблице пока не сохранена'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        version: '2.4',
        data: JSON.parse(snapshot)
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Создание архивной вкладки через GET-запрос (для прямой проверки)
    if (action === 'archive_stage' || action === 'save_stage_archive') {
      var tabName = (e && e.parameter && e.parameter.tab_name) || ('🏆 Итоги ' + new Date().toLocaleDateString('ru-RU'));
      var snapStr = PropertiesService.getScriptProperties().getProperty('tourney_snapshot');
      var snapPayload = snapStr ? JSON.parse(snapStr) : {};
      var createdTabName = createArchiveStageSheet(ss, snapPayload, tabName);
      logAudit(ss, (e && e.parameter && e.parameter.author) || 'Судья/Админ (GET)', 'Сохранение этапа в архив', 'Создана отдельная вкладка «' + createdTabName + '» со всеми результатами турнира');
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        version: '2.4',
        message: 'Этап успешно зафиксирован в отдельную вкладку «' + createdTabName + '»!',
        tab_name: createdTabName,
        sheets: ss.getSheets().map(function(s) { return s.getName(); }),
        archived_at: new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' })
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Стандартный пинг сервиса и проверка версии (CORS-friendly)
    return ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      version: '2.4',
      service: 'Drift Banda Emergency Backup Service (Single Dedicated Archive Tab v2.4)',
      has_archive_support: true,
      time: new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' }),
      sheets: ss.getSheets().map(function(s) { return s.getName(); })
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch(err) {
        payload = {};
      }
    }

    var action = payload.action || (e && e.parameter && e.parameter.action) || 'sync_tourney';

    // 1. Фиксация и сохранение завершённого этапа в отдельную вкладку
    if (action === 'archive_stage' || action === 'save_stage_archive') {
      var tabName = payload.archive_tab_name || (e && e.parameter && e.parameter.tab_name) || ('🏆 Итоги Этапа ' + new Date().toLocaleDateString('ru-RU'));

      // Если в payload часть данных отсутствует (например, усеченный POST), дополняем из последнего сохранённого снапшота
      var snapStr = PropertiesService.getScriptProperties().getProperty('tourney_snapshot');
      if (snapStr) {
        try {
          var snapObj = JSON.parse(snapStr);
          if ((!payload.battles || payload.battles.length === 0) && snapObj.battles) payload.battles = snapObj.battles;
          if ((!payload.battles_top24 || payload.battles_top24.length === 0) && snapObj.battles_top24) payload.battles_top24 = snapObj.battles_top24;
          if ((!payload.battles_top16 || payload.battles_top16.length === 0) && snapObj.battles_top16) payload.battles_top16 = snapObj.battles_top16;
          if ((!payload.battles_top8 || payload.battles_top8.length === 0) && snapObj.battles_top8) payload.battles_top8 = snapObj.battles_top8;
          if ((!payload.battles_top4 || payload.battles_top4.length === 0) && snapObj.battles_top4) payload.battles_top4 = snapObj.battles_top4;
          if ((!payload.qualification || payload.qualification.length === 0) && snapObj.qualification) payload.qualification = snapObj.qualification;
          if (!payload.podium && snapObj.podium) payload.podium = snapObj.podium;
          if (!payload.meta && snapObj.meta) payload.meta = snapObj.meta;
        } catch(errSnap) {}
      }

      var createdTabName = createArchiveStageSheet(ss, payload, tabName);
      logAudit(ss, (payload.meta && payload.meta.author) || 'Судья/Админ', 'Сохранение этапа в архив', 'Создана отдельная вкладка «' + createdTabName + '» со всеми результатами турнира');
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        version: '2.4',
        message: 'Этап успешно зафиксирован в отдельную вкладку «' + createdTabName + '»!',
        tab_name: createdTabName,
        sheets: ss.getSheets().map(function(s) { return s.getName(); }),
        archived_at: new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' })
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Очистка резервной таблицы по кнопке администратора
    if (action === 'clear_backup' || action === 'clear_all') {
      PropertiesService.getScriptProperties().deleteProperty('tourney_snapshot');
      syncQualiSheet(ss, [], payload.meta || {});
      syncBattlesSheet(ss, {}, payload.meta || {});
      logAudit(ss, (payload.meta && payload.meta.author) || 'Администратор', 'Очистка резервной таблицы', 'Все результаты квалификации и сетки в Google Таблице очищены');
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        message: 'Резервная Google Таблица успешно очищена!',
        cleared_at: new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' })
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Запрос бэкапа
    if (action === 'get_backup') {
      var snap = PropertiesService.getScriptProperties().getProperty('tourney_snapshot');
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        data: snap ? JSON.parse(snap) : null
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. Стандартная 30-секундная синхронизация текущего турнира
    // Сохранение снапшота для экстренного восстановления
    if (hasRealBracketData(payload)) {
      PropertiesService.getScriptProperties().setProperty('tourney_snapshot', JSON.stringify(payload));
    }

    var quali = payload.qualification || [];
    var meta = payload.meta || {};

    // Обновляем визуальный лист «🏁 Квалификация»
    syncQualiSheet(ss, quali, meta);

    // Обновляем визуальный лист «🏆 Сетка заездов» (Все стадии сквозным протоколом: ТОП-32, ТОП-24, ТОП-16, ТОП-8, Полуфиналы, Финал)
    syncBattlesSheet(ss, payload, meta);

    // Обновляем отдельную постоянную вкладку «📁 Полные результаты сеток» (Все стадии + Подиум)
    syncAllBracketsArchiveSheet(ss, payload, meta);

    // Добавляем запись в лист «🛡️ Аудит-лог»
    var b32Len = (payload.battles && payload.battles.length) || 0;
    var b24Len = (payload.battles_top24 && payload.battles_top24.length) || 0;
    var b16Len = (payload.battles_top16 && payload.battles_top16.length) || 0;
    var b8Len = (payload.battles_top8 && payload.battles_top8.length) || 0;
    var b4Len = (payload.battles_top4 && payload.battles_top4.length) || 0;
    var totalBattles = b32Len + b24Len + b16Len + b8Len + b4Len;
    logAudit(ss, meta.author || 'Судья/Админ', 'Синхронизация турнира (LIVE)', 'Пилотов квалы: ' + quali.length + ' | Всего заездов: ' + totalBattles + ' | Активный этап: ' + (meta.stage || '—'));

    return ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      message: 'Все результаты и стадии турнира успешно продублированы в Google Таблицу!',
      updated_at: new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' })
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function syncQualiSheet(ss, quali, meta) {
  var sheetName = '🏁 Квалификация';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();

  var nowStr = new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' });
  var tourneyName = (meta && meta.tournament_name) || 'DRIFT BANDA';
  
  // Верхний баннер квалификации
  sheet.appendRow(['🏁 ' + tourneyName + ' — ОФИЦИАЛЬНАЯ ТАБЛИЦА КВАЛИФИКАЦИИ']);
  sheet.getRange(1, 1, 1, 10).merge()
    .setBackground('#090d16')
    .setFontColor('#f59e0b')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  sheet.appendRow(['Синхронизация LIVE (каждые 30 сек) | Время обновления: ' + nowStr + ' (МСК) | Всего участников: ' + (quali ? quali.length : 0)]);
  sheet.getRange(2, 1, 1, 10).merge()
    .setBackground('#1e293b')
    .setFontColor('#94a3b8')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 22);

  var headers = [
    'Место', 'Семя (Seed)', 'Пилот (ФИО / Никнейм)', 'Автомобиль', 'Группа',
    'Попытка 1', 'Попытка 2', 'Лучший балл', 'Сумма', 'Статус'
  ];
  sheet.appendRow(headers);

  var headerRange = sheet.getRange(3, 1, 1, headers.length);
  headerRange.setBackground('#0f172a');
  headerRange.setFontColor('#fbbf24');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sheet.setRowHeight(3, 26);
  sheet.setFrozenRows(3);

  if (!quali || quali.length === 0) return;

  var rows = [];
  for (var i = 0; i < quali.length; i++) {
    var p = quali[i];
    var rank = p.rank || (i + 1);
    var statusText = (p.status === 'qualified' || rank <= 32) ? 'ТОП-32' : (p.status === 'eliminated' ? 'Выбыл' : (p.status || 'В ожидании'));
    if (p.is_health_group) statusText = 'Группа Здоровья';
    
    rows.push([
      rank,
      p.seed || (i + 1),
      p.name || '—',
      p.car || '—',
      p.group || 'A',
      p.run1 !== undefined ? p.run1 : 0,
      p.run2 !== undefined ? p.run2 : 0,
      p.best !== undefined ? p.best : 0,
      p.total !== undefined ? p.total : 0,
      statusText
    ]);
  }

  if (rows.length > 0) {
    var dataRange = sheet.getRange(4, 1, rows.length, headers.length);
    dataRange.setValues(rows);
    dataRange.setFontSize(10);
    sheet.getRange(4, 1, rows.length, 2).setHorizontalAlignment('center');
    sheet.getRange(4, 3, rows.length, 2).setHorizontalAlignment('left');
    sheet.getRange(4, 5, rows.length, 6).setHorizontalAlignment('center');
    sheet.getRange(4, 8, rows.length, 1).setFontWeight('bold');

    for (var rIdx = 0; rIdx < rows.length; rIdx++) {
      sheet.setRowHeight(4 + rIdx, 22);
    }
  }

  try { 
    sheet.autoResizeColumns(1, headers.length);
    if (sheet.getColumnWidth(3) < 180) sheet.setColumnWidth(3, 180);
    if (sheet.getColumnWidth(4) < 180) sheet.setColumnWidth(4, 180);
  } catch(e) {}
}

// ================================================================
// ОТДЕЛЬНАЯ ПОСТОЯННАЯ ВКЛАДКА С ПОЛНЫМИ СОХРАНЕНИЯМИ СЕТОК ЗАЕЗДОВ
// ================================================================

// Проверка наличия реальных данных сетки турнира (защита от затирания пустыми сбросами)
function hasRealBracketData(payload) {
  if (!payload) return false;
  if (payload.brackets_cleared === true || payload.is_bracket_cleared === true) return false;

  var stages = [
    payload.battles,
    payload.battles_top24,
    payload.battles_top16,
    payload.battles_top8,
    payload.battles_top4
  ];

  for (var s = 0; s < stages.length; s++) {
    var arr = stages[s];
    if (arr && arr.length > 0) {
      for (var i = 0; i < arr.length; i++) {
        var b = arr[i];
        if (!b) continue;
        if (b.winner) return true;
        if (b.status === 'finished') return true;
        if (b.judges && b.judges.some(function(j) { return j && j.vote; })) return true;
        var l = b.lead;
        var c = b.chase;
        if (l && l.name && l.name !== 'ОЖИДАНИЕ ПИЛОТА' && l.name !== '—' && !l.is_placeholder) return true;
        if (c && c.name && c.name !== 'ОЖИДАНИЕ ПИЛОТА' && c.name !== '—' && !c.is_placeholder) return true;
      }
    }
  }

  return false;
}

function syncAllBracketsArchiveSheet(ss, payload, meta) {
  var sheetName = '📁 Полный архив этапа';
  var sheet = ss.getSheetByName(sheetName);

  // ЗАЩИТА ОТ ЗАТИРАНИЯ ПУСТЫМИ ДАННЫМИ:
  // Если судья очистил сетку (или пришел пустой сброс), ни в коем случае не стираем
  // уже заполненный архив прошлого этапа!
  var isReal = hasRealBracketData(payload);
  if (!isReal) {
    if (sheet && sheet.getLastRow() > 5) {
      // Предыдущий этап сохранен, затирать пустышкой запрещено!
      logAudit(ss, (meta && meta.author) || 'Система', 'Защита архива', 'Сетка очищена в судейской. Вкладка «📁 Полные результаты сеток» защищена от затирания!');
      return;
    }
  }

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();

  var nowStr = new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' });
  var tourneyName = (meta && meta.tournament_name) || 'DRIFT BANDA';
  var author = (meta && meta.author) || 'Судейская коллегия';

  // 1. Главный баннер вкладки полного протокола
  sheet.appendRow(['📁 ' + tourneyName + ' — ПОЛНЫЙ ПРОТОКОЛ СЕТОК ЗАЕЗДОВ (ВСЕ СТАДИИ)']);
  sheet.getRange(1, 1, 1, 12).merge()
    .setBackground('#090d16')
    .setFontColor('#38bdf8')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  sheet.appendRow(['Синхронизация результатов LIVE | Обновлено: ' + nowStr + ' (МСК) | Все стадии турнира: ТОП-32, ТОП-24, ТОП-16, ТОП-8, Полуфиналы, Финал']);
  sheet.getRange(2, 1, 1, 12).merge()
    .setBackground('#1e293b')
    .setFontColor('#94a3b8')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 22);
  sheet.setFrozenRows(2);

  var currentRow = 3;

  // 2. Блок Подиума (если есть победители)
  var podium = payload.podium;
  if (!podium && payload.battles_top4 && payload.battles_top4.length >= 4) {
    podium = { first: null, second: null, third: null, fourth: null };
    var bBronze = payload.battles_top4[2];
    var bFinal = payload.battles_top4[3];
    if (bFinal && bFinal.winner) {
      var wLead = (bFinal.winner === 'lead');
      podium.first = wLead ? bFinal.lead : bFinal.chase;
      podium.second = wLead ? bFinal.chase : bFinal.lead;
    }
    if (bBronze && bBronze.winner) {
      var wLeadB = (bBronze.winner === 'lead');
      podium.third = wLeadB ? bBronze.lead : bBronze.chase;
      podium.fourth = wLeadB ? bBronze.chase : bBronze.lead;
    }
  }

  if (podium && (podium.first || podium.second || podium.third)) {
    sheet.appendRow(['']);
    sheet.setRowHeight(currentRow, 10);
    currentRow++;

    sheet.appendRow(['🎖️ ИТОГОВЫЙ ПОДИУМ ПОБЕДИТЕЛЕЙ ТУРНИРА']);
    sheet.getRange(currentRow, 1, 1, 12).merge()
      .setBackground('#78350f')
      .setFontColor('#fef08a')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 26);
    currentRow++;

    var pRows = [];
    if (podium.first && podium.first.name) pRows.push(['🥇 1-е МЕСТО (ЧЕМПИОН)', podium.first.name, podium.first.car || '—', '🏆 ЗОЛОТОЙ КУБОК']);
    if (podium.second && podium.second.name) pRows.push(['🥈 2-е МЕСТО (ВИЦЕ-ЧЕМПИОН)', podium.second.name, podium.second.car || '—', 'СЕРЕБРЯНЫЙ ПРИЗЁР']);
    if (podium.third && podium.third.name) pRows.push(['🥉 3-е МЕСТО (БРОНЗА)', podium.third.name, podium.third.car || '—', 'БРОНЗОВЫЙ ПРИЗЁР']);
    if (podium.fourth && podium.fourth.name) pRows.push(['4-е МЕСТО (ТОП-4)', podium.fourth.name, podium.fourth.car || '—', 'ПОЛУФИНАЛИСТ']);

    for (var pi = 0; pi < pRows.length; pi++) {
      var pr = pRows[pi];
      sheet.appendRow([pr[0], '', pr[1], '', '', pr[2], '', '', pr[3], '', '', '']);
      sheet.getRange(currentRow, 1, 1, 2).merge().setHorizontalAlignment('center').setFontWeight('bold');
      sheet.getRange(currentRow, 3, 1, 3).merge().setHorizontalAlignment('left').setFontWeight('bold');
      sheet.getRange(currentRow, 6, 1, 3).merge().setHorizontalAlignment('left');
      sheet.getRange(currentRow, 9, 1, 4).merge().setHorizontalAlignment('center').setFontWeight('bold');

      var bg = pi === 0 ? '#fef3c7' : (pi === 1 ? '#f1f5f9' : (pi === 2 ? '#ffedd5' : '#f8fafc'));
      var fg = pi === 0 ? '#78350f' : (pi === 1 ? '#1e293b' : (pi === 2 ? '#9a3412' : '#475569'));
      sheet.getRange(currentRow, 1, 1, 12).setBackground(bg).setFontColor(fg);
      sheet.setRowHeight(currentRow, 25);
      currentRow++;
    }
  }

  // 3. Таблицы всех стадий сеток
  var headers = [
    'Стадия', 'Хит №', 'Название пары',
    'Пилот (Лидер)', 'Автомобиль (Лидер)',
    'Пилот (Преследователь)', 'Автомобиль (Преследователь)',
    'Судья #1', 'Судья #2', 'Судья #3',
    'Победитель', 'Статус'
  ];

  var stageConfigs = [];
  var b32 = (payload && (payload.battles || payload.battles_top32)) || [];
  if (Array.isArray(b32) && b32.length > 0) stageConfigs.push({ title: '🟦 СЕТКА ТОП-32 (16 ПАРНЫХ ХИТОВ)', bgColor: '#1e3a8a', fgColor: '#bfdbfe', battles: b32, defaultStage: 'ТОП-32' });
  var b24 = (payload && payload.battles_top24) || [];
  if (Array.isArray(b24) && b24.length > 0) stageConfigs.push({ title: '🟪 СЕТКА ТОП-24 (8 ПАРНЫХ ХИТОВ)', bgColor: '#581c87', fgColor: '#e9d5ff', battles: b24, defaultStage: 'ТОП-24' });
  var b16 = (payload && payload.battles_top16) || [];
  if (Array.isArray(b16) && b16.length > 0) stageConfigs.push({ title: '🟩 СЕТКА ТОП-16 (8 ПАРНЫХ ХИТОВ)', bgColor: '#064e3b', fgColor: '#a7f3d0', battles: b16, defaultStage: 'ТОП-16' });
  var b8 = (payload && payload.battles_top8) || [];
  if (Array.isArray(b8) && b8.length > 0) stageConfigs.push({ title: '🟨 СЕТКА ТОП-8 — GREAT 8 (4 ПАРНЫХ ХИТА)', bgColor: '#78350f', fgColor: '#fef08a', battles: b8, defaultStage: 'ТОП-8' });
  var b4 = (payload && payload.battles_top4) || [];
  if (Array.isArray(b4) && b4.length > 0) {
    var semis = [];
    var finals = [];
    for (var k = 0; k < b4.length; k++) {
      var item = b4[k];
      if (item.stage === 'semifinal' || (item.id && item.id <= 302) || k < 2) semis.push(item);
      else finals.push(item);
    }
    if (semis.length > 0) stageConfigs.push({ title: '🟧 ПОЛУФИНАЛЫ — ТОП-4 (2 ХИТА)', bgColor: '#7c2d12', fgColor: '#fed7aa', battles: semis, defaultStage: 'Полуфинал' });
    if (finals.length > 0) stageConfigs.push({ title: '👑 ФИНАЛЫ: ЗАЕЗД ЗА 3-Е МЕСТО И ГРАНД-ФИНАЛ 🏆', bgColor: '#831843', fgColor: '#fbcfe8', battles: finals, defaultStage: 'Финал' });
  }

  function getJudgeVoteStr(b, idx, key) {
    if (b.judges && b.judges[idx] && b.judges[idx].vote) {
      var v = String(b.judges[idx].vote).toLowerCase();
      if (v === 'lead') return 'LEAD';
      if (v === 'chase') return 'CHASE';
      if (v === 'omt') return 'OMT';
      return v.toUpperCase();
    }
    if (b.judge_votes && b.judge_votes[key]) {
      var v2 = String(b.judge_votes[key]).toLowerCase();
      if (v2 === 'lead') return 'LEAD';
      if (v2 === 'chase') return 'CHASE';
      if (v2 === 'omt') return 'OMT';
      return v2.toUpperCase();
    }
    return '—';
  }

  for (var sIdx = 0; sIdx < stageConfigs.length; sIdx++) {
    var section = stageConfigs[sIdx];
    sheet.appendRow(['']);
    sheet.setRowHeight(currentRow, 12);
    currentRow++;

    sheet.appendRow([section.title]);
    var secRange = sheet.getRange(currentRow, 1, 1, 12);
    secRange.merge().setBackground(section.bgColor).setFontColor(section.fgColor).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 28);
    currentRow++;

    sheet.appendRow(headers);
    var headRange = sheet.getRange(currentRow, 1, 1, 12);
    headRange.setBackground('#0f172a').setFontColor('#f8fafc').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 24);
    currentRow++;

    var bList = section.battles;
    var rows = [];
    for (var i = 0; i < bList.length; i++) {
      var b = bList[i];
      var stageDisplay = section.defaultStage;
      if (b.stage === 'bronze') stageDisplay = '🥉 За 3-е место';
      else if (b.stage === 'final') stageDisplay = '🥇 Гранд-Финал';
      else if (b.stage === 'semifinal') stageDisplay = 'Полуфинал';
      else if (b.stage === 'top8') stageDisplay = 'ТОП-8';
      else if (b.stage === 'top16') stageDisplay = 'ТОП-16';
      else if (b.stage === 'top24') stageDisplay = 'ТОП-24';
      else if (b.stage === 'top32') stageDisplay = 'ТОП-32';

      var leadName = (b.lead && b.lead.name) ? b.lead.name : ((b.lead && b.lead.is_placeholder) ? '—' : 'Ожидание пилота');
      var leadCar = (b.lead && b.lead.car && b.lead.car !== '---') ? b.lead.car : '—';
      var chaseName = (b.chase && b.chase.name) ? b.chase.name : ((b.chase && b.chase.is_placeholder) ? '—' : 'Ожидание пилота');
      var chaseCar = (b.chase && b.chase.car && b.chase.car !== '---') ? b.chase.car : '—';

      var v1 = getJudgeVoteStr(b, 0, 'judge1');
      var v2 = getJudgeVoteStr(b, 1, 'judge2');
      var v3 = getJudgeVoteStr(b, 2, 'judge3');

      var winnerName = '—';
      if (b.winner === 'lead') winnerName = '🏆 ' + leadName;
      else if (b.winner === 'chase') winnerName = '🏆 ' + chaseName;
      else if (b.winner === 'omt') winnerName = '⚡ OMT (Перезаезд)';
      else if (b.winner) winnerName = b.winner;

      var statusStr = 'Ожидает';
      if (b.status === 'live') statusStr = '🔴 LIVE';
      else if (b.status === 'finished') statusStr = 'Завершён';
      else if (b.status === 'omt') statusStr = '⚡ OMT';
      else if (b.status === 'waiting' || b.status === 'upcoming') statusStr = 'Ожидает';

      rows.push([
        stageDisplay,
        b.pair_num || b.id || (i + 1),
        b.pair_title || ('Хит #' + (i + 1)),
        leadName,
        leadCar,
        chaseName,
        chaseCar,
        v1,
        v2,
        v3,
        winnerName,
        statusStr
      ]);
    }

    if (rows.length > 0) {
      var dataRange = sheet.getRange(currentRow, 1, rows.length, 12);
      dataRange.setValues(rows);
      dataRange.setFontSize(10).setVerticalAlignment('middle');
      sheet.getRange(currentRow, 1, rows.length, 3).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 4, rows.length, 4).setHorizontalAlignment('left');
      sheet.getRange(currentRow, 8, rows.length, 3).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 11, rows.length, 2).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 11, rows.length, 1).setFontWeight('bold');
      sheet.getRange(currentRow, 12, rows.length, 1).setFontWeight('bold');

      for (var rIdx = 0; rIdx < rows.length; rIdx++) sheet.setRowHeight(currentRow + rIdx, 22);
      currentRow += rows.length;
    }
  }

  // 4. Итоговая квалификация внизу полного протокола
  var quali = payload.qualification || [];
  if (quali.length > 0) {
    sheet.appendRow(['']);
    sheet.setRowHeight(currentRow, 14);
    currentRow++;

    sheet.appendRow(['🏁 ИТОГОВАЯ ТАБЛИЦА КВАЛИФИКАЦИИ']);
    sheet.getRange(currentRow, 1, 1, 10).merge()
      .setBackground('#0f172a')
      .setFontColor('#fbbf24')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 26);
    currentRow++;

    var qHeaders = [
      'Место', 'Семя (Seed)', 'Пилот (ФИО / Никнейм)', 'Автомобиль', 'Группа',
      'Попытка 1', 'Попытка 2', 'Лучший балл', 'Сумма', 'Статус'
    ];
    sheet.appendRow(qHeaders);
    sheet.getRange(currentRow, 1, 1, 10)
      .setBackground('#1e293b')
      .setFontColor('#94a3b8')
      .setFontWeight('bold')
      .setFontSize(9)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 22);
    currentRow++;

    var qRows = [];
    for (var qi = 0; qi < quali.length; qi++) {
      var qp = quali[qi];
      var qRank = qp.rank || (qi + 1);
      var qStatus = (qp.status === 'qualified' || qRank <= 32) ? 'ТОП-32' : (qp.status === 'eliminated' ? 'Выбыл' : (qp.status || 'В ожидании'));
      if (qp.is_health_group) qStatus = 'Группа Здоровья';
      qRows.push([
        qRank,
        qp.seed || (qi + 1),
        qp.name || '—',
        qp.car || '—',
        qp.group || 'A',
        qp.run1 !== undefined ? qp.run1 : 0,
        qp.run2 !== undefined ? qp.run2 : 0,
        qp.best !== undefined ? qp.best : 0,
        qp.total !== undefined ? qp.total : 0,
        qStatus
      ]);
    }

    if (qRows.length > 0) {
      var qDataRange = sheet.getRange(currentRow, 1, qRows.length, 10);
      qDataRange.setValues(qRows);
      qDataRange.setFontSize(9).setVerticalAlignment('middle');
      sheet.getRange(currentRow, 1, qRows.length, 2).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 3, qRows.length, 2).setHorizontalAlignment('left');
      sheet.getRange(currentRow, 5, qRows.length, 6).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 8, qRows.length, 1).setFontWeight('bold');

      for (var qIdx = 0; qIdx < qRows.length; qIdx++) sheet.setRowHeight(currentRow + qIdx, 20);
      currentRow += qRows.length;
    }
  }

  try {
    sheet.autoResizeColumns(1, 12);
    if (sheet.getColumnWidth(4) < 170) sheet.setColumnWidth(4, 170);
    if (sheet.getColumnWidth(5) < 170) sheet.setColumnWidth(5, 170);
    if (sheet.getColumnWidth(6) < 170) sheet.setColumnWidth(6, 170);
    if (sheet.getColumnWidth(7) < 170) sheet.setColumnWidth(7, 170);
    if (sheet.getColumnWidth(11) < 190) sheet.setColumnWidth(11, 190);
    if (sheet.getColumnWidth(12) < 110) sheet.setColumnWidth(12, 110);
  } catch(e) {}
}

function syncBattlesSheet(ss, payload, meta) {
  var sheetName = '🏆 Сетка заездов';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();

  var nowStr = new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' });
  var tourneyName = (meta && meta.tournament_name) || 'DRIFT BANDA';
  var activeStage = (meta && meta.stage) || 'top32';
  var stageTitles = {
    'quali': 'КВАЛИФИКАЦИЯ',
    'top32': 'ТОП-32',
    'top24': 'ТОП-24',
    'top16': 'ТОП-16',
    'top8': 'ТОП-8',
    'top4': 'ФИНАЛ & ТОП-4'
  };
  var activeStageStr = stageTitles[activeStage] || activeStage.toUpperCase();

  // Верхний баннер турнира
  sheet.appendRow(['🏆 ' + tourneyName + ' — ОФИЦИАЛЬНАЯ СЕТКА ЗАЕЗДОВ (СКВОЗНОЙ ПРОТОКОЛ)']);
  sheet.getRange(1, 1, 1, 12).merge()
    .setBackground('#090d16')
    .setFontColor('#fbbf24')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  sheet.appendRow(['Синхронизация LIVE (каждые 30 сек) | Время обновления: ' + nowStr + ' (МСК) | Текущий активный этап: ' + activeStageStr + ' | Автор: ' + ((meta && meta.author) || 'Судейская коллегия')]);
  sheet.getRange(2, 1, 1, 12).merge()
    .setBackground('#1e293b')
    .setFontColor('#94a3b8')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 22);
  sheet.setFrozenRows(2);

  var headers = [
    'Стадия', 'Хит №', 'Название пары',
    'Пилот (Лидер)', 'Автомобиль (Лидер)',
    'Пилот (Преследователь)', 'Автомобиль (Преследователь)',
    'Судья #1', 'Судья #2', 'Судья #3',
    'Победитель', 'Статус'
  ];

  // Формируем все стадии по порядку:
  var stageConfigs = [];

  // 1. ТОП-32
  var b32 = (payload && (payload.battles || payload.battles_top32)) || [];
  if (Array.isArray(b32) && b32.length > 0) {
    stageConfigs.push({
      title: '🟦 СЕТКА ТОП-32 (16 ПАРНЫХ ХИТОВ)',
      bgColor: '#1e3a8a',
      fgColor: '#bfdbfe',
      battles: b32,
      defaultStage: 'ТОП-32'
    });
  }

  // 2. ТОП-24 (если есть)
  var b24 = (payload && payload.battles_top24) || [];
  if (Array.isArray(b24) && b24.length > 0) {
    stageConfigs.push({
      title: '🟪 СЕТКА ТОП-24 (8 ПАРНЫХ ХИТОВ)',
      bgColor: '#581c87',
      fgColor: '#e9d5ff',
      battles: b24,
      defaultStage: 'ТОП-24'
    });
  }

  // 3. ТОП-16
  var b16 = (payload && payload.battles_top16) || [];
  if (Array.isArray(b16) && b16.length > 0) {
    stageConfigs.push({
      title: '🟩 СЕТКА ТОП-16 (8 ПАРНЫХ ХИТОВ)',
      bgColor: '#064e3b',
      fgColor: '#a7f3d0',
      battles: b16,
      defaultStage: 'ТОП-16'
    });
  }

  // 4. ТОП-8 (Great 8)
  var b8 = (payload && payload.battles_top8) || [];
  if (Array.isArray(b8) && b8.length > 0) {
    stageConfigs.push({
      title: '🟨 СЕТКА ТОП-8 — GREAT 8 (4 ПАРНЫХ ХИТА)',
      bgColor: '#78350f',
      fgColor: '#fef08a',
      battles: b8,
      defaultStage: 'ТОП-8'
    });
  }

  // 5. ТОП-4 (Полуфиналы) и 6. Заезд за 3-е место + Гранд-Финал
  var b4 = (payload && payload.battles_top4) || [];
  if (Array.isArray(b4) && b4.length > 0) {
    var semis = [];
    var finals = [];
    for (var k = 0; k < b4.length; k++) {
      var item = b4[k];
      if (item.stage === 'semifinal' || (item.id && item.id <= 302) || k < 2) {
        semis.push(item);
      } else {
        finals.push(item);
      }
    }

    if (semis.length > 0) {
      stageConfigs.push({
        title: '🟧 ПОЛУФИНАЛЫ — ТОП-4 (2 ХИТА)',
        bgColor: '#7c2d12',
        fgColor: '#fed7aa',
        battles: semis,
        defaultStage: 'Полуфинал'
      });
    }

    if (finals.length > 0) {
      stageConfigs.push({
        title: '👑 ФИНАЛЫ: ЗАЕЗД ЗА 3-Е МЕСТО И ГРАНД-ФИНАЛ 🏆',
        bgColor: '#831843',
        fgColor: '#fbcfe8',
        battles: finals,
        defaultStage: 'Финал'
      });
    }
  }

  if (stageConfigs.length === 0) {
    sheet.appendRow(['Сетка заездов пока пуста или находится на этапе формирования']);
    return;
  }

  function getJudgeVoteStr(b, idx, key) {
    if (b.judges && b.judges[idx] && b.judges[idx].vote) {
      var v = String(b.judges[idx].vote).toLowerCase();
      if (v === 'lead') return 'LEAD';
      if (v === 'chase') return 'CHASE';
      if (v === 'omt') return 'OMT';
      return v.toUpperCase();
    }
    if (b.judge_votes && b.judge_votes[key]) {
      var v2 = String(b.judge_votes[key]).toLowerCase();
      if (v2 === 'lead') return 'LEAD';
      if (v2 === 'chase') return 'CHASE';
      if (v2 === 'omt') return 'OMT';
      return v2.toUpperCase();
    }
    return '—';
  }

  var currentRow = 3;

  for (var sIdx = 0; sIdx < stageConfigs.length; sIdx++) {
    var section = stageConfigs[sIdx];

    // Разделитель секции
    sheet.appendRow([section.title]);
    var secRange = sheet.getRange(currentRow, 1, 1, 12);
    secRange.merge()
      .setBackground(section.bgColor)
      .setFontColor(section.fgColor)
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 28);
    currentRow++;

    // Заголовки колонок стадии
    sheet.appendRow(headers);
    var headRange = sheet.getRange(currentRow, 1, 1, 12);
    headRange.setBackground('#0f172a')
      .setFontColor('#f8fafc')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 24);
    currentRow++;

    var bList = section.battles;
    var rows = [];

    for (var i = 0; i < bList.length; i++) {
      var b = bList[i];
      var stageDisplay = section.defaultStage;
      if (b.stage === 'bronze') stageDisplay = '🥉 За 3-е место';
      else if (b.stage === 'final') stageDisplay = '🥇 Гранд-Финал';
      else if (b.stage === 'semifinal') stageDisplay = 'Полуфинал';
      else if (b.stage === 'top8') stageDisplay = 'ТОП-8';
      else if (b.stage === 'top16') stageDisplay = 'ТОП-16';
      else if (b.stage === 'top24') stageDisplay = 'ТОП-24';
      else if (b.stage === 'top32') stageDisplay = 'ТОП-32';

      var leadName = (b.lead && b.lead.name) ? b.lead.name : ((b.lead && b.lead.is_placeholder) ? '—' : 'Ожидание пилота');
      var leadCar = (b.lead && b.lead.car && b.lead.car !== '---') ? b.lead.car : '—';
      var chaseName = (b.chase && b.chase.name) ? b.chase.name : ((b.chase && b.chase.is_placeholder) ? '—' : 'Ожидание пилота');
      var chaseCar = (b.chase && b.chase.car && b.chase.car !== '---') ? b.chase.car : '—';

      var v1 = getJudgeVoteStr(b, 0, 'judge1');
      var v2 = getJudgeVoteStr(b, 1, 'judge2');
      var v3 = getJudgeVoteStr(b, 2, 'judge3');

      var winnerName = '—';
      if (b.winner === 'lead') winnerName = '🏆 ' + leadName;
      else if (b.winner === 'chase') winnerName = '🏆 ' + chaseName;
      else if (b.winner === 'omt') winnerName = '⚡ OMT (Перезаезд)';
      else if (b.winner) winnerName = b.winner;

      var statusStr = 'Ожидает';
      if (b.status === 'live') statusStr = '🔴 LIVE';
      else if (b.status === 'finished') statusStr = 'Завершён';
      else if (b.status === 'omt') statusStr = '⚡ OMT';
      else if (b.status === 'waiting' || b.status === 'upcoming') statusStr = 'Ожидает';

      rows.push([
        stageDisplay,
        b.pair_num || b.id || (i + 1),
        b.pair_title || ('Хит #' + (i + 1)),
        leadName,
        leadCar,
        chaseName,
        chaseCar,
        v1,
        v2,
        v3,
        winnerName,
        statusStr
      ]);
    }

    if (rows.length > 0) {
      var dataRange = sheet.getRange(currentRow, 1, rows.length, 12);
      dataRange.setValues(rows);
      dataRange.setFontSize(10);
      dataRange.setVerticalAlignment('middle');

      // Выравнивание и акценты
      sheet.getRange(currentRow, 1, rows.length, 3).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 4, rows.length, 4).setHorizontalAlignment('left');
      sheet.getRange(currentRow, 8, rows.length, 3).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 11, rows.length, 2).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 11, rows.length, 1).setFontWeight('bold');
      sheet.getRange(currentRow, 12, rows.length, 1).setFontWeight('bold');

      for (var rIdx = 0; rIdx < rows.length; rIdx++) {
        sheet.setRowHeight(currentRow + rIdx, 22);
      }
      currentRow += rows.length;
    }

    // Пустая строка-отступ между стадиями
    sheet.appendRow(['']);
    sheet.setRowHeight(currentRow, 14);
    currentRow++;
  }

  try {
    sheet.autoResizeColumns(1, 12);
    if (sheet.getColumnWidth(4) < 170) sheet.setColumnWidth(4, 170);
    if (sheet.getColumnWidth(5) < 170) sheet.setColumnWidth(5, 170);
    if (sheet.getColumnWidth(6) < 170) sheet.setColumnWidth(6, 170);
    if (sheet.getColumnWidth(7) < 170) sheet.setColumnWidth(7, 170);
    if (sheet.getColumnWidth(11) < 190) sheet.setColumnWidth(11, 190);
    if (sheet.getColumnWidth(12) < 110) sheet.setColumnWidth(12, 110);
  } catch(e) {}
}

// ================================================================
// ФУНКЦИЯ СОХРАНЕНИЯ ЭТАПА В ОТДЕЛЬНУЮ ПОСТОЯННУЮ ВКЛАДКУ (АРХИВ)
// ================================================================
function createArchiveStageSheet(ss, payload, tabName) {
  // Автоматически очищаем любые старые мусорные вкладки
  cleanupGarbageSheets(ss);

  var rawName = String(tabName || '📁 Полный архив этапа').trim();
  var safeName = rawName
    .replace(/[\\/?*\[\]:]/g, '_')
    .replace(/^'+|'+$/g, '')
    .trim();
  if (!safeName) safeName = '📁 Полный архив этапа';
  if (safeName.length > 80) safeName = safeName.substring(0, 80);

  // СТРОГО ОДНА ВКЛАДКА: если уже существует — очищаем и перезаписываем,
  // НИКОГДА не создаем дубликаты (2), (3) и не плодим листы!
  var sheet = ss.getSheetByName(safeName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(safeName);
  }
  var finalName = safeName;

  var meta = payload.meta || {};
  var nowStr = meta.archived_at || new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' });
  var tourneyName = meta.tournament_name || 'DRIFT BANDA';
  var author = meta.author || 'Судейская коллегия';

  // 1. Верхний баннер протокола
  sheet.appendRow(['🏆 ' + tourneyName + ' — ОФИЦИАЛЬНЫЙ ИТОГОВЫЙ ПРОТОКОЛ ЭТАПА (АРХИВ)']);
  sheet.getRange(1, 1, 1, 12).merge()
    .setBackground('#090d16')
    .setFontColor('#fbbf24')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);

  sheet.appendRow(['Фиксация результатов завершённого этапа | Дата и время: ' + nowStr + ' (МСК) | Судейская коллегия: ' + author]);
  sheet.getRange(2, 1, 1, 12).merge()
    .setBackground('#1e293b')
    .setFontColor('#94a3b8')
    .setFontSize(10)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 24);
  sheet.setFrozenRows(2);

  var currentRow = 3;

  // 2. Блок Подиума победителей
  var podium = payload.podium;
  if (podium && (podium.first || podium.second || podium.third)) {
    sheet.appendRow(['']);
    sheet.setRowHeight(currentRow, 10);
    currentRow++;

    sheet.appendRow(['🎖️ ИТОГОВЫЙ ПОДИУМ ПОБЕДИТЕЛЕЙ И ПРИЗЁРОВ ЭТАПА']);
    sheet.getRange(currentRow, 1, 1, 12).merge()
      .setBackground('#78350f')
      .setFontColor('#fef08a')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 26);
    currentRow++;

    var pRows = [];
    if (podium.first && podium.first.name) {
      pRows.push(['🥇 1-е МЕСТО (ЧЕМПИОН)', podium.first.name, podium.first.car || '—', '🏆 ЗОЛОТОЙ КУБОК ЭТАПА']);
    }
    if (podium.second && podium.second.name) {
      pRows.push(['🥈 2-е МЕСТО (ВИЦЕ-ЧЕМПИОН)', podium.second.name, podium.second.car || '—', 'СЕРЕБРЯНЫЙ ПРИЗЁР']);
    }
    if (podium.third && podium.third.name) {
      pRows.push(['🥉 3-е МЕСТО (БРОНЗА)', podium.third.name, podium.third.car || '—', 'БРОНЗОВЫЙ ПРИЗЁР']);
    }
    if (podium.fourth && podium.fourth.name) {
      pRows.push(['4-е МЕСТО (ТОП-4)', podium.fourth.name, podium.fourth.car || '—', 'ПОЛУФИНАЛИСТ']);
    }

    if (pRows.length > 0) {
      for (var pi = 0; pi < pRows.length; pi++) {
        var pr = pRows[pi];
        sheet.appendRow([pr[0], '', pr[1], '', '', pr[2], '', '', pr[3], '', '', '']);
        
        sheet.getRange(currentRow, 1, 1, 2).merge().setHorizontalAlignment('center').setFontWeight('bold');
        sheet.getRange(currentRow, 3, 1, 3).merge().setHorizontalAlignment('left').setFontWeight('bold');
        sheet.getRange(currentRow, 6, 1, 3).merge().setHorizontalAlignment('left');
        sheet.getRange(currentRow, 9, 1, 4).merge().setHorizontalAlignment('center').setFontWeight('bold');

        var bg = pi === 0 ? '#fef3c7' : (pi === 1 ? '#f1f5f9' : (pi === 2 ? '#ffedd5' : '#f8fafc'));
        var fg = pi === 0 ? '#78350f' : (pi === 1 ? '#1e293b' : (pi === 2 ? '#9a3412' : '#475569'));
        sheet.getRange(currentRow, 1, 1, 12).setBackground(bg).setFontColor(fg);
        sheet.setRowHeight(currentRow, 26);
        currentRow++;
      }
    }
  }

  // 3. Блок Сеток Заездов
  sheet.appendRow(['']);
  sheet.setRowHeight(currentRow, 12);
  currentRow++;

  var headers = [
    'Стадия', 'Хит №', 'Название пары',
    'Пилот (Лидер)', 'Автомобиль (Лидер)',
    'Пилот (Преследователь)', 'Автомобиль (Преследователь)',
    'Судья #1', 'Судья #2', 'Судья #3',
    'Победитель', 'Статус'
  ];

  var stageConfigs = [];
  var b32 = (payload && (payload.battles || payload.battles_top32)) || [];
  if (Array.isArray(b32) && b32.length > 0) {
    stageConfigs.push({ title: '🟦 СЕТКА ТОП-32 (16 ПАРНЫХ ХИТОВ)', bgColor: '#1e3a8a', fgColor: '#bfdbfe', battles: b32, defaultStage: 'ТОП-32' });
  }
  var b24 = (payload && payload.battles_top24) || [];
  if (Array.isArray(b24) && b24.length > 0) {
    stageConfigs.push({ title: '🟪 СЕТКА ТОП-24 (8 ПАРНЫХ ХИТОВ)', bgColor: '#581c87', fgColor: '#e9d5ff', battles: b24, defaultStage: 'ТОП-24' });
  }
  var b16 = (payload && payload.battles_top16) || [];
  if (Array.isArray(b16) && b16.length > 0) {
    stageConfigs.push({ title: '🟩 СЕТКА ТОП-16 (8 ПАРНЫХ ХИТОВ)', bgColor: '#064e3b', fgColor: '#a7f3d0', battles: b16, defaultStage: 'ТОП-16' });
  }
  var b8 = (payload && payload.battles_top8) || [];
  if (Array.isArray(b8) && b8.length > 0) {
    stageConfigs.push({ title: '🟨 СЕТКА ТОП-8 — GREAT 8 (4 ПАРНЫХ ХИТА)', bgColor: '#78350f', fgColor: '#fef08a', battles: b8, defaultStage: 'ТОП-8' });
  }
  var b4 = (payload && payload.battles_top4) || [];
  if (Array.isArray(b4) && b4.length > 0) {
    var semis = [];
    var finals = [];
    for (var k = 0; k < b4.length; k++) {
      var item = b4[k];
      if (item.stage === 'semifinal' || (item.id && item.id <= 302) || k < 2) semis.push(item);
      else finals.push(item);
    }
    if (semis.length > 0) stageConfigs.push({ title: '🟧 ПОЛУФИНАЛЫ — ТОП-4 (2 ХИТА)', bgColor: '#7c2d12', fgColor: '#fed7aa', battles: semis, defaultStage: 'Полуфинал' });
    if (finals.length > 0) stageConfigs.push({ title: '👑 ФИНАЛЫ: ЗАЕЗД ЗА 3-Е МЕСТО И ГРАНД-ФИНАЛ 🏆', bgColor: '#831843', fgColor: '#fbcfe8', battles: finals, defaultStage: 'Финал' });
  }

  function getJudgeVoteStr(b, idx, key) {
    if (b.judges && b.judges[idx] && b.judges[idx].vote) {
      var v = String(b.judges[idx].vote).toLowerCase();
      if (v === 'lead') return 'LEAD';
      if (v === 'chase') return 'CHASE';
      if (v === 'omt') return 'OMT';
      return v.toUpperCase();
    }
    if (b.judge_votes && b.judge_votes[key]) {
      var v2 = String(b.judge_votes[key]).toLowerCase();
      if (v2 === 'lead') return 'LEAD';
      if (v2 === 'chase') return 'CHASE';
      if (v2 === 'omt') return 'OMT';
      return v2.toUpperCase();
    }
    return '—';
  }

  for (var sIdx = 0; sIdx < stageConfigs.length; sIdx++) {
    var section = stageConfigs[sIdx];
    sheet.appendRow([section.title]);
    var secRange = sheet.getRange(currentRow, 1, 1, 12);
    secRange.merge().setBackground(section.bgColor).setFontColor(section.fgColor).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 28);
    currentRow++;

    sheet.appendRow(headers);
    var headRange = sheet.getRange(currentRow, 1, 1, 12);
    headRange.setBackground('#0f172a').setFontColor('#f8fafc').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 24);
    currentRow++;

    var bList = section.battles;
    var rows = [];
    for (var i = 0; i < bList.length; i++) {
      var b = bList[i];
      var stageDisplay = section.defaultStage;
      if (b.stage === 'bronze') stageDisplay = '🥉 За 3-е место';
      else if (b.stage === 'final') stageDisplay = '🥇 Гранд-Финал';
      else if (b.stage === 'semifinal') stageDisplay = 'Полуфинал';
      else if (b.stage === 'top8') stageDisplay = 'ТОП-8';
      else if (b.stage === 'top16') stageDisplay = 'ТОП-16';
      else if (b.stage === 'top24') stageDisplay = 'ТОП-24';
      else if (b.stage === 'top32') stageDisplay = 'ТОП-32';

      var leadName = (b.lead && b.lead.name) ? b.lead.name : ((b.lead && b.lead.is_placeholder) ? '—' : 'Ожидание пилота');
      var leadCar = (b.lead && b.lead.car && b.lead.car !== '---') ? b.lead.car : '—';
      var chaseName = (b.chase && b.chase.name) ? b.chase.name : ((b.chase && b.chase.is_placeholder) ? '—' : 'Ожидание пилота');
      var chaseCar = (b.chase && b.chase.car && b.chase.car !== '---') ? b.chase.car : '—';

      var v1 = getJudgeVoteStr(b, 0, 'judge1');
      var v2 = getJudgeVoteStr(b, 1, 'judge2');
      var v3 = getJudgeVoteStr(b, 2, 'judge3');

      var winnerName = '—';
      if (b.winner === 'lead') winnerName = '🏆 ' + leadName;
      else if (b.winner === 'chase') winnerName = '🏆 ' + chaseName;
      else if (b.winner === 'omt') winnerName = '⚡ OMT (Перезаезд)';
      else if (b.winner) winnerName = b.winner;

      var statusStr = 'Ожидает';
      if (b.status === 'live') statusStr = '🔴 LIVE';
      else if (b.status === 'finished') statusStr = 'Завершён';
      else if (b.status === 'omt') statusStr = '⚡ OMT';
      else if (b.status === 'waiting' || b.status === 'upcoming') statusStr = 'Ожидает';

      rows.push([
        stageDisplay,
        b.pair_num || b.id || (i + 1),
        b.pair_title || ('Хит #' + (i + 1)),
        leadName,
        leadCar,
        chaseName,
        chaseCar,
        v1,
        v2,
        v3,
        winnerName,
        statusStr
      ]);
    }

    if (rows.length > 0) {
      var dataRange = sheet.getRange(currentRow, 1, rows.length, 12);
      dataRange.setValues(rows);
      dataRange.setFontSize(10).setVerticalAlignment('middle');
      sheet.getRange(currentRow, 1, rows.length, 3).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 4, rows.length, 4).setHorizontalAlignment('left');
      sheet.getRange(currentRow, 8, rows.length, 3).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 11, rows.length, 2).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 11, rows.length, 1).setFontWeight('bold');
      sheet.getRange(currentRow, 12, rows.length, 1).setFontWeight('bold');

      for (var rIdx = 0; rIdx < rows.length; rIdx++) sheet.setRowHeight(currentRow + rIdx, 22);
      currentRow += rows.length;
    }

    sheet.appendRow(['']);
    sheet.setRowHeight(currentRow, 14);
    currentRow++;
  }

  // 4. Блок Итоговой Таблицы Квалификации
  var quali = payload.qualification || [];
  if (quali.length > 0) {
    sheet.appendRow(['🏁 ИТОГОВАЯ ТАБЛИЦА КВАЛИФИКАЦИИ ЭТАПА']);
    sheet.getRange(currentRow, 1, 1, 10).merge()
      .setBackground('#0f172a')
      .setFontColor('#fbbf24')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 26);
    currentRow++;

    var qHeaders = [
      'Место', 'Семя (Seed)', 'Пилот (ФИО / Никнейм)', 'Автомобиль', 'Группа',
      'Попытка 1', 'Попытка 2', 'Лучший балл', 'Сумма', 'Статус'
    ];
    sheet.appendRow(qHeaders);
    sheet.getRange(currentRow, 1, 1, 10)
      .setBackground('#1e293b')
      .setFontColor('#94a3b8')
      .setFontWeight('bold')
      .setFontSize(9)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 22);
    currentRow++;

    var qRows = [];
    for (var qi = 0; qi < quali.length; qi++) {
      var qp = quali[qi];
      var qRank = qp.rank || (qi + 1);
      var qStatus = (qp.status === 'qualified' || qRank <= 32) ? 'ТОП-32' : (qp.status === 'eliminated' ? 'Выбыл' : (qp.status || 'В ожидании'));
      if (qp.is_health_group) qStatus = 'Группа Здоровья';
      qRows.push([
        qRank,
        qp.seed || (qi + 1),
        qp.name || '—',
        qp.car || '—',
        qp.group || 'A',
        qp.run1 !== undefined ? qp.run1 : 0,
        qp.run2 !== undefined ? qp.run2 : 0,
        qp.best !== undefined ? qp.best : 0,
        qp.total !== undefined ? qp.total : 0,
        qStatus
      ]);
    }

    if (qRows.length > 0) {
      var qDataRange = sheet.getRange(currentRow, 1, qRows.length, 10);
      qDataRange.setValues(qRows);
      qDataRange.setFontSize(9).setVerticalAlignment('middle');
      sheet.getRange(currentRow, 1, qRows.length, 2).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 3, qRows.length, 2).setHorizontalAlignment('left');
      sheet.getRange(currentRow, 5, qRows.length, 6).setHorizontalAlignment('center');
      sheet.getRange(currentRow, 8, qRows.length, 1).setFontWeight('bold');

      for (var qIdx = 0; qIdx < qRows.length; qIdx++) sheet.setRowHeight(currentRow + qIdx, 20);
      currentRow += qRows.length;
    }
  }

  try {
    sheet.autoResizeColumns(1, 12);
    if (sheet.getColumnWidth(4) < 170) sheet.setColumnWidth(4, 170);
    if (sheet.getColumnWidth(5) < 170) sheet.setColumnWidth(5, 170);
    if (sheet.getColumnWidth(6) < 170) sheet.setColumnWidth(6, 170);
    if (sheet.getColumnWidth(7) < 170) sheet.setColumnWidth(7, 170);
    if (sheet.getColumnWidth(11) < 190) sheet.setColumnWidth(11, 190);
    if (sheet.getColumnWidth(12) < 110) sheet.setColumnWidth(12, 110);
  } catch(e) {}

  return finalName;
}

function logAudit(ss, author, action, details) {
  var sheetName = '🛡️ Аудит-лог';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = ['Время (МСК)', 'Автор / Роль', 'Действие', 'Детали'];
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#0f172a');
    hRange.setFontColor('#38bdf8');
    hRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var timeStr = new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' });
  sheet.appendRow([timeStr, author, action, details]);
  if (sheet.getLastRow() > 1000) {
    sheet.deleteRow(2);
  }
}
