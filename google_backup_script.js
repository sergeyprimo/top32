// ================================================================
// СКРИПТ АВАРИЙНОГО РЕЗЕРВНОГО КОНТУРА ДЛЯ GOOGLE ТАБЛИЦЫ
// Проект: DRIFT BANDA / ГАЗДЕРЖИМ.РФ
// ВАРИАНТ 1: ЕДИНЫЙ СКВОЗНОЙ ПРОТОКОЛ (ВСЕ СТАДИИ + КВАЛИФИКАЦИЯ)
//
// Инструкция по обновлению / установке:
// 1. Откройте вашу Google Таблицу (DRIFT BANDA - РЕЗЕРВНАЯ СЕТКА И КВАЛИФИКАЦИЯ)
// 2. В верхнем меню выберите: Расширения (Extensions) -> Apps Script
// 3. Сотрите весь старый код в редакторе и вставьте этот код целиком
// 4. Справа вверху нажмите: Развернуть (Deploy) -> Управление развертываниями (Manage deployments)
// 5. Нажмите на значок карандаша (Изменить) у текущего активного веб-приложения:
//    - Версия: Новая версия (New version)
//    - Нажмите кнопку "Развернуть" (Deploy)
// 6. Всё готово! Теперь каждые 30 секунд в Google Таблицу на один лист
//    выгружаются ТОП-32, ТОП-24, ТОП-16, ТОП-8, Полуфиналы и Финал,
//    а на соседний лист — вся таблица Квалификации.
// ================================================================

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = (e && e.parameter && e.parameter.action) || 'ping';

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
        data: JSON.parse(snapshot)
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Стандартный пинг сервиса
    return ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      service: 'Drift Banda Emergency Backup Service (All Stages Unified)',
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

    var action = payload.action || 'sync_tourney';

    // Очистка резервной таблицы по кнопке администратора
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

    if (action === 'get_backup') {
      var snap = PropertiesService.getScriptProperties().getProperty('tourney_snapshot');
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        data: snap ? JSON.parse(snap) : null
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Сохраняем полный снимок турнира в свойства скрипта (для мгновенного отката при сбоях)
    PropertiesService.getScriptProperties().setProperty('tourney_snapshot', JSON.stringify(payload));

    var quali = payload.qualification || [];
    var meta = payload.meta || {};

    // 1. Обновляем визуальный лист «🏁 Квалификация»
    syncQualiSheet(ss, quali, meta);

    // 2. Обновляем визуальный лист «🏆 Сетка заездов» (Все стадии сквозным протоколом)
    syncBattlesSheet(ss, payload, meta);

    // 3. Добавляем запись в лист «🛡️ Аудит-лог»
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
