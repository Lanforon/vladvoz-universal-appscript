/***** НАСТРОЙКИ ДЛЯ ПОЛЬЗОВАТЕЛЯ *****/

const ALLOWED_EMAILS = ['work@vladvoz.com']; // ИЗМЕНИТЕ НА ДЕЙСТВУЮЩИЙ EMAIL

const REGISTRY_FILE_ID = '1TEksg-gFc5rgPAcgUC7aOrVsJKhCrw4-UPUTSqxVaF8'; 
const REG_SHEET = 'REGISTRY';
const START_ROW = 2; // для реестра

const TARGET_FOLDER_ID = '14mUE3P63c79GqMHgWDy4GrKkcf13Ut7c'; // TOO 

// Ссылки в REGISTRY
const REG_MASTER_FACTORY_CELL = 'B1';
const REG_MASTER_NOFACT_CELL  = 'D1';
const REG_STYLE_MASTER_CELL   = 'F1';

// ИСПРАВЛЕНО: Новая структура колонок
const COLS = {
  fio: 1,        // A - ФИО
  order: 2,      // B - ID Геткурс
  // C - Пусто
  devUrl: 4,     // D - Ссылка DEV
  studentUrl: 5, // E - Ссылка STUDENT
  devMode: 6,    // F - Статус ('Фабрика' / 'Не Фабрика')
  aud1: 7,       // G - Аудитория 1
  exp1: 8,       // H - Эксперт 1
  aud2: 9,       // I - Аудитория 2
  exp2: 10,      // J - Эксперт 2
  aud3: 11,      // K - Аудитория 3
  exp3: 12       // L - Эксперт 3
};

const COL_B=2, COL_C=3, COL_D=4;






function onOpen() {
  const me = Session.getEffectiveUser().getEmail();
  if (!ALLOWED_EMAILS.includes(me)) return;

  const currentFile = SpreadsheetApp.getActive();
  const currentFileName = currentFile.getName();
  
  const menu = SpreadsheetApp.getUi().createMenu('[+] UTILIES [+]');
  
  // Если НЕ находимся в DEV файле - показываем кнопки создания DEV
  if (!/^DEV\s—\s/i.test(currentFileName) && !/^STUDENT\s—\s/i.test(currentFileName)) {
    menu
      .addSeparator()
      .addItem('🎯 СОЗДАТЬ DEV - ФАБРИКА', 'menuDevelopFactory')
      .addSeparator()
      .addItem('🎯 СОЗДАТЬ DEV - НЕ ФАБРИКА', 'menuDevelopNoFactory')
      .addSeparator();
  }
  
  // Показываем функции только в DEV файлах
  if (/^DEV\s—\s/i.test(currentFileName)) {
    menu
      .addSeparator()
      .addItem('[++] Создать STUDENT - для ученика', 'menuDeliverToStudent_AutoContext')
      .addSeparator()
      .addItem('[+][BCD] Забрать BCD колонки - у ученика', 'pasteSelectedValues_Bidirectional')
      .addItem('[-][BCD] ОТДАТЬ BCD', 'f2')
      .addItem('[+][EFG] Забрать EFG → E - у ученика', 'f1')
      .addSeparator()
      .addItem('[>] Раскрыть смыслы (> маркер)', 'menuExpandSurgically_Final') 
      .addSeparator()
      .addItem('[SYNC] Полная синхронизация ученику (без формул)', 'menuDeliverExpanded_Final')
      .addSeparator();
  }
  
  menu.addToUi();
}

/***** === ЗАБРАТЬ EFG У УЧЕНИКА СО СМЕЩЕНИЕМ В E ===*****/
function f1() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    
    const ssStud = SpreadsheetApp.openById(studentId);
    const shStud = ssStud.getActiveSheet();
    const sheetName = shStud.getName();
    
    const ssDev = SpreadsheetApp.openById(devId);
    const shDev = ssDev.getSheetByName(sheetName) || ssDev.insertSheet(sheetName);

    const lastRow = shStud.getLastRow();
    
    if (lastRow < 1) {
      SpreadsheetApp.getUi().alert('STUDENT файл пустой');
      return;
    }

    // Получаем только EFG данные из STUDENT
    const studValues = shStud.getRange(1, 5, lastRow, 3).getValues(); // E, F, G
    const studFormulas = shStud.getRange(1, 5, lastRow, 3).getFormulas(); // E, F, G
    
    // 1. Создаем массив для объединенного значения в E (затираем формулы)
    const devValuesE = studValues.map((row, rowIndex) => {
      // Объединяем значения E+F+G в одну строку (без формул)
      const combinedValue = row
        .map((value, colIndex) => studFormulas[rowIndex][colIndex] ? '' : value) // Заменяем формулы пустотами
        .filter(val => val) // Убираем пустые значения
        .join(' '); // Объединяем через пробел
      
      return [combinedValue]; // Возвращаем массив с одним элементом для столбца E
    });

    // 2. Создаем массивы для затирания формул в E, F, G
    const emptyValuesE = devValuesE; // E уже содержит значения без формул
    const emptyValuesF = Array(lastRow).fill().map(() => ['']); // Пустые значения для F
    const emptyValuesG = Array(lastRow).fill().map(() => ['']); // Пустые значения для G

    // 3. Записываем значения в E DEV (затираем формулы)
    shDev.getRange(1, 5, lastRow, 1).setValues(emptyValuesE);
    
    // 4. Затираем формулы в столбцах F и G DEV пустыми значениями
    shDev.getRange(1, 6, lastRow, 1).setValues(emptyValuesF);
    shDev.getRange(1, 7, lastRow, 1).setValues(emptyValuesG);

    SpreadsheetApp.getUi().alert(`✅ Значения E-F-G из STUDENT перенесены в E DEV, все формулы в E-F-G затерты`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка при переносе EFG в E: ' + (e.message || e));
  }
}

/***** === ОТДАТЬ BCD УЧЕНИКУ (ТОЛЬКО НЕПУСТЫЕ ЯЧЕЙКИ) ===*****/
function f2() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    
    const ssDev = SpreadsheetApp.openById(devId);
    const shDev = ssDev.getActiveSheet();
    const sheetName = shDev.getName();
    
    const ssStud = SpreadsheetApp.openById(studentId);
    const shStud = ssStud.getSheetByName(sheetName) || ssStud.insertSheet(sheetName);

    const lastRow = shDev.getLastRow();
    
    if (lastRow < 1) {
      SpreadsheetApp.getUi().alert('DEV файл пустой');
      return;
    }

    let copiedCount = 0;

    // Проходим по всем строкам DEV
    for (let r = 1; r <= lastRow; r++) {
      // Пропускаем сгруппированные строки в обеих таблицах
      if (isRowGrouped_(shDev, r) || isRowGrouped_(shStud, r)) continue;
      
      // Проверяем ячейки B, C, D в DEV
      const devCellB = shDev.getRange(r, 2); // B
      const devCellC = shDev.getRange(r, 3); // C
      const devCellD = shDev.getRange(r, 4); // D
      
      const devValueB = devCellB.getValue();
      const devValueC = devCellC.getValue();
      const devValueD = devCellD.getValue();
      
      // Копируем только непустые ячейки в STUDENT
      if (devValueB) {
        shStud.getRange(r, 2).setValue(devValueB);
        copiedCount++;
      }
      if (devValueC) {
        shStud.getRange(r, 3).setValue(devValueC);
        copiedCount++;
      }
      if (devValueD) {
        shStud.getRange(r, 4).setValue(devValueD);
        copiedCount++;
      }
    }

    if (copiedCount === 0) {
      SpreadsheetApp.getUi().alert('Не найдено данных в столбцах B-C-D DEV');
      return;
    }

    SpreadsheetApp.getUi().alert(`✅ Отдано ${copiedCount} ячеек B-C-D ученику (только непустые значения)`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка при отправке BCD ученику: ' + (e.message || e));
  }
}

function menuDeliverToStudent_AutoContext() {
  try {
    const { sheet, row } = resolveRegistryRowContext_();
    const order = String(sheet.getRange(row, COLS.order).getValue()||'').trim();
    if (!order) throw new Error('ID заказа пуст.');
    
    let studUrlExisting = String(sheet.getRange(row, COLS.studentUrl).getValue() || '').trim();
    let studId;
    
    if (studUrlExisting) {
      studId = fileIdFromUrl_(studUrlExisting);
      try { 
        DriveApp.getFileById(studId).getId(); 
      } catch (e) { 
        studId = null; 
      }
    }
    
    if (!studId) {
      // Используем DEV файл для создания STUDENT
      SpreadsheetApp.getUi().alert('🔄 Начинаю создание STUDENT файла...');

      const devUrl = String(sheet.getRange(row, COLS.devUrl).getValue() || '').trim();
      if (!devUrl) throw new Error('Сначала создайте DEV файл');
      
      const devId = fileIdFromUrl_(devUrl);
      const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
      const studFile = DriveApp.getFileById(devId).makeCopy(`STUDENT — ${order}`, folder);
      studId = studFile.getId();
      
      // Убираем формулы из STUDENT
      removeFormulasFromStudent_(studId);
      
      DriveApp.getFileById(studId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      const studUrl = `https://docs.google.com/spreadsheets/d/${studId}/edit`;
      sheet.getRange(row, COLS.studentUrl).setValue(studUrl);
    }
    
    const finalStudUrl = `https://docs.google.com/spreadsheets/d/${studId}/edit`;
    showLink_('STUDENT готов (создан из DEV, формулы удалены)', finalStudUrl, 'ПЕРЕЙТИ В STUD');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка создания STUDENT: ' + (e.message || e));
  }
}

function removeFormulasFromStudent_(studentId) {
  const ss = SpreadsheetApp.openById(studentId);
  const sheets = ss.getSheets();
  
  sheets.forEach(sh => {
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    
    if (lastRow > 0 && lastCol > 0) {
      const range = sh.getRange(1, 1, lastRow, lastCol);
      const formulas = range.getFormulas();
      const values = range.getValues();
      
      // Проходим по каждой ячейке и очищаем только те, где есть формулы
      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
          const formula = formulas[r][c];
          // Если есть формула (начинается с =) - очищаем только эту ячейку
          if (formula && formula.startsWith('=')) {
            const cell = sh.getRange(r + 1, c + 1);
            cell.clearContent(); // Очищаем содержимое, сохраняя форматирование
          }
        }
      }
    }
  });
}

/***** === 1. Создать DEV  ===*****/
function menuDevelopFactory()   { createDevOnly_AutoContext_('factory'); }
function menuDevelopNoFactory() { createDevOnly_AutoContext_('nofactory'); }
function createDevOnly_AutoContext_(mode) {
  const { sheet, row } = resolveRegistryRowContext_();
  const masterUrl = getMasterUrlByMode_(mode);
  if (!masterUrl) throw new Error(`В REGISTRY нет MASTER для режима ${mode}`);
  const masterId = fileIdFromUrl_(masterUrl);
  const order = String(sheet.getRange(row, COLS.order).getValue() || '').trim();
  if (!order) throw new Error('В колонке B (ID заказа) пусто.');

  const a1 = sheet.getRange(row, COLS.aud1).getValue() || '';
  const e1 = sheet.getRange(row, COLS.exp1).getValue() || '';
  const a2 = sheet.getRange(row, COLS.aud2).getValue() || '';
  const e2 = sheet.getRange(row, COLS.exp2).getValue() || '';
  const a3 = sheet.getRange(row, COLS.aud3).getValue() || '';
  const e3 = sheet.getRange(row, COLS.exp3).getValue() || '';

  const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const devFile = DriveApp.getFileById(masterId).makeCopy(`DEV — ${order}`, folder);
  const devId    = devFile.getId();

  applyAudienceExpert_(devId, {aud:[a1,a2,a3], exp:[e1,e2,e3]});
  clearAudienceColumnsIfMissing_(devId, {aud2:a2, aud3:a3});
  
  sheet.getRange(row, COLS.devUrl).setValue(`https://docs.google.com/spreadsheets/d/${devId}/edit`);
  
  const displayMode = mode === 'factory' ? 'Фабрика' : 'Не Фабрика';
  sheet.getRange(row, COLS.devMode).setValue(displayMode);

  showLink_('Перейди в DEV и дай отработать GPT.', `https://docs.google.com/spreadsheets/d/${devId}/edit`, 'ПЕРЕЙТИ В DEV');
}

function menuDeliverExpanded_Final() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    
    const ssDev = SpreadsheetApp.openById(devId);
    const shDev = ssDev.getActiveSheet();
    const sheetName = shDev.getName();
    
    const ssStud = SpreadsheetApp.openById(studentId);
    
    // Создаем УНИКАЛЬНОЕ имя для временной вкладки
    const timestamp = new Date().getTime(); // Используем timestamp для уникальности
    const tempSheetName = `temp_${timestamp}`;
    
    // Создаем новую вкладку как копию DEV с временным именем
    const newSheet = shDev.copyTo(ssStud);
    newSheet.setName(tempSheetName);
    
    // Удаляем формулы из новой вкладки (сохраняя стили)
    removeFormulasKeepStyles_(newSheet);
    
    // Удаляем старую вкладку если она существует
    const oldSheet = ssStud.getSheetByName(sheetName);
    if (oldSheet) {
      ssStud.deleteSheet(oldSheet);
    }
    
    // Переименовываем новую вкладку в оригинальное имя
    newSheet.setName(sheetName);
    
    // Активируем новую вкладку
    ssStud.setActiveSheet(newSheet);

    SpreadsheetApp.getUi().alert(`✅ STUDENT обновлен: вкладка "${sheetName}" заменена на версию без формул`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка при синхронизации DEV → STUDENT: ' + (e.message || e));
  }
}

function pasteSelectedValues_Bidirectional() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    
    const ssStud = SpreadsheetApp.openById(studentId);
    const shStud = ssStud.getActiveSheet();
    const sheetName = shStud.getName();
    
    const ssDev = SpreadsheetApp.openById(devId);
    const shDev = ssDev.getSheetByName(sheetName) || ssDev.insertSheet(sheetName);

    const lastRow = shStud.getLastRow();
    if (lastRow < 1) {
      SpreadsheetApp.getUi().alert('STUDENT файл пустой');
      return;
    }

    // Собираем данные из STUDENT (только несгруппированные строки)
    const bValues = [];
    const cValues = [];
    const dValues = [];
    const rowsToCopy = [];

    for (let r = 1; r <= lastRow; r++) {
      // Пропускаем сгруппированные строки
      if (isRowGrouped_(shStud, r)) continue;
      
      rowsToCopy.push(r);
      bValues.push([shStud.getRange(r, COL_B).getValue()]);
      cValues.push([shStud.getRange(r, COL_C).getValue()]);
      dValues.push([shStud.getRange(r, COL_D).getValue()]);
    }

    if (rowsToCopy.length === 0) {
      SpreadsheetApp.getUi().alert('Не найдено несгруппированных строк в STUDENT');
      return;
    }

    // Копируем в DEV (только несгруппированные строки)
    for (let i = 0; i < rowsToCopy.length; i++) {
      const row = rowsToCopy[i];
      if (!isRowGrouped_(shDev, row)) {
        shDev.getRange(row, COL_B).setValue(bValues[i][0]);
        shDev.getRange(row, COL_C).setValue(cValues[i][0]);
        shDev.getRange(row, COL_D).setValue(dValues[i][0]);
      }
    }

    SpreadsheetApp.getUi().alert(`✅ Скопировано ${rowsToCopy.length} строк B-C-D из STUDENT в DEV`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка при копировании BCD из STUDENT: ' + (e.message || e));
  }
}

/***** ======================= ПОМОЩНИКИ =======================*****/

function resolveDevStudentByContext_() {
  const { sheet, row } = resolveRegistryRowContext_();
  let devUrl = String(sheet.getRange(row, COLS.devUrl).getValue() || '').trim();
  let studentUrl = String(sheet.getRange(row, COLS.studentUrl).getValue() || '').trim();
  const cur = SpreadsheetApp.getActive();
  const curId = cur.getId();
  const curName = cur.getName();
  const thisUrl = `https://docs.google.com/spreadsheets/d/${curId}/edit`;
  
  if (/^STUDENT\s—\s/i.test(curName)) {
    if (studentUrl !== thisUrl) {
      sheet.getRange(row, COLS.studentUrl).setValue(thisUrl);
      studentUrl = thisUrl;
    }
  }
  
  if (/^DEV\s—\s/i.test(curName)) {
    if (devUrl !== thisUrl) {
      sheet.getRange(row, COLS.devUrl).setValue(thisUrl);
      devUrl = thisUrl;
    }
  }
  
  if (!devUrl) throw new Error('В реестре нет DEV. Сначала запусти «1. Создать DEV».');
  if (!studentUrl) throw new Error('В реестре нет STUDENT. Сначала запусти «2. DEV → STUDENT».');
  
  return { 
    devId: fileIdFromUrl_(devUrl), 
    studentId: fileIdFromUrl_(studentUrl) 
  };
}

function parseNumberedList_(text) {
  if (!text) return [];
  
  const cleanedText = String(text)
    .replace(/\r\n?/g, '\n')
    .replace(/\u00A0/g, ' ')
    .trim();

  if (!cleanedText) return [];

  const items = [];
  const lines = cleanedText.split('\n');
  
  for (const line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine) continue;

    // Ищем нумерованные пункты (1., 2., 3. и т.д.)
    const match = trimmedLine.match(/^\s*(\d{1,2})[\.\)]\s*(.+)$/);
    if (match) {
      items.push(match[2].trim());
    }
  }

  return items.length > 0 ? items : [cleanedText];
}

function pasteColsBCD_FromDevToStud_(devId, studId) {
  const srcSS = SpreadsheetApp.openById(devId);
  const dstSS = SpreadsheetApp.openById(studId);
  const src = srcSS.getSheets()[0];
  const dst = dstSS.getSheets()[0];
  if (!src.getLastRow()) return;
  const rows = src.getLastRow();
  
  ensureRowsAndCols_(dst, rows, 4); 

  const sourceRange = src.getRange(1, COL_B, rows, 3); // B:D
  const destinationRange = dst.getRange(1, COL_B, rows, 3);
  
  const values = sourceRange.getValues();
  destinationRange.setValues(values);
}

function copyFormattingBetweenSheets_(sourceSheet, targetSheet, lastRow, lastCol) {
  // Копируем форматирование строк
  for (let r = 1; r <= lastRow; r++) {
    const sourceRow = sourceSheet.getRange(r, 1, 1, lastCol);
    const targetRow = targetSheet.getRange(r, 1, 1, lastCol);
    sourceRow.copyTo(targetRow, {formatOnly: true});
  }
  
  // Копируем форматирование столбцов
  for (let c = 1; c <= lastCol; c++) {
    const sourceCol = sourceSheet.getRange(1, c, lastRow, 1);
    const targetCol = targetSheet.getRange(1, c, lastRow, 1);
    sourceCol.copyTo(targetCol, {formatOnly: true});
  }
}

function removeFormulasKeepStyles_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 1 || lastCol < 1) return;
  
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const formulas = range.getFormulas();
  
  // Проходим по каждой ячейке и очищаем только те, где есть формулы
  for (let r = 1; r <= lastRow; r++) {
    for (let c = 1; c <= lastCol; c++) {
      const formula = formulas[r-1][c-1];
      // Если есть формула - очищаем только содержимое
      if (formula && formula.startsWith('=')) {
        const cell = sheet.getRange(r, c);
        cell.clearContent(); // Очищает только содержимое, сохраняя стили
      }
    }
  }
}

function removeFormulasFromSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 1 || lastCol < 1) return;
  
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const formulas = range.getFormulas();
  const values = range.getValues();
  
  // Создаем массив значений без формул
  const valuesWithoutFormulas = values.map((row, rowIndex) => 
    row.map((value, colIndex) => {
      const formula = formulas[rowIndex][colIndex];
      // Если есть формула - оставляем пустую строку, иначе оставляем значение
      return formula && formula.startsWith('=') ? '' : value;
    })
  );
  
  // Записываем значения без формул
  range.setValues(valuesWithoutFormulas);
}

function copyBasicFormatting_(sourceSheet, targetSheet, lastRow, lastCol) {
  try {
    // Копируем настройки строк (высоту)
    for (let r = 1; r <= lastRow; r++) {
      const sourceRow = sourceSheet.getRange(r, 1);
      const targetRow = targetSheet.getRange(r, 1);
      targetSheet.setRowHeight(r, sourceSheet.getRowHeight(r));
    }
    
    // Копируем настройки столбцов (ширину)
    for (let c = 1; c <= lastCol; c++) {
      const sourceCol = sourceSheet.getRange(1, c);
      const targetCol = targetSheet.getRange(1, c);
      targetSheet.setColumnWidth(c, sourceSheet.getColumnWidth(c));
    }
    
    // Копируем базовые стили ячеек
    const sourceStyles = sourceSheet.getRange(1, 1, lastRow, lastCol).getTextStyles();
    const targetRange = targetSheet.getRange(1, 1, lastRow, lastCol);
    targetRange.setTextStyles(sourceStyles);
    
  } catch (e) {
    console.log('Частичное форматирование применено: ' + e.message);
  }
}

function showLink_(text, url, btn) {
  const html = HtmlService.createHtmlOutput(
    `<div style="font:14px/1.4 system-ui,Arial;padding:12px">
       <div style="margin-bottom:10px">${text}</div>
       <a href="${url}" target="_blank"
          style="display:inline-block;padding:8px 12px;background:#1a73e8;color:#fff;border-radius:6px;text-decoration:none;">
         ${btn || 'Перейти'}
       </a>
     </div>`
  ).setWidth(420).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(html, 'Уведомление');
}

function clearContentOnly_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow > 0 && lastCol > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clearContent(); 
  }
}

function removeFormulasFromRange_(range) {
  const formulas = range.getFormulas();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  
  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const formula = formulas[r][c];
      // Если есть формула - очищаем только содержимое
      if (formula && formula.startsWith('=')) {
        const cell = range.getCell(r + 1, c + 1);
        cell.clearContent(); // Очищает только содержимое, сохраняя стили
      }
    }
  }
}

function copyRowHeightsAndColumnWidths_(sourceSheet, targetSheet, lastRow, lastCol) {
  try {
    // Копируем высоты строк
    for (let r = 1; r <= lastRow; r++) {
      targetSheet.setRowHeight(r, sourceSheet.getRowHeight(r));
    }
    
    // Копируем ширины столбцов
    for (let c = 1; c <= lastCol; c++) {
      targetSheet.setColumnWidth(c, sourceSheet.getColumnWidth(c));
    }
  } catch (e) {
    console.log('Размеры скопированы частично: ' + e.message);
  }
}

function isRowGrouped_(sheet, rowIndex) {
  try {
    const rowGroups = sheet.getRowGroups();
    
    for (const group of rowGroups) {
      const startRow = group.getControlIndex() + 1; 
      const numRows = group.getNumRows();
      const endRow = startRow + numRows - 1;
      
      if (rowIndex > startRow && rowIndex <= endRow) {
        return true;
      }
    }
    return false;
  } catch (e) {
    console.log('Ошибка при проверке группировки:', e);
    return false;
  }
}

function ensureRowsAndCols_(sh, minRow, minCol){
  const maxR = sh.getMaxRows();
  if (maxR < minRow) sh.insertRowsAfter(maxR, minRow - maxR);
  const maxC = sh.getMaxColumns();
  if (maxC < minCol) sh.insertColumnsAfter(maxC, minCol - maxC);
}

function applyAudienceExpert_(fileId,{aud,exp}){
  const ss=SpreadsheetApp.openById(fileId);
  ss.getSheets().forEach(sh=>{
    try{ sh.getRange('B1:D1').setValues([aud]); }catch(e){}
    try{ sh.getRange('B2:D2').setValues([exp]); }catch(e){}
  });
}

function clearAudienceColumnsIfMissing_(fileId,{aud2,aud3}){
  const ss=SpreadsheetApp.openById(fileId);
  ss.getSheets().forEach(sh=>{
    const rows = sh.getMaxRows();
    if(!aud2) sh.getRange(1,COL_C,rows,1).clearContent();
    if(!aud3) sh.getRange(1,COL_D,rows,1).clearContent();
  });
}

function copyRowFormat_(sheet, srcRow, dstStartRow, count) {
    if (count <= 0) return;
    const maxCols = sheet.getMaxColumns();
    const sourceRange = sheet.getRange(srcRow, 1, 1, maxCols);
    for (let i = 0; i < count; i++) {
        const destRange = sheet.getRange(dstStartRow + i, 1, 1, maxCols);
        sourceRange.copyTo(destRange, { formatOnly: true });
    }
}

function resolveRegistryRowContext_() {
  let reg = SpreadsheetApp.getActive();
  let sheet = reg.getSheetByName(REG_SHEET);
  if (sheet) {
    const range = reg.getActiveRange();
    if (range) {
      const row = range.getRow();
      if (row>=START_ROW) return {reg, sheet, row};
    }
  }
  if (!REGISTRY_FILE_ID) throw new Error('Не задан REGISTRY_FILE_ID.');
  reg = SpreadsheetApp.openById(REGISTRY_FILE_ID);
  sheet = reg.getSheetByName(REG_SHEET);
  if (!sheet) throw new Error('В реестре нет листа REGISTRY.');
  const id = extractOrderIdFromFileName_(SpreadsheetApp.getActive().getName());
  if (!id) throw new Error('Не удалось определить ID заказа из имени файла.');
  const row = findRowByOrder_(sheet, id);
  if (row<START_ROW) throw new Error(`В REGISTRY не найдена строка с ID = ${id}.`);
  return {reg, sheet, row};
}

function getMasterUrlByMode_(mode) {
  const reg = SpreadsheetApp.openById(REGISTRY_FILE_ID).getSheetByName(REG_SHEET);
  if (!reg) throw new Error('Не найден лист REGISTRY.');
  const cell = (mode==='factory') ? REG_MASTER_FACTORY_CELL : REG_MASTER_NOFACT_CELL;
  return String(reg.getRange(cell).getValue()||'').trim();
}

function findRowByOrder_(sheet, orderId) {
  const rng = sheet.getRange(START_ROW, COLS.order, sheet.getLastRow() - START_ROW + 1, 1).getValues();
  for (let i=0;i<rng.length;i++) if (String(rng[i][0]).trim()===String(orderId).trim()) return START_ROW + i;
  return -1;
}

function extractOrderIdFromFileName_(name) {
  if (!name) return '';
  const parts = name.split('—').map(s=>s.trim());
  return parts.length>=2 ? parts[1] : '';
}

function fileIdFromUrl_(url) {
  const m = /\/d\/([a-zA-Z0-9\-_]+)/.exec(String(url)||'');
  if (!m) throw new Error('Не удалось извлечь fileId из URL: '+url);
  return m[1];
}