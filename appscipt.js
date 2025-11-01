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

const MARK_SELECT  = '>';
const MARK_FACTORY = 'ФАБРИКА';
const FORMULA_MARKER = '@@=';
const FORMULA_MARKER_LENGTH = FORMULA_MARKER.length;

const COL_A=1, COL_B=2, COL_C=3, COL_D=4, COL_E=5, COL_F=6, COL_G=7, COL_H=8, COL_I=9;

const SLEEP_MS = 1500; // только для телесуфлёра


/***** НАСТРОЙКИ ДЛЯ ПОЛЬЗОВАТЕЛЯ *****/










/***** === МЕНЮ === *****/
function onOpen() {
  const me = Session.getEffectiveUser().getEmail();
  if (!ALLOWED_EMAILS.includes(me)) return;

  SpreadsheetApp.getUi()
    .createMenu('Утилиты')
    .addItem('Формулы ⇄ Текст', 'toggleFormulasOnSelection')
    .addSeparator()
    .addItem('1. ФАБРИКА — Создать DEV', 'menuDevelopFactory')
    .addItem('1. НЕ ФАБРИКА — Создать DEV', 'menuDevelopNoFactory')
    .addItem('1. ОТДАТЬ УЧЕНИКУ — Создать STUD', 'menuDeliverToStudent_AutoContext')
    .addSeparator()
    .addItem("1. Выделенное вв STUDENT<>ADMIN", "pasteSelectedValues_Bidirectional")
    .addItem("1. Разделить Ячейку вниз по Номерам", "explodeNumberedListToRows")
    .addSeparator()
    .addItem('2, РАСКРЫТЬ СМЫСЛЫ в DEV', 'menuExpandSurgically_Final')
    .addItem('2. ОТДАТЬ СМЫСЛЫ в STUDENT', 'menuDeliverExpanded_Final')
    .addSeparator()
    .addItem('5. ОТДАТЬ ТЕЛЕСУФЛЕР', 'menuTeleprompter_InPlace')
    .addSeparator()
    .addItem('🔄 Добавить IF к GPT', 'f1') 
    .addToUi();
}

/***** Добавил асинхронное выполнение кода для раскрытия gpt формул (выполняется теперь одновременно) *****/

async function menuExpandSurgically_Final() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    
    const ssStud = SpreadsheetApp.openById(studentId);
    const shStud = ssStud.getActiveSheet();
    const sheetName = shStud.getName();
    
    const ssDev = SpreadsheetApp.openById(devId);
    const shDev = ssDev.getSheetByName(sheetName) || ssDev.insertSheet(sheetName);

    const groups = collectSelectedRows_WithParsedLists_(shStud);
    if (!groups.length) {
      SpreadsheetApp.getUi().alert('Не найдено строк с маркером "> ОТБЕРИТЕ" на активном листе.');
      return;
    }

    // Разворачиваем формулы в DEV
    await unfoldFormulasInColumnsAsync_(devId, [COL_E, COL_F, COL_G, COL_H]);
    SpreadsheetApp.flush();

    groups.sort((a, b) => b.rowIndex - a.rowIndex);

    const promises = groups.map(async (g) => {
      const r0 = g.rowIndex;
      const { k, B, C, D } = g.meta;
      if (!k || k < 1) return;

      // --- Шаг 1: Работа в STUDENT ---
      if (k > 1) {
        shStud.insertRowsAfter(r0, k - 1);
        copyRowFormat_(shStud, r0, r0 + 1, k - 1);
        
        // Копируем формулы из исходной строки во вставленные строки
        const sourceFormulas = shStud.getRange(r0, 1, 1, shStud.getLastColumn()).getFormulas()[0];
        for (let i = 1; i < k; i++) {
          const targetRange = shStud.getRange(r0 + i, 1, 1, sourceFormulas.length);
          const formulasToSet = sourceFormulas.map(formula => 
            formula ? adjustCellReferences_(formula, i) : ''
          );
          targetRange.setFormulas([formulasToSet]);
        }
      }

      // Обрабатываем ТОЛЬКО строки с маркером ">" - заполняем B-D данными
      if (g.hasSelectMarker) {
        const blockData = [];
        for(let i = 0; i < k; i++) {
          const bVal = B[i] !== undefined ? B[i] : (B.length === 1 ? B[0] : '');
          const cVal = C[i] !== undefined ? C[i] : (C.length === 1 ? C[0] : '');
          const dVal = D[i] !== undefined ? D[i] : (D.length === 1 ? D[0] : '');
          blockData.push([bVal, cVal, dVal]);
        }
        
        // Вставляем ТОЛЬКО данные в B-D
        if (!isRowGrouped_(shStud, r0)) {
          shStud.getRange(r0, COL_B, k, 3).setValues(blockData);
        }
      }
      SpreadsheetApp.flush();

      // --- Шаг 2: Работа в DEV ---
      if (k > 1) {
        shDev.insertRowsAfter(r0, k - 1);
        
        // СНАЧАЛА копируем формулы из исходной строки DEV во вставленные строки
        const sourceDevFormulas = shDev.getRange(r0, 1, 1, shDev.getLastColumn()).getFormulas()[0];
        for (let i = 1; i < k; i++) {
          const targetRange = shDev.getRange(r0 + i, 1, 1, sourceDevFormulas.length);
          const formulasToSet = sourceDevFormulas.map(formula => 
            formula ? adjustCellReferences_(formula, i) : ''
          );
          targetRange.setFormulas([formulasToSet]);
        }
      }

      // Теперь копируем только ЗНАЧЕНИЯ из STUDENT в DEV
      // Получаем значения из STUDENT
      const studValues = shStud.getRange(r0, COL_B, k, 18).getValues();
      
      // Получаем формулы из DEV чтобы понять где можно перезаписывать значения
      const devFormulas = shDev.getRange(r0, COL_B, k, 18).getFormulas();
      
      // Создаем массив для вставки - только значения там где нет формул
      const valuesToSet = studValues.map((row, rowIndex) => 
        row.map((value, colIndex) => 
          // Если в DEV есть формула - оставляем null (не изменяем), иначе берем значение из STUDENT
          devFormulas[rowIndex][colIndex] && devFormulas[rowIndex][colIndex].startsWith('=') ? null : value
        )
      );

      if (!isRowGrouped_(shDev, r0)) {
        // Используем setValues с null чтобы не изменять ячейки с формулами
        const range = shDev.getRange(r0, COL_B, k, 18);
        const currentValues = range.getValues();
        
        // Объединяем значения: где null - оставляем текущее значение, иначе берем новое
        const finalValues = currentValues.map((currentRow, rowIndex) => 
          currentRow.map((currentValue, colIndex) => 
            valuesToSet[rowIndex][colIndex] === null ? currentValue : valuesToSet[rowIndex][colIndex]
          )
        );
        
        range.setValues(finalValues);
      }

      return { row: r0, count: k, hasSelectMarker: g.hasSelectMarker };
    });

    const results = await Promise.all(promises);
    
    const expandedCount = results.filter(r => r.hasSelectMarker).length;
    const copiedCount = results.length - expandedCount;
    
    SpreadsheetApp.getUi().alert(`✅ Готово! Обработано ${results.length} строк:\n- Развернуто списков: ${expandedCount}\n- Скопировано как есть: ${copiedCount}\nФормулы в DEV сохранены!`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка [3. Раскрыть смыслы]: ' + (e.stack || e.message || e));
  }
}

/***** === АСИНХРОННАЯ функция для параллельного раскрытия формул в столбцах ===*****/
async function unfoldFormulasInColumnsAsync_(fileId, colIndexes) {
  // Открываем таблицу по IDв
  const ss = SpreadsheetApp.openById(fileId);
  // Получаем все листы таблицы
  const sheets = ss.getSheets();
  
  // Создаем массив промисов для каждого листа
  const sheetPromises = sheets.map(async (sh) => {
    // Получаем последнюю строку на листе
    const lastRow = sh.getLastRow();
    // Если лист пустой - пропускаем
    if (lastRow < 1) return;

    // Создаем массив промисов для КАЖДОГО СТОЛБЦА в этом листе
    const columnPromises = colIndexes.map(async (col) => {
      // Получаем диапазон для всего столбца
      const range = sh.getRange(1, col, lastRow, 1);
      // Получаем отображаемые значения (видимый текст в ячейках)
      const values = range.getDisplayValues();
      // Флаг для отслеживания изменений
      let changed = false;
      
      // Проходим по всем строкам в столбце
      for (let r = 0; r < values.length; r++) {
        const txt = values[r][0] || '';
        // Если текст начинается с маркера формулы "@@="
        if (txt.startsWith(FORMULA_MARKER)) {
          // Заменяем "@@=FORMULA" на "=FORMULA" - превращаем текст в активную формулу
          values[r][0] = '=' + txt.substring(FORMULA_MARKER_LENGTH);
          changed = true;
        }
      }
      
      // Если были изменения в этом столбце
      if (changed) {
        // Устанавливаем формулы обратно в ячейки
        // Теперь это активные формулы, которые начнут вычисляться
        range.setFormulas(values);
        // Небольшая задержка между обработкой столбцов для стабильности
        await Utilities.sleep(100);
      }
    });
    
    // Ожидаем завершения обработки ВСЕХ столбцов на этом листе ПАРАЛЛЕЛЬНО
    // Все столбцы (E, F, G, H) обрабатываются одновременно!
    await Promise.all(columnPromises);
  });
  
  // Ожидаем завершения обработки ВСЕХ листов ПАРАЛЛЕЛЬНО
  await Promise.all(sheetPromises);
}

/***** === Альтернативная версия с более агрессивным параллелизмом ===*****/
async function unfoldFormulasInColumnsAggressive_(fileId, colIndexes) {
  const ss = SpreadsheetApp.openById(fileId);
  const sheets = ss.getSheets();
  
  // Создаем один большой массив промисов для всех столбцов всех листов
  const allPromises = [];
  
  sheets.forEach(sh => {
    const lastRow = sh.getLastRow();
    if (lastRow < 1) return;

    // Для каждого столбца создаем отдельный промис
    colIndexes.forEach(col => {
      const promise = new Promise((resolve) => {
        try {
          const range = sh.getRange(1, col, lastRow, 1);
          const values = range.getDisplayValues();
          let changed = false;
          
          for (let r = 0; r < values.length; r++) {
            const txt = values[r][0] || '';
            if (txt.startsWith(FORMULA_MARKER)) {
              values[r][0] = '=' + txt.substring(FORMULA_MARKER_LENGTH);
              changed = true;
            }
          }
          
          if (changed) {
            range.setFormulas(values);
          }
          resolve(`Column ${col} processed`);
        } catch (e) {
          resolve(`Column ${col} error: ${e.message}`);
        }
      });
      
      allPromises.push(promise);
    });
  });
  
  // Запускаем ВСЕ операции параллельно без задержек
  await Promise.all(allPromises);
}

/***** === Альтернативная версия с параллельной обработкой столбцов ===*****/
function menuExpandSurgically_Parallel() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    
    const ssStud = SpreadsheetApp.openById(studentId);
    const shStud = ssStud.getActiveSheet();
    const sheetName = shStud.getName();
    
    const ssDev = SpreadsheetApp.openById(devId);
    const shDev = ssDev.getSheetByName(sheetName) || ssDev.insertSheet(sheetName);

    const groups = collectSelectedRows_WithParsedLists_(shStud);
    if (!groups.length) {
      SpreadsheetApp.getUi().alert('Не найдено строк с маркером "> ОТБЕРИТЕ" на активном листе.');
      return;
    }

    // Параллельное разворачивание формул в разных столбцах
    const formulaColumns = [COL_E, COL_F, COL_G, COL_H];
    const unfoldPromises = formulaColumns.map(col => {
      return new Promise((resolve) => {
        try {
          unfoldSingleColumnFormulas_(devId, col);
          resolve(`Column ${col} done`);
        } catch (e) {
          resolve(`Column ${col} error: ${e.message}`);
        }
      });
    });

    // Ждем завершения разворачивания формул
    Promise.all(unfoldPromises).then(() => {
      SpreadsheetApp.flush();
      
      groups.sort((a, b) => b.rowIndex - a.rowIndex);

      // Обрабатываем группы строк
      groups.forEach(g => {
        const r0 = g.rowIndex;
        const { k, B, C, D } = g.meta;
        if (!k || k < 1) return;

        // Работа со STUDENT
        if (k > 1) {
          shStud.insertRowsAfter(r0, k - 1);
          copyRowFormat_(shStud, r0, r0 + 1, k - 1);
        }
        
        const blockData = [];
        for(let i=0; i<k; i++) {
          const bVal = B[i] !== undefined ? B[i] : (B.length === 1 ? B[0] : '');
          const cVal = C[i] !== undefined ? C[i] : (C.length === 1 ? C[0] : '');
          const dVal = D[i] !== undefined ? D[i] : (D.length === 1 ? D[0] : '');
          blockData.push([bVal, cVal, dVal]);
        }
        shStud.getRange(r0, COL_B, k, 3).setValues(blockData);
        SpreadsheetApp.flush();

        // Работа с DEV
        if (k > 1) {
          shDev.insertRowsAfter(r0, k - 1);
        }
        
        const valuesBlockBCD = shStud.getRange(r0, COL_B, k, 3).getValues();
        shDev.getRange(r0, COL_B, k, 3).setValues(valuesBlockBCD);
        
        // Параллельная установка формул в разных столбцах
        const formulaPromises = formulaColumns.map((col, index) => {
          return new Promise((resolve) => {
            try {
              const templateFormula = shDev.getRange(r0, col, 1, 1).getFormulas()[0][0];
              const newFormulas = [];
              for (let i = 0; i < k; i++) {
                const newFormula = adjustCellReferences_(templateFormula, i);
                newFormulas.push([newFormula]);
              }
              shDev.getRange(r0, col, k, 1).setFormulas(newFormulas);
              resolve(`Column ${col} formulas set`);
            } catch (e) {
              resolve(`Column ${col} error: ${e.message}`);
            }
          });
        });
        
        Promise.all(formulaPromises).then(() => {
          console.log(`Formulas set for row ${r0}`);
        });
      });

      SpreadsheetApp.getUi().alert('✅ Готово! Смыслы раскрыты и синхронизированы в DEV (параллельная версия).');
      
    }).catch(error => {
      SpreadsheetApp.getUi().alert('Ошибка при разворачивании формул: ' + error);
    });

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка [3. Раскрыть смыслы]: ' + (e.stack || e.message || e));
  }
}

/***** === Вспомогательная функция для разворачивания одного столбца ===*****/
function unfoldSingleColumnFormulas_(fileId, colIndex) {
  const ss = SpreadsheetApp.openById(fileId);
  ss.getSheets().forEach(sh => {
    const lastRow = sh.getLastRow();
    if (lastRow < 1) return;

    const range = sh.getRange(1, colIndex, lastRow, 1);
    const values = range.getDisplayValues();
    let changed = false;
    
    for (let r = 0; r < values.length; r++) {
      const txt = values[r][0] || '';
      if (txt.startsWith(FORMULA_MARKER)) {
        values[r][0] = '=' + txt.substring(FORMULA_MARKER_LENGTH);
        changed = true;
      }
    }
    
    if (changed) {
      range.setFormulas(values);
    }
  });
}


/***** === 2. DEV → STUDENT (ИСПРАВЛЕННАЯ ВЕРСИЯ) ===*****/
function menuDeliverToStudent_AutoContext() {
  try {
    const { sheet, row } = resolveRegistryRowContext_();
    let devUrl = String(sheet.getRange(row, COLS.devUrl).getValue() || '').trim();
    const devId = devUrl ? fileIdFromUrl_(devUrl) : SpreadsheetApp.getActive().getId();
    if (!devId) throw new Error('Нет DEV.');
    const order = String(sheet.getRange(row, COLS.order).getValue()||'').trim();
    if (!order) throw new Error('ID заказа пуст.');
    let studUrlExisting = String(sheet.getRange(row, COLS.studentUrl).getValue() || '').trim();
    let studId;
    if (studUrlExisting) {
      studId = fileIdFromUrl_(studUrlExisting);
      try { DriveApp.getFileById(studId).getId(); }
      catch (e) { studId = null; }
    }
    if (!studId) {
      const styleUrl = String(sheet.getRange(REG_STYLE_MASTER_CELL).getValue()||'').trim();
      if (!styleUrl) throw new Error(`В REGISTRY!${REG_STYLE_MASTER_CELL} нет STYLE MASTER`);
      const styleId = fileIdFromUrl_(styleUrl);
      const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
      const studFile = DriveApp.getFileById(styleId).makeCopy(`STUDENT — ${order}`, folder);
      studId = studFile.getId();
      DriveApp.getFileById(studId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      const studUrl = `https://docs.google.com/spreadsheets/d/${studId}/edit`;
      sheet.getRange(row, COLS.studentUrl).setValue(studUrl);
    }
    
    pasteColsBCD_FromDevToStud_(devId, studId);
    
    const finalStudUrl = `https://docs.google.com/spreadsheets/d/${studId}/edit`;
    showLink_('STUDENT готов (скопированы B:D из DEV)', finalStudUrl, 'ПЕРЕЙТИ В STUD');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка DEV → STUDENT: ' + (e.message || e));
  }
}


/***** === 1. Создать DEV (ИСПРАВЛЕННАЯ ВЕРСИЯ) ===*****/
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


/***** === БЫСТРОЕ РАСКРЫТИЕ ФОРМУЛ (ИСПРАВЛЕННАЯ ВЕРСИЯ) ===*****/
function unfoldFormulasInColumns_(fileId, colIndexes) {
  const ss = SpreadsheetApp.openById(fileId);
  ss.getSheets().forEach(sh => {
    const lastRow = sh.getLastRow();
    if (lastRow < 1) return;

    colIndexes.forEach(col => {
      const range = sh.getRange(1, col, lastRow, 1);
      const values = range.getDisplayValues();
      let changed = false;
      for (let r = 0; r < values.length; r++) {
        const txt = values[r][0] || '';
        if (txt.startsWith(FORMULA_MARKER)) {
          values[r][0] = '=' + txt.substring(FORMULA_MARKER_LENGTH);
          changed = true;
        }
      }
      if (changed) {
        range.setFormulas(values);
      }
    });
  });
}


/***** === ПОМОЩНИК для протягивания ссылок === *****/
function adjustCellReferences_(text, offset) {
  if (typeof text !== 'string' || !text || offset === 0) return text;
  
  return text.replace(/&([BCD])(\d+)&/gi, (match, col, numStr) => {
    const n = parseInt(numStr, 10);
    if (isNaN(n) || n <= 20) {
      return match;
    }
    const newRow = n + offset;
    return `&${col.toUpperCase()}${newRow}&`;
  });
}


/***** === 4. Отдать раскрытое (DEV → STUDENT) ===*****/
function menuDeliverExpanded_Final() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    const ssDev = SpreadsheetApp.openById(devId);
    const shDev = ssDev.getActiveSheet();
    const sheetName = shDev.getName();
    const ssStud = SpreadsheetApp.openById(studentId);
    const shStud = ssStud.getSheetByName(sheetName);
    if (!shStud) throw new Error(`В файле STUDENT не найден лист с именем "${sheetName}"`);
    const lastRow = shDev.getLastRow();
    if (lastRow < 1) return;
    const rangeDev = shDev.getRange(1, COL_E, lastRow, 4); // E:H
    const values = rangeDev.getDisplayValues(); 
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (typeof values[r][c] === 'string' && values[r][c].startsWith(FORMULA_MARKER)) {
          values[r][c] = '';
        }
      }
    }
    shStud.getRange(1, COL_E, lastRow, 4).setValues(values);
    SpreadsheetApp.getUi().alert('✅ Готово! Данные из DEV (E:H) перенесены в STUDENT. Ячейки с @@= очищены.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка [4. Отдать раскрытое]: ' + (e.message || e));
  }
}

/***** === УТИЛИТА: Формулы ⇄ Текст (с @@=) ===*****/
function toggleFormulasOnSelection() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('Выделите ячейки и повторите.'); return; }
  const disp = range.getDisplayValues();
  const forms = range.getFormulas();
  for (let r=0;r<range.getNumRows();r++){
    for (let c=0;c<range.getNumColumns();c++){
      const cell = range.getCell(r+1,c+1);
      const f = forms[r][c]||'';
      const d = disp[r][c]||'';
      if (f.startsWith('=')) cell.setValue(FORMULA_MARKER + f.substring(1));
      else if (d.startsWith(FORMULA_MARKER)) cell.setFormula('=' + d.substring(FORMULA_MARKER_LENGTH));
    }
  }
}

/***** === 5. Телесуфлёр === *****/
function menuTeleprompter_InPlace() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    const shStud = SpreadsheetApp.openById(studentId).getActiveSheet();
    const ssDev  = SpreadsheetApp.openById(devId);
    const shDev  = ssDev.getSheetByName(shStud.getName()) || ssDev.insertSheet(shStud.getName());
    const picked=[];
    for(let r=2;r<=shStud.getLastRow();r++){
      const e=shStud.getRange(r,COL_E).getValue()||'';
      const f=shStud.getRange(r,COL_F).getValue()||'';
      const g=shStud.getRange(r,COL_G).getValue()||'';
      const txt=String(e||f||g).trim();
      if(txt) picked.push({rowIndex:r,B:txt});
    }
    if(!picked.length){ SpreadsheetApp.getUi().alert('Нет строк для телесуфлёра'); return; }
    picked.forEach(pr=>{
      ensureRowsAndCols_(shDev,pr.rowIndex,COL_H);
      shDev.getRange(pr.rowIndex,COL_B).setValue(pr.B);
      shDev.getRange(pr.rowIndex,COL_H).setFormula(FORMULA_MARKER+'GPT(...)');
      Utilities.sleep(SLEEP_MS);
      const vH=shDev.getRange(pr.rowIndex,COL_H).getValue();
      shStud.getRange(pr.rowIndex,COL_B).setValue(pr.B);
      shStud.getRange(pr.rowIndex,COL_H).setValue(vH);
    });
    SpreadsheetApp.getUi().alert('Телесуфлёр готов');
  } catch(e){ SpreadsheetApp.getUi().alert('Ошибка телесуфлёр: '+(e.message||e)); }
}

/***** === ПОМОЩНИКИ И СТАРЫЕ ФУНКЦИИ ===*****/

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

/***** === new code ===*****/


function explodeNumberedListToRows() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeRange = sheet.getActiveRange();
    
    if (!activeRange || activeRange.getNumRows() !== 1 || activeRange.getNumColumns() !== 1) {
      throw new Error('Выберите ОДНУ ячейку с нумерованным списком');
    }

    const cell = activeRange;
    const cellValue = cell.getDisplayValue();
    const row = cell.getRow();
    const col = cell.getColumn();

    if (!cellValue) {
      throw new Error('Выбранная ячейка пустая');
    }

    const leftCellCol = col - 1; // Столбец слева от выделенного
    if (leftCellCol < 1) {
      throw new Error('Нет ячейки слева от выделенной');
    }
    const leftCell = sheet.getRange(row, leftCellCol);
    const leftCellValue = leftCell.getDisplayValue();
    const hasOtberite = leftCellValue.toLowerCase().includes('отберите');

    // Парсим нумерованный список
    const items = parseNumberedList_(cellValue);
    
    if (items.length === 0) {
      throw new Error('Не найдено нумерованных пунктов в ячейке');
    }

    let startRow = row + 1; // Начинаем запись со следующей строки

    // Если есть "отберите" слева - создаем новые строки
    if (hasOtberite) {
      // Вставляем строки НИЖЕ исходной ячейки
      sheet.insertRowsAfter(row, items.length);
    } else {
      // Если нет "отберите" - проверяем, хватает ли существующих строк
      const lastRow = sheet.getLastRow();
      const availableRows = lastRow - row;
      
      if (availableRows < items.length) {
        // Если не хватает строк - добавляем только недостающие
        const rowsToAdd = items.length - availableRows;
        sheet.insertRowsAfter(lastRow, rowsToAdd);
      }
    }

    // Записываем каждый пункт в отдельную строку НИЖЕ исходной
    for (let i = 0; i < items.length; i++) {
      sheet.getRange(startRow + i, col).setValue(items[i]);
    }

    SpreadsheetApp.getUi().alert(`✅ Создано ${items.length} строк ниже${hasOtberite ? ' (с новыми строками)' : ''}`);

  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Ошибка: ${error.message}`);
  }
}

function parseNumberedList_(text) {
  const cleanedText = String(text || '')
    .replace(/\r\n?/g, '\n')
    .replace(/\u00A0/g, ' ')
    .trim();

  if (!cleanedText) return [];

  const items = [];
  const lines = cleanedText.split('\n');
  
  for (const line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine) continue;

    // Проверяем есть ли "отберите" (регистронезависимо)
    if (trimmedLine.toLowerCase().includes('отберите')) {
      items.push(trimmedLine);
    }
    // ИЛИ оставляем старую логику с нумерацией
    else {
      const match = trimmedLine.match(/^\s*(\d{1,2})[\.\)]\s*(.+)$/);
      if (match) {
        items.push(match[2].trim());
      }
    }
  }

  return items;
}

// Функция для показа логов при успехе
function showLogsAndSuccess_(logs, sourceType, destinationType, url, numRows, numCols) {
  const logText = logs.join('<br>');
  
  const html = HtmlService.createHtmlOutput(
    `<div style="font:14px/1.4 system-ui,Arial;padding:20px;max-height:400px;overflow-y:auto;">
       <div style="background:#d4edda;color:#155724;padding:15px;border-radius:8px;margin-bottom:15px;">
         <strong>✅ Успешно скопировано!</strong><br>
         📊 ${numRows}×${numCols} ячеек<br>
         📤 ${sourceType} → ${destinationType}
       </div>
       <div style="background:#f8f9fa;padding:15px;border-radius:8px;border:1px solid #ddd;">
         <strong>Детали выполнения:</strong><br>
         <div style="margin-top:10px;font-family:monospace;font-size:12px;">
           ${logText}
         </div>
       </div>
       <div style="margin-top:15px;text-align:center;">
         <a href="${url}" target="_blank"
            style="display:inline-block;padding:10px 20px;background:#1a73e8;color:#fff;border-radius:6px;text-decoration:none;">
           📂 Открыть ${destinationType}
         </a>
       </div>
     </div>`
  ).setWidth(600).setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Результат выполнения');
}

// Функция для показа логов при ошибке
function showLogsAndError_(logs, errorMessage) {
  const logText = logs.join('<br>');
  
  const html = HtmlService.createHtmlOutput(
    `<div style="font:14px/1.4 system-ui,Arial;padding:20px;max-height:400px;overflow-y:auto;">
       <div style="background:#f8d7da;color:#721c24;padding:15px;border-radius:8px;margin-bottom:15px;">
         <strong>❌ Ошибка:</strong> ${errorMessage}
       </div>
       <div style="background:#f8f9fa;padding:15px;border-radius:8px;border:1px solid #ddd;">
         <strong>Логи выполнения:</strong><br>
         <div style="margin-top:10px;font-family:monospace;font-size:12px;">
           ${logText}
         </div>
       </div>
       <div style="margin-top:15px;text-align:center;">
         <button onclick="google.script.host.close()"
                 style="padding:8px 16px;background:#6c757d;color:#fff;border:none;border-radius:6px;cursor:pointer;">
           Закрыть
         </button>
       </div>
     </div>`
  ).setWidth(600).setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Ошибка выполнения');
}

// Функция извлечения ID (улучшенная версия)
function extractOrderIdFromFileName_(name) {
  if (!name) return '';
  
  // Пробуем разные варианты разделителей
  const parts1 = name.split('—').map(s => s.trim()); // Длинное тире
  const parts2 = name.split('-').map(s => s.trim());  // Дефис
  const parts3 = name.split('–').map(s => s.trim());  // Короткое тире
  
  let orderId = '';
  if (parts1.length >= 2) orderId = parts1[1];
  else if (parts2.length >= 2) orderId = parts2[1]; 
  else if (parts3.length >= 2) orderId = parts3[1];
  
  // Убираем возможные лишние символы
  orderId = orderId.replace(/[^0-9]/g, '');
  
  return orderId;
}

/***** === new code ===*****/

function pasteSelectedValues_Bidirectional() {
  try {
    const currentFile = SpreadsheetApp.getActive();
    const currentFileName = currentFile.getName();
    const currentSheet = currentFile.getActiveSheet(); // Получаем активную вкладку
    const currentSheetName = currentSheet.getName(); // Получаем имя активной вкладки

    // Извлекаем ID заказа из названия файла
    const orderId = extractOrderIdFromFileName_(currentFileName);
    if (!orderId) throw new Error('Не удалось извлечь ID заказа из названия файла');

    // Определяем направление копирования
    let destinationName, isToStudent;
    if (/^DEV\s—\s/i.test(currentFileName)) {
      destinationName = `STUDENT — ${orderId}`;
      isToStudent = true;
    } else if (/^STUDENT\s—\s/i.test(currentFileName)) {
      destinationName = `DEV — ${orderId}`;
      isToStudent = false;
    } else {
      throw new Error('Имя файла должно начинаться с "DEV — " или "STUDENT — "');
    }

    // Ищем партнерский файл
    const files = DriveApp.getFilesByName(destinationName);
    if (!files.hasNext()) throw new Error(`Файл "${destinationName}" не найден`);
    
    const destinationId = files.next().getId();
    const dstSS = SpreadsheetApp.openById(destinationId);

    // Получаем или создаем вкладку с таким же именем в целевом файле
    let dstSheet;
    try {
      dstSheet = dstSS.getSheetByName(currentSheetName);
      if (!dstSheet) {
        // Если вкладки нет - создаем её
        dstSheet = dstSS.insertSheet(currentSheetName);
        console.log(`Создана новая вкладка: ${currentSheetName}`);
      }
    } catch (e) {
      throw new Error(`Ошибка при работе с вкладкой "${currentSheetName}": ${e.message}`);
    }

    // Копируем выделенный диапазон с активной вкладки
    const activeRange = currentSheet.getActiveRange();
    if (!activeRange) throw new Error('Не выделен диапазон для копирования');

    // Создаем целевой диапазон на соответствующей вкладке
    const destinationRange = dstSheet.getRange(
      activeRange.getRow(),
      activeRange.getColumn(),
      activeRange.getNumRows(),
      activeRange.getNumColumns()
    );
    
    // Проверяем и создаем строки/колонки если нужно
    ensureRowsAndCols_(dstSheet, 
      activeRange.getRow() + activeRange.getNumRows() - 1,
      activeRange.getColumn() + activeRange.getNumColumns() - 1
    );

    // Получаем данные из исходного диапазона
    const values = activeRange.getValues();
    const formulas = activeRange.getFormulas();
    const displayValues = activeRange.getDisplayValues();
    
    // Подготавливаем данные для вставки
    const dataToPaste = [];
    for (let i = 0; i < values.length; i++) {
      const row = [];
      for (let j = 0; j < values[i].length; j++) {
        if (isToStudent) {
          // В STUDENT копируем только значения (без формул)
          row.push(values[i][j]);
        } else {
          // В DEV копируем как есть, но корректно обрабатываем типы данных
          const hasFormula = formulas[i][j] && formulas[i][j] !== '';
          if (hasFormula) {
            // Если есть формула - используем её
            row.push(formulas[i][j]);
          } else {
            // Если нет формулы - используем оригинальное значение
            // Но проверяем, не является ли оно числом в текстовом формате
            const displayVal = displayValues[i][j];
            const originalVal = values[i][j];
            
            // Если это число, но отображается как текст (например "123.0")
            if (typeof originalVal === 'number' && String(originalVal) === displayVal) {
              row.push(originalVal);
            } else {
              row.push(originalVal);
            }
          }
        }
      }
      dataToPaste.push(row);
    }
    
    // Вставляем данные в целевой диапазон
    try {
      if (isToStudent) {
        // В STUDENT - только значения
        destinationRange.setValues(dataToPaste);
      } else {
        // В DEV - используем интеллектуальную вставку
        intelligentPaste_(destinationRange, dataToPaste, formulas);
      }
    } catch (e) {
      throw new Error(`Ошибка при вставке данных: ${e.message}`);
    }

    const direction = isToStudent ? 'DEV → STUDENT' : 'STUDENT → DEV';
    const copyType = isToStudent ? 'только значения' : 'значения и формулы';
    
    showLink_(
      `✅ Скопировано ${activeRange.getNumRows()}×${activeRange.getNumColumns()} ячеек\n` +
      `Вкладка: ${currentSheetName}\n` +
      `Направление: ${direction}\n` +
      `Тип: ${copyType}`,
      dstSS.getUrl(),
      'Открыть файл'
    );

  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Ошибка: ${error.message}`);
  }
}

// Новая функция для интеллектуальной вставки
function intelligentPaste_(destinationRange, dataToPaste, originalFormulas) {
  const hasAnyFormula = originalFormulas.some(row => 
    row.some(cell => cell && cell !== '')
  );
  
  if (hasAnyFormula) {
    // Если есть формулы, пробуем вставить как формулы
    try {
      destinationRange.setFormulas(dataToPaste);
      return;
    } catch (e) {
      // Если не получилось, вставляем как значения
      console.log('Не удалось вставить формулы, используем значения:', e);
    }
  }
  
  // Вставляем как значения
  const valuesOnly = dataToPaste.map(row => 
    row.map(cell => {
      // Если ячейка содержит формулу как текст (начинается с =), 
      // но это не настоящая формула, убираем =
      if (typeof cell === 'string' && cell.startsWith('=') && 
          !originalFormulas.flat().includes(cell)) {
        return cell.substring(1);
      }
      return cell;
    })
  );
  
  destinationRange.setValues(valuesOnly);
}

// Упрощенная функция извлечения ID
function extractOrderIdFromFileName_(name) {
  if (!name) return '';
  const parts = name.split('—').map(s => s.trim());
  return parts.length >= 2 ? parts[1].replace(/[^0-9]/g, '') : '';
}

/***** === end new code ===*****/


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

function copyRowFormat_(sheet, srcRow, dstStartRow, count) {
    if (count <= 0) return;
    const maxCols = sheet.getMaxColumns();
    const sourceRange = sheet.getRange(srcRow, 1, 1, maxCols);
    for (let i = 0; i < count; i++) {
        const destRange = sheet.getRange(dstStartRow + i, 1, 1, maxCols);
        sourceRange.copyTo(destRange, { formatOnly: true });
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
  if (!devUrl)      throw new Error('В реестре нет DEV. Сначала запусти «1. Создать DEV».');
  if (!studentUrl) throw new Error('В реестре нет STUDENT. Сначала запусти «2. DEV → STUDENT».');
  return { devId:fileIdFromUrl_(devUrl), studentId:fileIdFromUrl_(studentUrl) };
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


function f1() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    
    if (!range) {
      SpreadsheetApp.getUi().alert('Выделите диапазон и повторите.');
      return;
    }

    const formulas = range.getFormulas();
    let replacedCount = 0;

    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        const formula = formulas[r][c];
        
        if (formula && formula.toLowerCase().includes('gpt(')) {
          // Просто добавляем IF(E$2=""; перед GPT и закрываем в конце
          const newFormula = `=IF($C$7=""; ${formula.substring(1)}; "")`;
          
          const cell = range.getCell(r + 1, c + 1);
          cell.setFormula(newFormula);
          replacedCount++;
        }
      }
    }

    if (replacedCount > 0) {
      SpreadsheetApp.getUi().alert(`✅ Добавлен IF к ${replacedCount} формулам GPT`);
    } else {
      SpreadsheetApp.getUi().alert('Не найдено формул с gpt( в выделенном диапазоне');
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка: ' + (e.message || e));
  }
}
function f2() {
  
}
function f3() {
  
}
function f4() {
  
}
function f5() {
  
}
function f6() {
  
}
function f7() {
  
}
function f8() {
  
}
function f9() {
  
}
function f10() {
  
}
function f11() {
  
}
function f12() {
  
}
function f13() {
  
}
function f14() {
  
}
function f15() {
  
}
function collectSelectedRows_WithParsedLists_(shStud){
  const res = [];
  const last = shStud.getLastRow();
  if (last < 1) return res;

  const A = shStud.getRange(1, COL_A, last, 1).getDisplayValues().map(r=>String(r[0]||''));
  const B = shStud.getRange(1, COL_B, last, 1).getDisplayValues().map(r=>String(r[0]||''));
  const C = shStud.getRange(1, COL_C, last, 1).getDisplayValues().map(r=>String(r[0]||''));
  const D = shStud.getRange(1, COL_D, last, 1).getDisplayValues().map(r=>String(r[0]||''));

  for (let r = 1; r <= last; r++){
    // Пропускаем сгруппированные строки
    if (isRowGrouped_(shStud, r)) {
      continue;
    }

    const aClean = (A[r-1] || '').replace(/[\u200B\u200C\u200D\uFEFF]/g, '').replace(/\u00A0/g, ' ').trim();

    // Обрабатываем ТОЛЬКО строки с ">"
    const hasSelectMarker = aClean.includes(MARK_SELECT);
    if (!hasSelectMarker) {
      // Для строк без ">" - добавляем как есть (k=1)
      const meta = { 
        k: 1, 
        B: [B[r-1].trim()], 
        C: [C[r-1].trim()], 
        D: [D[r-1].trim()] 
      };
      res.push({ rowIndex: r, meta, hasSelectMarker: false });
      continue;
    }

    // Для строк с ">" - разбираем списки
    const listB = parseNumberedList_(B[r-1]);
    const listC = parseNumberedList_(C[r-1]);
    const listD = parseNumberedList_(D[r-1]);
    const valB = listB.length ? listB : (B[r-1].trim() ? [B[r-1].trim()] : []);
    const valC = listC.length ? listC : (C[r-1].trim() ? [C[r-1].trim()] : []);
    const valD = listD.length ? listD : (D[r-1].trim() ? [D[r-1].trim()] : []);
    const k = Math.max(valB.length, valC.length, valD.length, 1);
    
    const meta = { k, B: valB, C: valC, D: valD };
    res.push({ rowIndex: r, meta, hasSelectMarker: true });
  }
  return res;
}