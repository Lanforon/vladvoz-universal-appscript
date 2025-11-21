/***** НАСТРОЙКИ ДЛЯ ПОЛЬЗОВАТЕЛЯ *****/

const ALLOWED_EMAILS = ['work@vladvoz.com']; // ИЗМЕНИТЕ НА ДЕЙСТВУЮЩИЙ EMAIL

const REGISTRY_FILE_ID = '1TEksg-gFc5rgPAcgUC7aOrVsJKhCrw4-UPUTSqxVaF8'; 
const REG_SHEET = 'REGISTRY';
const START_ROW = 2; // для реестра
const NameMainTable = "БАЗА"

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
  expert: 7,     // G - Эксперт → B2/C2/D2/E2/F2/G2
  aud1: 8,       // H - Аудитория 1 → B1
  aud2: 9,       // I - Аудитория 2 → C1  
  aud3: 10,      // J - Аудитория 3 → D1
  aud4: 11,      // K - Аудитория 4 → E2
  aud5: 12,      // L - Аудитория 5 → F2
  aud6: 13,      // M - Аудитория 6 → G2
  expertProgram: 14 // N - Программа эксперта (уходит в B4)
};

const COL_A = 1, COL_B = 2, COL_C = 3, COL_D = 4, COL_E = 5, COL_F = 6, COL_G = 7, COL_H = 8;

// КОНСТАНТЫ ДЛЯ ЦЕЛЕВЫХ ЯЧЕЕК В DEV
const DEV_AUD1_CELL = 'B1'; // Аудитория 1
const DEV_AUD2_CELL = 'C1'; // Аудитория 2  
const DEV_AUD3_CELL = 'D1'; // Аудитория 3
const DEV_AUD4_CELL = 'E2'; // Аудитория 4
const DEV_AUD5_CELL = 'F2'; // Аудитория 5
const DEV_AUD6_CELL = 'G2'; // Аудитория 6
const DEV_EXPERT_CELL_BASE = 'B2'; // Базовая ячейка для эксперта
const DEV_EXPERT_PROGRAM_CELL = 'B4'; // Программа эксперта

const MARK_SELECT = '>';










function onOpen() {
  const currentFile = SpreadsheetApp.getActive();
  const currentFileName = currentFile.getName();
  
  const menu = SpreadsheetApp.getUi().createMenu('Технический');
  
  // Для таблицы БАЗА
  if (/БАЗА/i.test(currentFileName)) {
    menu
      .addSeparator()
      .addItem('СОЗДАТЬ DEV - КЛУБ', 'menuDevelopFactory')
      .addSeparator()
      .addItem('СОЗДАТЬ DEV - НЕ КЛУБ', 'menuDevelopNoFactory')
      .addSeparator();
  }
  
  // Для таблиц DEV 
  if (/DEV/i.test(currentFileName)) {
    menu
      .addItem('Создать STUDENT — для ученика', 'menuDeliverToStudent_AutoContext')
      .addSeparator()
      .addItem('ШАГ 1-4 — Отдать BCD [DEV > STUD] ', 'f2')
      .addItem('ШАГ 1-4 — Забрать BCD [STUD > DEV]', 'pasteSelectedValues_Bidirectional')
      .addItem('ШАГ 5 — Раскрыть > строки [в DEV]', 'menuExpandSurgically_Final') 
      .addItem('ШАГ 5 — Отдать ВКЛАДКУ [DEV > STUD]', 'menuDeliverExpanded_Final')
      .addSeparator()
      .addItem('ШАГ 6 — Забрать EFG [STUD > DEV] → в Е', 'f1')
      .addItem('ШАГ 6 — Отдать ВКЛАДКУ [DEV > STUD] → в Е', 'f1')
      .addSeparator()
      .addItem('🔄 Добавить IF к GPT', 'f3')
      .addItem('🔍 Проверить ERROR ячейки', 'f5'); 
  }

  menu.addToUi();
}

// Обработчик для запрета удаления строк и столбцов в STUDENT файлах
function onChange(e) {
  try {
    const source = e.source;
    const currentFileName = source.getName();
    
    // Проверяем, является ли файл STUDENT файлом
    if (/STUDENT/i.test(currentFileName)) {
      const changeType = e.changeType;
      
      // Проверяем тип изменения - удаление строк или столбцов
      if (changeType === 'REMOVE_ROW' || changeType === 'REMOVE_COLUMN') {
        // Показываем сообщение пользователю
        SpreadsheetApp.getUi().alert(
          '❌ Запрещено удалять!', 
          'В файлах STUDENT запрещено удалять строки и столбцы!\n\nМожно:\n• Редактировать содержимое ячеек\n• Добавлять новые строки/столбцы\n• Изменять форматирование\n\nЗапрещено:\n• Удалять строки\n• Удалять столбцы', 
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        
        // Предлагаем отменить действие
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
          'Восстановить структуру?',
          'Рекомендуется немедленно отменить удаление (Ctrl+Z).\nХотите показать инструкцию?',
          ui.ButtonSet.YES_NO
        );
        
        if (response === ui.Button.YES) {
          showUndoInstructions();
        }
      }
    }
  } catch (error) {
    console.error('Ошибка в onChange:', error);
  }
}

function showUndoInstructions() {
  const message = 
    '📋 Инструкция по отмене удаления:\n\n' +
    'Windows:\n• Нажмите Ctrl + Z\n\n' +
    'Mac:\n• Нажмите Cmd + Z\n\n' +
    'Или через меню:\n• Правка → Отменить\n\n' +
    'Это восстановит удаленные строки/столбцы.';
  
  SpreadsheetApp.getUi().alert('↩️ Как отменить удаление', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function onEdit(e) {

}

function menuExpandSurgically_Final() {
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

    // Собираем строки с маркером ">" в колонке A из STUDENT
    const rowsWithMarker = [];
    const aValues = shStud.getRange(1, 1, lastRow, 1).getDisplayValues();
    
    for (let r = 0; r < aValues.length; r++) {
      const aValue = String(aValues[r][0] || '').trim();
      // Ищем строки с маркером ">" в колонке A и пропускаем сгруппированные
      if (aValue.includes(MARK_SELECT) && !isRowGrouped_(shStud, r + 1)) {
        rowsWithMarker.push(r + 1);
      }
    }

    console.log('Найдено строк с маркером ">" в колонке A:', rowsWithMarker);

    if (rowsWithMarker.length === 0) {
      SpreadsheetApp.getUi().alert('Не найдено строк с маркером ">" в колонке A STUDENT');
      return;
    }

    // Обрабатываем каждую строку с маркером (ТОЛЬКО В DEV)
    let expandedCount = 0;
    
    // Обрабатываем строки в обратном порядке чтобы не сбивать нумерацию
    rowsWithMarker.reverse().forEach(row => {
      // ПОЛУЧАЕМ ДАННЫЕ ИЗ DEV (ИЗМЕНЕНИЕ ЗДЕСЬ)
      const aValue = shDev.getRange(row, 1).getValue();
      const bValue = shDev.getRange(row, 2).getValue();
      const cValue = shDev.getRange(row, 3).getValue();
      const dValue = shDev.getRange(row, 4).getValue();
      
      console.log(`Строка ${row}: A="${aValue}", B="${bValue}", C="${cValue}", D="${dValue}"`);
      
      // Парсим нумерованные списки из колонок B, C, D DEV
      const bItems = parseNumberedList_(bValue);
      const cItems = parseNumberedList_(cValue);
      const dItems = parseNumberedList_(dValue);
      
      const maxItems = Math.max(bItems.length, cItems.length, dItems.length, 1);
      
      console.log(`Строка ${row}: B items=${bItems.length}, C items=${cItems.length}, D items=${dItems.length}, max=${maxItems}`);
      
      if (maxItems > 1) {
        console.log(`Раскрываем строку ${row} на ${maxItems} элементов`);
        
        shDev.insertRowsAfter(row, maxItems - 1);
        
        copyRowFormat_(shDev, row, row + 1, maxItems - 1);
        
        // Получаем формулы из исходной строки DEV
        const sourceDevFormulas = shDev.getRange(row, 1, 1, shDev.getLastColumn()).getFormulas()[0];
        
        // Дублируем формулы во все новые строки DEV с адаптацией ссылок
        for (let i = 1; i < maxItems; i++) {
          const targetRange = shDev.getRange(row + i, 1, 1, sourceDevFormulas.length);
          const formulasToSet = sourceDevFormulas.map(formula => 
            formula ? adjustCellReferences_(formula, i) : ''
          );
          targetRange.setFormulas([formulasToSet]);
        }
        
        // --- СПЕЦИАЛЬНАЯ ОБРАБОТКА ДЛЯ СТОЛБЦОВ E-H В DEV --
        // Получаем формулы шаблона из столбцов E-H исходной строки
        const templateFormulasEFGH = shDev.getRange(row, COL_E, 1, 4).getFormulas()[0];
        const newBlockFormulasEFGH = [];
        
        // Создаем адаптированные формулы для всех строк (включая исходную)
        for (let i = 0; i < maxItems; i++) {
          const newRow = templateFormulasEFGH.map(formulaText => 
            adjustCellReferences_(formulaText, i)
          );
          newBlockFormulasEFGH.push(newRow);
        }
        
        // Устанавливаем формулы для столбцов E-H во всех строках блока
        shDev.getRange(row, COL_E, maxItems, 4).setFormulas(newBlockFormulasEFGH);
        
        // Заполняем данные ТОЛЬКО В DEV (только в столбцы A-D, чтобы не перезаписать формулы)
        for (let i = 0; i < maxItems; i++) {
          const targetRow = row + i;
          shDev.getRange(targetRow, 1).setValue(aValue); // Колонка A без изменений
          shDev.getRange(targetRow, 2).setValue(bItems[i] || '');
          shDev.getRange(targetRow, 3).setValue(cItems[i] || '');
          shDev.getRange(targetRow, 4).setValue(dItems[i] || '');
        }
        
        expandedCount++;
      }
    });

    SpreadsheetApp.getUi().alert(`✅ Раскрыто ${expandedCount} строк с маркером ">" в DEV\nФормулы автоматически продублированы!`);

  } catch (e) {
    console.error('Ошибка:', e);
    SpreadsheetApp.getUi().alert('Ошибка при раскрытии смыслов: ' + (e.message || e));
  }
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

    // Получаем только EFG из STUDENT
    const studValues = shStud.getRange(1, 5, lastRow, 3).getValues(); 
    const studFormulas = shStud.getRange(1, 5, lastRow, 3).getFormulas(); 
    
    //  Создаем массив для объединенного значения в E (затираем формулы)
    const devValuesE = studValues.map((row, rowIndex) => {
      // Объединяем значения E+F+G в одну строку (без формул)
      const combinedValue = row
        .map((value, colIndex) => studFormulas[rowIndex][colIndex] ? '' : value) // Заменяем формулы пустотами
        .filter(val => val) // Убираем пустые значения
        .join(' '); // Объединяем через пробел
      
      return [combinedValue]; // Возвращаем массив с одним элементом для столбца E
    });

    // Создаем массивы для затирания формул в E, F, G
    const emptyValuesE = devValuesE; // E уже содержит значения без формул
    const emptyValuesF = Array(lastRow).fill().map(() => ['']); // Пустые значения для F
    const emptyValuesG = Array(lastRow).fill().map(() => ['']); // Пустые значения для G

    //  Записываем значения в E DEV (затираем формулы)
    shDev.getRange(1, 5, lastRow, 1).setValues(emptyValuesE);
    
    //  Затираем формулы в столбцах F и G DEV пустыми значениями
    shDev.getRange(1, 6, lastRow, 1).setValues(emptyValuesF);
    shDev.getRange(1, 7, lastRow, 1).setValues(emptyValuesG);

    SpreadsheetApp.getUi().alert(`✅ Значения E-F-G из STUDENT перенесены в E DEV, все формулы в E-F-G затерты`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка при переносе EFG в E: ' + (e.message || e));
  }
}

function f5() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  let errorCells = [];
  
  // Проверяем каждую ячейку в диапазоне данных
  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      const value = values[row][col];
      
      // Проверяем на ошибки Google Sheets
      if (value === '#ERROR!' || value === '#N/A' || value === '#VALUE!' || 
          value === '#REF!' || value === '#DIV/0!' || value === '#NUM!' || 
          value === '#NAME?' || value === '#NULL!') {
        
        const cellNotation = `${String.fromCharCode(65 + col)}${row + 1}`;
        errorCells.push(cellNotation);
      }
    }
  }
  
  // Формируем сообщение для пользователя
  if (errorCells.length === 0) {
    SpreadsheetApp.getUi().alert('✅ Ошибок нет');
  } else {
    const message = `ОШИБКИ: ${errorCells.join(' ')}`;
    SpreadsheetApp.getUi().alert(message);
  }
}

/***** === ОТДАТЬ BCD УЧЕНИКУ (ТОЛЬКО НЕПУСТЫЕ ЯЧЕЙКИ) ===*****/
function f2() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    
    const ssDev = SpreadsheetApp.openById(devId);
    const shDev = ssDev.getActiveSheet();
    const sheetName = shDev.getName();
    
    // Извлекаем ID из имени DEV файла
    const devFileName = ssDev.getName();
    const idMatch = devFileName.match(/DEV\s—\s(\d+)/);
    if (!idMatch) {
      throw new Error('Не удалось извлечь ID из имени DEV файла. Формат: "DEV — 111"');
    }
    const devIdNumber = idMatch[1];
    
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

    updateDatabaseWithDeliveryInfo_(devIdNumber);

    SpreadsheetApp.getUi().alert(`✅ Отдано ${copiedCount} ячеек B-C-D ученику (только непустые значения)\n\nID ${devIdNumber} записан в базу`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка при отправке BCD ученику: ' + (e.message || e));
  }
}

/***** === ДОБАВИТЬ IF К GPT ФОРМУЛАМ ===*****/
function f3() {
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

/***** === СОЗДАНИЕ STUDENT ФАЙЛА ===*****/
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
      const currentFile = SpreadsheetApp.getActive();
      const currentFileId = currentFile.getId();
      const currentFileName = currentFile.getName();
      
      if (!/^DEV\s—\s/i.test(currentFileName)) {
        throw new Error('Текущий файл не является DEV файлом. Откройте DEV файл и запустите функцию снова.');
      }
      
      const devUrlInRegistry = String(sheet.getRange(row, COLS.devUrl).getValue() || '').trim();
      const devIdInRegistry = fileIdFromUrl_(devUrlInRegistry);
      
      if (currentFileId !== devIdInRegistry) {
        throw new Error('Текущий DEV файл не соответствует записи в реестре. Откройте правильный DEV файл.');
      }
      
      const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
      const studFile = DriveApp.getFileById(currentFileId).makeCopy(`STUDENT — ${order}`, folder);
      studId = studFile.getId();
      
      // Убираем формулы из STUDENT
      removeFormulasFromStudent_(studId);
      
      DriveApp.getFileById(studId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      const studUrl = `https://docs.google.com/spreadsheets/d/${studId}/edit`;
      sheet.getRange(row, COLS.studentUrl).setValue(studUrl);
      
      // Обновляем данные в реестре
      sheet.getRange(row, COLS.studentUrl).setValue(studUrl);
      SpreadsheetApp.flush(); // Принудительно сохраняем изменения
      
      const ssDev = SpreadsheetApp.openById(currentFileId);
      const shDev = ssDev.getActiveSheet();
      const sheetName = shDev.getName();
      
      const ssStud = SpreadsheetApp.openById(studId);
      const shStud = ssStud.getSheetByName(sheetName) || ssStud.insertSheet(sheetName);

      const lastRow = shDev.getLastRow();
      
      if (lastRow >= 1) {
        let copiedCount = 0;

        // Проходим по всем строкам DEV и копируем BCD в STUDENT
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

        console.log(`Автоматически скопировано ${copiedCount} ячеек BCD при создании STUDENT`);
      }
    }
    
    const finalStudUrl = `https://docs.google.com/spreadsheets/d/${studId}/edit`;
    showLink_('STUDENT готов (создан из текущего DEV, формулы удалены, данные BCD скопированы)', finalStudUrl, 'ПЕРЕЙТИ В STUD');
    
  } catch (e) {
    console.error('Ошибка при создании STUDENT:', e);
    SpreadsheetApp.getUi().alert('Ошибка создания STUDENT: ' + (e.message || e));
  }
}

/***** === УДАЛЕНИЕ ФОРМУЛ ИЗ STUDENT ===*****/
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
      
      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
          const formula = formulas[r][c];
          
          // Проверяем, нужно ли сохранить эту ячейку
          const shouldPreserve = 
            // Не стирать C8 и D8
            (r + 1 === 8 && (c + 1 === 3 || c + 1 === 4)) ||
            // Не стирать EFG с 1 по 14 строку
            (r + 1 >= 1 && r + 1 <= 14 && c + 1 >= 5 && c + 1 <= 7);
          
          if (formula && formula.startsWith('=') && !shouldPreserve) {
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

  const expert = sheet.getRange(row, COLS.expert).getValue() || '';
  const aud1 = sheet.getRange(row, COLS.aud1).getValue() || '';
  const aud2 = sheet.getRange(row, COLS.aud2).getValue() || '';
  const aud3 = sheet.getRange(row, COLS.aud3).getValue() || '';
  const aud4 = sheet.getRange(row, COLS.aud4).getValue() || '';
  const aud5 = sheet.getRange(row, COLS.aud5).getValue() || '';
  const aud6 = sheet.getRange(row, COLS.aud6).getValue() || '';
  const expertProgram = sheet.getRange(row, COLS.expertProgram).getValue() || '';

  console.log('Создание DEV с данными:', {
    expert, aud1, aud2, aud3, aud4, aud5, aud6, expertProgram
  });

  const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const devFile = DriveApp.getFileById(masterId).makeCopy(`DEV — ${order}`, folder);
  const devId = devFile.getId();

  applyAudienceExpert_(devId, {
    expert: expert,
    aud1: aud1,
    aud2: aud2, 
    aud3: aud3,
    aud4: aud4,
    aud5: aud5,
    aud6: aud6,
    expertProgram: expertProgram
  });
  
  // Очищаем только старые незаполненные аудитории
  clearAudienceColumnsIfMissing_(devId, {
    aud2: aud2,
    aud3: aud3
  });
  
  sheet.getRange(row, COLS.devUrl).setValue(`https://docs.google.com/spreadsheets/d/${devId}/edit`);
  
  const displayMode = mode === 'factory' ? 'Отправить STUDENT' : 'Не Фабрика';
  sheet.getRange(row, COLS.devMode).setValue(displayMode);

  // ФИНАЛЬНАЯ ПРОВЕРКА
  const ssDev = SpreadsheetApp.openById(devId);
  const shDev = ssDev.getSheets()[0];
  

  const resultMessage = 
    "DEV создан!";

  showLink_(resultMessage, `https://docs.google.com/spreadsheets/d/${devId}/edit`, 'ПЕРЕЙТИ В DEV');
}

function menuDeliverExpanded_Final() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();
    
    const ssDev = SpreadsheetApp.openById(devId);
    const ssStud = SpreadsheetApp.openById(studentId);
    const shDev = ssDev.getActiveSheet();
    const sheetName = shDev.getName();
    
    // Шаг 1: Извлекаем ID из имени DEV файла
    const devFileName = ssDev.getName();
    const idMatch = devFileName.match(/DEV\s—\s(\d+)/);
    if (!idMatch) {
      throw new Error('Не удалось извлечь ID из имени DEV файла. Формат: "DEV — 111"');
    }
    const devIdNumber = idMatch[1];
    
    // Шаг 2: Обрабатываем исходную вкладку в DEV - заменяем формулы значениями
    processFormulasInPlace_(shDev);
    
    // Шаг 3: Создаем УНИКАЛЬНОЕ имя для временной вкладки в DEV
    const timestamp = new Date().getTime();
    const tempSheetName = `temp_${timestamp}`;
    
    // Создаем временную вкладку как копию исходной в DEV (уже без "плохих" формул)
    const tempSheet = shDev.copyTo(ssDev);
    tempSheet.setName(tempSheetName);
    
    try {
      // Шаг 4: Очищаем ВСЕ оставшиеся формулы во временной вкладке
      removeFormulasKeepStyles_(tempSheet);
      
      // Шаг 5: Копируем полностью очищенную временную вкладку в STUDENT
      const newSheetInStudent = tempSheet.copyTo(ssStud);
      const tempSheetNameInStudent = `temp_student_${timestamp}`;
      newSheetInStudent.setName(tempSheetNameInStudent);
      
      // Удаляем старую вкладку в STUDENT если существует
      const oldSheet = ssStud.getSheetByName(sheetName);
      if (oldSheet) {
        ssStud.deleteSheet(oldSheet);
      }
      
      // Переименовываем новую вкладку в оригинальное имя
      newSheetInStudent.setName(sheetName);
      
      // Активируем новый лист в STUDENT
      ssStud.setActiveSheet(newSheetInStudent);
      
      // Шаг 6: Записываем информацию в базу
      updateDatabaseWithDeliveryInfo_(devIdNumber);
      
      SpreadsheetApp.getUi().alert(`✅ STUDENT обновлен: вкладка "${sheetName}" заменена на версию без формул\n\nID ${devIdNumber} записан в базу`);
      
    } finally {
      ssDev.deleteSheet(tempSheet);
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка при синхронизации DEV → STUDENT: ' + (e.message || e));
  }
}

function updateDatabaseWithDeliveryInfo_(devIdNumber) {
  try {
    const files = DriveApp.getFilesByName(NameMainTable);
    if (!files.hasNext()) {
      console.log('Таблица "БАЗА" не найдена');
      return;
    }
    
    const baseFile = files.next();
    const ssBase = SpreadsheetApp.openById(baseFile.getId());
    const shBase = ssBase.getSheets()[0]; // Берем первую вкладку
    
    const data = shBase.getDataRange().getValues();
    
    // Ищем строку с совпадающим ID в столбце B (индекс 1)
    let targetRow = -1;
    for (let i = 0; i < data.length; i++) {
      const rowId = String(data[i][1] || '').trim(); // Столбец B
      if (rowId === devIdNumber) {
        targetRow = i + 1; // +1 потому что индексы начинаются с 1 в Google Sheets
        break;
      }
    }
    
    if (targetRow === -1) {
      console.log(`ID ${devIdNumber} не найден в столбце B таблицы "БАЗА"`);
      return;
    }
    
    // Записываем в столбец F (индекс 5) сообщение
    const message = "написать Влад сделал";
    shBase.getRange(targetRow, 6).setValue(message); // Столбец F
    
    console.log(`Записано в базу: строка ${targetRow}, столбец F - "${message}"`);
    
  } catch (e) {
    console.error('Ошибка при обновлении базы:', e);
    throw new Error('Не удалось обновить базу: ' + (e.message || e));
  }
}

function processFormulasInPlace_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 1 || lastCol < 1) return;
  
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const formulas = range.getFormulas();
  const values = range.getValues();
  
  let hasChanges = false;
  
  for (let r = 0; r < lastRow; r++) {
    for (let c = 0; c < lastCol; c++) {
      const formula = formulas[r][c];
      
      if (formula && formula.startsWith('=')) {
        const value = values[r][c];
        
        if (!isErrorValue_(value)) {
          values[r][c] = value;
          hasChanges = true;
        }
      }
    }
  }
  if (hasChanges) {
    range.setValues(values);
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

    const dataToCopy = [];

    for (let r = 1; r <= lastRow; r++) {
      if (isRowGrouped_(shStud, r)) continue;
      
      dataToCopy.push({
        row: r,
        bValue: shStud.getRange(r, COL_B).getValue(),
        cValue: shStud.getRange(r, COL_C).getValue(),
        dValue: shStud.getRange(r, COL_D).getValue()
      });
    }

    if (dataToCopy.length === 0) {
      SpreadsheetApp.getUi().alert('Не найдено несгруппированных строк в STUDENT');
      return;
    }

    // Копируем в DEV (только ячейки без формул)
    let copiedCount = 0;
    
    for (const data of dataToCopy) {
      const targetRow = data.row;
      
      if (!isRowGrouped_(shDev, targetRow)) {
        // Проверяем каждую ячейку на наличие формулы
        const rangeB = shDev.getRange(targetRow, COL_B);
        const rangeC = shDev.getRange(targetRow, COL_C);
        const rangeD = shDev.getRange(targetRow, COL_D);
        
        // Копируем только если в ячейке нет формулы
        if (!hasFormula_(rangeB)) {
          rangeB.setValue(data.bValue);
          copiedCount++;
        }
        if (!hasFormula_(rangeC)) {
          rangeC.setValue(data.cValue);
          copiedCount++;
        }
        if (!hasFormula_(rangeD)) {
          rangeD.setValue(data.dValue);
          copiedCount++;
        }
      }
    }

    SpreadsheetApp.getUi().alert(`✅ Обновлено ${copiedCount} ячеек B-C-D из STUDENT в DEV`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка при копировании BCD из STUDENT: ' + (e.message || e));
  }
}

// Проверяет, содержит ли ячейка формулу
function hasFormula_(range) {
  try {
    const formula = range.getFormula();
    return formula && formula.startsWith('=');
  } catch (e) {
    return false;
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

/***** === ФУНКЦИЯ ДЛЯ ПАРСИНГА СПИСКОВ ===*****/
function parseNumberedListEnhanced_(text) {
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

    // 1. "1. текст", "2. текст"
    const matchDot = trimmedLine.match(/^\s*(\d{1,2})\.\s*(.+)$/);
    // 2. "1) текст", "2) текст"  
    const matchBracket = trimmedLine.match(/^\s*(\d{1,2})\)\s*(.+)$/);
    // 3. "1 текст", "2 текст"
    const matchNumber = trimmedLine.match(/^\s*(\d{1,2})\s+(.+)$/);
    // 4. Любой текст с переносами строк
    const hasMultipleLines = lines.length > 1;

    if (matchDot) {
      items.push(matchDot[2].trim());
    } else if (matchBracket) {
      items.push(matchBracket[2].trim());
    } else if (matchNumber) {
      items.push(matchNumber[2].trim());
    } else if (hasMultipleLines) {
      // Если есть несколько строк, но без нумерации - берем все
      items.push(trimmedLine);
    }
  }

  // Если ничего не найдено, но есть текст - возвращаем как один элемент
  return items.length > 0 ? items : [cleanedText];
}


function parseNumberedListSimple_(text) {
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

    // Ищем нумерованные пункты
    const match = trimmedLine.match(/^\s*(\d{1,2})[\.\)]\s*(.+)$/);
    if (match) {
      items.push(match[2].trim());
    } else {
      // Если не нашли нумерацию, но есть текст - добавляем как есть
      items.push(trimmedLine);
    }
  }

  return items.length > 0 ? items : [cleanedText];
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
  
  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c] && formulas[r][c].startsWith('=')) {
        const cell = sheet.getRange(r + 1, c + 1);
        cell.clearContent(); 
      }
    }
  }
}

function isErrorValue_(value) {
  if (value === null || value === undefined) return false;
  
  const stringValue = value.toString();
  const errorPatterns = [
    '#ERROR!',
    '#DIV/0!',
    '#N/A',
    '#NAME?',
    '#NUM!',
    '#VALUE!',
    '#REF!',
    '#NULL!'
  ];
  
  return errorPatterns.some(pattern => stringValue.includes(pattern));
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
       <a href="${url}" target="_blank" onclick="google.script.host.close()"
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
      
      if (rowIndex >= startRow && rowIndex <= endRow) {
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

function applyAudienceExpert_(fileId, data) {
  const ss = SpreadsheetApp.openById(fileId);
  const sheets = ss.getSheets();
  
  sheets.forEach(sh => {
    try { 
      console.log('Применяем данные:', data);
      
      // === 1. АУДИТОРИИ ===
      // Аудитории 1-3 в строку 1 (B1, C1, D1)
      const row1Audiences = [
        data.aud1 || '',
        data.aud2 || '', 
        data.aud3 || ''
      ];
      console.log('Аудитории строка 1 (B1:D1):', row1Audiences);
      sh.getRange('B1:D1').setValues([row1Audiences]);
      
      // Аудитории 4-6 в строку 2 (E2, F2, G2) - ВАЖНО: это АУДИТОРИИ, а не эксперт!
      const row2Audiences = [
        data.aud4 || '',
        data.aud5 || '',
        data.aud6 || ''
      ];
      console.log('Аудитории строка 2 (E2:G2):', row2Audiences);
      sh.getRange('E2:G2').setValues([row2Audiences]);
      
      // === 2. ЭКСПЕРТ ===
      // Эксперт распределяется в строку 3 (B3, C3, D3, E3, F3, G3) для соответствующих аудиторий
      const expert = data.expert || '';
      console.log('Эксперт для распределения:', expert);
      
      // Определяем для каких аудиторий нужно добавить эксперта
      const expertColumns = [];
      
      if (data.aud1 && data.aud1.toString().trim() !== '') expertColumns.push(2); // B3
      if (data.aud2 && data.aud2.toString().trim() !== '') expertColumns.push(3); // C3
      if (data.aud3 && data.aud3.toString().trim() !== '') expertColumns.push(4); // D3
      if (data.aud4 && data.aud4.toString().trim() !== '') expertColumns.push(5); // E3
      if (data.aud5 && data.aud5.toString().trim() !== '') expertColumns.push(6); // F3
      if (data.aud6 && data.aud6.toString().trim() !== '') expertColumns.push(7); // G3
      
      console.log('Колонки для эксперта (строка 3):', expertColumns);
      
      // Записываем эксперта в строку 3 для заполненных аудиторий
      expertColumns.forEach(col => {
        sh.getRange(3, col).setValue(expert);
        console.log(`Записан эксперт в ячейку ${String.fromCharCode(64 + col)}3`);
      });
      
      // === 3. ПРОГРАММА ЭКСПЕРТА ===
      if (data.expertProgram) {
        console.log('Программа эксперта в B4:', data.expertProgram);
        sh.getRange('B4').setValue(data.expertProgram);
      }
      
      // === ПРОВЕРКА РЕЗУЛЬТАТА ===
      console.log('=== ФИНАЛЬНЫЙ РЕЗУЛЬТАТ В DEV ===');
      console.log('СТРОКА 1 (Аудитории):');
      console.log('B1:', sh.getRange('B1').getValue());
      console.log('C1:', sh.getRange('C1').getValue());
      console.log('D1:', sh.getRange('D1').getValue());
      
      console.log('СТРОКА 2 (Аудитории 4-6):');
      console.log('E2:', sh.getRange('E2').getValue());
      console.log('F2:', sh.getRange('F2').getValue());
      console.log('G2:', sh.getRange('G2').getValue());
      
      console.log('СТРОКА 3 (Эксперт):');
      console.log('B3:', sh.getRange('B3').getValue());
      console.log('C3:', sh.getRange('C3').getValue());
      console.log('D3:', sh.getRange('D3').getValue());
      console.log('E3:', sh.getRange('E3').getValue());
      console.log('F3:', sh.getRange('F3').getValue());
      console.log('G3:', sh.getRange('G3').getValue());
      
      console.log('СТРОКА 4 (Программа):');
      console.log('B4:', sh.getRange('B4').getValue());
      
    } catch(e) {
      console.log('Ошибка при применении данных к листу:', e.message);
      throw e;
    }
  });
}

function clearAudienceColumnsIfMissing_(fileId, data) {
  const ss = SpreadsheetApp.openById(fileId);
  const sheets = ss.getSheets();
  
  sheets.forEach(sh => {
    // Очищаем только старые аудитории (1-3) если не заполнены
    // Новые аудитории (4-6) не очищаем - они просто остаются пустыми если нет данных
    if (!data.aud2) sh.getRange('C1').clearContent();
    if (!data.aud3) sh.getRange('D1').clearContent();
    
    console.log('Очистка выполнена для незаполненных аудиторий 1-3');
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

function parseNumberedList_(text) {
  if (!text) return [];
  
  const cleanedText = String(text)
    .replace(/\r\n?/g, '\n')
    .replace(/\u00A0/g, ' ')
    .trim();

  if (!cleanedText) return [];

  const items = [];
  const lines = cleanedText.split('\n');
  
  let currentItem = '';
  let currentNumber = null;
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    // Проверяем, начинается ли строка с нового нумерованного пункта
    const matchNumber = line.match(/^\s*(\d{1,2})[\.\)]\s*(.*)$/) || 
                       line.match(/^\s*(\d{1,2})\s+(.*)$/);
    
    if (matchNumber) {
      const number = parseInt(matchNumber[1]);
      const content = matchNumber[2].trim();
      
      // Если у нас уже есть собранный элемент, сохраняем его
      if (currentItem !== '') {
        items.push(currentItem.trim());
      }
      
      // Начинаем новый элемент
      currentNumber = number;
      currentItem = content;
      
      // Проверяем следующий элемент - если он тоже нумерованный, то это отдельный пункт
      if (i < lines.length - 1) {
        const nextLine = lines[i + 1].trim();
        const nextMatch = nextLine.match(/^\s*(\d{1,2})[\.\)]\s*/) || 
                         nextLine.match(/^\s*(\d{1,2})\s+/);
        
        if (nextMatch && parseInt(nextMatch[1]) === number + 1) {
          // Следующий элемент имеет следующий номер - заканчиваем текущий
          items.push(currentItem.trim());
          currentItem = '';
          currentNumber = null;
        }
      }
    } else {
      // Это продолжение текущего элемента
      if (currentItem !== '') {
        currentItem += '\n' + line;
      } else {
        // Если нет текущего элемента, начинаем новый
        currentItem = line;
      }
    }
  }
  
  // Добавляем последний элемент
  if (currentItem !== '') {
    items.push(currentItem.trim());
  }

  // Если не нашли структурированных элементов, возвращаем весь текст как один элемент
  return items.length > 0 ? items : [cleanedText];
}

function adjustCellReferences_(formula, rowOffset) {
  if (!formula || !formula.startsWith('=')) return formula;
  
  return formula.replace(/([A-Z])(\d+)/g, function(match, col, row) {
    const newRow = parseInt(row) + rowOffset;
    return col + newRow;
  });
}

function unfoldFormulasInColumnsAsync_(fileId, columns) {
  console.log('Разворачиваем формулы в колонках:', columns);
  return Promise.resolve();
}