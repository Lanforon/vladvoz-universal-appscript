/***** НАСТРОЙКИ ДЛЯ ПОЛЬЗОВАТЕЛЯ *****/

const ALLOWED_EMAILS = ['']; // ИЗМЕНИТЕ НА ДЕЙСТВУЮЩИЙ EMAIL

const REGISTRY_FILE_ID = '1p8sBJylRf5-UuDAkXcxoq60Xwrta_L7EoY4EBb_OO5s'; 
const REG_SHEET = 'REGISTRY';
const START_ROW = 2; // для реестра
const NameMainTable = "MAIN"

const TARGET_FOLDER_ID = '1Zp5-PxFMYFH0eC7PWdr9LgsrYDR6eJEG'; // TOO 

// Ссылки в REGISTRY
const REG_MASTER_FACTORY_CELL = 'B1';
const REG_MASTER_NOFACT_CELL  = 'D1';
const REG_STYLE_MASTER_CELL   = 'F1';


const COLS = {
  fio: 1,           // A - ФИО
  order: 2,         // B - ID Геткурс
  // C - Пусто
  devUrl: 4,        // D - Ссылка DEV
  studentUrl: 5,    // E - Ссылка STUDENT
  devMode: 6,       // F - Статус ('Фабрика' / 'Не Фабрика')
  
  // Старые аудитории 1-3
  aud1: 7,          // G - Аудитория 1 → B1
  expert1: 8,       // H - Эксперт 1 → B2
  aud2: 9,          // I - Аудитория 2 → C1  
  expert2: 10,      // J - Эксперт 2 → C2
  aud3: 11,         // K - Аудитория 3 → D1
  expert3: 12,      // L - Эксперт 3 → D2
  
  // Новые аудитории 4-6
  aud4: 13,         // M - Аудитория 4 → E2
  expert4: 14,      // N - Эксперт 4 → E3 / Программа эксперта → B4
  aud5: 15,         // O - Аудитория 5 → F2
  expert5: 16,      // P - Эксперт 5 → F3
  aud6: 17,         // Q - Аудитория 6 → G2
  expert6: 18       // R - Эксперт 6 → G3
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
  if (currentFileName === NameMainTable) {
    menu
      .addSeparator()
      .addItem('СОЗДАТЬ DEV - КЛУБ', 'menuDevelopFactory')
      .addSeparator()
      .addItem('СОЗДАТЬ DEV - НЕ КЛУБ', 'menuDevelopNoFactory')
      .addSeparator()
      .addItem('🔄 ОБНОВИТЬ ИЗ РЕЕСТРА', 'f7')
      .addSeparator();
  } else if (/БАЗА/i.test(currentFileName)) {
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

function onChange(e) {
  try {
    const source = e.source;
    const currentFileName = source.getName();
    
    if (/STUDENT/i.test(currentFileName)) {
      const changeType = e.changeType;
      
      if (changeType === 'REMOVE_ROW' || changeType === 'REMOVE_COLUMN') {
        SpreadsheetApp.getUi().alert(
          '❌ Запрещено удалять!', 
          'В файлах STUDENT запрещено удалять строки и столбцы!\n\nМожно:\n• Редактировать содержимое ячеек\n• Добавлять новые строки/столбцы\n• Изменять форматирование\n\nЗапрещено:\n• Удалять строки\n• Удалять столбцы', 
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        
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

    const rowsWithMarker = [];
    const aValues = shStud.getRange(1, 1, lastRow, 1).getDisplayValues();
    
    for (let r = 0; r < aValues.length; r++) {
      const aValue = String(aValues[r][0] || '').trim();
      if (aValue.includes(MARK_SELECT) && !isRowGrouped_(shStud, r + 1)) {
        rowsWithMarker.push(r + 1);
      }
    }

    console.log('Найдено строк с маркером ">" в колонке A:', rowsWithMarker);

    if (rowsWithMarker.length === 0) {
      SpreadsheetApp.getUi().alert('Не найдено строк с маркером ">" в колонке A STUDENT');
      return;
    }

    let expandedCount = 0;
    
    rowsWithMarker.reverse().forEach(row => {
      const aValue = shDev.getRange(row, 1).getValue();
      const bValue = shDev.getRange(row, 2).getValue();
      const cValue = shDev.getRange(row, 3).getValue();
      const dValue = shDev.getRange(row, 4).getValue();
      
      console.log(`Строка ${row}: A="${aValue}", B="${bValue}", C="${cValue}", D="${dValue}"`);
      
      const bItems = parseNumberedList_(bValue);
      const cItems = parseNumberedList_(cValue);
      const dItems = parseNumberedList_(dValue);
      
      const maxItems = Math.max(bItems.length, cItems.length, dItems.length, 1);
      
      console.log(`Строка ${row}: B items=${bItems.length}, C items=${cItems.length}, D items=${dItems.length}, max=${maxItems}`);
      
      if (maxItems > 1) {
        console.log(`Раскрываем строку ${row} на ${maxItems} элементов`);
        
        shDev.insertRowsAfter(row, maxItems - 1);
        
        copyRowFormat_(shDev, row, row + 1, maxItems - 1);
        
        const sourceDevFormulas = shDev.getRange(row, 1, 1, shDev.getLastColumn()).getFormulas()[0];
        
        for (let i = 1; i < maxItems; i++) {
          const targetRange = shDev.getRange(row + i, 1, 1, sourceDevFormulas.length);
          const formulasToSet = sourceDevFormulas.map(formula => 
            formula ? adjustCellReferences_(formula, i) : ''
          );
          targetRange.setFormulas([formulasToSet]);
        }
        
        const templateFormulasEFGH = shDev.getRange(row, COL_E, 1, 4).getFormulas()[0];
        const newBlockFormulasEFGH = [];
        
        for (let i = 0; i < maxItems; i++) {
          const newRow = templateFormulasEFGH.map(formulaText => 
            adjustCellReferences_(formulaText, i)
          );
          newBlockFormulasEFGH.push(newRow);
        }
        
        shDev.getRange(row, COL_E, maxItems, 4).setFormulas(newBlockFormulasEFGH);
        
        for (let i = 0; i < maxItems; i++) {
          const targetRow = row + i;
          shDev.getRange(targetRow, 1).setValue(aValue); 
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

    const studValues = shStud.getRange(1, 5, lastRow, 3).getValues(); 
    const studFormulas = shStud.getRange(1, 5, lastRow, 3).getFormulas(); 
    
    const devValuesE = studValues.map((row, rowIndex) => {
      const combinedValue = row
        .map((value, colIndex) => studFormulas[rowIndex][colIndex] ? '' : value) 
        .filter(val => val) 
        .join(' '); 
      
      return [combinedValue]; 
    });

    const emptyValuesE = devValuesE; 
    const emptyValuesF = Array(lastRow).fill().map(() => ['']); 
    const emptyValuesG = Array(lastRow).fill().map(() => ['']); 

    shDev.getRange(1, 5, lastRow, 1).setValues(emptyValuesE);
    
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
  
  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      const value = values[row][col];
      
      if (value === '#ERROR!' || value === '#N/A' || value === '#VALUE!' || 
          value === '#REF!' || value === '#DIV/0!' || value === '#NUM!' || 
          value === '#NAME?' || value === '#NULL!') {
        
        const cellNotation = `${String.fromCharCode(65 + col)}${row + 1}`;
        errorCells.push(cellNotation);
      }
    }
  }
  
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

    for (let r = 1; r <= lastRow; r++) {
      if (isRowGrouped_(shDev, r) || isRowGrouped_(shStud, r)) continue;
      
      const devCellB = shDev.getRange(r, 2); // B
      const devCellC = shDev.getRange(r, 3); // C
      const devCellD = shDev.getRange(r, 4); // D
      
      const devValueB = devCellB.getValue();
      const devValueC = devCellC.getValue();
      const devValueD = devCellD.getValue();
      
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

/***** === НОВАЯ ФУНКЦИЯ: ОБНОВИТЬ ИЗ РЕЕСТРА ===*****/
function f7() {
  try {
    const currentFile = SpreadsheetApp.getActive();
    const currentFileName = currentFile.getName();
    
    // Проверяем, что находимся в таблице "БАЗА"
    if (currentFileName !== NameMainTable) {
      throw new Error(`Эта функция работает только в таблице "${NameMainTable}"`);
    }
    
    const activeSheet = currentFile.getActiveSheet();
    const activeRange = currentFile.getActiveRange();
    
    if (!activeRange) {
      throw new Error('Выберите ячейку в столбце D или E');
    }
    
    const activeColumn = activeRange.getColumn();
    const activeRow = activeRange.getRow();
    
    // Проверяем, что активная ячейка в столбце D или E
    if (activeColumn !== 4 && activeColumn !== 5) { // D=4, E=5
      throw new Error('Выберите ячейку в столбце D (ссылка DEV) или E (ссылка STUDENT)');
    }
    
    // Получаем URL из активной ячейки
    const url = activeRange.getValue();
    if (!url || typeof url !== 'string') {
      throw new Error('В выбранной ячейке нет ссылки');
    }
    
    // Извлекаем ID файла из URL
    const targetFileId = fileIdFromUrl_(url);
    
    // Получаем MAIN таблицу из B1
    const mainTableUrl = activeSheet.getRange('B1').getValue();
    if (!mainTableUrl || typeof mainTableUrl !== 'string') {
      throw new Error('В ячейке B1 нет ссылки на MAIN таблицу');
    }
    
    const mainFileId = fileIdFromUrl_(mainTableUrl);
    
    // Получаем ячейки для копирования из F1
    const cellsToCopy = activeSheet.getRange('F1').getValue();
    if (!cellsToCopy || typeof cellsToCopy !== 'string') {
      throw new Error('В ячейке F1 не указаны ячейки для копирования (формат: "B90, C90")');
    }
    
    // Парсим ячейки из F1
    const cellReferences = cellsToCopy.split(',').map(cell => cell.trim());
    if (cellReferences.length === 0) {
      throw new Error('Не удалось распознать ячейки в F1. Формат: "B90, C90"');
    }
    
    console.log('Копируем ячейки:', cellReferences);
    
    // Открываем MAIN таблицу
    const mainSS = SpreadsheetApp.openById(mainFileId);
    const mainSheets = mainSS.getSheets();
    
    // Открываем целевую таблицу
    const targetSS = SpreadsheetApp.openById(targetFileId);
    const targetSheets = targetSS.getSheets();
    
    let totalCopied = 0;
    
    // Копируем данные из каждой указанной ячейки
    for (const cellRef of cellReferences) {
      console.log(`Копируем ячейку ${cellRef}`);
      
      // Копируем во все листы целевой таблицы
      for (let i = 0; i < targetSheets.length; i++) {
        const targetSheet = targetSheets[i];
        const mainSheet = mainSheets[i] || mainSheets[0]; // Если листов меньше, берем первый
        
        try {
          // Получаем значение из MAIN таблицы
          const value = mainSheet.getRange(cellRef).getValue();
          const formula = mainSheet.getRange(cellRef).getFormula();
          
          // Записываем в целевую таблицу
          if (formula && formula.startsWith('=')) {
            targetSheet.getRange(cellRef).setFormula(formula);
          } else {
            targetSheet.getRange(cellRef).setValue(value);
          }
          
          console.log(`Скопировано в ${targetSheet.getName()}: ${cellRef}`);
          totalCopied++;
          
        } catch (e) {
          console.log(`Ошибка при копировании ${cellRef} в лист ${targetSheet.getName()}: ${e.message}`);
        }
      }
    }
    
    SpreadsheetApp.getUi().alert(`✅ Обновлено ${totalCopied} ячеек из MAIN таблицы в целевую таблицу`);
    
  } catch (e) {
    console.error('Ошибка в updateFromRegistry:', e);
    SpreadsheetApp.getUi().alert('Ошибка при обновлении из реестра: ' + (e.message || e));
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
  try {
    const { sheet, row } = resolveRegistryRowContext_();
    const masterUrl = getMasterUrlByMode_(mode);
    if (!masterUrl) throw new Error(`В REGISTRY нет MASTER для режима ${mode}`);
    const masterId = fileIdFromUrl_(masterUrl);
    const order = String(sheet.getRange(row, COLS.order).getValue() || '').trim();
    if (!order) throw new Error('В колонке B (ID заказа) пусто.');

    // Получаем данные из реестра
    // Проверяем, что строка существует
    if (row > sheet.getLastRow()) {
      throw new Error('Строка ' + row + ' не существует в таблице');
    }
    
    const aud1 = sheet.getRange(row, 7).getValue() || '';
    const expert1 = sheet.getRange(row, 8).getValue() || '';
    const aud2 = sheet.getRange(row, 9).getValue() || '';
    const expert2 = sheet.getRange(row, 10).getValue() || '';
    const aud3 = sheet.getRange(row, 11).getValue() || '';
    const expert3 = sheet.getRange(row, 12).getValue() || '';

    const aud4 = sheet.getRange(row, 13).getValue() || '';
    const aud5 = sheet.getRange(row, 15).getValue() || '';
    const aud6 = sheet.getRange(row, 17).getValue() || '';

    const expert4 = sheet.getRange(row, 14).getValue() || '';
    const expert5 = sheet.getRange(row, 16).getValue() || '';
    const expert6 = sheet.getRange(row, 18).getValue() || '';


    
    const expertProgram = sheet.getRange(row, COLS.expertProgram || 14).getValue() || ''; // N → B4

    console.log('=== ДАННЫЕ ИЗ MAIN ===');
    console.log('Аудитория 1:', aud1);
    console.log('Эксперт 1:', expert1);
    console.log('Аудитория 2:', aud2);
    console.log('Эксперт 2:', expert2);
    console.log('Аудитория 3:', aud3);
    console.log('Эксперт 3:', expert3);

    const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    const devFile = DriveApp.getFileById(masterId).makeCopy(`DEV — ${order}`, folder);
    const devId = devFile.getId();

    // Передаем все данные
    applyAudienceExpert_(devId, {
      aud1,      // MAIN G
      expert1,   // MAIN H
      aud2,      // MAIN I
      expert2,   // MAIN J
      aud3,      // MAIN K
      expert3,   // MAIN L
      aud4       // MAIN M
    });


    
    // Очищаем только старые незаполненные аудитории
    clearAudienceColumnsIfMissing_(devId, {
      aud2: aud2,
      expert2: expert2,
      aud3: aud3,
      expert3: expert3
    });
    
    sheet.getRange(row, COLS.devUrl).setValue(`https://docs.google.com/spreadsheets/d/${devId}/edit`);
    
    const displayMode = mode === 'factory' ? 'Отправить STUDENT' : 'Не Фабрика';
    sheet.getRange(row, COLS.devMode).setValue(displayMode);

    const resultMessage = "DEV создан!";
    showLink_(resultMessage, `https://docs.google.com/spreadsheets/d/${devId}/edit`, 'ПЕРЕЙТИ В DEV');
    
  } catch (e) {
    console.error('Ошибка при создании DEV:', e);
    SpreadsheetApp.getUi().alert('Ошибка создания DEV: ' + (e.message || e));
  }
}

function menuDeliverExpanded_Final() {
  try {
    const { devId, studentId } = resolveDevStudentByContext_();

    const ssDev = SpreadsheetApp.openById(devId);
    const ssStud = SpreadsheetApp.openById(studentId);
    const shDev = ssDev.getActiveSheet();
    const sheetName = shDev.getName();

    // 1️⃣ DEV — аккуратно заменяем формулы ТОЛЬКО с валидным значением
    processFormulasInPlace_(shDev);

    // 2️⃣ DEV → temp
    const tempSheet = shDev.copyTo(ssDev).setName(`temp_${Date.now()}`);

    try {
      // 3️⃣ temp — жёстко удаляем формулы + чистим ошибки
      removeFormulasKeepStyles_(tempSheet);

      // 4️⃣ STUDENT — copy → delete → rename
      const newSheet = tempSheet.copyTo(ssStud);
      newSheet.setName(`__new_${sheetName}`);

      const old = ssStud.getSheetByName(sheetName);
      if (old) ssStud.deleteSheet(old);

      newSheet.setName(sheetName);
      ssStud.setActiveSheet(newSheet);

      SpreadsheetApp.getUi().alert(
        '✅ STUDENT обновлён\n\n' +
        '• DEV: формулы сохранены\n' +
        '• STUDENT: без формул и ошибок'
      );

    } finally {
      ssDev.deleteSheet(tempSheet);
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка DEV → STUDENT: ' + (e.message || e));
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

  let changed = false;

  for (let r = 0; r < lastRow; r++) {
    for (let c = 0; c < lastCol; c++) {
      const formula = formulas[r][c];
      if (!formula) continue;

      const value = values[r][c];

      if (value === '' || value === null) {
        continue;
      }

      if (!isErrorValue_(value)) {
        values[r][c] = value;
        changed = true;
      }
    }
  }

  if (changed) {
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
  const range = sheet.getDataRange();
  const values = range.getValues();
  const formulas = range.getFormulas();

  let changed = false;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[0].length; c++) {
      const v = values[r][c];

      if (isErrorValue_(v)) {
        values[r][c] = "";
        changed = true;
        continue;
      }

      if (typeof v === 'string' && v.startsWith('@@=')) {
        values[r][c] = "";
        changed = true;
        continue;
      }

      if (formulas[r][c]) {
        values[r][c] = v;
        changed = true;
      }
    }
  }

  range.setValues(values);
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
  
  const valuesWithoutFormulas = values.map((row, rowIndex) => 
    row.map((value, colIndex) => {
      const formula = formulas[rowIndex][colIndex];
      return formula && formula.startsWith('=') ? '' : value;
    })
  );
  
  range.setValues(valuesWithoutFormulas);
}

function copyBasicFormatting_(sourceSheet, targetSheet, lastRow, lastCol) {
  try {
    for (let r = 1; r <= lastRow; r++) {
      const sourceRow = sourceSheet.getRange(r, 1);
      const targetRow = targetSheet.getRange(r, 1);
      targetSheet.setRowHeight(r, sourceSheet.getRowHeight(r));
    }
    
    for (let c = 1; c <= lastCol; c++) {
      const sourceCol = sourceSheet.getRange(1, c);
      const targetCol = targetSheet.getRange(1, c);
      targetSheet.setColumnWidth(c, sourceSheet.getColumnWidth(c));
    }
    
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
      if (formula && formula.startsWith('=')) {
        const cell = range.getCell(r + 1, c + 1);
        cell.clearContent(); 
      }
    }
  }
}

function copyRowHeightsAndColumnWidths_(sourceSheet, targetSheet, lastRow, lastCol) {
  try {
    for (let r = 1; r <= lastRow; r++) {
      targetSheet.setRowHeight(r, sourceSheet.getRowHeight(r));
    }
    
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

function applyAudienceExpert_(devId, data) {
  const ss = SpreadsheetApp.openById(devId);
  const sh = ss.getSheets()[0];
  const lastRow = sh.getLastRow();

  if (data.aud1) {
    sh.getRange('B2').setValue(data.aud1);
    sh.getRange('D2').setValue(data.aud1);
    sh.getRange('G2').setValue(data.aud1);
  }

  if (data.expert1) {
    sh.getRange('B1').setValue(data.expert1);
  }

  applyOrClear_(sh, data.aud2, 'C1', 3, lastRow);

  applyOrClear_(sh, data.expert2, 'D1', 4, lastRow);

  if (data.aud3) {
    sh.getRange('E2').setValue(data.aud3);
  }

  if (data.expert3) {
    sh.getRange('F2').setValue(data.expert3);
  }

  if (data.aud4) {
    sh.getRange('G2').setValue(data.aud4);
  }
}

function applyOrClear_(sheet, value, cell, col, lastRow) {
  if (value) {
    sheet.getRange(cell).setValue(value);
  } else if (lastRow > 0) {
    sheet.getRange(1, col, lastRow).clearContent();
  }
}


function clearFormulasInColumnFromRow_(sheet, columnLetter, startRow, endRow) {
  if (startRow > endRow) return;
  
  try {
    const range = sheet.getRange(`${columnLetter}${startRow}:${columnLetter}${endRow}`);
    const formulas = range.getFormulas();
    
    let clearedCount = 0;
    for (let i = 0; i < formulas.length; i++) {
      if (formulas[i][0] && formulas[i][0].startsWith('=')) {
        const cell = sheet.getRange(startRow + i, columnToIndex_(columnLetter));
        cell.clearContent();
        clearedCount++;
      }
    }
    
    if (clearedCount > 0) {
      console.log(`  Затерто ${clearedCount} формул в ${columnLetter}${startRow}:${columnLetter}${endRow}`);
    }
  } catch (e) {
    console.log(`  Ошибка при очистке колонки ${columnLetter}:`, e.message);
  }
}

function columnToIndex_(columnLetter) {
  return columnLetter.charCodeAt(0) - 64;
}

function isColumnEmpty_(sheet, columnLetter, startRow, endRow) {
  const range = sheet.getRange(`${columnLetter}${startRow}:${columnLetter}${endRow}`);
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && String(values[i][0]).trim() !== '') {
      return false;
    }
  }
  return true;
}

// Функция для гарантии, что колонка пустая (очищает любые значения/формулы)
function ensureColumnIsEmpty_(sheet, columnLetter, lastRow) {
  const range = sheet.getRange(`${columnLetter}1:${columnLetter}${lastRow}`);
  const formulas = range.getFormulas();
  const values = range.getValues();
  
  let hasContent = false;
  for (let i = 0; i < formulas.length; i++) {
    if (formulas[i][0] && formulas[i][0].startsWith('=')) {
      hasContent = true;
      break;
    }
    if (values[i][0] && String(values[i][0]).trim() !== '') {
      hasContent = true;
      break;
    }
  }
  
  if (hasContent) {
    range.clearContent();
    console.log(`✓ Гарантировано, что колонка ${columnLetter} пустая`);
  }
}

// Новая функция для очистки формул вниз по столбцу
function clearFormulasInColumn_(sheet, columnLetter, lastRow) {
  if (lastRow <= 1) return; // Если только заголовки, нечего очищать
  
  // Начинаем с 4 строки (после программы эксперта) или с 5, если нужно пропустить первые строки
  const startRow = 5; // Начинаем с 5 строки, чтобы не трогать заголовки и программу
  if (startRow > lastRow) return;
  
  const range = sheet.getRange(`${columnLetter}${startRow}:${columnLetter}${lastRow}`);
  const formulas = range.getFormulas();
  
  let clearedCount = 0;
  for (let i = 0; i < formulas.length; i++) {
    if (formulas[i][0] && formulas[i][0].startsWith('=')) {
      // Очищаем только формулу, оставляя значения
      const cell = sheet.getRange(startRow + i, columnToIndex_(columnLetter));
      cell.clearContent(); // Очищает содержимое (формулу), но не форматирование
      clearedCount++;
    }
  }
  
  console.log(`✓ Очищено ${clearedCount} формул в колонке ${columnLetter} (строки ${startRow}-${lastRow})`);
}


function clearAudienceColumnsIfMissing_(fileId, data) {
  const ss = SpreadsheetApp.openById(fileId);
  const sheets = ss.getSheets();
  
  sheets.forEach(sh => {
    try {
      const lastRow = sh.getLastRow();
      const lastRowToClear = Math.max(lastRow || 1, 100);
      const startRowForFormulas = 5;
      
      if (!data.aud2 || String(data.aud2).trim() === '') {
        clearFormulasInColumnFromRow_(sh, 'C', startRowForFormulas, lastRowToClear);
        console.log('✓ Колонка C: аудитория 2 не заполнена → формулы затерты');
      }
      
      if (!data.aud3 || String(data.aud3).trim() === '') {
        clearFormulasInColumnFromRow_(sh, 'D', startRowForFormulas, lastRowToClear);
        console.log('✓ Колонка D: аудитория 3 не заполнена → формулы затерты');
      }
      
    } catch (e) {
      console.log('Ошибка при очистке колонок:', e.message);
    }
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