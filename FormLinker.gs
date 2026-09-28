/**
 * Google Forms Response Linker
 *
 * Portfolio-safe version of an Apps Script workflow that discovers Google Forms
 * in configured Drive folders and links selected response destinations to the
 * active spreadsheet. Organization-specific labels and identifiers have been
 * replaced with generic examples.
 *
 * Workflow:
 * 1. Paste the Group A Forms folder URL into Form Link Setup!B2.
 * 2. Paste the Group B Forms folder URL into Form Link Setup!B3.
 * 3. Use Form Tools > Scan form folders.
 * 4. Review the detected Forms and planned tab names.
 * 5. Use Form Tools > Link checked forms.
 *
 * This file intentionally does not define onOpen().
 * The main project can call addFormLinkerMenu_() from its existing onOpen().
 */

const CONFIG = {

  setupSheetName: 'Form Link Setup',

  groupAFolderUrlCell: 'B2',

  groupBFolderUrlCell: 'B3',

  headerRow: 6,

  firstDataRow: 7,

  reviewMarker: 'Review Quiz',

  miniTabPrefix: 'Quiz',

  groups: ['Group A', 'Group B'],

  columnCount: 7,

};



function addFormLinkerMenu_() {

  SpreadsheetApp.getUi()

    .createMenu('Form Tools')

    .addItem('Set up / repair template', 'setupFormLinker')

    .addSeparator()

    .addItem('1. Scan form folders', 'scanFormFolders')

    .addItem('2. Link checked forms', 'linkCheckedForms')

    .addSeparator()

    .addItem('Clear form list', 'clearFormList')

    .addToUi();

}



function setupFormLinker() {

  const sheet = getOrCreateSetupSheet_();



  // Preserve values from an older setup layout when possible.

  const oldLayout = sheet.getRange('A2').getDisplayValue() === 'Quiz group';

  const oldGroup = String(sheet.getRange('B2').getDisplayValue()).trim();

  const oldFolderUrl = String(sheet.getRange('B3').getDisplayValue()).trim();



  sheet.getRange('A1:G1').breakApart();

  sheet.getRange('A1:G1').merge();

  sheet.getRange('A1').setValue('Google Forms Linker');



  sheet.getRange('A2:G4').clearContent();

  sheet.getRange('B2:B3').clearDataValidations();



  sheet.getRange('A2').setValue('Group A folder URL');

  sheet.getRange('A3').setValue('Group B folder URL');

  sheet.getRange('A4').setValue(

    'Paste both form-folder URLs, then use Form Tools → Scan form folders.'

  );



  if (oldLayout && oldFolderUrl) {

    if (oldGroup === 'Group B') {

      sheet.getRange(CONFIG.groupBFolderUrlCell).setValue(oldFolderUrl);

    } else {

      sheet.getRange(CONFIG.groupAFolderUrlCell).setValue(oldFolderUrl);

    }

  }



  sheet.getRange(CONFIG.headerRow, 1, 1, CONFIG.columnCount).setValues([[

    'Link?',

    'Group',

    'Form name',

    'Planned tab name',

    'Current destination',

    'Status',

    'Form ID',

  ]]);



  sheet.setFrozenRows(CONFIG.headerRow);



  sheet.setColumnWidth(1, 70);

  sheet.setColumnWidth(2, 75);

  sheet.setColumnWidth(3, 340);

  sheet.setColumnWidth(4, 180);

  sheet.setColumnWidth(5, 160);

  sheet.setColumnWidth(6, 320);



  // Show working columns and hide the technical Form ID column.

  sheet.showColumns(1, CONFIG.columnCount);

  sheet.hideColumns(7);



  sheet.getRange('A1:G1').setFontWeight('bold');

  sheet.getRange(CONFIG.headerRow, 1, 1, CONFIG.columnCount).setFontWeight('bold');



  sheet.getRange(CONFIG.groupAFolderUrlCell).setNote(

    'Paste the Google Drive folder URL containing the Group A Forms.'

  );

  sheet.getRange(CONFIG.groupBFolderUrlCell).setNote(

    'Paste the Google Drive folder URL containing the Group B Forms.'

  );



  SpreadsheetApp.getUi().alert(

    'Template ready',

    'Paste the Group A folder URL into B2 and the Group B folder URL into B3, then choose Form Tools → Scan form folders.',

    SpreadsheetApp.getUi().ButtonSet.OK

  );

}



function scanFormFolders() {

  const ui = SpreadsheetApp.getUi();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = getOrCreateSetupSheet_();

  ensureSetupLayout_(sheet);



  const folderInputs = [

    {

      group: 'Group A',

      url: String(sheet.getRange(CONFIG.groupAFolderUrlCell).getDisplayValue()).trim(),

    },

    {

      group: 'Group B',

      url: String(sheet.getRange(CONFIG.groupBFolderUrlCell).getDisplayValue()).trim(),

    },

  ];



  if (folderInputs.every(input => !input.url)) {

    ui.alert(

      'Folder URLs missing',

      'Paste the Group A folder URL into B2 and the Group B folder URL into B3 first.',

      ui.ButtonSet.OK

    );

    return;

  }



  clearFormRows_(sheet);



  const allForms = [];

  const scanMessages = [];



  for (const input of folderInputs) {

    if (!input.url) {

      scanMessages.push(`${input.group}: folder URL is missing.`);

      continue;

    }



    let folder;

    try {

      folder = DriveApp.getFolderById(extractFolderId_(input.url));

    } catch (error) {

      scanMessages.push(

        `${input.group}: folder could not be opened. Check the URL and your access. ${error.message}`

      );

      continue;

    }



    const groupForms = [];

    const files = folder.getFilesByType(MimeType.GOOGLE_FORMS);



    while (files.hasNext()) {

      const file = files.next();



      try {

        const form = FormApp.openById(file.getId());

        groupForms.push({

          group: input.group,

          id: form.getId(),

          title: String(form.getTitle() || '').trim() || file.getName(),

          destinationId: safeGetDestinationId_(form),

          error: '',

        });

      } catch (error) {

        groupForms.push({

          group: input.group,

          id: file.getId(),

          title: file.getName(),

          destinationId: '',

          error: error.message,

        });

      }

    }



    groupForms.sort((a, b) => a.title.localeCompare(b.title, 'ja'));



    if (groupForms.length === 0) {

      scanMessages.push(

        `${input.group}: no Google Forms were found directly inside the folder.`

      );

      continue;

    }



    const reviewForms = groupForms.filter(

      form => !form.error && form.title.includes(CONFIG.reviewMarker)

    );

    const miniForms = groupForms.filter(

      form => !form.error && !form.title.includes(CONFIG.reviewMarker)

    );



    const plannedNameById = {};



    miniForms.forEach((form, index) => {

      plannedNameById[form.id] =

        `${CONFIG.miniTabPrefix} ${index + 1}, ${input.group}`;

    });



    reviewForms.forEach(form => {

      plannedNameById[form.id] = `Review Quiz, ${input.group}`;

    });



    const reviewProblem = reviewForms.length === 1

      ? ''

      : reviewForms.length === 0

        ? `ERROR: No Form name contains “${CONFIG.reviewMarker}” in the ${input.group} folder.`

        : `ERROR: ${reviewForms.length} Form names contain “${CONFIG.reviewMarker}” in the ${input.group} folder.`;



    groupForms.forEach(form => {

      let status = form.error ? `ERROR: ${form.error}` : 'Ready';

      if (reviewProblem && !form.error) status = reviewProblem;



      allForms.push({

        ...form,

        plannedTabName: plannedNameById[form.id] || '',

        status,

        checked: !form.error && !reviewProblem,

      });

    });



    if (reviewProblem) {

      scanMessages.push(`${input.group}: ${reviewProblem.replace(/^ERROR:\s*/, '')}`);

    } else {

      scanMessages.push(

        `${input.group}: ${groupForms.length} Forms found — ` +

        `${miniForms.length} quiz form(s) and 1 review quiz form.`

      );

    }

  }



  if (allForms.length === 0) {

    ui.alert(

      'No Forms found',

      scanMessages.join('\n') || 'No Forms could be loaded.',

      ui.ButtonSet.OK

    );

    return;

  }



  const rows = allForms.map(form => [

    form.checked,

    form.group,

    form.title,

    form.plannedTabName,

    describeDestination_(form.destinationId, ss.getId()),

    form.status,

    form.id,

  ]);



  const outputRange = sheet.getRange(

    CONFIG.firstDataRow,

    1,

    rows.length,

    CONFIG.columnCount

  );



  outputRange.setValues(rows);



  sheet.getRange(CONFIG.firstDataRow, 1, rows.length, 1).insertCheckboxes();

  sheet.getRange(CONFIG.firstDataRow, 1, rows.length, 1)

    .setValues(rows.map(row => [row[0]]));



  const hasProblems = allForms.some(form => form.status.startsWith('ERROR:')) ||

    folderInputs.some(input => !input.url);



  ui.alert(

    hasProblems ? 'Folders scanned with a warning' : 'Both folders scanned',

    scanMessages.join('\n') +

      '\n\nReview the list, then choose Form Tools → Link checked forms.',

    ui.ButtonSet.OK

  );

}



// Compatibility alias for existing buttons or menu items.

function scanForms() {

  scanFormFolders();

}



function linkCheckedForms() {

  const lock = LockService.getDocumentLock();



  if (!lock.tryLock(5000)) {

    SpreadsheetApp.getUi().alert(

      'The tool is already running. Please try again.'

    );

    return;

  }



  try {

    linkCheckedFormsLocked_();

  } finally {

    lock.releaseLock();

  }

}



function linkCheckedFormsLocked_() {

  const ui = SpreadsheetApp.getUi();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName(CONFIG.setupSheetName);



  if (!sheet || sheet.getLastRow() < CONFIG.firstDataRow) {

    ui.alert(

      'Nothing to link',

      'Scan form folders first.',

      ui.ButtonSet.OK

    );

    return;

  }



  const rowCount = sheet.getLastRow() - CONFIG.firstDataRow + 1;

  const values = sheet.getRange(

    CONFIG.firstDataRow,

    1,

    rowCount,

    CONFIG.columnCount

  ).getValues();



  const selected = values

    .map((row, index) => ({

      rowNumber: CONFIG.firstDataRow + index,

      checked: row[0] === true,

      group: String(row[1] || '').trim(),

      scannedTitle: String(row[2] || ''),

      plannedTabName: String(row[3] || ''),

      formId: String(row[6] || ''),

    }))

    .filter(item => item.checked && item.formId);



  if (selected.length === 0) {

    ui.alert(

      'Nothing selected',

      'Tick at least one Form in the Link? column.',

      ui.ButtonSet.OK

    );

    return;

  }



  const opened = [];



  for (const item of selected) {

    if (!CONFIG.groups.includes(item.group)) {

      sheet.getRange(item.rowNumber, 6).setValue(

        'ERROR: Group must be Group A or Group B.'

      );

      continue;

    }



    try {

      const form = FormApp.openById(item.formId);



      opened.push({

        ...item,

        form,

        currentTitle: String(form.getTitle() || '').trim() || item.scannedTitle,

        destinationId: safeGetDestinationId_(form),

      });

    } catch (error) {

      sheet.getRange(item.rowNumber, 6).setValue(

        `ERROR: ${error.message}`

      );

    }

  }



  if (opened.length === 0) {

    ui.alert(

      'No Forms could be opened',

      'Check the Status column.',

      ui.ButtonSet.OK

    );

    return;

  }



  const groupProblems = [];



  for (const group of CONFIG.groups) {

    const groupItems = opened.filter(item => item.group === group);



    if (groupItems.length === 0) continue;



    const reviewForms = groupItems.filter(

      item => item.currentTitle.includes(CONFIG.reviewMarker)

    );



    if (reviewForms.length !== 1) {

      groupProblems.push(

        `${group}: exactly one selected Form must contain ` +

        `“${CONFIG.reviewMarker}”. Found: ${reviewForms.length}.`

      );

      continue;

    }



    const minis = groupItems

      .filter(item => !item.currentTitle.includes(CONFIG.reviewMarker))

      .sort((a, b) => a.currentTitle.localeCompare(b.currentTitle, 'ja'));



    minis.forEach((item, index) => {

      item.plannedTabName =

        `${CONFIG.miniTabPrefix} ${index + 1}, ${group}`;

    });



    reviewForms[0].plannedTabName = `Review Quiz, ${group}`;

  }



  if (groupProblems.length > 0) {

    ui.alert(

      'Review Quiz check failed',

      groupProblems.join('\n') +

        '\n\nCorrect the Form names or checkbox selection, then scan again.',

      ui.ButtonSet.OK

    );

    return;

  }



  opened.forEach(item => {

    sheet.getRange(item.rowNumber, 3).setValue(item.currentTitle);

    sheet.getRange(item.rowNumber, 4).setValue(item.plannedTabName);

  });



  const targetNames = opened.map(

    item => sanitizeSheetName_(item.plannedTabName)

  );



  if (new Set(targetNames).size !== targetNames.length) {

    ui.alert(

      'Duplicate tab names',

      'The planned tab names are not unique. Nothing was linked.',

      ui.ButtonSet.OK

    );

    return;

  }



  const moving = opened.filter(

    item => item.destinationId && item.destinationId !== ss.getId()

  );



  if (moving.length > 0) {

    const preview = moving

      .slice(0, 10)

      .map(item => `• [${item.group}] ${item.currentTitle}`)

      .join('\n');



    const more = moving.length > 10

      ? `\n…and ${moving.length - 10} more`

      : '';



    const choice = ui.alert(

      'Some Forms are already linked elsewhere',

      `${moving.length} selected Form(s) are linked to another spreadsheet. ` +

      'Continuing will change the response destination so future responses ' +

      `are sent to this spreadsheet instead.\n\n${preview}${more}\n\nContinue?`,

      ui.ButtonSet.YES_NO

    );



    if (choice !== ui.Button.YES) return;

  }



  let successCount = 0;

  let errorCount = 0;



  for (const item of opened) {

    const destinationCell = sheet.getRange(item.rowNumber, 5);

    const statusCell = sheet.getRange(item.rowNumber, 6);



    statusCell.setValue('Working…');

    SpreadsheetApp.flush();



    try {

      let responseSheet = findResponseSheetForForm_(ss, item.formId);



      if (item.destinationId !== ss.getId()) {

        item.form.setDestination(

          FormApp.DestinationType.SPREADSHEET,

          ss.getId()

        );



        responseSheet = waitForResponseSheet_(

          ss,

          item.formId,

          15,

          500

        );

      } else if (!responseSheet) {

        responseSheet = waitForResponseSheet_(

          ss,

          item.formId,

          5,

          400

        );

      }



      if (!responseSheet) {

        throw new Error(

          'The response tab was not found after linking.'

        );

      }



      const safeName = sanitizeSheetName_(item.plannedTabName);

      const existing = ss.getSheetByName(safeName);



      if (

        existing &&

        existing.getSheetId() !== responseSheet.getSheetId()

      ) {

        throw new Error(

          `A different tab already uses the name “${safeName}”.`

        );

      }



      if (responseSheet.getName() !== safeName) {

        responseSheet.setName(safeName);

      }



      // Keep the tab's short group name, while using the original Form title

      // for the Google Sheets response table shown above the data.

      let tableStatus = '';



      try {

        const responseTableName = renameResponseTable_(

          ss,

          responseSheet,

          item.currentTitle

        );



        tableStatus = responseTableName

          ? `; table “${responseTableName}”`

          : '; table name unchanged (response table not found)';

      } catch (tableError) {

        tableStatus =

          `; table-name warning: ${tableError.message}`;

      }



      destinationCell.setValue('This spreadsheet');

      statusCell.setValue(

        `Linked as “${safeName}”${tableStatus}`

      );



      successCount++;

    } catch (error) {

      statusCell.setValue(`ERROR: ${error.message}`);

      errorCount++;

    }

  }



  ui.alert(

    'Finished',

    `${successCount} Form(s) linked successfully.` +

      (errorCount

        ? `\n${errorCount} Form(s) had an error; check the Status column.`

        : ''),

    ui.ButtonSet.OK

  );

}



function clearFormList() {

  const sheet = SpreadsheetApp

    .getActiveSpreadsheet()

    .getSheetByName(CONFIG.setupSheetName);



  if (!sheet) return;

  clearFormRows_(sheet);

}



function getOrCreateSetupSheet_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();



  return ss.getSheetByName(CONFIG.setupSheetName) ||

    ss.insertSheet(CONFIG.setupSheetName, 0);

}



function ensureSetupLayout_(sheet) {

  if (

    sheet.getRange('A1').getDisplayValue() !== 'Google Forms Linker' ||

    sheet.getRange('A2').getDisplayValue() !== 'Group A folder URL' ||

    sheet.getRange('A3').getDisplayValue() !== 'Group B folder URL'

  ) {

    setupFormLinker();

  }

}



function clearFormRows_(sheet) {

  const maxRows = sheet.getMaxRows();



  if (maxRows >= CONFIG.firstDataRow) {

    const range = sheet.getRange(

      CONFIG.firstDataRow,

      1,

      maxRows - CONFIG.firstDataRow + 1,

      CONFIG.columnCount

    );



    range.clearContent();

    range.clearDataValidations();

  }

}



function extractFolderId_(input) {

  const value = String(input || '').trim();

  const urlMatch = value.match(

    /\/folders\/([a-zA-Z0-9_-]+)/

  );



  if (urlMatch) return urlMatch[1];



  if (/^[a-zA-Z0-9_-]{10,}$/.test(value)) {

    return value;

  }



  throw new Error(

    'The folder URL or folder ID is not valid.'

  );

}



function safeGetDestinationId_(form) {

  try {

    return form.getDestinationId() || '';

  } catch (error) {

    // Forms with no response destination can throw here.

    return '';

  }

}



function describeDestination_(

  destinationId,

  currentSpreadsheetId

) {

  if (!destinationId) return 'Not linked';

  if (destinationId === currentSpreadsheetId) {

    return 'This spreadsheet';

  }



  return 'Another spreadsheet';

}



function findResponseSheetForForm_(spreadsheet, formId) {

  for (const sheet of spreadsheet.getSheets()) {

    let formUrl = null;



    try {

      formUrl = sheet.getFormUrl();

    } catch (error) {

      continue;

    }



    if (!formUrl) continue;



    try {

      if (FormApp.openByUrl(formUrl).getId() === formId) {

        return sheet;

      }

    } catch (error) {

      // Ignore inaccessible or stale Form links.

    }

  }



  return null;

}



function waitForResponseSheet_(

  spreadsheet,

  formId,

  attempts,

  delayMs

) {

  for (let attempt = 0; attempt < attempts; attempt++) {

    SpreadsheetApp.flush();



    const sheet = findResponseSheetForForm_(

      spreadsheet,

      formId

    );



    if (sheet) return sheet;

    Utilities.sleep(delayMs);

  }



  return null;

}



function sanitizeSheetName_(name) {

  let safe = String(name || '')

    .replace(/[\\/:?*\[\]]/g, '')

    .replace(/[\u0000-\u001F\u007F]/g, '')

    .trim()

    .replace(/^'+|'+$/g, '')

    .trim();



  if (!safe) safe = 'Responses';

  return safe.slice(0, 100);

}



/**

 * Renames the Google Sheets table shown above a Form response area.

 *

 * Requires the Advanced Google Sheets service:

 * Apps Script editor -> Services (+) -> Google Sheets API -> Add.

 */

function renameResponseTable_(

  spreadsheet,

  responseSheet,

  originalFormTitle

) {

  const spreadsheetId = spreadsheet.getId();

  const sheetId = responseSheet.getSheetId();



  let table = null;



  // The Form response sheet/table can take a moment to appear.

  for (let attempt = 0; attempt < 10; attempt++) {

    const result = Sheets.Spreadsheets.get(

      spreadsheetId,

      {

        fields:

          'sheets(properties(sheetId),tables(tableId,name,range))',

      }

    );



    const apiSheet = (result.sheets || []).find(

      sheet =>

        sheet.properties &&

        sheet.properties.sheetId === sheetId

    );



    const tables =

      apiSheet && apiSheet.tables ? apiSheet.tables : [];



    if (tables.length > 0) {

      // A Google Form response sheet normally has one table.

      table = tables[0];

      break;

    }



    Utilities.sleep(300);

  }



  if (!table) return '';



  const baseName = sanitizeTableName_(originalFormTitle);

  const uniqueName = makeUniqueTableName_(

    spreadsheetId,

    baseName,

    table.tableId

  );



  Sheets.Spreadsheets.batchUpdate(

    {

      requests: [

        {

          updateTable: {

            table: {

              tableId: table.tableId,

              name: uniqueName,

            },

            fields: 'name',

          },

        },

      ],

    },

    spreadsheetId

  );



  return uniqueName;

}



/**

 * Cleans a Form title so it is valid as a Google Sheets table name.

 * Spaces are allowed. Unsupported symbols are removed.

 */

function sanitizeTableName_(name) {

  let safe = String(name || '')

    // Keep Unicode letters/numbers, spaces and underscores.

    .replace(/[^\p{L}\p{N}_ ]/gu, '')

    .replace(/\s+/g, ' ')

    .trim();



  if (!safe) safe = 'Form Responses';



  // Table names cannot begin with a number.

  if (/^\d/.test(safe)) safe = `_${safe}`;



  // Avoid reserved Boolean names and cell-reference-style names.

  if (

    /^(TRUE|FALSE)$/i.test(safe) ||

    /^[A-Z]{1,3}\d+$/i.test(safe) ||

    /^R\d+C\d+$/i.test(safe)

  ) {

    safe = `_${safe}`;

  }



  return safe.slice(0, 255);

}



/**

 * Google Sheets table names must be unique throughout the spreadsheet.

 */

function makeUniqueTableName_(

  spreadsheetId,

  baseName,

  currentTableId

) {

  const result = Sheets.Spreadsheets.get(

    spreadsheetId,

    {

      fields: 'sheets(tables(tableId,name))',

    }

  );



  const usedNames = new Set();



  for (const sheet of result.sheets || []) {

    for (const table of sheet.tables || []) {

      if (

        table.tableId !== currentTableId &&

        table.name

      ) {

        usedNames.add(

          String(table.name).toLowerCase()

        );

      }

    }

  }



  if (!usedNames.has(baseName.toLowerCase())) {

    return baseName;

  }



  let number = 2;



  while (true) {

    const suffix = ` ${number}`;

    const candidate =

      baseName.slice(

        0,

        Math.max(1, 255 - suffix.length)

      ) + suffix;



    if (!usedNames.has(candidate.toLowerCase())) {

      return candidate;

    }



    number++;

  }

}
