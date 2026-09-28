/**
 * Google Sheets Assessment Processing Tool
 *
 * Automates name formatting, score conversion, and score transfer
 * between Google Sheets.
 *
 * Portfolio version:
 * - Spreadsheet IDs have been removed.
 * - Internal organization terminology has been generalized.
 * - Configuration values are grouped below.
 */

// ========================================
// Configuration
// ========================================

const CONFIG = {
  SOURCE_TABS: {
    LEVEL_8: "Level 8",
    LEVEL_2: "Level 2",
  },

  TARGET_SHEET_NAME: "Grades",

  SPREADSHEET_IDS: {
    2024: "YOUR_2024_SPREADSHEET_ID",
    2025: "YOUR_2025_SPREADSHEET_ID",
  },
};


// ========================================
// Custom Menu
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Assessment Tools")
    .addItem("Format Student Names", "fixNames")
    .addItem("Post Assessment Scores", "postScores")
    .addToUi();

  addFormLinkerMenu_();
}


// ========================================
// Name Formatting
// ========================================

/**
 * Standardizes names in column D of the active sheet.
 *
 * - Converts names to title case.
 * - Removes unnecessary whitespace.
 * - Ensures spacing before parentheses.
 */
function fixNames() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();

  const range = sheet.getRange("D2:D");
  const values = range.getValues();

  const formattedValues = values.map(row => {
    return row.map(name => {
      return name
        .toLowerCase()
        .replace(/\b[a-z](?=[a-z]{1})/g, letter => letter.toUpperCase())
        .trim()
        .replace(/(\S+)\(/, "$1 (")
        .replace(/　|\s\s/, " ");
    });
  });

  range.setValues(formattedValues);
}


// ========================================
// Score Processing
// ========================================

/**
 * Retrieves assessment results, converts raw scores,
 * and posts the converted values to the appropriate
 * column in the target spreadsheet.
 */
function postScores() {
  let convertedScores = {};

  const sourceSpreadsheet = SpreadsheetApp.getActive();

  // Process Level 8 assessment scores.
  const level8Values = fetchValues(
    sourceSpreadsheet,
    CONFIG.SOURCE_TABS.LEVEL_8
  );

  for (const value of level8Values) {
    const studentName = value.slice(0, 1).flat()[0];
    const rawScore = value.slice(5, 6);

    convertedScores = {
      ...convertedScores,
      [studentName]: {
        score: scoreConversion(rawScore, 8),
        level8: true,
      },
    };
  }

  // Process Level 2 assessment scores.
  const level2Values = fetchValues(
    sourceSpreadsheet,
    CONFIG.SOURCE_TABS.LEVEL_2
  );

  for (const value of level2Values) {
    const studentName = value.slice(0, 1).flat()[0];
    const rawScore = value.slice(5, 6);

    convertedScores = {
      ...convertedScores,
      [studentName]: {
        score: scoreConversion(rawScore, 2),
        level2: true,
      },
    };
  }

  const currentYear = new Date().getFullYear();

  const targetSpreadsheetId =
    CONFIG.SPREADSHEET_IDS[currentYear] ||
    CONFIG.SPREADSHEET_IDS[2025];

  const targetSpreadsheet =
    SpreadsheetApp.openById(targetSpreadsheetId);

  const gradeSheet =
    targetSpreadsheet.getSheetByName(CONFIG.TARGET_SHEET_NAME);

  const dateValues = gradeSheet
    .getRange("A1:1")
    .getValues()
    .flat();

  const currentMonthColumn =
    getCurrentMonthColumn(dateValues);

  const lastGradeRow = gradeSheet.getLastRow();

  const gradeRange = gradeSheet.getRange(
    3,
    currentMonthColumn + 9,
    lastGradeRow - 2,
    1
  );

  const studentNames = gradeSheet
    .getRange("B3:B")
    .getValues()
    .map(row => row[0]);

  const outputArray =
    generateOutputArray(studentNames, convertedScores);

  gradeRange.setValues(outputArray);
}


// ========================================
// Data Retrieval
// ========================================

/**
 * Retrieves assessment data from a specified sheet
 * and excludes rows containing invalid score values.
 *
 * @param {Spreadsheet} spreadsheet
 * @param {string} tabName
 * @return {Array<Array<*>>}
 */
function fetchValues(spreadsheet, tabName) {
  const sheet = spreadsheet.getSheetByName(tabName);

  if (!sheet) {
    throw new Error(`Sheet not found: ${tabName}`);
  }

  const values = sheet
    .getDataRange()
    .getValues()
    .filter(row => row[5] !== "#N/A");

  // Remove the header row.
  values.shift();

  return values;
}


// ========================================
// Output Generation
// ========================================

/**
 * Matches student names with converted assessment scores.
 *
 * @param {Array<string>} studentNames
 * @param {Object} studentScores
 * @return {Array<Array<number|string>>}
 */
function generateOutputArray(studentNames, studentScores) {
  return studentNames.map(studentName => {
    try {
      const score = studentScores[studentName]?.score;

      return score !== undefined
        ? [score]
        : [""];
    } catch (error) {
      console.error(
        `Unable to process score for ${studentName}: ${error.message}`
      );

      return [""];
    }
  });
}


// ========================================
// Date Handling
// ========================================

/**
 * Finds the column containing the current month and year.
 *
 * @param {Array<*>} dateValues
 * @return {number}
 */
function getCurrentMonthColumn(dateValues) {
  const currentDate = new Date();

  const convertedDates = dateValues.map(
    value => new Date(value)
  );

  return convertedDates.findIndex(date => {
    return (
      date.getMonth() === currentDate.getMonth() &&
      date.getFullYear() === currentDate.getFullYear()
    );
  });
}


// ========================================
// Score Conversion
// ========================================

/**
 * Converts raw assessment scores to a standardized
 * rating between 0 and 5.
 *
 * @param {number|Array<number>} score
 * @param {number} level
 * @return {number}
 */
function scoreConversion(score, level) {
  const numericScore = Number(
    Array.isArray(score) ? score[0] : score
  );

  /*
   * Level 2
   *
   * Raw Score    Rating
   * 0            0
   * 1–4          1
   * 5–8          2
   * 9–12         3
   * 13–16        4
   * 17–20        5
   */

  if (level === 2) {
    return numericScore >= 17 ? 5 :
           numericScore >= 13 ? 4 :
           numericScore >= 9  ? 3 :
           numericScore >= 5  ? 2 :
           numericScore >= 1  ? 1 :
           0;
  }

  /*
   * Level 8
   *
   * Raw Score    Rating
   * 0–2          0
   * 3–6          1
   * 7–10         2
   * 11–14        3
   * 15–18        4
   * 19–20        5
   */

  if (level === 8) {
    return numericScore >= 19 ? 5 :
           numericScore >= 15 ? 4 :
           numericScore >= 11 ? 3 :
           numericScore >= 7  ? 2 :
           numericScore >= 3  ? 1 :
           0;
  }

  throw new Error(`Unsupported assessment level: ${level}`);
}
