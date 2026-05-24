/*
 * Demo / sanitized Google Apps Script utility for spreadsheet workflow automation.
 
 * This script inserts a new contributor task block into a structured Google Sheet.
 * It finds the boundary of the current month section, inserts three rows,
 * preserves merged cells and formatting, and copies required formulas from
 * existing task rows.
 */

const CONTRIBUTOR_TASK_CONFIG = {

  SEARCH_START_ROW: 9,
  TASK_BLOCK_ROW_COUNT: 3,

  NAME_COLUMN: 1,
  TASK_TYPE_COLUMN: 2,
  RATE_COLUMN: 7,
  TOTAL_COLUMN: 8,

  MONTH_SECTION_REGEX: /(^[A-Z]+|^[a-z]+)[a-z]+\s*\(\d+.+\)/i //For format like August (21.07.24-20.08.24)
};

function addContributorTaskBlock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Additional Tasks');

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Add contributor?', SpreadsheetApp.getUi().ButtonSet.YES_NO);

  if (response === ui.Button.NO) 
    {
      return response;
    }

  const row = findCurrentTaskSectionEndRow(sheet);

  insertTaskBlockRows(sheet, row);
  copyTaskTypeFormula(sheet, row);
  setRateFormulas(sheet, row);
  copyTotalFormula(sheet, row);
}

//The script starts checking from row 9 to skip the current month header. Once it finds the next month header, it inserts the new contributor block right before it
function findCurrentTaskSectionEndRow(sheet) {
  const columnToCheck = sheet.getRange('A:A70').getValues();

  for (let i = CONTRIBUTOR_TASK_CONFIG.SEARCH_START_ROW; i < columnToCheck.length; i++) 
    {
      const valueToCheck = sheet.getRange(i, CONTRIBUTOR_TASK_CONFIG.NAME_COLUMN).getValue();

      const isMonthSection = CONTRIBUTOR_TASK_CONFIG.MONTH_SECTION_REGEX.test(valueToCheck);

      if (isMonthSection) 
        {
          return i;
        }
    }

  return null;
}

//Inserts a new3-row task block (there are always 3 rows in the block) and prepares merged cells
function insertTaskBlockRows(sheet, row) {
  sheet.insertRowsAfter(row - 1, CONTRIBUTOR_TASK_CONFIG.TASK_BLOCK_ROW_COUNT);

  sheet.getRange(row, CONTRIBUTOR_TASK_CONFIG.NAME_COLUMN, CONTRIBUTOR_TASK_CONFIG.TASK_BLOCK_ROW_COUNT, 1).mergeVertically();

  sheet.getRange(row, CONTRIBUTOR_TASK_CONFIG.NAME_COLUMN).setBorder(true, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(row, CONTRIBUTOR_TASK_CONFIG.TASK_TYPE_COLUMN, CONTRIBUTOR_TASK_CONFIG.TASK_BLOCK_ROW_COUNT, 1).mergeVertically();
}

//This function and those below copy the formulas from the previous block so that the sheet continues to work correctly with the new block
function copyTaskTypeFormula(sheet, row) {
  for (let i = row; i > 20; i--) 
    {
      const formula = sheet.getRange(i, CONTRIBUTOR_TASK_CONFIG.TASK_TYPE_COLUMN).getFormula();

      if (/^=/.test(formula)) 
        {
          sheet.getRange(i, CONTRIBUTOR_TASK_CONFIG.TASK_TYPE_COLUMN).copyTo(sheet.getRange(row, CONTRIBUTOR_TASK_CONFIG.TASK_TYPE_COLUMN), SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);

          break;
        }
    }
}


function setRateFormulas(sheet, row) {
  for (let i = row; i <= row + CONTRIBUTOR_TASK_CONFIG.TASK_BLOCK_ROW_COUNT - 1; i++)
    {
      const formula =
        `=IFERROR(IFS(` +
        `OR(R[0]C[-1]="w",R[0]C[-1]="words"),` +
        `FILTER('Contributors'!R3C13:R118C13,` +
        `'Contributors'!R3C1:R118C1=R${row}C1,` +
        `NOT('Contributors'!R3C1:R118C1="")),` +
        `OR(R[0]C[-1]="min",R[0]C[-1]="h"),` +
        `FILTER('Contributors'!R3C10:R118C10,` +
        `'Contributors'!R3C1:R118C1=R${row}C1,` +
        `NOT('Contributors'!R3C1:R118C1=""))),` +
        `"")`;

      sheet.getRange(i, CONTRIBUTOR_TASK_CONFIG.RATE_COLUMN).setFormulaR1C1(formula);
    }
}


function copyTotalFormula(sheet, row) {
  for (let i = row; i > 8; i--) 
    {
      const formula = sheet.getRange(i, CONTRIBUTOR_TASK_CONFIG.TOTAL_COLUMN).getFormula();

      if (/^=/.test(formula)) 
        {
          sheet.getRange(i, CONTRIBUTOR_TASK_CONFIG.TOTAL_COLUMN).copyTo(sheet.getRange(row, CONTRIBUTOR_TASK_CONFIG.TOTAL_COLUMN, CONTRIBUTOR_TASK_CONFIG.TASK_BLOCK_ROW_COUNT, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

          break;
        }
    }
}