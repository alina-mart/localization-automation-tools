/*
 * Demo / sanitized Google Apps Script tool for localization workflows.
 
 * What it does:
 * 1. Parses an HTML table export from an internal tool pasted into a working sheet.
 * 2. Cleans HTML tags and formatting artifacts.
 * 3. Splits source text by preserved line break markers.
 * 4. Compares source strings against a translation memory sheet.
 * 5. Populates existing translations into locale columns.
 * 6. Highlights duplicate and new source strings.
 */

const CONFIG = {

  SOURCE_START_ROW: 4,
  SOURCE_COLUMN: 1,

  LOCALE_START_COLUMN: 5,
  LOCALE_COLUMN_COUNT: 8,
  MAX_ROWS_TO_PROCESS: 250,

  DUPLICATE_LABEL: 'DUPLICATE_SOURCE',
  NEW_TEXT_LABEL: 'NEW_SOURCE_TEXT',

  LINE_BREAK_MARKER: '=br='
};


function parseHtmlExportAndPopulateTranslations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workingSheet = ss.getSheetByName('Working Sheet');
  const translationMemorySheet = ss.getSheetByName('Translation Memory');

  if (!workingSheet || !translationMemorySheet) {
    throw new Error('Required sheets are missing.');
  }

  clearWorkingArea(workingSheet);
  parseHtmlExportToSheet(workingSheet);
  splitSourceRowsByLineBreakMarker(workingSheet); //Splits rows in the EN-column by '=br=' tag
  
  const sourceValues = getSourceValues(workingSheet);
  const translationMemory = translationMemorySheet.getRange('A1:I1000').getValues();

  const lookupResults = buildTranslationLookupResults(sourceValues, translationMemory); //Matches source strings against translation memory

  writeTranslationResults(workingSheet, lookupResults);
  clearLabelsForEmptySourceRows(workingSheet, sourceValues); //Removes labels for rows with empty source strings (after the HTML export is parsed there will be empty rows between paragraphs)
  highlightStatusLabels(workingSheet);
  hideHelperColumns(workingSheet);
}

//Clears previous parsed content so that LM do not need to do it manually before running the script
function clearWorkingArea(sheet) {
  const lastColumn = sheet.getDataRange().getLastColumn();
  const lastRow = sheet.getLastRow();

  if (lastRow >= CONFIG.SOURCE_START_ROW) 
    {
      sheet.getRange(CONFIG.SOURCE_START_ROW, 1, lastRow - CONFIG.SOURCE_START_ROW + 1, lastColumn).clearContent();
    }

  const localeRange = sheet.getRange(CONFIG.SOURCE_START_ROW, CONFIG.LOCALE_START_COLUMN, CONFIG.MAX_ROWS_TO_PROCESS, CONFIG.LOCALE_COLUMN_COUNT);

  localeRange.setBackgrounds(localeRange.getBackgrounds().map(row => row.map(() => null)));

  localeRange.setNotes(localeRange.getNotes().map(row => row.map(() => '')));
}

//The HTML export is pasted into A1 so the formula parses it into A4
function parseHtmlExportToSheet(sheet) {
  const outputCell = sheet.getRange('A4');

  const formula =`=ArrayFormula(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(ArrayFormula(SPLIT(TRANSPOSE(SPLIT(` +
    `REGEXREPLACE(REGEXREPLACE(A1,"</t.>",CHAR(10)),` +
    `CHAR(10)&CHAR(10)&"</table>",""),` +
    `CHAR(10)&CHAR(10),0)),CHAR(10),0)),` +
    `"<br>","${CONFIG.LINE_BREAK_MARKER}"),` +
    `"&nbsp;",CHAR(160)),"<.*?>",""))`;


  outputCell.setFormula(formula);

  SpreadsheetApp.flush(); //Force spreadsheet recalculation before continuing

  const lastColumn = sheet.getDataRange().getLastColumn();
  const lastParsedRow = findLastParsedRow(sheet);

  if (lastParsedRow >= CONFIG.SOURCE_START_ROW) 
    {
      const parsedRange = sheet.getRange(CONFIG.SOURCE_START_ROW, 1, lastParsedRow - CONFIG.SOURCE_START_ROW + 1, lastColumn); //Copy the parsed content (formula result)
      
      parsedRange.copyTo(parsedRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false); //Paste the parsed content to the working sheet so that it is not a formula but actual values 
    }
}


function findLastParsedRow(sheet) {
  const values = sheet.getRange(CONFIG.SOURCE_START_ROW, CONFIG.SOURCE_COLUMN, CONFIG.MAX_ROWS_TO_PROCESS, 1).getValues();

  for (let i = 0; i < values.length; i++) 
    {
      if (values[i][0] === '') 
        {
          return CONFIG.SOURCE_START_ROW + i - 1;
        }
    }

  return CONFIG.SOURCE_START_ROW + values.length - 1;
}

//Splits rows in the EN-column by '=br=' tag
function splitSourceRowsByLineBreakMarker(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.SOURCE_START_ROW + 1) 
    {
      return; //If there are no rows to process, exit the function
    }

  const sourceRange = sheet.getRange(CONFIG.SOURCE_START_ROW + 1, CONFIG.SOURCE_COLUMN, lastRow - CONFIG.SOURCE_START_ROW, 1);

  const sourceValues = sourceRange.getValues();
  const splitResults = [];

  sourceValues.forEach(row => 
    {
      const cellValue = row[0];

      if (cellValue && typeof cellValue === 'string') 
        {
          cellValue.split(CONFIG.LINE_BREAK_MARKER).forEach(value => splitResults.push([value]));
        }
    });

  if (splitResults.length === 0) 
    {
      return;
    }

  sheet.getRange(CONFIG.SOURCE_START_ROW + 1, CONFIG.SOURCE_COLUMN, splitResults.length, 1).setValues(splitResults);

  const rowsToClear = lastRow - CONFIG.SOURCE_START_ROW - splitResults.length;

  if (rowsToClear > 0) 
    {
      sheet.deleteRows(CONFIG.SOURCE_START_ROW + 1 + splitResults.length, rowsToClear);
    }
}


function getSourceValues(sheet) {
  return sheet.getRange(CONFIG.SOURCE_START_ROW, CONFIG.SOURCE_COLUMN, CONFIG.MAX_ROWS_TO_PROCESS, 1).getValues();
}

//Matches source strings against translation memory to find already translated strings, duplicates and new strings that were not translated yet
function buildTranslationLookupResults(sourceValues, translationMemory) {
  return sourceValues.map(sourceRow => {
    const sourceText = sourceRow[0];

    const matches = translationMemory.filter(memoryRow => memoryRow[0] === sourceText);

    if (matches.length > 1) 
      {
        return [CONFIG.DUPLICATE_LABEL, ...new Array(CONFIG.LOCALE_COLUMN_COUNT - 1).fill(null)]; //Returns an array with the duplicate label and empty locales
      }

    if (matches.length === 0) //If the source string is not found in the translation memory, it is a new string that was not translated yet
      {
        return [CONFIG.NEW_TEXT_LABEL, ...new Array(CONFIG.LOCALE_COLUMN_COUNT - 1).fill(null)];
      }

    return matches[0].slice(1, CONFIG.LOCALE_COLUMN_COUNT + 1); //If the source string is found in the translation memory (only once), return the translation for the corresponding locale
  });
}


function writeTranslationResults(sheet, lookupResults) {
  sheet.getRange(CONFIG.SOURCE_START_ROW, CONFIG.LOCALE_START_COLUMN, lookupResults.length, CONFIG.LOCALE_COLUMN_COUNT).setValues(lookupResults);
}

//Removes labels for rows with empty source strings (after the HTML export is parsed there will be empty rows between paragraphs)
function clearLabelsForEmptySourceRows(sheet, sourceValues) {
  sourceValues.forEach((row, index) => {
    if (row[0] === '') 
      {
        sheet.getRange(CONFIG.SOURCE_START_ROW + index, CONFIG.LOCALE_START_COLUMN).clearContent(); //Clear the content of the locale column
      }
  });
}


function highlightStatusLabels(sheet) {
  const statusRange = sheet.getRange(CONFIG.SOURCE_START_ROW, CONFIG.LOCALE_START_COLUMN, CONFIG.MAX_ROWS_TO_PROCESS, 1);

  const statusValues = statusRange.getValues();

  const backgrounds = statusValues.map(row => {
    if (row[0] === CONFIG.DUPLICATE_LABEL) 
      {
        return ['#fce5cd']; //Highlight duplicate rows in light orange
      }

    if (row[0] === CONFIG.NEW_TEXT_LABEL) 
      {
        return ['#f4cccc']; //Highlight new rows in light red
      }

    return [null];
  });

  statusRange.setBackgrounds(backgrounds);
}


function hideHelperColumns(sheet) {
  sheet.hideColumns(2, 3);
}
