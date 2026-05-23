/*
 * Demo / sanitized Google Apps Script tool for text change highlighting.
 
 * This script compares selected OLD text cells with corresponding NEW text cells.
 * Added words in the NEW text are highlighted in bold.
 * Removed words in the OLD text are highlighted with strikethrough.
 */

const CHANGE_HIGHLIGHTER_CONFIG = {

  NEW_RU_COLUMN: 6,
  NEW_EN_COLUMN: 7,

  OLD_RU_COLUMN: 16,
  OLD_EN_COLUMN: 17
};

function compareTextRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Text Database');
  const selectedRanges = sheet.getActiveRangeList().getRanges();

  let lastProcessedRow;
  let selectedColumn;

  for (let i = 0; i < selectedRanges.length; i++) 
    {
      const selectedRange = selectedRanges[i];
      const rangeRows = getRangeRows(selectedRange);
      selectedColumn = selectedRange.getColumn();

      if (!isValidOldTextColumn(selectedColumn)) 
        {
          SpreadsheetApp.getUi().alert('Not so fast, cowboy','Wrong range selected. Please select the range from the columns "OLD RU" or "OLD EN"', SpreadsheetApp.getUi().ButtonSet.OK);
          return;
        }

      const newTextColumn = getNewTextColumn(selectedColumn);

      for (let row = rangeRows.firstRow; row <= rangeRows.lastRow; row++) 
        {
          const oldTextCell = sheet.getRange(row, selectedColumn);
          const newTextCell = sheet.getRange(row, newTextColumn);

          const oldText = oldTextCell.getValue().toString();
          const newText = newTextCell.getValue().toString();

          if (!oldText && !newText) 
            {
              continue;
            }

          highlightNewTextChanges(newTextCell, oldText, newText);
          highlightOldTextChanges(oldTextCell, oldText, newText);

          lastProcessedRow = row;
        }
    }

  if (lastProcessedRow && selectedColumn) 
    {
      sheet.getRange(lastProcessedRow, selectedColumn).activate();
    }
}

//Checks whether the selected column is one of the OLD text columns
function isValidOldTextColumn(column) {
  return (
    column === CHANGE_HIGHLIGHTER_CONFIG.OLD_RU_COLUMN || column === CHANGE_HIGHLIGHTER_CONFIG.OLD_EN_COLUMN);
}

//Returns the matching NEW text column for the selected OLD text column
function getNewTextColumn(selectedColumn) {
  if (selectedColumn === CHANGE_HIGHLIGHTER_CONFIG.OLD_RU_COLUMN) 
    {
      return CHANGE_HIGHLIGHTER_CONFIG.NEW_RU_COLUMN;
    }

  if (selectedColumn === CHANGE_HIGHLIGHTER_CONFIG.OLD_EN_COLUMN) 
    {
      return CHANGE_HIGHLIGHTER_CONFIG.NEW_EN_COLUMN;
    }

  return null;
}

//Gets first and last row numbers from a selected range
function getRangeRows(range) {
  const firstRow = range.getRow();
  const lastRow = firstRow + range.getNumRows() - 1;

  return { firstRow, lastRow };
}


function highlightNewTextChanges(newTextCell, oldText, newText) {
  const newWords = newText.split(' ');
  const oldWords = oldText.split(' ');

  const addedWordRanges = getChangedWordRanges(newWords, oldWords, true);
  const unchangedWordRanges = getChangedWordRanges(newWords, oldWords, false);

  applyTextStyle(newTextCell, addedWordRanges, SpreadsheetApp.newTextStyle().setBold(true).build());
  applyTextStyle(newTextCell, unchangedWordRanges, SpreadsheetApp.newTextStyle().setBold(false).build());
}


function highlightOldTextChanges(oldTextCell, oldText, newText) {
  const oldWords = oldText.split(' ');
  const newWords = newText.split(' ');

  const removedWordRanges = getChangedWordRanges(oldWords, newWords, true);
  const unchangedWordRanges = getChangedWordRanges(oldWords, newWords, false);

  applyTextStyle(oldTextCell, removedWordRanges, SpreadsheetApp.newTextStyle().setStrikethrough(true).build());
  applyTextStyle(oldTextCell, unchangedWordRanges, SpreadsheetApp.newTextStyle().setStrikethrough(false).build());
}


function getChangedWordRanges(baseWords, compareWords, shouldBeMissing) {
  const ranges = [];

  for (let i = 0; i < baseWords.length; i++) 
    {
      const word = baseWords[i];

      if (!word) 
        {
          continue;
        }

      const wordIsMissing = compareWords.indexOf(word) === -1;

      if (wordIsMissing === shouldBeMissing) 
        {
          const indexList = getAllIndexes(baseWords, word);

          for (let j = 0; j < indexList.length; j++) 
            {
              const wordRange = getWordCharacterRange(baseWords, indexList[j]);

              if (wordRange.start !== wordRange.end) 
                {
                  ranges.push(wordRange);
                }
            }
        }
    }

  return ranges;
}

//Calculates start and end character positions for a word by its word index
function getWordCharacterRange(words, wordIndex) {
  if (wordIndex === 0) 
    {
      return {
        start: 0,
        end: words[0].length
      };
    }

  let start = 0;

  for (let i = 0; i < wordIndex; i++) 
    {
      start += words[i].length + 1;
    }

  return {
    start: start,
    end: start + words[wordIndex].length
  };
}


function applyTextStyle(cell, ranges, textStyle) {
  let richTextValue = cell.getRichTextValue();

  ranges.forEach(({ start, end }) => {
    richTextValue = richTextValue.copy().setTextStyle(start, end, textStyle).build();
  });

  cell.setRichTextValue(richTextValue);
}

//Returns all indexes of a value in an array because the indexOf method returns only the first index but I need all of them
function getAllIndexes(array, value) {
  const indexes = [];

  for (let i = 0; i < array.length; i++) 
    {
      if (array[i] === value) 
        {
          indexes.push(i);
        }
    }

  return indexes;
}